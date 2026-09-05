Attribute VB_Name = "VBT_BACK"
Option Explicit
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
'                                Call AddStoredMeasurement(MeasStoreName_Ary(TestSeqNum), Store_Rtn_Meas(StoreIndex))
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
'                                    Dict_VT_Value = GetStoredMeasurement(DictKey_StoreVT)
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
'                                Call AddStoredMeasurement(MeasStoreName_Ary(TestSeqNum), Store_Rtn_Meas(StoreIndex))
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
'                                Call AddStoredMeasurement(MeasStoreName_Ary(TestSeqNum), Store_Rtn_Meas(StoreIndex))
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
Public Function DACTrimPrint(Optional TestSeqNum As Integer = 0, Optional Trimname As String, Optional TargetVal As Double) As Double
Dim i, j, n As Integer
Dim v() As Double
Dim k() As Variant
Dim Tk() As String
Dim site As Variant
Dim StrIdx() As String
Dim StrPrint As String
Dim TargetIdx As New SiteVariant


    If DACTargetStr.Exists(Trimname) Then
        TargetIdx = DACTargetStr.item(Trimname)
    Else
        TheExec.Datalog.WriteComment "No Trim Value for TrimName: " & Trimname
        Exit Function
    End If
    k = DACTrimValue.Keys
    For Each site In TheExec.sites
        StrIdx = Split(TargetIdx, "@")
        TheExec.Datalog.WriteComment vbNullString
        
        'Added for Osprey Metrology to print special Osprey Metrology Gain_err in datalog 20170315
        If LCase(TheExec.DataManager.instancename) Like "*gtrim*" Then
            TheExec.Datalog.WriteComment "Site" & site & " : TrimName=" & Trimname & " : TargetCode[MSB:LSB]=" & StrIdx(1) & " : Gain_err= " & StrIdx(4) & " : TargetValue= " & TargetVal
        Else
            TheExec.Datalog.WriteComment "Site" & site & " : TrimName=" & Trimname & " : TargetCode[MSB:LSB]=" & StrIdx(1) & " : TrimValue= " & StrIdx(4) & " : TargetValue= " & TargetVal
        End If
        
        Tk = Filter(k, "@site" & site, True)
        For i = 0 To UBound(Tk)
            StrPrint = vbNullString
            v = DACTrimValue.item(Tk(i))
            StrIdx = Split(Tk(i), "@")
            n = UBound(v)
            For j = 0 To UBound(v)
                If j = n Then
                    'StrPrint = StrPrint & " MeasValue" & j & "= " & v(j)
                    ''Added for Osprey Metrology to print special Osprey Metrology Gain_err in datalog 20170315
                    If LCase(TheExec.DataManager.instancename) Like "*gtrim*" Then
                        StrPrint = " Gain_err= " & v(j) & StrPrint
                    Else
                        StrPrint = " TrimValue= " & v(j) & StrPrint
                    End If
                Else
                    If j < n Then
                        StrPrint = StrPrint & " MeasValue" & j & "= " & v(j)
                    End If
                End If
            Next j
            StrPrint = "TrimCode[MSB:LSB]=" & StrIdx(1) & " :" & StrPrint
            TheExec.Datalog.WriteComment StrPrint
        Next i
        TheExec.Datalog.WriteComment vbNullString
    Next site
    DACTrimValue.RemoveAll
End Function

Public Function PrintAllTrimResult() As Long
Dim k() As Variant
Dim Tk() As String
Dim site As Variant
Dim index As Variant
Dim StrIdx() As String
Dim TargetIdx As New SiteVariant
    TheExec.Datalog.WriteComment "==================================="
    TheExec.Datalog.WriteComment "     TrimResult Print Start        "
    TheExec.Datalog.WriteComment "==================================="
    k = DACTargetStr.Keys
    For Each site In TheExec.sites
        TheExec.Datalog.WriteComment vbNullString
        TheExec.Datalog.WriteComment "Site" & site
        For Each index In k
            If Not index Like "DSPWF@*" Then
                TargetIdx = DACTargetStr.item(index)
                StrIdx = Split(TargetIdx, "@")
                TheExec.Datalog.WriteComment "TrimName=" & FormatNumeric(StrIdx(0), -28) & " : TargetCode[MSB:LSB]=" & FormatNumeric(StrIdx(1), -8) & " : TargetValue= " & StrIdx(4)
            End If
        Next index
    Next site
    TheExec.Datalog.WriteComment vbNullString
    TheExec.Datalog.WriteComment "==================================="
    TheExec.Datalog.WriteComment "       TrimResult Print End        "
    TheExec.Datalog.WriteComment "==================================="
End Function

Public Function Opt_DdrLpBkFunc3(DqSwpPat As Pattern, DqsSwpPat As Pattern, _
                            DisableComparePins As PinList, DisableConnectPins As PinList, _
                            DigCap_Pin As PinList, NoOfBists As Integer, _
                            DqSwpNoOfBits As Long, DqsSwpNoOfBits As Long, _
                            Optional DispCaptStrm As Boolean = False, _
                            Optional DigSrc_pin As PinList, Optional DigSrc_DataWidth As Long, Optional DigSrc_Sample_Size As String, _
                            Optional DqDigSrc_Equation As String, Optional DqsDigSrc_Equation As String, _
                            Optional digsrc_assignment As String, _
                            Optional CUS_Str_DigSrcData As String, _
                            Optional DigCap_DSPWaveSetting As CalculateMethodSetup_DSPWave = 0, _
                            Optional EyeTestRegName As String, _
                            Optional DigCap_Sample_Size_Dq As Long, _
                            Optional CUS_Str_DigCapData_Dq As String, _
                            Optional DigCap_Sample_Size_Dqs As Long, _
                            Optional CUS_Str_DigCapData_Dqs As String, _
                            Optional Interpose_PrePat As String, _
                            Optional SweepVtStr As String, _
                            Optional Calc_Eqn As String, _
                                                        Optional BV_Enable As Boolean, _
                            Optional Validating_ As Boolean) As Long

    Dim i As Long
    Dim site As Variant
    Dim Pat As String
    Dim EyeStrobes As Long
    Dim DqSwpWf As New DSPWave, DqsSwpWf As New DSPWave
    Dim Testname_CZ_Vt As String: Testname_CZ_Vt = vbNullString
    Dim Instname_split() As String
    Dim TempStr() As String
    Dim p As Long
    Dim BistIdx As Long
    Dim PatCnt As Long
    Dim PatNames() As String
    Dim DSP_Eye_StartBit_DQ As New DSPWave
    Dim DSP_Eye_BitLength_DQ As New DSPWave
    Dim DSP_Eye_StartBit_DQS As New DSPWave
    Dim DSP_Eye_BitLength_DQS As New DSPWave
    Dim DSP_Eye_Width As New DSPWave
    Dim DQ_EYE_Data As New DSPWave
    Dim DQS_EYE_Data As New DSPWave
    Dim TestName As String

    ''''' Sweep Vt from SweepVtStr
    Dim SplitByColon() As String
    Dim SourceIndexStr As String, SourceIndex As Long
    Dim StartVal As Double, StepVal As Double, FinalVal As Double
    Dim ReplaceStr() As String
        
    glb_TestInstance = vbNullString
    glb_TestInstance = UCase(TheExec.DataManager.instancename)
        
    'Dim Inst_Name_Str As String: Inst_Name_Str = glb_TestInstance
    
    ''''' Speed up the first run test time
    If Validating_ Then
        Call PrLoadPattern(DqsSwpPat.value)
        Call PrLoadPattern(DqSwpPat.value)
        Exit Function    ''''' Exit after validation
    End If
    
    If TheHdw.DSP.ExecutionMode = tlDSPModeHostDebug Then
        TheHdw.DSP.ExecutionMode = tlDSPModeAutomatic
    End If
    
    On Error GoTo errHandler
       
    Call GetFlowTName
       
    If BV_Enable Then
    Else
        TheHdw.Digital.ApplyLevelsTiming True, True, True, tlPowered
    End If
    
    EyeStrobes = DqSwpNoOfBits / NoOfBists
    
    If (DisableConnectPins <> "") Then TheHdw.Digital.Pins(DisableComparePins).Disconnect
    If (DisableComparePins <> "") Then TheHdw.Digital.Pins(DisableComparePins).DisableCompare = True
    
    ''''' Input Parsing Function to get every EYE start point bit and EYE bit length
    Call Opt_Input_Parsing(CUS_Str_DigCapData_Dq, EyeTestRegName, DSP_Eye_StartBit_DQ, DSP_Eye_BitLength_DQ)
    Call Opt_Input_Parsing(CUS_Str_DigCapData_Dqs, EyeTestRegName, DSP_Eye_StartBit_DQS, DSP_Eye_BitLength_DQS)
    
    If AMP_EYE_VT_CZ_Flag = True Then
        If SweepVtStr <> "" Then
        SplitByColon = Split(SweepVtStr, ":")
        SourceIndexStr = SplitByColon(0)
        SourceIndex = TheExec.Flow.var(SourceIndexStr).value
        StartVal = SplitByColon(1)
        StepVal = SplitByColon(2)
        FinalVal = StartVal + SourceIndex * StepVal
        
            If InStr(UCase(Interpose_PrePat), ":VT:") <> 0 Then
                
                ''''' Purpose to only update VT value and keep the other interpose setting the same
                ReplaceStr = Split(Interpose_PrePat, "VT")
                If InStr(ReplaceStr(1), ";") Then
                    TempStr = Split(ReplaceStr(1), ";")
                    
                    For p = 0 To UBound(TempStr)
                        If p = 0 Then
                            Interpose_PrePat = ReplaceStr(0) & "VT:" & CStr(FinalVal)
                        Else
                            Interpose_PrePat = Interpose_PrePat & ";" & TempStr(p)
                        End If
                    Next p
                Else
                    Interpose_PrePat = ReplaceStr(0) & "VT:" & CStr(FinalVal)
                End If
               
                ''''' For Char TestName
                FinalVal = Format(FinalVal, "0.000")
                Instname_split = Split(glb_TestInstance, "_")
                If FinalVal < 0 Then
                    Testname_CZ_Vt = Replace(CStr(FinalVal), "-", "m")
                Else
                    Testname_CZ_Vt = CStr(FinalVal)
                End If
                Testname_CZ_Vt = Replace(Testname_CZ_Vt, ".", "p")
                'Testname_CZ_Vt = "_" & Instname_split(10) & "_" & Instname_split(1) & "_" & Instname_split(11) & "_" & "VT" & "_" & Testname_CZ_Vt & "_" & Instname_split(UBound(Instname_split))
                Testname_CZ_Vt = "_" & Instname_split(1) & "_" & Instname_split(9) & "_" & Instname_split(10) & "_" & "VT" & "_" & Testname_CZ_Vt & "_" & Instname_split(UBound(Instname_split)) ' update190925 for eye plot
            End If
        End If
    End If

    If Interpose_PrePat <> "" Then
        Call SetForceCondition(Interpose_PrePat & ";STOREPREPAT")
    End If
     
    ''''' Offline Simulation Data
    If TheExec.TesterMode = testModeOffline Then
        DqSwpWf.CreateConstant 0, DigCap_Sample_Size_Dq          ''''' New added to create space for DqSwpWf . Placed TheExec before pattern run for DSP optimization
        DqsSwpWf.CreateConstant 0, DigCap_Sample_Size_Dqs      ''''' New added to create space for DqsSwpWf . Placed TheExec before pattern run for DSP optimization
        
        For Each site In TheExec.sites
            For i = 0 To DigCap_Sample_Size_Dq - 1
               DqSwpWf.Element(i) = Round(Rnd())
            Next i
            
            For i = 0 To DigCap_Sample_Size_Dqs - 1
                DqsSwpWf.Element(i) = Round(Rnd())
            Next i
        Next site
    End If
      
    ''''' Capture Setup and Pattern Run
    PatNames() = TheExec.DataManager.Raw.GetPatternsInSet(DqSwpPat.value, PatCnt)
    Call DigCapSetup(PatNames(0), DigCap_Pin, "Capture_Code_0", DigCap_Sample_Size_Dq, DqSwpWf)
    Call TheHdw.Patterns(PatNames(0)).test(pfAlways, 0)
    Call Update_BC_PassFail_Flag(True)

    If DebugPrintEnable = True Then TheExec.Datalog.WriteComment "  Pattern(" & PatCnt & "): " & PatNames(0) & vbNullString
 
    PatNames() = TheExec.DataManager.Raw.GetPatternsInSet(DqsSwpPat.value, PatCnt)
    Call DigCapSetup(PatNames(0), DigCap_Pin, "Capture_Code_1", DigCap_Sample_Size_Dqs, DqsSwpWf)
    Call TheHdw.Patterns(PatNames(0)).test(pfAlways, 0)
    Call Update_BC_PassFail_Flag(True)

    If DebugPrintEnable = True Then TheExec.Datalog.WriteComment "  Pattern(" & PatCnt & "): " & PatNames(0) & vbNullString
    ''''' End of Capture Setup and Pattern Run

    ''''' DSP DQ/DQS Filter and EYE Width Calculation
    Call rundsp.DSP_Opt_EYE(DqSwpWf, DSP_Eye_StartBit_DQ, DSP_Eye_BitLength_DQ, DqsSwpWf, DSP_Eye_StartBit_DQS, DSP_Eye_BitLength_DQS, NoOfBists, DQ_EYE_Data, DQS_EYE_Data, DSP_Eye_Width)
    
    ''''' DQ/DQS all Capture Code Limits
    Call DigCapDataProcessByDSP(CUS_Str_DigCapData_Dq, DqSwpWf, DigCap_Sample_Size_Dq, 0, , , DigCap_Pin.value)
    Call DigCapDataProcessByDSP(CUS_Str_DigCapData_Dqs, DqsSwpWf, DigCap_Sample_Size_Dqs, 0, , , DigCap_Pin.value)

    ''''' EYE Width Limits
    For BistIdx = 0 To NoOfBists - 1
        If LCase(glb_TestInstance) Like "*cacs*_ck*" Then
        
            If AMP_EYE_VT_CZ_Flag = True Then
                'TestName = Report_TName_From_Instance("calc", "", "EYE_CACS_CK_" & Testname_CZ_Vt & "DDR" & CStr(BistIdx), 0)
                TheExec.Flow.TestLimit resultVal:=DSP_Eye_Width.Element(BistIdx), Tname:="DDR" & CStr(BistIdx) & "_EYE_CACS_CK" & Testname_CZ_Vt, ForceResults:=tlForceFlow
            Else
                TestName = Report_TName_From_Instance("calc", vbNullString, "EYE_CACS_CK_" & "DDR" & CStr(BistIdx), 0)
                TheExec.Flow.TestLimit resultVal:=DSP_Eye_Width.Element(BistIdx), Tname:=TestName, ForceResults:=tlForceFlow
            End If
            
            'TheExec.Flow.TestLimit resultVal:=DSP_Eye_Width.Element(BistIdx), TName:="DDR" & CStr(BistIdx) & "_EYE_CACS_CK" & Testname_CZ_Vt, ForceResults:=tlForceFlow
        Else
            If AMP_EYE_VT_CZ_Flag = True Then
                TheExec.Flow.TestLimit resultVal:=DSP_Eye_Width.Element(BistIdx), Tname:="DDR" & CStr(BistIdx \ 2) & "_EYE_DQ_DQS" & CStr(BistIdx Mod 2) & Testname_CZ_Vt, ForceResults:=tlForceFlow
                'TestName = Report_TName_From_Instance("calc", "", "EYE_DQ_DQS_" & CStr(BistIdx Mod 2) & Testname_CZ_Vt & "DDR" & CStr(BistIdx \ 2), 0)
            Else
                TestName = Report_TName_From_Instance("calc", vbNullString, "EYE_DQ_DQS_" & CStr(BistIdx Mod 2) & "DDR" & CStr(BistIdx \ 2), 0)
                TheExec.Flow.TestLimit resultVal:=DSP_Eye_Width.Element(BistIdx), Tname:=TestName, ForceResults:=tlForceFlow
            End If
            
            'TheExec.Flow.TestLimit resultVal:=DSP_Eye_Width.Element(BistIdx), TName:="DDR" & CStr(BistIdx \ 2) & "_EYE_DQ_DQS" & CStr(BistIdx Mod 2) & Testname_CZ_Vt, ForceResults:=tlForceFlow
        End If
                Call Update_BC_PassFail_Flag
    Next BistIdx
    
    If (DisableConnectPins <> "") Then TheHdw.Digital.Pins(DisableComparePins).Connect
    If (DisableComparePins <> "") Then TheHdw.Digital.Pins(DisableComparePins).DisableCompare = False
        
    ''''' EYE Printing : Display the captured information, as well eye diagrams in the captured order
    If DispCaptStrm Then

        Dim BitStrM As String
        Dim EyeSt As Integer
        
        For Each site In TheExec.sites.Active

            ''''' DQ_EYE_Data Sweep
            BitStrM = CStr(DQ_EYE_Data(site).Element(0))
            For i = 1 To DqSwpNoOfBits - 1
                BitStrM = BitStrM & CStr(DQ_EYE_Data(site).Element(i))
            Next i
            TheExec.Datalog.WriteComment "Site " & site & ": 1st Sweep " & DqSwpNoOfBits & " bits(LSB->MSB) = " & BitStrM

            ''''' DQS_EYE_Data Sweep
            BitStrM = CStr(DQS_EYE_Data(site).Element(0))
            For i = 1 To DqsSwpNoOfBits - 1
                BitStrM = BitStrM & CStr(DQS_EYE_Data(site).Element(i))
            Next i
            TheExec.Datalog.WriteComment "        2nd Sweep " & DqsSwpNoOfBits & " bits(LSB->MSB) = " & BitStrM

            For BistIdx = 0 To NoOfBists - 1
                EyeSt = BistIdx * EyeStrobes

                BitStrM = CStr(DQ_EYE_Data(site).Element(EyeSt))
                For i = 1 To EyeStrobes - 1
                    BitStrM = BitStrM & CStr(DQ_EYE_Data(site).Element(EyeSt + i))
                Next i

                If LCase(glb_TestInstance) Like "*cacs*_ck*" Then
                    TheExec.Datalog.WriteComment "         CACS Eye, DDR" & BistIdx & "eye0" & ": " & BitStrM
                Else
                    TheExec.Datalog.WriteComment "         DQ    Eye, DDR" & CInt(BistIdx \ 2) & "eye" & (BistIdx Mod 2) & ": " & BitStrM
                End If

                BitStrM = CStr(DQS_EYE_Data(site).Element(EyeSt))
                For i = 1 To EyeStrobes - 1
                    BitStrM = BitStrM & CStr(DQS_EYE_Data(site).Element(EyeSt + i))
                Next i

                If LCase(glb_TestInstance) Like "*cacs*_ck*" Then
                    TheExec.Datalog.WriteComment "         CK   Eye, DDR" & BistIdx & "eye0" & ": " & BitStrM
                Else
                    TheExec.Datalog.WriteComment "         DQS   Eye, DDR" & CInt(BistIdx \ 2) & "eye" & (BistIdx Mod 2) & ": " & BitStrM
                End If

            Next BistIdx
        Next site

    End If
    ''''' End of EYE Printing
 
    Pat = DqSwpPat.value & "," & DqsSwpPat.value
    Shmoo_Pattern = DqSwpPat.value & "," & DqsSwpPat.value
    DebugPrintFunc Pat
    
     ''''' Process calculate equation by dictionary
     If Calc_Eqn <> "" Then
         Call ProcessCalcEquation(Calc_Eqn)
     End If
    
''     Added to print out performance mode power level
     
     
     If glb_TestInstance Like "DDR_*" Then '''20190509
            Call CUS_DDR_DCS_PrintOut
     End If
     
   
    If Interpose_PrePat <> "" Then
        Call SetForceCondition("RESTOREPREPAT")
    End If
    'Alarm check from Sicily,20200423, Oscar
    ' Check implicit alarms
    TheHdw.Alarms.Check
    
    
Exit Function

errHandler:
    TheExec.Datalog.WriteComment "Error in Opt_DdrLpBkFunc3"
    If AbortTest Then Exit Function Else Resume Next
  
End Function

Public Function DigSrc_DigCap_Universal_func(Optional patset As Pattern, _
    Optional DigCap_Pin As PinList, Optional DigCap_DataWidth As Long, Optional DigCap_Sample_Size As Long, _
    Optional DigSrc_pin As PinList, Optional DigSrc_DataWidth As Long, Optional DigSrc_Sample_Size As Long, Optional DigSrc_Equation As String, Optional digsrc_assignment As String, Optional DigSrc_FlowForLoopIntegerName As String = vbNullString, _
    Optional CUS_Str_MainProgram As String = vbNullString, Optional CUS_Str_DigCapData As String = vbNullString, Optional CUS_Str_DigSrcData As String = vbNullString, _
    Optional Interpose_PrePat As String, Optional Interpose_PostTest As String, Optional Validating_ As Boolean) As Long
    
    Dim PatCount As Long, PattArray() As String
   
    If Validating_ Then
        Call PrLoadPattern(patset.value)
        Exit Function    ' Exit after validation
    End If

    Call HardIP_InitialSetupForPatgen

    Dim InDSPWave As New DSPWave
    Dim OutDspWave() As New DSPWave
    Dim ShowDec As String, ShowOut As String
    Dim site As Variant
    Dim patt As Variant
    Dim Pat As String
    Dim HighLimitVal() As Double, LowLimitVal() As Double
    Dim i As Long, j As Long, k As Long

''    Dim RTN_InterposeString As String
    On Error GoTo errHandler
    
    ''20141219 Get use-limit from flow table
    Call GetFlowSingleUseLimit(HighLimitVal, LowLimitVal)
    
    TheHdw.Digital.ApplyLevelsTiming True, True, True, tlPowered

    '' 20160923 - Add Interpose_PrePat entry point
    If Interpose_PrePat <> "" Then
        Call SetForceCondition(Interpose_PrePat & ";STOREPREPAT")
    End If
    
    Call HardIP_InitialSetupForPatgen
    gl_TName_Pat = patset.value
    TheHdw.Patterns(patset).Load
    Call PATT_GetPatListFromPatternSet(patset.value, PattArray, PatCount)
    
    ReDim OutDspWave(PatCount - 1) As New DSPWave
    
   For i = 0 To PatCount - 1
        Pat = CStr(PattArray(i))
        
        TheHdw.Patterns(Pat).Load

        Call GeneralDigSrcSetting(Pat, DigSrc_pin, DigSrc_Sample_Size, DigSrc_DataWidth, DigSrc_Equation, digsrc_assignment, _
                                               DigSrc_FlowForLoopIntegerName, CUS_Str_DigSrcData, InDSPWave)
        
        
        Call GeneralDigCapSetting(Pat, DigCap_Pin, DigCap_Sample_Size / PatCount, OutDspWave(i))
        
        'Call PrintDigCapSetting(DigCap_Pin, DigCap_Sample_Size, CUS_Str_DigCapData)
    
    
        Call TheHdw.Patterns(Pat).test(pfAlways, 0)

        If DebugPrintEnable = True Then TheExec.Datalog.WriteComment "  Pattern(" & PatCount & "): " & Pat & vbNullString
        
        TheHdw.Digital.Patgen.HaltWait ' haltwait at patten end
    
        Call SetForceCondition(Interpose_PostTest)
    Next i
        
        Dim OutDspWave_final As New DSPWave
        OutDspWave_final.CreateConstant 0, DigCap_Sample_Size, DspLong
        For Each site In TheExec.sites.Active
            For i = 0 To PatCount - 1
                For j = 0 To 11
                    OutDspWave_final.Element(i * 12 + j) = OutDspWave(i).Element(j)
                Next j
            Next i
        Next site
        '' 20160211 - Process DigCapData by using DSP
        If DigCap_Sample_Size <> 0 Then
            Dim DigCapPinAry() As String, NumberPins As Long
            Call TheExec.DataManager.DecomposePinList(DigCap_Pin, DigCapPinAry(), NumberPins)
            
            If NumberPins > 1 Then
                Call CreateSimulateDataDSPWave_Parallel(OutDspWave_final, DigCap_Sample_Size)
                Call Checker_StoreDigCapAllToDictionary(CUS_Str_DigCapData, OutDspWave_final, NumberPins)
                Call DigCapDataProcessByDSP_Parallel(CUS_Str_DigCapData, OutDspWave_final, DigCap_Sample_Size, NumberPins)
            ElseIf NumberPins = 1 Then
                Call CreateSimulateDataDSPWave(OutDspWave_final, DigCap_Sample_Size, DigCap_DataWidth)
                Call Checker_StoreDigCapAllToDictionary(CUS_Str_DigCapData, OutDspWave_final, NumberPins)
                Call DigCapDataProcessByDSP(CUS_Str_DigCapData, OutDspWave_final, DigCap_Sample_Size, DigCap_DataWidth)
            End If
        End If

    
    DebugPrintFunc patset.value  ' print all debug information
    
    If Interpose_PrePat <> "" Then
        Call SetForceCondition("RESTOREPREPAT")
    End If
    
    Exit Function

errHandler:
    TheExec.Datalog.WriteComment "error in DigSrc_DigCap_Universal_func"
    If AbortTest Then Exit Function Else Resume Next
  
End Function

Public Function TMPS_Voltage_Print(PowerName As String) As Long
If gl_Disable_HIP_debug_log = False Then
    TheExec.Datalog.WriteComment "*******************************"
    TheExec.Datalog.WriteComment "Set " & PowerName & " : " & TheHdw.DCVS.Pins(PowerName).Voltage.value
    TheExec.Datalog.WriteComment "*******************************"
End If

End Function



Public Function ReMeasImpedByAveTrimCode(Optional patset As String, Optional CPUA_Flag_In_Pat As Boolean, _
    Optional MeasR_Pins_SingleEnd As String, Optional MeasR_Pins_Differential As String, Optional StrForceVolt As String, Optional DigSrc_pin As PinList, Optional DigSrc_Sample_Size As Long, _
    Optional TrimStoreName As String, _
    Optional Fixed_DigSrc_DataWidth As Long, Optional Fixed_DigSrc_Sample_Size As Long, Optional Fixed_DigSrc_Equation As String, Optional Fixed_DigSrc_Assignment As String, _
    Optional b_PD_Mode As Boolean = True) As Long
    
    Dim PatCount As Long, PattArray() As String
    Dim InDSPWave As New DSPWave
    Dim site As Variant
    Dim Pat As String
    Dim i As Long, j As Long, k As Long
    Dim MeasureImped As New PinListData

    Dim PatWithPinsInfo(3) As MeasTrimImpedInfo
    Dim Pin_Ary() As String, Pin_Cnt As Long, pin As Variant, TempPins As String
    Dim b_IsDifferential As Boolean
    Dim SingleEndSplitByAdd() As String, DifferentialSplitByAdd() As String, PatternSplitByAdd() As String
    Dim InfoCounter As Long
    
    On Error GoTo ErrorHandler
    
    Call HardIP_InitialSetupForPatgen
    
    SingleEndSplitByAdd = Split(MeasR_Pins_SingleEnd, "+")
    DifferentialSplitByAdd = Split(MeasR_Pins_Differential, "+")
    PatternSplitByAdd = Split(patset, "+")
    
    For InfoCounter = 0 To UBound(PatWithPinsInfo)
        If MeasR_Pins_SingleEnd <> "" Then
            TheExec.DataManager.DecomposePinList SingleEndSplitByAdd(InfoCounter), Pin_Ary, Pin_Cnt
            b_IsDifferential = False
            
        ElseIf MeasR_Pins_Differential <> "" Then
            TheExec.DataManager.DecomposePinList DifferentialSplitByAdd(InfoCounter), Pin_Ary, Pin_Cnt
            
            For i = 0 To Pin_Cnt - 1
                If InStr(UCase(Pin_Ary(i)), "_P") <> 0 Then
                    If i = 0 Then
                        TempPins = Pin_Ary(i)
                    Else
                        TempPins = TempPins & "," & Pin_Ary(i)
                    End If
                End If
            Next i
            Pin_Cnt = Pin_Cnt / 2
            ReDim Pin_Ary(Pin_Cnt) As String
            Pin_Ary = Split(TempPins, ",")
            b_IsDifferential = True
        End If
        
        TheHdw.Patterns(PatternSplitByAdd(InfoCounter)).Load
        Call PATT_GetPatListFromPatternSet(PatternSplitByAdd(InfoCounter), PattArray, PatCount)
        PatWithPinsInfo(InfoCounter).Pat = PattArray(0)
        PatWithPinsInfo(InfoCounter).MeasPinsAry = Pin_Ary
        PatWithPinsInfo(InfoCounter).IsDifferential = b_IsDifferential
    Next InfoCounter

    Dim SplitForceVolt() As String
    SplitForceVolt = Split(StrForceVolt, ",")
    Dim ForceVolt As String
    Call HIP_Evaluate_ForceVal(SplitForceVolt)
    For i = 0 To UBound(SplitForceVolt)
        If i = 0 Then
            ForceVolt = SplitForceVolt(i)
        Else
            ForceVolt = ForceVolt & "," & SplitForceVolt(i)
        End If
    Next i
    
    TheHdw.Digital.Patgen.Halt
    
    TheHdw.Digital.ApplyLevelsTiming True, True, True, tlPowered
    
    
    Dim InitialFixedDSPWave As New DSPWave
    Dim FinalTrimDSPWave_BIN As New DSPWave
    Dim FinalTrimDSPWave_DEC As New DSPWave
    Dim Trim_DigSrc_Sample_Size As Long

    Trim_DigSrc_Sample_Size = DigSrc_Sample_Size - Fixed_DigSrc_Sample_Size
    
    FinalTrimDSPWave_DEC.CreateConstant 0, 1, DspLong
    FinalTrimDSPWave_BIN.CreateConstant 0, Trim_DigSrc_Sample_Size, DspLong
    
    If Fixed_DigSrc_Equation <> "" Then
        For Each site In TheExec.sites.Active
            Call Create_DigSrc_Data(DigSrc_pin, Fixed_DigSrc_DataWidth, Fixed_DigSrc_Sample_Size, Fixed_DigSrc_Equation, Fixed_DigSrc_Assignment, InitialFixedDSPWave, site)
        Next site
        
        If (TheExec.TesterMode = testModeOffline) Then
            FinalTrimDSPWave_DEC.CreateConstant 18, 1, DspLong
        Else
            FinalTrimDSPWave_DEC = GetStoredCaptureData(TrimStoreName)
        End If
''        FinalTrimDSPWave_DEC = FinalTrimDSPWave_DEC.ConvertDataTypeTo(DspLong)
''        FinalTrimDSPWave_BIN = FinalTrimDSPWave_BIN.ConvertDataTypeTo(DspLong)
        Call rundsp.DSPWaveDecToBinary(FinalTrimDSPWave_DEC, Trim_DigSrc_Sample_Size, FinalTrimDSPWave_BIN)
       
        Call rundsp.CombineDSPWave(InitialFixedDSPWave, FinalTrimDSPWave_BIN, Fixed_DigSrc_Sample_Size, Trim_DigSrc_Sample_Size, InDSPWave)
    End If
    Dim OutputTrimCode As String
    For Each site In TheExec.sites
        OutputTrimCode = vbNullString
        For k = 0 To InDSPWave(site).SampleSize - 1
            OutputTrimCode = OutputTrimCode & CStr(InDSPWave(site).Element(k))
        Next k
        If gl_Disable_HIP_debug_log = False Then TheExec.Datalog.WriteComment ("Site_" & site & " Output Trim Code = " & OutputTrimCode)
    Next site

    For InfoCounter = 0 To UBound(PatWithPinsInfo)
        For Each pin In PatWithPinsInfo(InfoCounter).MeasPinsAry
            Call SetupDigSrcDspWave(PatWithPinsInfo(InfoCounter).Pat, DigSrc_pin, "TrimCodeImped", DigSrc_Sample_Size, InDSPWave)
    
            Call TheHdw.Patterns(PatWithPinsInfo(InfoCounter).Pat).start
            Call SubMeasR(CPUA_Flag_In_Pat, CStr(pin), ForceVolt, MeasureImped, PatWithPinsInfo(InfoCounter).IsDifferential, b_PD_Mode)
            TheExec.Flow.TestLimit resultVal:=MeasureImped, unit:=unitCustom, customUnit:="ohm", Tname:="SourceAverCode" & "_Pin_" & pin, ForceResults:=tlForceFlow
        Next pin
    Next InfoCounter
    Exit Function
    
ErrorHandler:
    TheExec.Datalog.WriteComment "error in ReMeasImpedByAveTrimCode function"
    If AbortTest Then Exit Function Else Resume Next
End Function

Public Function TMPS_Bin2Dec(ByRef DataOut_85C As DSPWave, Optional DSPWave_Dict As DSPWave) As Long

Dim i As Integer
Dim Data_Temp As String
Dim site As Variant
    For Each site In TheExec.sites
        For i = 0 To (DSPWave_Dict(site).SampleSize - 1)
            Data_Temp = Data_Temp & (DSPWave_Dict(site).Element(i))
        Next i
            DataOut_85C(site).Element(0) = Bin2Dec_rev(Data_Temp)
            Data_Temp = vbNullString
    Next site

End Function

Public Function TMPS_Dec2Bin(ByRef Read_Code As DSPWave, Optional DSPWave_Dict As DSPWave, Optional dspwavesize) As Long
Dim TempVal As Long
Read_Code.CreateConstant 0, dspwavesize

Dim i As Integer
Dim Data_Temp As String
Dim site As Variant

    For Each site In TheExec.sites
        TempVal = DSPWave_Dict(site).Element(0)
        For i = 0 To dspwavesize - 1
            Read_Code.Element(i) = TempVal Mod 2
            TempVal = TempVal \ 2
        Next i
    Next site
End Function
Public Function TrimCodeDig_SeaHawk(Optional patset As Pattern, Optional TestSequence As String, Optional CPUA_Flag_In_Pat As Boolean, _
    Optional MeasureF_Pin As PinList, Optional DigSrc_pin As PinList, Optional DigSrc_Sample_Size As Long, Optional CUS_Str_DigCapData As String, _
    Optional TrimPrcocessAll As Boolean = False, Optional UseMinimumTrimCode As Boolean = False, Optional PreCheckMinMaxTrimCode As Boolean = False, _
    Optional TrimTarget As Double, Optional TrimTargetTolerance As Double = 0, Optional TrimStart As String, Optional TrimFormat As String, _
    Optional TrimStoreName As String, Optional TrimFuseName As String, Optional TrimFuseTypeName As String, Optional Interpose_PrePat As String, Optional DigCap_Pin As PinList, Optional DigCap_Sample_Size As Long, _
    Optional Validating_ As Boolean, Optional Interpose_PostTest As String, Optional TrimOffset As String, Optional TrimBase As String, Optional DigSrc_Sample_Size_Real As String, Optional DigCap_DataWidth As Long, Optional Digcap_key As String) As Long

    TheHdw.DSP.ExecutionMode = tlDSPModeHostDebug
    TheHdw.DSSC.MoveMode = tlDSSCMoveModeDatabus
    
        glb_TestInstance = vbNullString
    glb_TestInstance = UCase(TheExec.DataManager.instancename)
    
    Dim SourceTrimCode_SiteVariant As New SiteVariant
    
    Dim SourceTrimCode_temp As Long
    Dim SourceTrimCode_final As Long
    Dim TrimBase_temp() As String
    Dim TrimBase_Num As Long
    Dim TrimBase_Item() As String
    Dim TrimBase_DSPWave As New DSPWave
    Dim R As Integer
    Dim TrimBase_DSPWave_Final() As New DSPWave
    Dim OutputTname_format() As String
    Dim TestNameInput As String
    
    TrimBase_temp = Split(TrimBase, ";")
    TrimBase_Num = UBound(TrimBase_temp)
    
    Dim DigSrc_Sample_Size_Real_Temp() As String    ' Fix 20190812
    Dim Divide_Result As Long
    DigSrc_Sample_Size_Real_Temp = Split(DigSrc_Sample_Size_Real, "@")
    Divide_Result = DigSrc_Sample_Size_Real_Temp(1) / DigSrc_Sample_Size_Real_Temp(0)
          
    Dim Bin_arry() As Long
    ReDim Bin_arry(DigSrc_Sample_Size_Real_Temp(0) - 1)
    Dim InDspWave_New As New DSPWave
    Dim Dec_Trim_Temp As Long
    Dim t As Integer
    
    Dim PatCount As Long, PattArray() As String
    Dim OutDspWave As New DSPWave
               TrimTarget = 0.3 ''Just for Flag    2017/Dec
    If Validating_ Then
        Call PrLoadPattern(patset.value)
        Exit Function    ' Exit after validation
    End If
    Call HardIP_InitialSetupForPatgen
    Dim Ts As Variant, TestSequenceArray() As String
    Dim InitialDSPWave As New DSPWave, PastDSPWave As New DSPWave, InDSPWave As New DSPWave

    Dim site As Variant
    Dim Pat As String
    Dim i As Long, j As Long, k As Long, z As Long
    'Dim Inst_Name_Str As String: Inst_Name_Str = glb_TestInstance
    
    Dim CapValue As New PinListData, CapValue_V1 As New PinListData, CapValue_V2 As New PinListData
    CapValue.AddPin ("CapValueString")
    On Error GoTo ErrorHandler

    Call GetFlowTName
 
    TestSequenceArray = Split(TestSequence, ",")
    TheHdw.Digital.Patgen.Halt
    
    TheHdw.Digital.ApplyLevelsTiming True, True, True, tlPowered
    
    If Interpose_PrePat <> "" Then ''''180109 update
        Call SetForceCondition(Interpose_PrePat & ";STOREPREPAT")
    End If
    
    TheHdw.Patterns(patset).Load

    Call PATT_GetPatListFromPatternSet(patset.value, PattArray, PatCount)

    
    '' 20160425 - Check format from TrimFormat
    Dim StrSeparatebyComma() As String
    Dim ExecutionMax As Long
    StrSeparatebyComma = Split(TrimFormat, ";")
    ExecutionMax = UBound(StrSeparatebyComma)
    Dim StrSeparatebyEqual() As String, StrSeparatebyColon() As String '' Get Src bit
    Dim SrcStartBit As Long, SrcEndBit As Long
    
    Dim b_HigherThanTarget As New SiteBoolean
    b_HigherThanTarget = False
    
    Dim LastSectionV1V2_Index As Long
    LastSectionV1V2_Index = 0
    Dim OutputTrimCode As String
    Dim SourceTrimCode As String
    Dim TestLimitIndex As Long '
    
'    Dim TrimStart_1st() As String
    Dim Dec_TrimStart_1st As Long
    
    Dim StoredTargetTrimCode As New DSPWave
    Dim b_MatchTagetCap As New SiteBoolean
    Dim b_DisplayCap As New SiteBoolean
    b_MatchTagetCap = False
    b_DisplayCap = False
    
    StoredTargetTrimCode.CreateConstant 0, DigSrc_Sample_Size, DspLong
    Dim StoreEachTrimCode() As New DSPWave
    ReDim StoreEachTrimCode(DigSrc_Sample_Size + 1) As New DSPWave
    
    Dim DigCapData() As New PinListData
    ReDim DigCapData(DigSrc_Sample_Size + 1) As New PinListData
    
    Dim StoreEachIndex As Long
    
    ''20161128-Stop trim code process
    Dim b_StopTrimCodeProcess As New SiteBoolean
    b_StopTrimCodeProcess = False
    
    For i = 0 To UBound(StoreEachTrimCode)
        StoreEachTrimCode(i).CreateConstant 0, DigSrc_Sample_Size, DspLong
    Next i

    If TrimStart <> "" And TrimStart Like "*&*" Then
        TrimStart = Replace(TrimStart, "&", vbNullString)
    End If

'' ===============================   TrimStart_1st = TrimStart===============================================
    Dec_TrimStart_1st = Bin2Dec(TrimStart)
    
    InitialDSPWave.CreateConstant Dec_TrimStart_1st, 1, DspLong

    Call rundsp.CreateFlexibleDSPWave(InitialDSPWave, DigSrc_Sample_Size, InDSPWave)

    'Call SetupDigSrcDspWave(PattArray(0), DigSrc_pin, "TrimCodeCap", DigSrc_Sample_Size, InDspWave)
    TheExec.Datalog.WriteComment ("========First Time Setup========")
    
    For Each site In TheExec.sites
        SourceTrimCode = vbNullString
        For k = 0 To InDSPWave(site).SampleSize - 1
            SourceTrimCode = SourceTrimCode & CStr(InDSPWave(site).Element(k))
        Next k
        TheExec.Datalog.WriteComment ("Site_" & site & " Initial Source Trim Code = " & SourceTrimCode)
        SourceTrimCode = StrReverse(SourceTrimCode)
        Dec_Trim_Temp = Bin2Dec(SourceTrimCode)
        TrimOffset = CInt(TrimOffset)
        Dec_Trim_Temp = Dec_Trim_Temp + TrimOffset
        
        'TheExec.Datalog.WriteComment ("Site_" & Site & " Initial Source Trim Code = " & SourceTrimCode)
    Next site
    
    InDspWave_New.CreateConstant 0, DigSrc_Sample_Size_Real_Temp(1)
                
    Call Dec2Bin(Dec_Trim_Temp, Bin_arry())
    
    For z = 0 To UBound(Bin_arry)
        InDspWave_New.Element(z) = Bin_arry(UBound(Bin_arry) - z)
        For t = 1 To Divide_Result - 1
            InDspWave_New.Element(z + ((UBound(Bin_arry) + 1) * t)) = InDspWave_New.Element(z)
        Next t
    Next z
    
    Call SetupDigSrcDspWave(PattArray(0), DigSrc_pin, "TrimCodeCap", CLng(DigSrc_Sample_Size_Real_Temp(1)), InDspWave_New)
    
    For Each site In TheExec.sites
        StoreEachTrimCode(0)(site) = InDSPWave(site).Copy
    Next site
    
'=========Set Up DigCap parameter=============================
    Dim Decompose_DigCapData() As String
    Dim OutDspWave_elementnum As Long
    Dim Digcap_width() As String
    Dim Digcap_sum As Long
    
    
    If Digcap_key <> "" Then
        Decompose_DigCapData() = Split(CUS_Str_DigCapData, ",")
        For i = 1 To UBound(Decompose_DigCapData)
            Digcap_width() = Split(Decompose_DigCapData(i), ":")
            Digcap_sum = Digcap_sum + Digcap_width(0)
            If InStr(Decompose_DigCapData(i), Digcap_key) <> 0 Then
                OutDspWave_elementnum = Digcap_sum - 1
                TheExec.Datalog.WriteComment "capture digcap name and bits  " & Decompose_DigCapData(i)
                Exit For
            End If
        Next i
    Else
        OutDspWave_elementnum = 0
    End If
     
    
    Call GeneralDigCapSetting(PattArray(0), DigCap_Pin, DigCap_Sample_Size, OutDspWave)
    Call TheHdw.Patterns(PattArray(0)).start
'     For Each Site In TheExec.sites
        CapValue.value = OutDspWave.Element(OutDspWave_elementnum)
'     Next Site
    Call PrintDigCapSetting(DigCap_Pin, DigCap_Sample_Size, CUS_Str_DigCapData)
    ''Update Interpose_PreMeas 20170801
    Dim TestSeqNum As Integer
    TestSeqNum = 0
           
    TheHdw.Digital.Patgen.HaltWait
'    For Each Site In TheExec.sites
    DigCapData(0) = CapValue
    b_HigherThanTarget = CapValue.Math.Subtract(TrimTarget).compare(GreaterThan, 0)
    PastDSPWave = InDSPWave
    TestNameInput = "Cap_Value_"
    TestLimitIndex = 0
'    Next Site

    Dim CUS_Str_MainProgram As String: CUS_Str_MainProgram = vbNullString
    
    For Each site In TheExec.sites
        If b_MatchTagetCap(site) = False And b_DisplayCap(site) = False Then
            TheExec.Datalog.WriteComment ("Site " & site & " Output CapValue = " & OutDspWave(site).Element(0))
        End If
    Next site
    
    b_MatchTagetCap = CapValue.Math.Subtract(TrimTarget).Abs.compare(LessThanOrEqualTo, TrimTargetTolerance)
    b_DisplayCap = b_DisplayCap.LogicalOr(b_MatchTagetCap)
    
    For Each site In TheExec.sites
        If b_DisplayCap(site) = True And StoredTargetTrimCode(site).CalcSum = 0 Then
            StoredTargetTrimCode(site) = InDSPWave(site).Copy
            b_StopTrimCodeProcess(site) = True
        End If
    Next site
    
    TheExec.Datalog.WriteComment ("======================================================================================")
    
    Dim b_KeepGoing As New SiteBoolean
    
    Dim PreviousCap As New PinListData
   
    Dim b_ControlNextBit As Boolean
    b_ControlNextBit = False
    Dim b_FirstExecution As Boolean
    b_FirstExecution = False
    StoreEachIndex = 1
    
    If PreCheckMinMaxTrimCode = False Then
        b_KeepGoing = True
    End If
    
    If b_KeepGoing.All(False) Then
    Else
        For i = 0 To ExecutionMax
            If TrimPrcocessAll = False Then
                If b_StopTrimCodeProcess.All(True) Then
                    Exit For
                End If
            End If
                    
            StrSeparatebyEqual = Split(StrSeparatebyComma(i), "=")
            StrSeparatebyColon = Split(StrSeparatebyEqual(1), ":")
            SrcStartBit = StrSeparatebyColon(0)
            SrcEndBit = StrSeparatebyColon(1)
                    
            If i = 0 Then
                b_FirstExecution = True
            Else
                b_FirstExecution = False
                SrcStartBit = SrcStartBit + 1
            End If
                
            For j = SrcStartBit To (SrcEndBit + 1) Step -1
                ''===============up for src bit step
        
        '            For Each Site In TheExec.sites
        '                CapValue = CStr(OutDspWave(Site).Element(0))
        '            Next Site
                If b_FirstExecution = True Then
                    b_ControlNextBit = True
                    If j = SrcEndBit Then ''=============Trim from Format untill same======1124
                        b_ControlNextBit = False
                    End If
                Else
                    ''20160716-Control next bit to 1 no matter first or last progress
                    b_ControlNextBit = True
                    If j = SrcEndBit Then
                        b_ControlNextBit = False
                    End If
                End If
        
                If b_FirstExecution = True And j = SrcEndBit Then
                    Call rundsp.SetupTrimCodeBit(PastDSPWave, True, j, b_ControlNextBit, InDSPWave)
                Else
                    Call rundsp.SetupTrimCodeBit(PastDSPWave, b_HigherThanTarget, j, b_ControlNextBit, InDSPWave)
                End If
        
        '            Dim Dec_Trim_Temp As Long
                    
                For Each site In TheExec.sites
                    SourceTrimCode = vbNullString
                    For k = 0 To InDSPWave(site).SampleSize - 1
                        SourceTrimCode = SourceTrimCode & CStr(InDSPWave(site).Element(k))
                    Next k
                    SourceTrimCode = StrReverse(SourceTrimCode)
                    Dec_Trim_Temp = Bin2Dec(SourceTrimCode)
                    TrimOffset = CInt(TrimOffset)
                    Dec_Trim_Temp = Dec_Trim_Temp + TrimOffset
                    
        '                Dim DigSrc_Sample_Size_Real_Temp() As String
        '                Dim Divide_Result As Long
        '                DigSrc_Sample_Size_Real_Temp = Split(DigSrc_Sample_Size_Real, "@")
        '                Divide_Result = DigSrc_Sample_Size_Real_Temp(1) / DigSrc_Sample_Size_Real_Temp(0)
        '
        '
        '                Dim Bin_arry() As Long
        '                ReDim Bin_arry(DigSrc_Sample_Size_Real_Temp(0) - 1)
        '                Dim InDspWave_New As New DSPWave
                    InDspWave_New.CreateConstant 0, DigSrc_Sample_Size_Real_Temp(1)
                    
                    Call Dec2Bin(Dec_Trim_Temp, Bin_arry())
                    
                    Dim Array_code As String
                    'Dim t As Integer
                    
                    Array_code = vbNullString
                    
                    For z = 0 To UBound(Bin_arry)
                        InDspWave_New.Element(z) = Bin_arry(UBound(Bin_arry) - z)
                        For t = 1 To Divide_Result - 1
                            InDspWave_New.Element(z + ((UBound(Bin_arry) + 1) * t)) = InDspWave_New.Element(z)
                        Next t
                        Array_code = Array_code & InDspWave_New(site).Element(z)
                    Next z
                    TheExec.Datalog.WriteComment "Site " & site & "  ,Digcap Decimal value:  " & Dec_Trim_Temp & "  ,Array Code:" & Array_code
                Next site
                    
         
                Call SetupDigSrcDspWave(PattArray(0), DigSrc_pin, "TrimCodeCap", CLng(DigSrc_Sample_Size_Real_Temp(1)), InDspWave_New)
                   
                For Each site In TheExec.sites
                    StoreEachTrimCode(StoreEachIndex)(site) = InDSPWave(site).Copy
                Next site
                    
                '' Debug use
                '' ==============================================================================================
                '' 20160716 - Modify trim code rule
                If b_FirstExecution = True Then
                    If j = SrcEndBit Then
                        TheExec.Datalog.WriteComment ("Setup Bit " & j & " to 0")
                    Else
                        TheExec.Datalog.WriteComment ("Setup Bit " & j & ", Trim Code Bit " & j - 1)
                    End If
                Else
                    If j = SrcEndBit Then
                        TheExec.Datalog.WriteComment ("Setup Bit " & j)
                    Else
                        TheExec.Datalog.WriteComment ("Setup Bit " & j & ", Trim Code Bit " & j - 1)
                    End If
                End If
                    
                For Each site In TheExec.sites
                    If b_MatchTagetCap(site) = False And b_DisplayCap(site) = False Then
                        If b_KeepGoing(site) = True Then

                            SourceTrimCode_SiteVariant(site) = vbNullString
                            
                            For k = 0 To InDSPWave(site).SampleSize - 1

                                SourceTrimCode_SiteVariant = SourceTrimCode_SiteVariant & CStr(InDSPWave(site).Element(k))
                            Next k
                            Dim OutputDec As String

                            TheExec.Datalog.WriteComment ("Site_" & site & " Source Trim Code = " & SourceTrimCode_SiteVariant(site))
                        End If
                    End If
                Next site
                
                Set OutDspWave = Nothing
                Call GeneralDigCapSetting(PattArray(0), DigCap_Pin, DigCap_Sample_Size, OutDspWave)
                Call TheHdw.Patterns(PattArray(0)).start
                    
                    ''Update Interpose_PreMeas 20170801
                TestSeqNum = 0
                                
                
                TheHdw.Digital.Patgen.HaltWait
                
                CapValue.value = OutDspWave.Element(OutDspWave_elementnum)
                b_HigherThanTarget = CapValue.Math.Subtract(TrimTarget).compare(GreaterThan, 0)
                PastDSPWave = InDSPWave
        
                TestLimitIndex = TestLimitIndex + 1
                    
                    '' 20160712 - Modify to use WriteComment to display output frequency.
                For Each site In TheExec.sites
                    If b_KeepGoing(site) = True Then
        '                    TheExec.Datalog.WriteComment ("Site " & Site & " Output CapValue = " & CapValue.Pins(Site).Value(Site))
                        TheExec.Datalog.WriteComment ("Site " & site & " Output CapValue = " & OutDspWave(site).Element(0))
                    End If
                Next site
               
                b_MatchTagetCap = CapValue.Math.Subtract(TrimTarget).Abs.compare(LessThanOrEqualTo, TrimTargetTolerance)
                b_DisplayCap = b_DisplayCap.LogicalOr(b_MatchTagetCap)
                For Each site In TheExec.sites
                    If b_KeepGoing(site) = True Then
                        If b_MatchTagetCap(site) = True And StoredTargetTrimCode(site).CalcSum = 0 Then
                            StoredTargetTrimCode(site) = InDSPWave(site).Copy
                            b_StopTrimCodeProcess(site) = True
                        End If
                    End If
                Next site
               ''20161128-Stop trim code process if found out match code of all site
                If TrimPrcocessAll = False Then
                    If b_StopTrimCodeProcess.All(True) Then
                        Exit For
                    End If
                End If
                TheExec.Datalog.WriteComment ("======================================================================================")
            Next j
        
        
                    
                    
                    
            If TrimOffset <> "" Then
                    
            'Dim SourceTrimCodeArray(8) As Long
                TrimOffset = CInt(TrimOffset)
                If TrimBase <> "" Then
                        ''Dim TrimBase_SiteNum() As String
                        
        ''                    Dim gDictDSPWaves As Scriptin.Dictionary
        ''                    Set gDictDSPWaves = New Scripting.Dictionary
                        
                    TrimBase_temp = Split(TrimBase, ";")
                    TrimBase_Num = UBound(TrimBase_temp)
        
                    ReDim TrimBase_DSPWave_Final(TrimBase_Num)
                    TrimBase_DSPWave.CreateConstant 0, DigSrc_Sample_Size_Real_Temp(0)
                    TrimBase_DSPWave_Final(0).CreateConstant 0, DigSrc_Sample_Size_Real_Temp(0)
                    

                    For R = 0 To TrimBase_Num
                        TrimBase_Item = Split(TrimBase_temp(R), ":")
                        For Each site In TheExec.sites

                            SourceTrimCode_SiteVariant(site) = StrReverse(SourceTrimCode_SiteVariant(site))   ' for Bin to Dec need to reverse
                            SourceTrimCode_temp = Bin2Dec(SourceTrimCode_SiteVariant(site))
                            SourceTrimCode_temp = SourceTrimCode_temp + TrimOffset - 256
                            SourceTrimCode_final = TrimBase_Item(1) - SourceTrimCode_temp
                            
                            
                            Dim printDec As String
                            printDec = SourceTrimCode_final
                            'TrimBase_DSPWave(Site).Element(0) = SourceTrimCode_final
                            'SourceTrimCode = Dec2Bin(SourceTrimCode_temp, SourceTrimCodeArray())
                            '///////////////////////////////////Add to src 9 bit binary 20190528/////////////////////////////////
                            If SourceTrimCode_final < 0 Then
                                If SourceTrimCode_final <= 0 - 2 ^ (DigSrc_Sample_Size_Real_Temp(0) - 1) Then
                                    TheExec.Datalog.WriteComment ("Your number is too small and current bits are not enough to save")
                                    GoTo ErrorHandler
                                Else
                                    SourceTrimCode_final = SourceTrimCode_final + 2 ^ (DigSrc_Sample_Size_Real_Temp(0))
                                End If
                            End If
                       
                            Dim Bin_arry_Base() As Long
                            ReDim Bin_arry_Base(DigSrc_Sample_Size_Real_Temp(0) - 1)
                            Dim Array_code_Base As String
                            
                            Array_code_Base = vbNullString
                            Call Dec2Bin(SourceTrimCode_final, Bin_arry_Base())
                    
                            For z = 0 To UBound(Bin_arry_Base)
                                TrimBase_DSPWave.Element(z) = Bin_arry_Base(UBound(Bin_arry) - z)
                                Array_code_Base = Array_code_Base & TrimBase_DSPWave(site).Element(z)
                            Next
                            '////////////////////////////////////////////////////////////////////////////////////////////////////
                            TheExec.Datalog.WriteComment ("Site_" & site & " Final Source Trim Code = " & SourceTrimCode_SiteVariant(site) & " & " & TrimBase_Item(0) & " = " & Array_code_Base & ",Decimal value: " & printDec)

                            'theexec.Datalog.WriteComment ("Site_" & Site & " Final Source Trim Code = " & SourceTrimCode & " & " & TrimBase_Item(0) & " = " & SourceTrimCode_final)
                        Next site
                        
                        TrimBase_DSPWave_Final(R) = TrimBase_DSPWave
                        Call AddStoredCaptureData(TrimBase_Item(0), TrimBase_DSPWave_Final(R))
                       '' TrimBase_DSPWave(R).CreateConstant , SourceTrimCode_final
                       '' AddStoredCaptureData TrimBase_Item(0) + CStr(site), TrimBase_DSPWave(R)
                       '' TheExec.Datalog.WriteComment ("Site_" & site & " : " & TrimBase_Item(0) & " = " & SourceTrimCode_final)
                    Next R
                End If
            Else
                For Each site In TheExec.sites
                    If CapValue = 1 Then

                        SourceTrimCode_SiteVariant(site) = vbNullString
                        SourceTrimCode_SiteVariant(site) = SourceTrimCode_SiteVariant(site) & "0"
                        For k = 1 To InDSPWave(site).SampleSize - 1
                            SourceTrimCode_SiteVariant(site) = SourceTrimCode_SiteVariant(site) & CStr(InDSPWave(site).Element(k))
                        Next k
                        InDSPWave(site).Element(0) = 0
                    Else

                        
                        SourceTrimCode_SiteVariant(site) = vbNullString
                        For k = 0 To InDSPWave(site).SampleSize - 1
                            SourceTrimCode_SiteVariant(site) = SourceTrimCode_SiteVariant(site) & CStr(InDSPWave(site).Element(k))
                        Next k
                    End If

                    TheExec.Datalog.WriteComment ("Site_" & site & " Final Source Trim Code = " & SourceTrimCode_SiteVariant(site))
                Next site
            End If
        Next i
    End If
    
    Call HardIP_WriteFuncResult(, , glb_TestInstance)
    
    '//////////////////////////////////////////    add for capture bit more than one and print
    
    If DigCap_DataWidth <> 0 Then
        Dim DigCapPinAry() As String, NumberPins As Long
        Dim CUS_Str_DigCapData_temp As String
        Call TheExec.DataManager.DecomposePinList(DigCap_Pin, DigCapPinAry(), NumberPins)
    
        If NumberPins > 1 Then
            Call CreateSimulateDataDSPWave_Parallel(OutDspWave, DigCap_Sample_Size)
            Call Checker_StoreDigCapAllToDictionary(CUS_Str_DigCapData, OutDspWave, NumberPins)
            Call DigCapDataProcessByDSP_Parallel(CUS_Str_DigCapData, OutDspWave, DigCap_Sample_Size, NumberPins, , DigCap_Pin.value)
    
        ElseIf NumberPins = 1 Then
            Call CreateSimulateDataDSPWave(OutDspWave, DigCap_Sample_Size, DigCap_DataWidth)
            Call Checker_StoreDigCapAllToDictionary(CUS_Str_DigCapData, OutDspWave, NumberPins)
    '                                                                                                                                                                Dim VBT_LIB_HardIP_ProfileMark_1429 As Long: VBT_LIB_HardIP_ProfileMark_1429 = ProfileMarkEnter(2, instance_name & "_" & "MeasC&Module=VBT_LIB_HardIP&ProcName=Meas_FreqVoltCurr_Universal_func&LineNumber=1427")    ' Profile Mark
            Call DigCapDataProcessByDSP(CUS_Str_DigCapData, OutDspWave, DigCap_Sample_Size, DigCap_DataWidth, CUS_Str_MainProgram, , DigCap_Pin.value)
    '                                                                                                                                                                ProfileMarkLeave VBT_LIB_HardIP_ProfileMark_1429    ' Profile Mark
        End If
    End If
    '//////////////////////////////////////////
    
    If TrimStoreName <> "" Then
       Call Checker_StoreDigCapAllToDictionary(TrimStoreName, InDSPWave)
    End If
    
    Dim ConvertedDataWf As New DSPWave

    rundsp.ConvertToLongAndSerialToParrel InDSPWave, DigSrc_Sample_Size, ConvertedDataWf
    
    Call GetFlowTName
    
    If gl_UseStandardTestName_Flag = True Then
        Call Report_ALG_TName_From_Instance(OutputTname_format, "C", StrSeparatebyEqual(0), gl_Tname_Meas_FromFlow(TheExec.Flow.TestLimitIndex))
        TestNameInput = Merge_TName(OutputTname_format)
    Else
        TestNameInput = glb_TestInstance & "DDR_Sweep"
    End If
        
        
    TheExec.Flow.TestLimit ConvertedDataWf.Element(0), Tname:=TestNameInput, PinName:="Voffset_Trim_Dec", ForceResults:=tlForceFlow
        
    
'    TheExec.Flow.TestLimit ConvertedDataWf.Element(0), TName:=TheExec.DataManager.instanceName, PinName:="SEPVM_Trim_Dec", ForceResults:=tlForceFlow
    
    Call SetupDigSrcDspWave(PattArray(0), DigSrc_pin, "TrimCodeCap", DigSrc_Sample_Size, StoredTargetTrimCode)
    Call TheHdw.Patterns(PattArray(0)).start

    ''Update Interpose_PreMeas 20170801
    TestSeqNum = 0
   
    TheHdw.Digital.Patgen.HaltWait
    
    If Interpose_PrePat <> "" Then '''180109 update
        Call SetForceCondition("RESTOREPREPAT")
    End If
    
    Call SetForceCondition(Interpose_PostTest)
    
    Dim sl_FUSE_Val As New SiteLong

    DebugPrintFunc patset.value
    
    TheHdw.DSSC.MoveMode = tlDSSCMoveModeIIM
    
    Exit Function
    
ErrorHandler:
    TheExec.Datalog.WriteComment "error in TrimCodeDig_SeaHawk function"
    If AbortTest Then Exit Function Else Resume Next
    
    
End Function





Public Function TrimCodeBasicDig(Optional patset As Pattern, Optional DigSrc_pin As PinList, Optional DigSrc_Sample_Size As String, _
Optional TrimTarget As Double, Optional TrimStart As String, Optional TrimFormat As String, Optional TrimStoreName As String, _
Optional IncreaseFlag As Boolean = True, Optional BinarySearchFlag As Boolean = True, Optional TrimPrcocessAll As Boolean = True, _
Optional Interpose_PrePat As String, Optional DigCap_Pin As PinList, Optional DigCap_Sample_Size As Long, Optional Validating_ As Boolean, _
Optional Interpose_PostTest As String, Optional TrimOffset As String, Optional TrimBase As String, Optional DigSrc_Sample_Size_Real As String, Optional digsrc_assignment As String) As Long
'Dylan Edited 20190615

    If Validating_ Then
        Call PrLoadPattern(patset.value)
        Exit Function    ' Exit after validation
    End If
    On Error GoTo ErrorHandler
    
    Dim site As Variant
    Dim PatCount As Long

    glb_TestInstance = vbNullString
    glb_TestInstance = UCase(TheExec.DataManager.instancename)
    
    Dim TrimBaseStr() As String
    Dim TrimBaseNum() As String
    Dim OffsetDelta As Integer
    'Dim Inst_Name_Str As String: Inst_Name_Str = glb_TestInstance
    Dim OutputTname_format() As String
    Dim TestNameInput As String
    
    Dim LimitCodeStart As Integer
    Dim LimitCodeEnd As Integer
    Dim i, j, k, x As Integer
    Dim SrcEndBit As Integer
    Dim RegSeparate As Integer
    Dim SrcStartBit As Integer
    Dim ExecutionMax As Integer
    Dim DecTrimStart As Integer
    Dim InDSPWave As New DSPWave
    Dim OutDspWave As New DSPWave
    Dim InDspWave_New As New DSPWave
    Dim InitialDSPWave As New DSPWave
    Dim CaptureDSPWave() As New DSPWave
    ReDim CaptureDSPWave(0)
    Dim ProcessDoneDSPWave() As New DSPWave
    ReDim ProcessDoneDSPWave(0)
    Dim ConvertedDataWf As New DSPWave
    
    Dim PattArray() As String
    Dim SourceTrimCode As String
    Dim CapValue As New PinListData
    Dim StrSeparatebyEqual() As String
    Dim StrSeparatebyColon() As String
    Dim StrSeparatebyComma() As String
    Dim RegSeparatebyComma() As String
    Dim EachRegSize As New SiteLong
    Dim InitStateCode As New SiteLong
    Dim TrimOriginalSize As New SiteLong
    Dim TrimScanPoint As New SiteLong
    Dim TrimOffsetPoint As New SiteLong
    Dim CaptureDSPFlag As New SiteBoolean
    Dim b_HigherThanTarget As New SiteBoolean
    Dim b_StopTrimCodeProcess As New SiteBoolean
    Dim DigSrc_Sample_Size_Real_Temp() As String
    Dim assignment As New DSPWave
    Dim DigSrc_Assignment_Temp() As String
    Dim AssignmentDSPWave As New DSPWave
    Dim Src_dig As New SiteBoolean
    Dim b_ControlNextBit As Boolean
    b_ControlNextBit = False
    
                
    Call HardIP_InitialSetupForPatgen
    Call GetFlowTName
    TheHdw.Patterns(patset).Load
    TheHdw.Digital.ApplyLevelsTiming True, True, True, tlPowered
    Call PATT_GetPatListFromPatternSet(patset.value, PattArray, PatCount)
    
    If Interpose_PrePat <> "" Then
        Call SetForceCondition(Interpose_PrePat & ";STOREPREPAT")
    End If
    
    CapValue.AddPin ("CapValueString")
    StrSeparatebyComma = Split(TrimFormat, ";")
    ExecutionMax = UBound(StrSeparatebyComma)
    RegSeparatebyComma = Split(DigSrc_Sample_Size_Real, ";")
    Src_dig = False
    assignment.CreateConstant 0, 1, DspLong
    AssignmentDSPWave.CreateConstant 0, 1, DspLong
    
    If digsrc_assignment <> "" Then
    
    DigSrc_Assignment_Temp = Split(digsrc_assignment, "+")
    AssignmentDSPWave.CreateConstant 0, UBound(DigSrc_Assignment_Temp) + 1, DspLong
    For i = 0 To UBound(DigSrc_Assignment_Temp)
        If LCase(DigSrc_Assignment_Temp(i)) <> "sweep" Then
          assignment = GetStoredCaptureData(DigSrc_Assignment_Temp(i))
          AssignmentDSPWave.Element(i) = 1
          Src_dig = True
        End If
    Next
    
    End If
    
    For i = 0 To ExecutionMax
        StrSeparatebyEqual = Split(StrSeparatebyComma(0), "=")
        StrSeparatebyColon = Split(StrSeparatebyEqual(1), ":")
        SrcEndBit = StrSeparatebyColon(1)
        SrcStartBit = StrSeparatebyColon(0)
        DigSrc_Sample_Size_Real_Temp = Split(RegSeparatebyComma(i), "@")
        CaptureDSPWave(0).CreateConstant 0, 1, DspLong                  ' Avoid sweep fail which any site
        RegSeparate = DigSrc_Sample_Size_Real_Temp(1) / DigSrc_Sample_Size_Real_Temp(0)
        InDSPWave.CreateConstant 0, CLng(DigSrc_Sample_Size_Real_Temp(1)), DspLong
        InDspWave_New.CreateConstant 0, CLng(DigSrc_Sample_Size_Real_Temp(1)), DspLong
        
        If BinarySearchFlag = False Then                                ' This initial method Only support linear search mode
            If TrimStart = "" Then
                If IncreaseFlag = True Then                             ' Based on IncreaseFlag state to determine start point
                    TrimStart = 0
                Else
                    TrimStart = CStr(2 ^ (SrcStartBit + 1) - 1)
                End If
            End If
        End If
        
'''        For Each site In TheExec.sites.Active
            CaptureDSPFlag = False
            If UBound(StrSeparatebyColon) > 1 Then
                InitStateCode = StrSeparatebyColon(2)
            End If
            EachRegSize = CLng(DigSrc_Sample_Size_Real_Temp(0))
            TrimScanPoint = CLng(TrimStart)
            TrimOffsetPoint = CLng(TrimOffset)
            TrimOriginalSize = CLng(SrcStartBit) + 1
'''        Next site
        
        
        
        TrimStart = CStr(CInt(TrimStart) + CInt(TrimOffset))
        InitialDSPWave.CreateConstant TrimStart, 1, DspLong                           ' Define first trim code from TrimStart
        rundsp.CreateFlexibleDSPWave InitialDSPWave, CLng(DigSrc_Sample_Size_Real_Temp(0)), InDSPWave
        rundsp.ElementTransformer InDSPWave, CLng(DigSrc_Sample_Size_Real_Temp(0)), CLng(DigSrc_Sample_Size_Real_Temp(1))
        
               
        If BinarySearchFlag = True Then                                               ' Binary Search
            For j = SrcStartBit + 1 To SrcEndBit Step -1
            
               If j = SrcEndBit Then
                   b_ControlNextBit = False
            
               Else
                    b_ControlNextBit = True
            
               End If
               
            
                rundsp.ReAssignmentDSPWave InDSPWave, RegSeparate, InDspWave_New, Src_dig, assignment, AssignmentDSPWave              ' ReAssignment DSPWave element to each register
                
                
                
                '*****************************************************For HardIP_D2D debug*****************************************************
'''''''''''                Dim Constant As Integer
'''''''''''                Dim ForConstantCode As String
'''''''''''                Dim ForConstantSplit() As String
'''''''''''                ForConstantCode = "1,0,1,0,1,0,1,0"
'''''''''''                ForConstantCode = StrReverse(ForConstantCode)
'''''''''''
'''''''''''                ForConstantSplit = Split(ForConstantCode, ",")
'''''''''''                If theexec.DataManager.instanceName = "D2D_ZCPDD2DIMPCL_PP_SHKA0_C_FULP_AN_AMXX_DLL_JTG_PRG_ALLFV_SI_D2DIMPCL_ZCPD_NV" Then
'''''''''''                    Constant = 0
'''''''''''                    For k = 0 To 7                                    '|TrimCode|Constanct|TrimCode|Constanct|Constanct|
'''''''''''                        If CStr(ForConstantSplit(Constant)) = "0" Then
'''''''''''                            InDspWave_New(0).Element(k) = 0
'''''''''''                        Else
'''''''''''                            InDspWave_New(0).Element(k) = 1
'''''''''''                        End If
'''''''''''                        Constant = Constant + 1
'''''''''''                    Next k
'''''''''''                    Constant = 0
'''''''''''                    For k = 8 To 15
'''''''''''                        If CStr(ForConstantSplit(Constant)) = "0" Then
'''''''''''                            InDspWave_New(0).Element(k) = 0
'''''''''''                        Else
'''''''''''                            InDspWave_New(0).Element(k) = 1
'''''''''''                        End If
'''''''''''                        Constant = Constant + 1
'''''''''''                    Next k
'''''''''''                    Constant = 0
'''''''''''                    For k = 24 To 31
'''''''''''                        If CStr(ForConstantSplit(Constant)) = "0" Then
'''''''''''                            InDspWave_New(0).Element(k) = 0
'''''''''''                        Else
'''''''''''                            InDspWave_New(0).Element(k) = 1
'''''''''''                        End If
'''''''''''                        Constant = Constant + 1
'''''''''''                    Next k
'''''''''''                End If
                '*****************************************************For HardIP_D2D debug*****************************************************
                
                
                
                
                
                For Each site In TheExec.sites.Active
                    SourceTrimCode = vbNullString
                    For k = InDspWave_New.SampleSize - 1 To 0 Step -1
                        SourceTrimCode = SourceTrimCode & CStr(InDspWave_New.Element(k))
                    Next k
                    
                    If j = SrcStartBit + 1 Then
                        TheExec.Datalog.WriteComment ("InitialBit , Trim Code Bit " & SourceTrimCode)
                    Else
                        TheExec.Datalog.WriteComment ("Setup Bit " & (j) & ", Trim Code Bit " & SourceTrimCode)
                    End If
                Next site
                Dim tempVarArray As Variant
                Dim Digstart_name As String
                tempVarArray = TheHdw.DSSC.Pins(DigSrc_pin).Pattern(PattArray(1)).Source.Labels.list ''20210609 temp
                Digstart_name = tempVarArray(0)
                
                Call SetupDigSrcDspWave(PattArray(1), DigSrc_pin, Digstart_name, CLng(DigSrc_Sample_Size_Real_Temp(1)), InDspWave_New)   '''' not LSB to MSB ???
                Set OutDspWave = Nothing
                Call GeneralDigCapSetting(PattArray(1), DigCap_Pin, DigCap_Sample_Size, OutDspWave)
                
                Dim DSPwave_temp As New DSPWave
                Dim W As Long
                
                
                For Each site In TheExec.sites.Active
                
                DSPwave_temp.CreateConstant 0, 1, DspLong
                
                
                
                DSPwave_temp = InDspWave_New.ConvertStreamTo(tldspParallel, CLng(DigSrc_Sample_Size_Real_Temp(0)), 0, Bit0IsMsb)
                
                For W = 0 To DSPwave_temp.SampleSize - 1
                
                TheExec.Datalog.WriteComment CStr(DSPwave_temp.Element(W))
                
                Next
                Next site
                            
                Call TheHdw.Patterns(LCase(patset.value)).start
                TheHdw.Digital.Patgen.HaltWait
                
                For Each site In TheExec.sites.Active
                    CapValue.value = OutDspWave.Element(0)
                    ' TrimTarget is your expected transfer-point
                    b_HigherThanTarget = CapValue.Math.Subtract(TrimTarget).compare(EqualTo, 0)
                    
                    '///////////////// fail stop ////////////////////
                    If TrimPrcocessAll = False Then
                        If b_HigherThanTarget = True Then
                            If CaptureDSPFlag = False Then
                                CaptureDSPWave(0) = InDspWave_New.Copy                  ' For each site arry(0) is uncalculate value
                                CaptureDSPFlag = True
                            End If
                            b_StopTrimCodeProcess(site) = True
                        End If
                    Else
                    '/////////////////  do all  ////////////////////
                        If TrimPrcocessAll = True Then
                           CaptureDSPWave(0) = InDspWave_New.Copy
                        End If
                    End If
                    '///////////////////////////////////////////////
                    TheExec.Datalog.WriteComment ("Site " & site & " Output CapValue = " & CapValue.value)
                Next site
    
                TheExec.Datalog.WriteComment ("======================================================================================")

                If j <> SrcEndBit Then '20190530 Need to include initial bit so endbit don't need to do transform
                    If j - 1 = SrcEndBit Then
                      b_ControlNextBit = False
                      
                     End If
                      
                     
                    rundsp.SetupBinaryTrimCodeBit InDspWave_New, b_HigherThanTarget, j - 1, InitStateCode, TrimOffsetPoint, TrimOriginalSize, InDSPWave, b_ControlNextBit, AssignmentDSPWave, EachRegSize
                    rundsp.ReAssignmentDSPWave InDSPWave, RegSeparate, InDspWave_New, Src_dig, assignment, AssignmentDSPWave      ' ReAssignment DSPWave element to each register
                    
         
                    
                End If
                
                If TrimPrcocessAll = False Then                                     ' Immediately stop if all site capture done
                    If b_StopTrimCodeProcess.All(True) Then
                        Exit For
                    End If
                End If
                
            Next j
           
            
        Else                                                                        ' Linear Search
            If IncreaseFlag = True Then
                LimitCodeEnd = CInt((2 ^ (SrcStartBit + 1) - 1))
                LimitCodeStart = TrimStart
            Else                                                                    ' Decrease sweep
                LimitCodeEnd = TrimStart
                LimitCodeStart = 0
            End If
            
            For j = LimitCodeStart To LimitCodeEnd
            
                rundsp.ReAssignmentDSPWave InDSPWave, RegSeparate, InDspWave_New, Src_dig, assignment, AssignmentDSPWave   ' ReAssignment DSPWave element to each register
                
                For Each site In TheExec.sites.Active
                    SourceTrimCode = vbNullString
                    For k = InDspWave_New.SampleSize - 1 To 0 Step -1
                        SourceTrimCode = SourceTrimCode & CStr(InDspWave_New.Element(k))
                    Next k
                    If j = LimitCodeStart Then
                        b_StopTrimCodeProcess = False                               ' Inititial b_StopTrimCodeProcess flag
                        TheExec.Datalog.WriteComment ("InitialCode, Trim Code Bit " & SourceTrimCode)
                    Else
                        TheExec.Datalog.WriteComment ("LinearSweep " & TrimScanPoint(0) & "," & " Trim Code Bit " & SourceTrimCode)
                    End If
                Next site
                
                Call SetupDigSrcDspWave(PattArray(0), DigSrc_pin, "TrimCodeCap", CLng(DigSrc_Sample_Size_Real_Temp(1)), InDspWave_New)
                Call GeneralDigCapSetting(PattArray(0), DigCap_Pin, DigCap_Sample_Size, OutDspWave)
                
                Call TheHdw.Patterns(PattArray(0)).start
                TheHdw.Digital.Patgen.HaltWait
                
                For Each site In TheExec.sites.Active
                    CapValue.value = OutDspWave.Element(0)
                    ' TrimTarget is your expected transfer-point
                    b_HigherThanTarget = CapValue.Math.Subtract(TrimTarget).compare(EqualTo, 0)
                    If b_HigherThanTarget = True Then
                    
                        If CaptureDSPFlag = False Then
                            CaptureDSPWave(0) = InDspWave_New.Copy                  ' For each site arry(0) is uncalculate value
                            CaptureDSPFlag = True
                        End If
                        
                        If TrimPrcocessAll = False Then
                            b_StopTrimCodeProcess(site) = True
                        End If
                    End If
                    TheExec.Datalog.WriteComment ("Site " & site & " Output CapValue = " & OutDspWave(site).Element(0))
                Next site
    
                TheExec.Datalog.WriteComment ("======================================================================================")
                
                If j <> LimitCodeEnd Then
'                    Judgment addition "1" or "0" based on b_HigherThanTarget value
                    rundsp.SetupLinearTrimCodeBit IncreaseFlag, TrimScanPoint, b_HigherThanTarget, EachRegSize, InDSPWave, TrimPrcocessAll
                End If
                
                If TrimPrcocessAll = False Then                                     ' Immediately stop if all site capture done
                    If b_StopTrimCodeProcess.All(True) Then
                        Exit For
                    End If
                End If
            Next j
 
        End If
        
        
        Instance_Data.patset = patset
        Call HardIP_WriteFuncResult(, , glb_TestInstance)
        rundsp.ConvertToLongAndSerialToParrel InDSPWave, EachRegSize, ConvertedDataWf
        
        Call GetFlowTName
        
        If gl_UseStandardTestName_Flag = True Then
            Call Report_ALG_TName_From_Instance(OutputTname_format, "C", TrimStoreName, gl_Tname_Meas_FromFlow(TheExec.Flow.TestLimitIndex))
            TestNameInput = Merge_TName(OutputTname_format)
                  
        Else
            TestNameInput = glb_TestInstance & "DDR_Sweep"
        End If
        
        
        TheExec.Flow.TestLimit ConvertedDataWf.Element(0), Tname:=TestNameInput, PinName:="D2D_Trim_Dec", ForceResults:=tlForceFlow
        
        'TheExec.Flow.TestLimit ConvertedDataWf.Element(0), TName:=TheExec.DataManager.instanceName, PinName:="SEPVM_Trim_Dec", ForceResults:=tlForceFlow
            
           
        If TrimBase <> "" Then
            TrimBaseStr = Split(TrimBase, ";")
            OffsetDelta = CInt(TrimOffset) + 2 ^ CInt(StrSeparatebyColon(0))
            For j = 0 To UBound(TrimBaseStr)
                TrimBaseNum = Split(TrimBaseStr(j), ":")
                ReDim Preserve CaptureDSPWave(UBound(CaptureDSPWave) + 1)
                ReDim Preserve ProcessDoneDSPWave(UBound(ProcessDoneDSPWave) + 1)                   'This dspwave will save as to Dictionary
                
                CaptureDSPWave(UBound(CaptureDSPWave)).CreateConstant 0, 1, DspLong
                ProcessDoneDSPWave(UBound(ProcessDoneDSPWave)).CreateConstant 0, 1, DspLong
                                
                If CaptureDSPFlag.Any(True) Then
                    rundsp.CalculateDSPWaveforTrimCode CaptureDSPWave(0), EachRegSize, CaptureDSPWave(UBound(CaptureDSPWave)), _
                                                       CInt(OffsetDelta), CInt(TrimBaseNum(1)), ProcessDoneDSPWave(UBound(ProcessDoneDSPWave))
                                                       
                    
                    
                    
'''''                    TheExec.Flow.TestLimit CaptureDSPWave(UBound(CaptureDSPWave)).Element(0), TName:=TheExec.DataManager.instanceName, PinName:="SEPVM_Trim_Dec", ForceResults:=tlForceFlow
                                                       
                    AddStoredCaptureData TrimBaseNum(0), ProcessDoneDSPWave(UBound(ProcessDoneDSPWave))
                End If
            Next j
        End If
    Next i
    
    
    If TrimStoreName <> "" Then
       Call Checker_StoreDigCapAllToDictionary(TrimStoreName, InDSPWave)
    End If
    
        
    If Interpose_PrePat <> "" Then '''180109 update
        Call SetForceCondition("RESTOREPREPAT")
    End If

    Call SetForceCondition(Interpose_PostTest)
        
    Exit Function
    
ErrorHandler:
    TheExec.Datalog.WriteComment "error in TrimCodeBasicDig function"
    If AbortTest Then Exit Function Else Resume Next
    
    
End Function


Public Function PCIE_Eye_Diagram_0() As Long
Dim i, j As Integer
Dim site As Variant
Dim Eye_Diagram_Binary_Lane0(62) As New SiteVariant
Dim Eye_Diagram_Binary_lane1(62) As New SiteVariant
Dim Eye_Diagram_Binary_Lane2(62) As New SiteVariant
Dim Eye_Diagram_Binary_Lane3(62) As New SiteVariant
Dim Eye_Diagram_Binary_Lane4(62) As New SiteVariant
Dim Eye_Diagram_Binary_Lane5(62) As New SiteVariant

glb_TestInstance = vbNullString
glb_TestInstance = UCase(TheExec.DataManager.instancename)
        
TheExec.Datalog.WriteComment "<" & glb_TestInstance & ">"

For Each site In TheExec.sites
    For i = -31 To 31
        For j = 1 To 32
            Eye_Diagram_Binary_Lane0(i + 31)(site) = Eye_Diagram_Binary_Lane0(i + 31)(site) & mid(Eye_Diagram_Binary(i + 31)(site), 5 * j - 4, 1)
            'Eye_Diagram_Binary_Lane1(i + 31)(Site) = Eye_Diagram_Binary_Lane1(i + 31)(Site) & Mid(Eye_Diagram_Binary(i + 31)(Site), 5 * j - 3, 1)
           ' Eye_Diagram_Binary_Lane2(i + 31)(Site) = Eye_Diagram_Binary_Lane2(i + 31)(Site) & Mid(Eye_Diagram_Binary(i + 31)(Site), 5 * j - 2, 1)
            'Eye_Diagram_Binary_Lane3(i + 31)(Site) = Eye_Diagram_Binary_Lane3(i + 31)(Site) & Mid(Eye_Diagram_Binary(i + 31)(Site), 5 * j - 1, 1)
           ' Eye_Diagram_Binary_Lane4(i + 31)(Site) = Eye_Diagram_Binary_Lane4(i + 31)(Site) & Mid(Eye_Diagram_Binary(i + 31)(Site), 5 * j, 1)
        Next j
        Call TheExec.Datalog.WriteComment("Site(" & site & "), Lane 0, Binary string = " & Eye_Diagram_Binary_Lane0(i + 31)(site))
      '  Call TheExec.DataLog.WriteComment("Site(" & Site & "), Lane 0, Binary string = " & Eye_Diagram_Binary_Lane0(i + 31)(Site) & " Lane 1, Binary string = " & Eye_Diagram_Binary_lane1(i + 31)(Site) & " Lane 2, Binary string = " & Eye_Diagram_Binary_Lane2(i + 31)(Site) & " Lane 3, Binary string = " & Eye_Diagram_Binary_Lane3(i + 31)(Site) & " Lane 4, Binary string = " & Eye_Diagram_Binary_Lane4(i + 31)(Site) & " h0dac_off : " & i)
     '   Eye_Diagram_Binary(i + 31)(Site) = ""
        Next i
Next site

Dim TempDec() As New SiteDouble
Dim First_src_code As New SiteVariant
Dim End_src_code As New SiteVariant
Dim vertical_width As New SiteVariant
Dim Zero_counter As New SiteVariant
Dim horizontal_width As New SiteVariant
Dim Temp_counter As Integer
Dim timing_res_start As New SiteVariant
Dim timing_res_end As New SiteVariant
Dim timing_res_start_temp As Long
Dim timing_res_end_temp As Long
ReDim TempDec(62)

TheExec.Datalog.WriteComment "<" & glb_TestInstance & ">"

For Each site In TheExec.sites
    For i = -31 To 31
' horizontal_width = 0
'    Zero_counter = 0

        'Call TheExec.DataLog.WriteComment("Site(" & Site & ") Binary string = " & Eye_Diagram_Binary(i + 31)(Site) & " h0dac_off : " & i)
            'Bin2Dec
            Dim x As Integer
            Dim iLen As Integer
                iLen = Len(Eye_Diagram_Binary_Lane0(i + 31)(site)) - 1
                For x = 0 To iLen
                    TempDec(i + 31) = TempDec(i + 31) + mid(Eye_Diagram_Binary_Lane0(i + 31)(site), iLen - x + 1, 1) * 2 ^ x
                Next
                'process vertical width
                If TempDec(i + 31) <= 4294967295# Then
                        If First_src_code <> 0 Then
                            End_src_code = i
                        Else
                            First_src_code = i
                        End If
                        If First_src_code < 0 Then
                        
                        vertical_width = End_src_code - First_src_code + 1
                        Else
                        vertical_width = End_src_code - First_src_code
                        End If
                End If
                'process   the  Max Zero horizontal
                Temp_counter = 0
                Dim Temp_Counter_Act As Long
                Dim Total_Zero_Count As Long
                Total_Zero_Count = 0
                Temp_Counter_Act = 0
                timing_res_start_temp = 0
                timing_res_end_temp = 0
                For x = 0 To iLen
                        If mid(Eye_Diagram_Binary_Lane0(i + 31)(site), iLen - x + 1, 1) = 0 Then
                            Temp_counter = Temp_counter + 1
                            Total_Zero_Count = Total_Zero_Count + 1
                            If x = iLen Then
                                If Temp_counter > Temp_Counter_Act Then
                                    Temp_Counter_Act = Temp_counter
                                End If
                            End If
                        ElseIf mid(Eye_Diagram_Binary(i + 31)(site), iLen - x + 1, 1) = 1 Then
                            If Temp_counter > Temp_Counter_Act Then
                                Temp_Counter_Act = Temp_counter
                                timing_res_end_temp = 32 - (x - Temp_Counter_Act + 1)
                                timing_res_start_temp = timing_res_end_temp - Temp_Counter_Act + 1
                            End If
                            Temp_counter = 0
                        End If
                Next x
               If horizontal_width < Temp_Counter_Act Then
                           horizontal_width = Temp_Counter_Act
                           timing_res_end = timing_res_end_temp
                           timing_res_start = timing_res_start_temp
                End If
                If Zero_counter < Total_Zero_Count Then
                           Zero_counter = Total_Zero_Count
                End If
            '  Eye_Diagram_Binary(i + 31)(Site) = ""
        Next i
        
         '//////////////////////// for all 1 eye by csho/////////////////
      If horizontal_width = "" Then
         horizontal_width = 0
         End If
      If timing_res_end = "" Then
         timing_res_end = 0
          End If
      If timing_res_start = "" Then
         timing_res_start = 0
          End If
      If Zero_counter = "" Then
         Zero_counter = 0
          End If
    '/////////////////////////////////////////////////////////////////////////
        
Next site


    TheExec.Flow.TestLimit resultVal:=vertical_width, Tname:="vertical_width_lane0"
    TheExec.Flow.TestLimit resultVal:=First_src_code, Tname:="vertical_width_Start_lane0"
    TheExec.Flow.TestLimit resultVal:=End_src_code, Tname:="vertical_width_End_lane0"
    TheExec.Flow.TestLimit resultVal:=horizontal_width, Tname:="horizontal_width_lane0"
    TheExec.Flow.TestLimit resultVal:=timing_res_start, Tname:="horizontal_width_Start_lane0"
    TheExec.Flow.TestLimit resultVal:=timing_res_end, Tname:="horizontal_width_End_lane0"
    TheExec.Flow.TestLimit resultVal:=Zero_counter, Tname:="Max_Zero_lane0"
    

    Call PCIE_Eye_Diagram_1
    Call PCIE_Eye_Diagram_2
    Call PCIE_Eye_Diagram_3
    Call PCIE_Eye_Diagram_4
    'Call PCIE_Eye_Diagram_5
End Function

'errHandler:
'    TheExec.DataLog.WriteComment "error in Meas_FreqVoltCurr_Universal_func"
'    If AbortTest Then Exit Function Else Resume Next
'
'End Function

Public Function PCIE_Eye_Diagram_1() As Long
Dim i, j As Integer
Dim site As Variant
Dim Eye_Diagram_Binary_Lane0(62) As New SiteVariant
Dim Eye_Diagram_Binary_lane1(62) As New SiteVariant
Dim Eye_Diagram_Binary_Lane2(62) As New SiteVariant
Dim Eye_Diagram_Binary_Lane3(62) As New SiteVariant
Dim Eye_Diagram_Binary_Lane4(62) As New SiteVariant

glb_TestInstance = vbNullString
glb_TestInstance = UCase(TheExec.DataManager.instancename)

TheExec.Datalog.WriteComment "<" & glb_TestInstance & ">"

For Each site In TheExec.sites
    For i = -31 To 31
        For j = 1 To 32
            'Eye_Diagram_Binary_Lane0(i + 31)(Site) = Eye_Diagram_Binary_Lane0(i + 31)(Site) & Mid(Eye_Diagram_Binary(i + 31)(Site), 5 * j - 4, 1)
            Eye_Diagram_Binary_lane1(i + 31)(site) = Eye_Diagram_Binary_lane1(i + 31)(site) & mid(Eye_Diagram_Binary(i + 31)(site), 5 * j - 3, 1)
           ' Eye_Diagram_Binary_Lane2(i + 31)(Site) = Eye_Diagram_Binary_Lane2(i + 31)(Site) & Mid(Eye_Diagram_Binary(i + 31)(Site), 5 * j - 2, 1)
            'Eye_Diagram_Binary_Lane3(i + 31)(Site) = Eye_Diagram_Binary_Lane3(i + 31)(Site) & Mid(Eye_Diagram_Binary(i + 31)(Site), 5 * j - 1, 1)
           ' Eye_Diagram_Binary_Lane4(i + 31)(Site) = Eye_Diagram_Binary_Lane4(i + 31)(Site) & Mid(Eye_Diagram_Binary(i + 31)(Site), 5 * j, 1)
        Next j
        Call TheExec.Datalog.WriteComment("Site(" & site & "), Lane 1, Binary string = " & Eye_Diagram_Binary_lane1(i + 31)(site))
      '  Call TheExec.DataLog.WriteComment("Site(" & Site & "), Lane 0, Binary string = " & Eye_Diagram_Binary_lane1(i + 31)(Site) & " Lane 1, Binary string = " & Eye_Diagram_Binary_lane1(i + 31)(Site) & " Lane 2, Binary string = " & Eye_Diagram_Binary_Lane2(i + 31)(Site) & " Lane 3, Binary string = " & Eye_Diagram_Binary_Lane3(i + 31)(Site) & " Lane 4, Binary string = " & Eye_Diagram_Binary_Lane4(i + 31)(Site) & " h0dac_off : " & i)
     '   Eye_Diagram_Binary(i + 31)(Site) = ""
        Next i
Next site

Dim TempDec() As New SiteDouble
Dim First_src_code As New SiteVariant
Dim End_src_code As New SiteVariant
Dim vertical_width As New SiteVariant
Dim Zero_counter As New SiteVariant
Dim horizontal_width As New SiteVariant
Dim Temp_counter As Integer
Dim timing_res_start As New SiteVariant
Dim timing_res_end As New SiteVariant
Dim timing_res_start_temp As Long
Dim timing_res_end_temp As Long
ReDim TempDec(62)

TheExec.Datalog.WriteComment "<" & glb_TestInstance & ">"

For Each site In TheExec.sites
    For i = -31 To 31
' horizontal_width = 0
'    Zero_counter = 0

        'Call TheExec.DataLog.WriteComment("Site(" & Site & ") Binary string = " & Eye_Diagram_Binary(i + 31)(Site) & " h0dac_off : " & i)
            'Bin2Dec
            Dim x As Integer
            Dim iLen As Integer
                iLen = Len(Eye_Diagram_Binary_lane1(i + 31)(site)) - 1
                For x = 0 To iLen
                    TempDec(i + 31) = TempDec(i + 31) + mid(Eye_Diagram_Binary_lane1(i + 31)(site), iLen - x + 1, 1) * 2 ^ x
                Next
                'process vertical width
                If TempDec(i + 31) <= 4294967295# Then
                        If First_src_code <> 0 Then
                            End_src_code = i
                        Else
                            First_src_code = i
                        End If
                        If First_src_code < 0 Then
                        
                        vertical_width = End_src_code - First_src_code + 1
                        Else
                        vertical_width = End_src_code - First_src_code
                        End If
                End If
                'process   the  Max Zero horizontal
                Temp_counter = 0
                Dim Temp_Counter_Act As Long
                Dim Total_Zero_Count As Long
                Total_Zero_Count = 0
                Temp_Counter_Act = 0
                timing_res_start_temp = 0
                timing_res_end_temp = 0
                For x = 0 To iLen
                        If mid(Eye_Diagram_Binary_lane1(i + 31)(site), iLen - x + 1, 1) = 0 Then
                            Temp_counter = Temp_counter + 1
                            Total_Zero_Count = Total_Zero_Count + 1
                            If x = iLen Then
                                If Temp_counter > Temp_Counter_Act Then
                                    Temp_Counter_Act = Temp_counter
                                End If
                            End If
                        ElseIf mid(Eye_Diagram_Binary(i + 31)(site), iLen - x + 1, 1) = 1 Then
                            If Temp_counter > Temp_Counter_Act Then
                                Temp_Counter_Act = Temp_counter
                                timing_res_end_temp = 32 - (x - Temp_Counter_Act + 1)
                                timing_res_start_temp = timing_res_end_temp - Temp_Counter_Act + 1
                            End If
                            Temp_counter = 0
                        End If
                Next x
               If horizontal_width < Temp_Counter_Act Then
                           horizontal_width = Temp_Counter_Act
                           timing_res_end = timing_res_end_temp
                           timing_res_start = timing_res_start_temp
                End If
                If Zero_counter < Total_Zero_Count Then
                           Zero_counter = Total_Zero_Count
                End If
            ' Eye_Diagram_Binary(i + 31)(Site) = ""
        Next i
        
        
         '//////////////////////// for all 1 eye by csho/////////////////
      If horizontal_width = "" Then
         horizontal_width = 0
         End If
      If timing_res_end = "" Then
         timing_res_end = 0
          End If
      If timing_res_start = "" Then
         timing_res_start = 0
          End If
      If Zero_counter = "" Then
         Zero_counter = 0
          End If
    '/////////////////////////////////////////////////////////////////////////
Next site



    TheExec.Flow.TestLimit resultVal:=vertical_width, Tname:="vertical_width_lane1"
    TheExec.Flow.TestLimit resultVal:=First_src_code, Tname:="vertical_width_Start_lane1"
    TheExec.Flow.TestLimit resultVal:=End_src_code, Tname:="vertical_width_End_lane1"
    TheExec.Flow.TestLimit resultVal:=horizontal_width, Tname:="horizontal_width_lane1"
    TheExec.Flow.TestLimit resultVal:=timing_res_start, Tname:="horizontal_width_Start_lane1"
    TheExec.Flow.TestLimit resultVal:=timing_res_end, Tname:="horizontal_width_End_lane1"
    TheExec.Flow.TestLimit resultVal:=Zero_counter, Tname:="Max_Zero_lane1"
End Function

Public Function PCIE_Eye_Diagram_2() As Long
Dim i, j As Integer
Dim site As Variant
Dim Eye_Diagram_Binary_Lane0(62) As New SiteVariant
Dim Eye_Diagram_Binary_lane1(62) As New SiteVariant
Dim Eye_Diagram_Binary_Lane2(62) As New SiteVariant
Dim Eye_Diagram_Binary_Lane3(62) As New SiteVariant
Dim Eye_Diagram_Binary_Lane4(62) As New SiteVariant

glb_TestInstance = vbNullString
glb_TestInstance = UCase(TheExec.DataManager.instancename)

TheExec.Datalog.WriteComment "<" & glb_TestInstance & ">"

For Each site In TheExec.sites
    For i = -31 To 31
        For j = 1 To 32
          '  Eye_Diagram_Binary_Lane0(i + 31)(Site) = Eye_Diagram_Binary_Lane0(i + 31)(Site) & Mid(Eye_Diagram_Binary(i + 31)(Site), 5 * j - 4, 1)
           ' Eye_Diagram_Binary_Lane1(i + 31)(Site) = Eye_Diagram_Binary_Lane1(i + 31)(Site) & Mid(Eye_Diagram_Binary(i + 31)(Site), 5 * j - 3, 1)
         Eye_Diagram_Binary_Lane2(i + 31)(site) = Eye_Diagram_Binary_Lane2(i + 31)(site) & mid(Eye_Diagram_Binary(i + 31)(site), 5 * j - 2, 1)
          ' Eye_Diagram_Binary_Lane3(i + 31)(Site) = Eye_Diagram_Binary_Lane3(i + 31)(Site) & Mid(Eye_Diagram_Binary(i + 31)(Site), 5 * j - 1, 1)
          '  Eye_Diagram_Binary_Lane4(i + 31)(Site) = Eye_Diagram_Binary_Lane4(i + 31)(Site) & Mid(Eye_Diagram_Binary(i + 31)(Site), 5 * j, 1)
        Next j
        Call TheExec.Datalog.WriteComment("Site(" & site & "), Lane 2, Binary string = " & Eye_Diagram_Binary_Lane2(i + 31)(site))
'        Call TheExec.DataLog.WriteComment("Site(" & Site & "), Lane 0, Binary string = " & Eye_Diagram_Binary_Lane2(i + 31)(Site) & " Lane 1, Binary string = " & Eye_Diagram_Binary_lane1(i + 31)(Site) & " Lane 2, Binary string = " & Eye_Diagram_Binary_Lane2(i + 31)(Site) & " Lane 3, Binary string = " & Eye_Diagram_Binary_Lane3(i + 31)(Site) & " Lane 4, Binary string = " & Eye_Diagram_Binary_Lane4(i + 31)(Site) & " h0dac_off : " & i)
     '   Eye_Diagram_Binary(i + 31)(Site) = ""
        Next i
Next site

Dim TempDec() As New SiteDouble
Dim First_src_code As New SiteVariant
Dim End_src_code As New SiteVariant
Dim vertical_width As New SiteVariant
Dim Zero_counter As New SiteVariant
Dim horizontal_width As New SiteVariant
Dim Temp_counter As Integer
Dim timing_res_start As New SiteVariant
Dim timing_res_end As New SiteVariant
Dim timing_res_start_temp As Long
Dim timing_res_end_temp As Long
ReDim TempDec(62)

TheExec.Datalog.WriteComment "<" & glb_TestInstance & ">"

For Each site In TheExec.sites
    For i = -31 To 31
' horizontal_width = 0
'    Zero_counter = 0

        'Call TheExec.DataLog.WriteComment("Site(" & Site & ") Binary string = " & Eye_Diagram_Binary(i + 31)(Site) & " h0dac_off : " & i)
            'Bin2Dec
            Dim x As Integer
            Dim iLen As Integer
                iLen = Len(Eye_Diagram_Binary_Lane2(i + 31)(site)) - 1
                For x = 0 To iLen
                    TempDec(i + 31) = TempDec(i + 31) + mid(Eye_Diagram_Binary_Lane2(i + 31)(site), iLen - x + 1, 1) * 2 ^ x
                Next
                'process vertical width
                If TempDec(i + 31) <= 4294967295# Then
                        If First_src_code <> 0 Then
                            End_src_code = i
                        Else
                            First_src_code = i
                        End If
                        If First_src_code < 0 Then
                        
                        vertical_width = End_src_code - First_src_code + 1
                        Else
                        vertical_width = End_src_code - First_src_code
                        End If
                End If
                'process   the  Max Zero horizontal
                Temp_counter = 0
                Dim Temp_Counter_Act As Long
                Dim Total_Zero_Count As Long
                Total_Zero_Count = 0
                Temp_Counter_Act = 0
                timing_res_start_temp = 0
                timing_res_end_temp = 0
                For x = 0 To iLen
                        If mid(Eye_Diagram_Binary_Lane2(i + 31)(site), iLen - x + 1, 1) = 0 Then
                            Temp_counter = Temp_counter + 1
                            Total_Zero_Count = Total_Zero_Count + 1
                            If x = iLen Then
                                If Temp_counter > Temp_Counter_Act Then
                                    Temp_Counter_Act = Temp_counter
                                End If
                            End If
                        ElseIf mid(Eye_Diagram_Binary(i + 31)(site), iLen - x + 1, 1) = 1 Then
                            If Temp_counter > Temp_Counter_Act Then
                                Temp_Counter_Act = Temp_counter
                                timing_res_end_temp = 32 - (x - Temp_Counter_Act + 1)
                                timing_res_start_temp = timing_res_end_temp - Temp_Counter_Act + 1
                            End If
                            Temp_counter = 0
                        End If
                Next x
               If horizontal_width < Temp_Counter_Act Then
                           horizontal_width = Temp_Counter_Act
                           timing_res_end = timing_res_end_temp
                           timing_res_start = timing_res_start_temp
                End If
                If Zero_counter < Total_Zero_Count Then
                           Zero_counter = Total_Zero_Count
                End If
            '  Eye_Diagram_Binary(i + 31)(Site) = ""
        Next i
        
         '//////////////////////// for all 1 eye by csho/////////////////
      If horizontal_width = "" Then
         horizontal_width = 0
         End If
      If timing_res_end = "" Then
         timing_res_end = 0
          End If
      If timing_res_start = "" Then
         timing_res_start = 0
          End If
      If Zero_counter = "" Then
         Zero_counter = 0
          End If
    '/////////////////////////////////////////////////////////////////////////
        
Next site


    TheExec.Flow.TestLimit resultVal:=vertical_width, Tname:="vertical_width_lane2"
    TheExec.Flow.TestLimit resultVal:=First_src_code, Tname:="vertical_width_Start_lane2"
    TheExec.Flow.TestLimit resultVal:=End_src_code, Tname:="vertical_width_End_lane2"
    TheExec.Flow.TestLimit resultVal:=horizontal_width, Tname:="horizontal_width_lane2"
    TheExec.Flow.TestLimit resultVal:=timing_res_start, Tname:="horizontal_width_Start_lane2"
    TheExec.Flow.TestLimit resultVal:=timing_res_end, Tname:="horizontal_width_End_lane2"
    TheExec.Flow.TestLimit resultVal:=Zero_counter, Tname:="Max_Zero_lane2"
End Function

Public Function PCIE_Eye_Diagram_3() As Long
Dim i, j As Integer
Dim site As Variant
Dim Eye_Diagram_Binary_Lane0(62) As New SiteVariant
Dim Eye_Diagram_Binary_lane1(62) As New SiteVariant
Dim Eye_Diagram_Binary_Lane2(62) As New SiteVariant
Dim Eye_Diagram_Binary_Lane3(62) As New SiteVariant
Dim Eye_Diagram_Binary_Lane4(62) As New SiteVariant

glb_TestInstance = vbNullString
glb_TestInstance = UCase(TheExec.DataManager.instancename)
        
TheExec.Datalog.WriteComment "<" & glb_TestInstance & ">"

For Each site In TheExec.sites
    For i = -31 To 31
        For j = 1 To 32
'            Eye_Diagram_Binary_Lane0(i + 31)(Site) = Eye_Diagram_Binary_Lane0(i + 31)(Site) & Mid(Eye_Diagram_Binary(i + 31)(Site), 5 * j - 4, 1)
'            Eye_Diagram_Binary_Lane1(i + 31)(Site) = Eye_Diagram_Binary_Lane1(i + 31)(Site) & Mid(Eye_Diagram_Binary(i + 31)(Site), 5 * j - 3, 1)
'            Eye_Diagram_Binary_Lane2(i + 31)(Site) = Eye_Diagram_Binary_Lane2(i + 31)(Site) & Mid(Eye_Diagram_Binary(i + 31)(Site), 5 * j - 2, 1)
            Eye_Diagram_Binary_Lane3(i + 31)(site) = Eye_Diagram_Binary_Lane3(i + 31)(site) & mid(Eye_Diagram_Binary(i + 31)(site), 5 * j - 1, 1)
'            Eye_Diagram_Binary_Lane4(i + 31)(Site) = Eye_Diagram_Binary_Lane4(i + 31)(Site) & Mid(Eye_Diagram_Binary(i + 31)(Site), 5 * j, 1)
        Next j
        Call TheExec.Datalog.WriteComment("Site(" & site & "), Lane 3, Binary string = " & Eye_Diagram_Binary_Lane3(i + 31)(site))
'        Call TheExec.DataLog.WriteComment("Site(" & Site & "), Lane 0, Binary string = " & Eye_Diagram_Binary_Lane3(i + 31)(Site) & " Lane 1, Binary string = " & Eye_Diagram_Binary_lane1(i + 31)(Site) & " Lane 2, Binary string = " & Eye_Diagram_Binary_Lane2(i + 31)(Site) & " Lane 3, Binary string = " & Eye_Diagram_Binary_Lane3(i + 31)(Site) & " Lane 4, Binary string = " & Eye_Diagram_Binary_Lane4(i + 31)(Site) & " h0dac_off : " & i)
     '   Eye_Diagram_Binary(i + 31)(Site) = ""
        Next i
Next site

Dim TempDec() As New SiteDouble
Dim First_src_code As New SiteVariant
Dim End_src_code As New SiteVariant
Dim vertical_width As New SiteVariant
Dim Zero_counter As New SiteVariant
Dim horizontal_width As New SiteVariant
Dim Temp_counter As Integer
Dim timing_res_start As New SiteVariant
Dim timing_res_end As New SiteVariant
Dim timing_res_start_temp As Long
Dim timing_res_end_temp As Long
ReDim TempDec(62)

TheExec.Datalog.WriteComment "<" & glb_TestInstance & ">"

For Each site In TheExec.sites
    For i = -31 To 31
' horizontal_width = 0
'    Zero_counter = 0

        'Call TheExec.DataLog.WriteComment("Site(" & Site & ") Binary string = " & Eye_Diagram_Binary(i + 31)(Site) & " h0dac_off : " & i)
            'Bin2Dec
            Dim x As Integer
            Dim iLen As Integer
                iLen = Len(Eye_Diagram_Binary_Lane3(i + 31)(site)) - 1
                For x = 0 To iLen
                    TempDec(i + 31) = TempDec(i + 31) + mid(Eye_Diagram_Binary_Lane3(i + 31)(site), iLen - x + 1, 1) * 2 ^ x
                Next
                'process vertical width
                If TempDec(i + 31) <= 4294967295# Then
                        If First_src_code <> 0 Then
                            End_src_code = i
                        Else
                            First_src_code = i
                        End If
                        If First_src_code < 0 Then
                        
                        vertical_width = End_src_code - First_src_code + 1
                        Else
                        vertical_width = End_src_code - First_src_code
                        End If
                 
                        
                End If
                'process   the  Max Zero horizontal
                Temp_counter = 0
                Dim Temp_Counter_Act As Long
                Dim Total_Zero_Count As Long
                Total_Zero_Count = 0
                Temp_Counter_Act = 0
                timing_res_start_temp = 0
                timing_res_end_temp = 0
                For x = 0 To iLen
                        If mid(Eye_Diagram_Binary_Lane3(i + 31)(site), iLen - x + 1, 1) = 0 Then
                            Temp_counter = Temp_counter + 1
                            Total_Zero_Count = Total_Zero_Count + 1
                            If x = iLen Then
                                If Temp_counter > Temp_Counter_Act Then
                                    Temp_Counter_Act = Temp_counter
                                
                                
                                End If
                            End If
                        ElseIf mid(Eye_Diagram_Binary(i + 31)(site), iLen - x + 1, 1) = 1 Then
                            If Temp_counter > Temp_Counter_Act Then
                                Temp_Counter_Act = Temp_counter
                                timing_res_end_temp = 32 - (x - Temp_Counter_Act + 1)
                                timing_res_start_temp = timing_res_end_temp - Temp_Counter_Act + 1
                            End If
                            Temp_counter = 0
                        End If
                Next x
               If horizontal_width < Temp_Counter_Act Then
                           horizontal_width = Temp_Counter_Act
                           timing_res_end = timing_res_end_temp
                           timing_res_start = timing_res_start_temp
                End If
                If Zero_counter < Total_Zero_Count Then
                           Zero_counter = Total_Zero_Count
                End If
             ' Eye_Diagram_Binary(i + 31)(Site) = ""
        Next i
        
   '//////////////////////// for all 1 eye by csho/////////////////
      If horizontal_width = "" Then
         horizontal_width = 0
         End If
      If timing_res_end = "" Then
         timing_res_end = 0
          End If
      If timing_res_start = "" Then
         timing_res_start = 0
          End If
      If Zero_counter = "" Then
         Zero_counter = 0
          End If
    '/////////////////////////////////////////////////////////////////////////
Next site

 



    TheExec.Flow.TestLimit resultVal:=vertical_width, Tname:="vertical_width_lane3"
    TheExec.Flow.TestLimit resultVal:=First_src_code, Tname:="vertical_width_Start_lane3"
    TheExec.Flow.TestLimit resultVal:=End_src_code, Tname:="vertical_width_End_lane3"
    TheExec.Flow.TestLimit resultVal:=horizontal_width, Tname:="horizontal_width_lane3"
    TheExec.Flow.TestLimit resultVal:=timing_res_start, Tname:="horizontal_width_Start_lane3"
    TheExec.Flow.TestLimit resultVal:=timing_res_end, Tname:="horizontal_width_End_lane3"
    TheExec.Flow.TestLimit resultVal:=Zero_counter, Tname:="Max_Zero_lane3"
End Function

Public Function PCIE_Eye_Diagram_4() As Long
Dim i, j As Integer
Dim site As Variant
Dim Eye_Diagram_Binary_Lane0(62) As New SiteVariant
Dim Eye_Diagram_Binary_lane1(62) As New SiteVariant
Dim Eye_Diagram_Binary_Lane2(62) As New SiteVariant
Dim Eye_Diagram_Binary_Lane3(62) As New SiteVariant
Dim Eye_Diagram_Binary_Lane4(62) As New SiteVariant

glb_TestInstance = vbNullString
glb_TestInstance = UCase(TheExec.DataManager.instancename)
        
TheExec.Datalog.WriteComment "<" & glb_TestInstance & ">"

For Each site In TheExec.sites
    For i = -31 To 31
        For j = 1 To 32
'            Eye_Diagram_Binary_Lane0(i + 31)(Site) = Eye_Diagram_Binary_Lane0(i + 31)(Site) & Mid(Eye_Diagram_Binary(i + 31)(Site), 5 * j - 4, 1)
'            Eye_Diagram_Binary_Lane1(i + 31)(Site) = Eye_Diagram_Binary_Lane1(i + 31)(Site) & Mid(Eye_Diagram_Binary(i + 31)(Site), 5 * j - 3, 1)
'            Eye_Diagram_Binary_Lane2(i + 31)(Site) = Eye_Diagram_Binary_Lane2(i + 31)(Site) & Mid(Eye_Diagram_Binary(i + 31)(Site), 5 * j - 2, 1)
'            Eye_Diagram_Binary_Lane3(i + 31)(Site) = Eye_Diagram_Binary_Lane3(i + 31)(Site) & Mid(Eye_Diagram_Binary(i + 31)(Site), 5 * j - 1, 1)
            Eye_Diagram_Binary_Lane4(i + 31)(site) = Eye_Diagram_Binary_Lane4(i + 31)(site) & mid(Eye_Diagram_Binary(i + 31)(site), 5 * j - 1, 1)
        Next j
        Call TheExec.Datalog.WriteComment("Site(" & site & "), Lane 4, Binary string = " & Eye_Diagram_Binary_Lane4(i + 31)(site))
'       Call TheExec.DataLog.WriteComment("Site(" & Site & "), Lane 0, Binary string = " & Eye_Diagram_Binary_Lane4(i + 31)(Site) & " Lane 1, Binary string = " & Eye_Diagram_Binary_lane1(i + 31)(Site) & " Lane 2, Binary string = " & Eye_Diagram_Binary_Lane2(i + 31)(Site) & " Lane 3, Binary string = " & Eye_Diagram_Binary_Lane3(i + 31)(Site) & " Lane 4, Binary string = " & Eye_Diagram_Binary_Lane4(i + 31)(Site) & " h0dac_off : " & i)
     '   Eye_Diagram_Binary(i + 31)(Site) = ""
        Next i
Next site

Dim TempDec() As New SiteDouble
Dim First_src_code As New SiteVariant
Dim End_src_code As New SiteVariant
Dim vertical_width As New SiteVariant
Dim Zero_counter As New SiteVariant
Dim horizontal_width As New SiteVariant
Dim Temp_counter As Integer
Dim timing_res_start As New SiteVariant
Dim timing_res_end As New SiteVariant
Dim timing_res_start_temp As Long
Dim timing_res_end_temp As Long
ReDim TempDec(62)

TheExec.Datalog.WriteComment "<" & glb_TestInstance & ">"

For Each site In TheExec.sites
    For i = -31 To 31
' horizontal_width = 0
'    Zero_counter = 0

        'Call TheExec.DataLog.WriteComment("Site(" & Site & ") Binary string = " & Eye_Diagram_Binary(i + 31)(Site) & " h0dac_off : " & i)
            'Bin2Dec
            Dim x As Integer
            Dim iLen As Integer
                iLen = Len(Eye_Diagram_Binary_Lane4(i + 31)(site)) - 1
                For x = 0 To iLen
                    TempDec(i + 31) = TempDec(i + 31) + mid(Eye_Diagram_Binary_Lane4(i + 31)(site), iLen - x + 1, 1) * 2 ^ x
                Next
                'process vertical width
                If TempDec(i + 31) <= 4294967295# Then
                        If First_src_code <> 0 Then
                            End_src_code = i
                        Else
                            First_src_code = i
                        End If
                        If First_src_code < 0 Then
                        
                        vertical_width = End_src_code - First_src_code + 1
                        Else
                        vertical_width = End_src_code - First_src_code
                        End If
                End If
                'process   the  Max Zero horizontal
                Temp_counter = 0
                Dim Temp_Counter_Act As Long
                Dim Total_Zero_Count As Long
                Total_Zero_Count = 0
                Temp_Counter_Act = 0
                timing_res_start_temp = 0
                timing_res_end_temp = 0
                For x = 0 To iLen
                        If mid(Eye_Diagram_Binary_Lane4(i + 31)(site), iLen - x + 1, 1) = 0 Then
                            Temp_counter = Temp_counter + 1
                            Total_Zero_Count = Total_Zero_Count + 1
                            If x = iLen Then
                                If Temp_counter > Temp_Counter_Act Then
                                    Temp_Counter_Act = Temp_counter
                                End If
                            End If
                        ElseIf mid(Eye_Diagram_Binary(i + 31)(site), iLen - x + 1, 1) = 1 Then
                            If Temp_counter > Temp_Counter_Act Then
                                Temp_Counter_Act = Temp_counter
                                timing_res_end_temp = 32 - (x - Temp_Counter_Act + 1)
                                timing_res_start_temp = timing_res_end_temp - Temp_Counter_Act + 1
                            End If
                            Temp_counter = 0
                        End If
                Next x
               If horizontal_width < Temp_Counter_Act Then
                           horizontal_width = Temp_Counter_Act
                           timing_res_end = timing_res_end_temp
                           timing_res_start = timing_res_start_temp
                End If
                If Zero_counter < Total_Zero_Count Then
                           Zero_counter = Total_Zero_Count
                End If
              'Eye_Diagram_Binary(i + 31)(site) = ""
        Next i
        
    '//////////////////////// for all 1 eye by csho/////////////////
      If horizontal_width = "" Then
         horizontal_width = 0
         End If
      If timing_res_end = "" Then
         timing_res_end = 0
          End If
      If timing_res_start = "" Then
         timing_res_start = 0
          End If
      If Zero_counter = "" Then
         Zero_counter = 0
          End If
    '/////////////////////////////////////////////////////////////////////////
Next site


    TheExec.Flow.TestLimit resultVal:=vertical_width, Tname:="vertical_width_lane4"
    TheExec.Flow.TestLimit resultVal:=First_src_code, Tname:="vertical_width_Start_lane4"
    TheExec.Flow.TestLimit resultVal:=End_src_code, Tname:="vertical_width_End_lane4"
    TheExec.Flow.TestLimit resultVal:=horizontal_width, Tname:="horizontal_width_lane4"
    TheExec.Flow.TestLimit resultVal:=timing_res_start, Tname:="horizontal_width_Start_lane4"
    TheExec.Flow.TestLimit resultVal:=timing_res_end, Tname:="horizontal_width_End_lane4"
    TheExec.Flow.TestLimit resultVal:=Zero_counter, Tname:="Max_Zero_lane4"
End Function







Public Function PCIE_Eye_Diagram_5() As Long
Dim i, j As Integer
Dim site As Variant
Dim Eye_Diagram_Binary_Lane0(62) As New SiteVariant
Dim Eye_Diagram_Binary_lane1(62) As New SiteVariant
Dim Eye_Diagram_Binary_Lane2(62) As New SiteVariant
Dim Eye_Diagram_Binary_Lane3(62) As New SiteVariant
Dim Eye_Diagram_Binary_Lane4(62) As New SiteVariant
Dim Eye_Diagram_Binary_Lane5(62) As New SiteVariant
    
glb_TestInstance = vbNullString
glb_TestInstance = UCase(TheExec.DataManager.instancename)
        
TheExec.Datalog.WriteComment "<" & glb_TestInstance & ">"

For Each site In TheExec.sites
    For i = -31 To 31
        For j = 1 To 32
'            Eye_Diagram_Binary_Lane0(i + 31)(Site) = Eye_Diagram_Binary_Lane0(i + 31)(Site) & Mid(Eye_Diagram_Binary(i + 31)(Site), 5 * j - 4, 1)
'            Eye_Diagram_Binary_Lane1(i + 31)(Site) = Eye_Diagram_Binary_Lane1(i + 31)(Site) & Mid(Eye_Diagram_Binary(i + 31)(Site), 5 * j - 3, 1)
'            Eye_Diagram_Binary_Lane2(i + 31)(Site) = Eye_Diagram_Binary_Lane2(i + 31)(Site) & Mid(Eye_Diagram_Binary(i + 31)(Site), 5 * j - 2, 1)
'            Eye_Diagram_Binary_Lane3(i + 31)(Site) = Eye_Diagram_Binary_Lane3(i + 31)(Site) & Mid(Eye_Diagram_Binary(i + 31)(Site), 5 * j - 1, 1)
'            Eye_Diagram_Binary_Lane4(i + 31)(site) = Eye_Diagram_Binary_Lane4(i + 31)(site) & Mid(Eye_Diagram_Binary(i + 31)(site), 6 * j - 1, 1)
            Eye_Diagram_Binary_Lane5(i + 31)(site) = Eye_Diagram_Binary_Lane5(i + 31)(site) & mid(Eye_Diagram_Binary(i + 31)(site), 6 * j - 1, 1)
        Next j
        Call TheExec.Datalog.WriteComment("Site(" & site & "), Lane 5, Binary string = " & Eye_Diagram_Binary_Lane5(i + 31)(site))
'       Call TheExec.DataLog.WriteComment("Site(" & Site & "), Lane 0, Binary string = " & Eye_Diagram_Binary_Lane4(i + 31)(Site) & " Lane 1, Binary string = " & Eye_Diagram_Binary_lane1(i + 31)(Site) & " Lane 2, Binary string = " & Eye_Diagram_Binary_Lane2(i + 31)(Site) & " Lane 3, Binary string = " & Eye_Diagram_Binary_Lane3(i + 31)(Site) & " Lane 4, Binary string = " & Eye_Diagram_Binary_Lane4(i + 31)(Site) & " h0dac_off : " & i)
     '   Eye_Diagram_Binary(i + 31)(Site) = ""
        Next i
Next site

Dim TempDec() As New SiteDouble
Dim First_src_code As New SiteVariant
Dim End_src_code As New SiteVariant
Dim vertical_width As New SiteVariant
Dim Zero_counter As New SiteVariant
Dim horizontal_width As New SiteVariant
Dim Temp_counter As Integer
Dim timing_res_start As New SiteVariant
Dim timing_res_end As New SiteVariant
Dim timing_res_start_temp As Long
Dim timing_res_end_temp As Long
ReDim TempDec(62)

TheExec.Datalog.WriteComment "<" & glb_TestInstance & ">"

For Each site In TheExec.sites
    For i = -31 To 31
' horizontal_width = 0
'    Zero_counter = 0

        'Call TheExec.DataLog.WriteComment("Site(" & Site & ") Binary string = " & Eye_Diagram_Binary(i + 31)(Site) & " h0dac_off : " & i)
            'Bin2Dec
            Dim x As Integer
            Dim iLen As Integer
                iLen = Len(Eye_Diagram_Binary_Lane5(i + 31)(site)) - 1
                For x = 0 To iLen
                    TempDec(i + 31) = TempDec(i + 31) + mid(Eye_Diagram_Binary_Lane5(i + 31)(site), iLen - x + 1, 1) * 2 ^ x
                Next
                'process vertical width
                If TempDec(i + 31) < 4294967295# Then
                        If First_src_code <> 0 Then
                            End_src_code = i
                        Else
                            First_src_code = i
                        End If
                        If First_src_code < 0 Then
                        
                        vertical_width = End_src_code - First_src_code + 1
                        Else
                        vertical_width = End_src_code - First_src_code
                        End If
                End If
                'process   the  Max Zero horizontal
                Temp_counter = 0
                Dim Temp_Counter_Act As Long
                Dim Total_Zero_Count As Long
                Total_Zero_Count = 0
                Temp_Counter_Act = 0
                timing_res_start_temp = 0
                timing_res_end_temp = 0
                For x = 0 To iLen
                        If mid(Eye_Diagram_Binary_Lane5(i + 31)(site), iLen - x + 1, 1) = 0 Then
                            Temp_counter = Temp_counter + 1
                            Total_Zero_Count = Total_Zero_Count + 1
                            If x = iLen Then
                                If Temp_counter > Temp_Counter_Act Then
                                    Temp_Counter_Act = Temp_counter
                                End If
                            End If
                        ElseIf mid(Eye_Diagram_Binary(i + 31)(site), iLen - x + 1, 1) = 1 Then
                            If Temp_counter > Temp_Counter_Act Then
                                Temp_Counter_Act = Temp_counter
                                timing_res_end_temp = 32 - (x - Temp_Counter_Act + 1)
                                timing_res_start_temp = timing_res_end_temp - Temp_Counter_Act + 1
                            End If
                            Temp_counter = 0
                        End If
                Next x
               If horizontal_width < Temp_Counter_Act Then
                           horizontal_width = Temp_Counter_Act
                           timing_res_end = timing_res_end_temp
                           timing_res_start = timing_res_start_temp
                End If
                If Zero_counter < Total_Zero_Count Then
                           Zero_counter = Total_Zero_Count
                End If
              Eye_Diagram_Binary(i + 31)(site) = vbNullString
        Next i
Next site


    TheExec.Flow.TestLimit resultVal:=vertical_width, Tname:="vertical_width_lane5"
    TheExec.Flow.TestLimit resultVal:=First_src_code, Tname:="vertical_width_Start_lane5"
    TheExec.Flow.TestLimit resultVal:=End_src_code, Tname:="vertical_width_End_lane5"
    TheExec.Flow.TestLimit resultVal:=horizontal_width, Tname:="horizontal_width_lane5"
    TheExec.Flow.TestLimit resultVal:=timing_res_start, Tname:="horizontal_width_Start_lane5"
    TheExec.Flow.TestLimit resultVal:=timing_res_end, Tname:="horizontal_width_End_lane5"
    TheExec.Flow.TestLimit resultVal:=Zero_counter, Tname:="Max_Zero_lane5"
End Function



Public Function LPDPRX_Eye_Diagram() As Long
Dim i As Integer
Dim site As Variant
Dim TempDec() As New SiteDouble
Dim First_src_code As New SiteVariant
Dim End_src_code As New SiteVariant
Dim vertical_width As New SiteVariant
Dim Zero_counter As New SiteVariant
Dim horizontal_width As New SiteVariant
Dim Temp_counter As Integer
Dim timing_res_start As New SiteVariant
Dim timing_res_end As New SiteVariant
Dim timing_res_start_temp As Long
Dim timing_res_end_temp As Long
ReDim TempDec(62)

glb_TestInstance = vbNullString
glb_TestInstance = UCase(TheExec.DataManager.instancename)
        
TheExec.Datalog.WriteComment "<" & glb_TestInstance & ">"
For Each site In TheExec.sites
' horizontal_width = 0
'    Zero_counter = 0
    For i = -31 To 31
        Call TheExec.Datalog.WriteComment("Site(" & site & ") Binary string = " & Eye_Diagram_Binary(i + 31)(site) & " h0dac_off : " & i)
            'Bin2Dec
            Dim x As Integer
            Dim iLen As Integer
                iLen = Len(Eye_Diagram_Binary(i + 31)(site)) - 1
                For x = 0 To iLen
                    TempDec(i + 31) = TempDec(i + 31) + mid(Eye_Diagram_Binary(i + 31)(site), iLen - x + 1, 1) * 2 ^ x
                Next
                'process vertical width
                If TempDec(i + 31) <= 4294967295# Then
                        If First_src_code <> 0 Then
                            End_src_code = i
                        Else
                            First_src_code = i
                        End If
                        If First_src_code < 0 Then
                        
                        vertical_width = End_src_code - First_src_code + 1
                        Else
                        vertical_width = End_src_code - First_src_code
                        End If
                End If
                'process   the  Max Zero horizontal
                Temp_counter = 0
                Dim Temp_Counter_Act As Long
                Dim Total_Zero_Count As Long
                Total_Zero_Count = 0
                Temp_Counter_Act = 0
                timing_res_start_temp = 0
                timing_res_end_temp = 0
                For x = 0 To iLen
                        If mid(Eye_Diagram_Binary(i + 31)(site), iLen - x + 1, 1) = 0 Then
                            Temp_counter = Temp_counter + 1
                            Total_Zero_Count = Total_Zero_Count + 1
                            If x = iLen Then
                                If Temp_counter > Temp_Counter_Act Then
                                    Temp_Counter_Act = Temp_counter
                                End If
                            End If
                        ElseIf mid(Eye_Diagram_Binary(i + 31)(site), iLen - x + 1, 1) = 1 Then
                            If Temp_counter > Temp_Counter_Act Then
                                Temp_Counter_Act = Temp_counter
                                timing_res_end_temp = 32 - (x - Temp_Counter_Act + 1)
                                timing_res_start_temp = timing_res_end_temp - Temp_Counter_Act + 1
                            End If
                            Temp_counter = 0
                        End If
                Next x
               If horizontal_width < Temp_Counter_Act Then
                           horizontal_width = Temp_Counter_Act
                           timing_res_end = timing_res_end_temp
                           timing_res_start = timing_res_start_temp
                End If
                If Zero_counter < Total_Zero_Count Then
                           Zero_counter = Total_Zero_Count
                End If
              Eye_Diagram_Binary(i + 31)(site) = vbNullString
        Next i
         '//////////////////////// for all 1 eye by csho/////////////////
      If horizontal_width = "" Then
         horizontal_width = 0
         End If
      If timing_res_end = "" Then
         timing_res_end = 0
          End If
      If timing_res_start = "" Then
         timing_res_start = 0
          End If
      If Zero_counter = "" Then
         Zero_counter = 0
          End If
    '/////////////////////////////////////////////////////////////////////////
Next site

For Each site In TheExec.sites
    If vertical_width <> Empty Or horizontal_width <> Empty Then
    
        TheExec.Flow.TestLimit resultVal:=vertical_width, Tname:="vertical_width"
        TheExec.Flow.TestLimit resultVal:=First_src_code, Tname:="vertical_width_Start"
        TheExec.Flow.TestLimit resultVal:=End_src_code, Tname:="vertical_width_End"
        TheExec.Flow.TestLimit resultVal:=horizontal_width, Tname:="horizontal_width"
        TheExec.Flow.TestLimit resultVal:=timing_res_start, Tname:="horizontal_width_Start"
        TheExec.Flow.TestLimit resultVal:=timing_res_end, Tname:="horizontal_width_End"
        TheExec.Flow.TestLimit resultVal:=Zero_counter, Tname:="Max_Zero"
    Else
        TheExec.Datalog.WriteComment "*****NO SIGNAL OPENING*****"
    End If
   Next site


End Function
Public Function Enable_HIP_Datalog_Format_CZ()

        With TheExec.Datalog
                .Setup.DatalogSetup.DisableInstanceNameInPTR = False
                .Setup.DatalogSetup.DisablePinNameInPTR = False
                .Setup.DatalogSetup.DisableChannelNumberInPTR = True
                .Setup.DatalogSetup.PTR_InstanceNameIsTINameOnly = True
                ''' 20210709 for datalog format alignment
'''    .Setup.Shared.ascii.Columns.EnableCustomWidths = True
'''    .Setup.Shared.ascii.Columns.Parametric.TestName.Width = 200
'''    .Setup.Shared.ascii.Columns.Parametric.Measured.Width = 16
'''    .Setup.Shared.ascii.Columns.Functional.TestName.Width = 180
'''    .Setup.Shared.ascii.Columns.Functional.pattern.Width = 100
                .ApplySetup
        End With
        
        If False Then 'TheExec.Flow.EnableWord("Dummy") = True Then
                TheHdw.DCVI.Pins("All_DCVI").Alarm(tlDCVIAlarmOverRange) = tlAlarmOff
                TheHdw.DCVI.Pins("All_DCVI").Alarm(tlDCVIAlarmMode) = tlAlarmOff
                TheHdw.DCVI.Pins("All_DCVI").Alarm(tlDCVIAlarmCapture) = tlAlarmOff
        End If
End Function

Public Function Set_SEPVM_Ref_Level_Div()

    
    With TheHdw.DCVI.Pins("HSC_SEPVM_TEST_N")
        .Gate = False
        .mode = tlDCVIModeVoltage
        .Voltage = 0
        .VoltageRange.value = pc_Def_VFI_UVI80_VoltageRange ''--------need to change the clamp value
        .SetCurrentAndRange 0.002, 0.02
        .Connect tlDCVIConnectDefault
        .Gate = True
    End With
    
    With TheHdw.DCVI.Pins("HSC_SEPVM_TEST_P_SRC")
'        .Gate = False
'        .mode = tlDCVIModeVoltage
'        .Voltage = 6
'        .VoltageRange.Value = pc_Def_VFI_UVI80_VoltageRange
'        .SetCurrentAndRange 0.02, 0.2
'        .Connect tlDCVIConnectDefault
'        .Gate = True
        .mode = tlDCVIModeVoltage
        ''20170509 - Comment this
''            .Voltage = ForceV
        .Voltage = 6
        .VoltageRange.value = pc_Def_VFI_UVI80_VoltageRange
        ''20161018 - Swap current and current range sequence to avoid mode alarm
''            .Current = MI_TestCond_UVI80(i).CurrentRange
''            .CurrentRange.Value = MI_TestCond_UVI80(i).CurrentRange
        .SetCurrentAndRange 0.02, 0.02
        .Connect tlDCVIConnectDefault
        .Gate = True
    End With
    'thehdw.Wait 100 * ms
End Function
Public Function SEPVM_Ref_measurement(UVI_P_SRC As PinList, UVI_P As PinList, UVI_N As PinList, Vdiff_High As PinList, Vdiff_Low As PinList, Source_P_Voltage As Double, Source_N_Voltage As Double, _
TargetVoltage As Double, TheHdwWait As Double)
'UVI_P_SRC=HSC_SEPVM_TEST_P_SRC  @  UVI_P=HSC_SEPVM_TEST_P_Meas  @  UVI_N=HSC_SEPVM_TEST_N  @ Vdiff_High=HSC_SEPVM_P_High  @  Vdiff_Low=HSC_SEPVM_N_Low
'@ Source_P_Voltage=6  @  Source_N_Voltage=0  @  TargetVoltage=0.75
    On Error GoTo errHandler
    Dim Vdiff_Pin As String: Vdiff_Pin = UVI_P + "," + UVI_N
    Dim Disconnect_Pin As String: Disconnect_Pin = UVI_P_SRC + "," + UVI_N + "," + UVI_P
    Dim MeasureVolt_P As New PinListData
    Dim MeasureVolt_N As New PinListData
'===============================================================================================================
'for 3. Force UVI80-1 6V==================================================================================
'===============================================================================================================
      With TheHdw.DCVI.Pins(UVI_P_SRC)
          .Gate = False
          .mode = tlDCVIModeVoltage
          .Voltage = Source_P_Voltage
          .VoltageRange.value = pc_Def_VFI_UVI80_VoltageRange
          .SetCurrentAndRange 0.2, 0.2
          .Connect tlDCVIConnectHighForce
          .Gate = True
      End With
'===============================================================================================================
'for 3. Force UVI80-2 0V==================================================================================
'===============================================================================================================
      With TheHdw.DCVI.Pins(UVI_P)
          .Gate = False
          .mode = tlDCVIModeVoltage
'          .Voltage = 0 '-----------------------------------------------------------------------------------sense line only
          .VoltageRange.value = pc_Def_VFI_UVI80_VoltageRange
          .SetCurrentAndRange 0.002, 0.02
          .Connect tlDCVIConnectDefault
          .Gate = True
      End With
'===============================================================================================================
'for 3. Force UVI80-N 0V========================================================================================
'===============================================================================================================
      With TheHdw.DCVI.Pins(UVI_N)
          .Gate = False
          .mode = tlDCVIModeVoltage
          .Voltage = Source_N_Voltage
          .VoltageRange.value = pc_Def_VFI_UVI80_VoltageRange
          .SetCurrentAndRange 0.002, 0.02
          .Connect tlDCVIConnectDefault
          .Gate = True
      End With
'===============================================================================================================
'for 4. Use UVI80-2/UVI80-N VDIFF mode to measure the voltage across HSC_SEPVM_TEST_P/N=========================
'===============================================================================================================
'     TheHdw.DCVI.Pins(Vdiff_Pin).Connect '----------------------------------- Gate on all DCVIs
'
'     TheHdw.DCDiffMeter.Pins(Vdiff_High).LowSide.Pins = (Vdiff_Low) '----------------------- Specify the low side of the DCDiffMeter
'
'     With TheHdw.DCDiffMeter.Pins(Vdiff_High) ' ---------------------------------------------------- Set up the DCDiffMeter
'         .Connect tlDCDiffMeterConnectDefault
'         .VoltageRange = TargetVoltage
'     End With
'
'     TheHdw.Wait (TheHdwWait) ' --------------------------------------------------------------------------------- Program a wait time
'===============================================================================================================
'===============================================================================================================
'for 4.1. Use SINGLE END mode to measure the voltage across HSC_SEPVM_TEST_P/N=========================
'===============================================================================================================
    TheHdw.Wait (TheHdwWait) ' --------------------------------------------------------------------------------- Program a wait time
    MeasureVolt_P = TheHdw.DCVI.Pins(UVI_P).Meter.Read(tlStrobe, pc_Def_VFI_UVI80_ReadPoint)
    MeasureVolt_N = TheHdw.DCVI.Pins(UVI_N).Meter.Read(tlStrobe, pc_Def_VFI_UVI80_ReadPoint)

'===============================================================================================================


'===============================================================================================================
'Program the DCDiffMeter to make a measurement==================================================================
'===============================================================================================================
     Dim Results As New SiteDouble
'     Results = TheHdw.DCDiffMeter.Pins(Vdiff_High).Read(tlStrobe, 100, -1, tlDCDiffMeterReadingFormatAverage)
     Results = MeasureVolt_P.Math.Subtract(MeasureVolt_N)
     TheExec.Flow.TestLimit resultVal:=Results, PinName:="HSC_SEPVM_TEST_Vdiff", unit:=unitVolt, ForceResults:=tlForceFlow
     
     Dim ErrorValue As New PinListData: ErrorValue.AddPin ("ErrorValue")
'     ErrorValue = Results.Pins(Vdiff_High).Subtract(TargetVoltage)
     ErrorValue = Results.Subtract(TargetVoltage)
     TheExec.Flow.TestLimit resultVal:=ErrorValue.Pins, unit:=unitVolt, ForceResults:=tlForceFlow
'===============================================================================================================
'for Gate off and disconnect the DCVI===========================================================================
'===============================================================================================================
     With TheHdw.DCVI.Pins(Disconnect_Pin)
        .mode = tlDCVIModeVoltage
        .Gate = False
        .Disconnect
     End With
     
'     With TheHdw.DCDiffMeter.Pins(Vdiff_High) ' ---------------------------------------------------- Set up the DCDiffMeter
'         .Disconnect tlDCDiffMeterConnectDefault
'         .VoltageRange = TargetVoltage
'     End With
     
     Exit Function

errHandler:
    If isDebugMode Then TheExec.AddOutput "Error in SEPVM_Ref_measurement"
    TheExec.Datalog.WriteComment "Error in SEPVM_Ref_measurement"
    If AbortTest Then Exit Function Else Resume Next

End Function


Public Function SEPVM_Ref2_Calibration(UVI_P As PinList, UVI_N As PinList, Vdiff_High As PinList, Vdiff_Low As PinList, Source_P_Voltage As Double, Source_N_Voltage As Double, _
TargetVoltage As Double, TheHdwWait As Double)
'UVI_P=HSC_SEPVM_TEST_P_Meas  @  UVI_N=HSC_SEPVM_TEST_N  @ Vdiff_High=HSC_SEPVM_P_High  @  Vdiff_Low=HSC_SEPVM_N_Low
'@ Source_P_Voltage=0.75  @ Source_N_Voltage=0  @  TargetVoltage=0.75  @ ErrorValueTarget=0.00022
    Dim ErrorValueTarget As Long
    
    ErrorValueTarget = 0.00022
    Dim SEPDSP As New DSPWave
    
    
    On Error GoTo errHandler
    Dim Vdiff_Pin As String: Vdiff_Pin = UVI_P + "," + UVI_N
        Dim Results As New SiteDouble
        Dim ErrorValue As New PinListData: ErrorValue.AddPin ("ErrorValue")
        Dim site As Variant
        Dim ForceCalibration As New SiteDouble: ForceCalibration = Source_P_Voltage ''--------let original value =0.75mV
        Dim MeasureVolt_P As New PinListData
        Dim MeasureVolt_N As New PinListData
    On Error GoTo errHandler
'===============================================================================================================
'for step3 Force UVI80-2 0.75V==================================================================================
'===============================================================================================================
    With TheHdw.DCVI.Pins(UVI_P)
        .Gate = False
        .mode = tlDCVIModeVoltage
        .Voltage = Source_P_Voltage
        .VoltageRange.value = pc_Def_VFI_UVI80_VoltageRange
        .SetCurrentAndRange 0.002, 0.02
        .Connect tlDCVIConnectDefault
        .Gate = True
    End With
    'thehdw.Wait 100 * ms
    
'===============================================================================================================
'for step3 Force UVI80-N 0V=====================================================================================
'===============================================================================================================
    With TheHdw.DCVI.Pins(UVI_N)
        .Gate = False
        .mode = tlDCVIModeVoltage
        .Voltage = Source_N_Voltage
        .VoltageRange.value = pc_Def_VFI_UVI80_VoltageRange ''--------need to change the clamp value
        .SetCurrentAndRange 0.002, 0.02
        .Connect tlDCVIConnectDefault
        .Gate = True
    End With
    'thehdw.Wait 100 * ms
'===============================================================================================================
'for step4. Use UVI80-2/UVI80-N VDIFF mode to measure the voltage across HSC_SEPVM_TEST_P/N=====================
'===============================================================================================================
                       
'    TheHdw.DCVI.Pins(Vdiff_Pin).Gate = False
'    TheHdw.DCVI.Pins(Vdiff_Pin).Connect ' ---------------Gate on all DCVIs
'    TheHdw.DCDiffMeter.Pins(Vdiff_High).LowSide.Pins = (Vdiff_Low) ' ---Specify the low side of the DCDiffMeter
'    TheHdw.DCVI.Pins(Vdiff_Pin).Gate = True
'
'
'    With TheHdw.DCDiffMeter.Pins(Vdiff_High) ' ---------------------------------Set up the DCDiffMeter
'        .Connect tlDCDiffMeterConnectDefault
'        .VoltageRange = Source_P_Voltage
'    End With
'    TheHdw.Wait 10 * ms
    TheHdw.Wait (TheHdwWait) '--------------------------------------------------------------- Program a wait time
    
    MeasureVolt_P = TheHdw.DCVI.Pins(UVI_P).Meter.Read(tlStrobe, pc_Def_VFI_UVI80_ReadPoint)
    MeasureVolt_N = TheHdw.DCVI.Pins(UVI_N).Meter.Read(tlStrobe, pc_Def_VFI_UVI80_ReadPoint)
    
'===============================================================================================================
'Program the DCDiffMeter to make a measurement==================================================================
'===============================================================================================================
'    Results = TheHdw.DCDiffMeter.Pins(Vdiff_High).Read(tlStrobe, 100, -1, tlDCDiffMeterReadingFormatAverage)
    Results = MeasureVolt_P.Math.Subtract(MeasureVolt_N)
    ErrorValue = Results.Subtract(Source_P_Voltage) '-------------------------------------------- the value of differ from 0.75mV
    
      For Each site In TheExec.sites
      
         Dim LoopCount As New SiteDouble: LoopCount = 0
         Dim step As Integer: step = 1
         
          Do While Abs(ErrorValue) > ErrorValueTarget And LoopCount < 11 And ForceCalibration < 7 '-------------------------------------------- the value of error target 220E-06
                            
                LoopCount = LoopCount + step
            
                ForceCalibration = ForceCalibration - ErrorValue '----------------------------- for loop error value add the last result
                With TheHdw.DCVI.Pins(UVI_P)
                    .Voltage = ForceCalibration '---------------------------------------------- for force last result add error value
                End With
                With TheHdw.DCVI.Pins(UVI_N)
                    .Voltage = 0
                End With
                  TheHdw.Wait (TheHdwWait) ' -------------------------------------------------------Program a wait time
                MeasureVolt_P = TheHdw.DCVI.Pins(UVI_P).Meter.Read(tlStrobe, pc_Def_VFI_UVI80_ReadPoint)
                MeasureVolt_N = TheHdw.DCVI.Pins(UVI_N).Meter.Read(tlStrobe, pc_Def_VFI_UVI80_ReadPoint)
                
'                Results = TheHdw.DCDiffMeter.Pins(Vdiff_High).Read(tlStrobe, 100, -1, tlDCDiffMeterReadingFormatAverage)
                Results = MeasureVolt_P.Math.Subtract(MeasureVolt_N)
                ErrorValue = Results.Subtract(TargetVoltage)
                
                If TheExec.TesterMode = testModeOffline Then '---------------------------------for avoid offline in to infinite loop
                  ErrorValue = 0.0001
                End If
                
                TheExec.Datalog.WriteComment "site " & site & " LoopCount " & LoopCount & "  ForceVoltage" & ForceCalibration & "  Results" & Results & "  ErrorValue" & ErrorValue & " "
                    
          Loop
          
      Next site
'Alarm *ForceVoltage out of range* or *Loop out of range*
'===============================================================================================================
'Print measured and error value=================================================================================
'===============================================================================================================
    
    TheExec.Flow.TestLimit resultVal:=Results, PinName:="HSC_SEPVM_TEST_Vdiff", unit:=unitVolt, ForceResults:=tlForceFlow
    TheExec.Flow.TestLimit resultVal:=ErrorValue.Pins, unit:=unitVolt, ForceResults:=tlForceFlow
    For Each site In TheExec.sites
    If LoopCount > 9 Or ForceCalibration > 2 Then
    
        TheExec.Datalog.WriteComment "site " & site & " LoopCount " & LoopCount & "  ForceVoltage" & ForceCalibration & " Alarm *ForceVoltage out of range* or *Loop out of range*"
    Else
    TheExec.Datalog.WriteComment "site " & site & " LoopCount " & LoopCount & "  ForceVoltage" & ForceCalibration & "  "
    End If
    
    Next site
'===============================================================================================================
'Gate off and disconnect the DCVI===============================================================================
'===============================================================================================================
                    
    With TheHdw.DCVI.Pins(Vdiff_Pin)
       .mode = tlDCVIModeVoltage
       .Gate = False
       .Disconnect
    End With


    Exit Function

errHandler:
    If isDebugMode Then TheExec.AddOutput "Error in SEPVM_Ref2_measurement"
    TheExec.Datalog.WriteComment "Error in SEPVM_Ref2_measurement"
    If AbortTest Then Exit Function Else Resume Next

End Function

Public Function ReSet_SEPVM_Ref_Level_Div()

    With TheHdw.DCVI.Pins("HSC_SEPVM_TEST_P_SRC,HSC_SEPVM_TEST_N,HSC_SEPVM_TEST_P_MEA")
        .Gate(tlDCVIGateHiZ) = False
        TheHdw.Wait 0.001
        .VoltageRange.value = pc_Def_VFI_UVI80_VoltageRange ''--------need to change the clamp value
        .SetCurrentAndRange 0.02, 0.02
        .Disconnect
        .mode = tlDCVIModeCurrent
    End With
    TheHdw.Wait 10 * ms
    
'===============================================================================================================
'for step3 Force UVI80-N 0V=====================================================================================
'===============================================================================================================
'    With thehdw.DCVI.Pins("HSC_SEPVM_TEST_N")
'        .Gate = False
''        .mode = tlDCVIModeVoltage
'        .Voltage = 0
'        .VoltageRange.Value = pc_Def_VFI_UVI80_VoltageRange ''--------need to change the clamp value
''        .SetCurrentAndRange 0.002, 0.02
'        .Disconnect tlDCVIConnectDefault
''        .Gate = True
'    End With
'
'    With thehdw.DCVI.Pins("HSC_SEPVM_TEST_P_MEAS")
'        .Gate = False
''        .mode = tlDCVIModeVoltage
'        .Voltage = 0
'        .VoltageRange.Value = pc_Def_VFI_UVI80_VoltageRange ''--------need to change the clamp value
''        .SetCurrentAndRange 0.002, 0.02
'        .Disconnect tlDCVIConnectDefault
''        .Gate = True
'    End With
    
End Function

Public Function HIP_eFuse_Read_TMPS_Coeff(FuseType As String, m_catename As String, Dict_Store_Code_Name As String, dspwavesize As Long, Optional Efuse_Read_Dec_Flag As Boolean = False, Optional Dict_Store_Dec_Name As String = vbNullString) As Long

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

    On Error GoTo errHandler

    Read_Code.CreateConstant 0, dspwavesize

    If Efuse_Read_Dec_Flag = True Then
        Read_Value.CreateConstant 0, 1
    End If
    
    '20210406 Modify for New Efuse
    Efuse_Value = GetEfuseHipValue(FuseType, m_catename)

    For Each site In TheExec.sites

        'Efuse_Value(site) = auto_eFuse_GetReadDecimal(FuseType, m_catename, True)
'''''        Efuse_Value(Site) = CLng(Site) + 8

        If Efuse_Read_Dec_Flag = True Then
            Read_Value.Element(0) = Efuse_Value(site)
        End If

        TempVal = Efuse_Value(site)
        For i = 0 To dspwavesize - 1
            Read_Code.Element(i) = TempVal Mod 2
            TempVal = TempVal \ 2
        Next i

        If Efuse_Value(site) = 0 Then
        'If Read out value = 0 then bin out
            Efuse_Value_Chk(site) = 0
        Else
            Efuse_Value_Chk(site) = 1
        End If

    Next site

    'TheExec.Flow.TestLimit resultVal:=Efuse_Value_Chk, lowVal:=1, hiVal:=1, Tname:="NonZero_Val_Chk", ForceResults:=tlForceFlow

    Call AddStoredCaptureData(Dict_Store_Code_Name, Read_Code)

    If Efuse_Read_Dec_Flag = True Then
        Call AddStoredCaptureData(Dict_Store_Dec_Name, Read_Value)
    End If

    Exit Function

errHandler:
    TheExec.Datalog.WriteComment "Error in HIP_eFuse_Read"
    If AbortTest Then Exit Function Else Resume Next

End Function

Public Function ADC_Trim(patset As Pattern, CPUA_Flag_In_Pat As Boolean, _
    MeasureV_PinS As String, _
    DigSrc_pin As PinList, DigSrc_DataWidth As Long, DigSrc_Sample_Size As Long, _
    DigSrc_Equation As String, digsrc_assignment As String, _
    Optional TargetValue_Volt As Double, Optional CUS_Str As String, Optional Validating_ As Boolean) As Long

'' Step 1 : trim code is 32 bit, show out measured volt and trimed code, target volt is 1.1v
'' Step 2 : start from 0x8 and add algorithm to decide +/- direction
'' while decimal < 2 ^ DigSrc_Sample_Size
'' convert decimal to binary reverse
'' input the binary reverse data to digSrc_assignment


    If Validating_ Then
        Call PrLoadPattern(patset.value)
        Exit Function    ' Exit after validation
    End If

    Call HardIP_InitialSetupForPatgen
    Dim x As Long
    Dim InDSPWave As New DSPWave
    Dim SrcOut As String
    Dim site As Variant
    Dim Pat As String
    Dim i As Integer
    Dim ShowDec As String
    Dim ShowOut As String
    Dim TrimBits As String
    Dim b_TestDone As Boolean
    Dim SourceNum As Integer
    Dim k As Integer
    Dim MeasureVoltage As New PinListData
    Dim data As New SiteLong
    'Dim data As Integer
    Dim PassFlag_ADC As New SiteBoolean
    Dim opbank As eFuseBdfBank '20210406 Add for new Efuse
    Dim field As eFuseBdfField  '20210406 Add for new Efuse

    gl_TName_Pat = patset.value

    On Error GoTo errHandler
      For Each site In TheExec.sites.Active
            Src_DSPWave.CreateConstant 0, DigSrc_Sample_Size
    Next site

    b_TestDone = False
    SourceNum = 0

    If DigSrc_Sample_Size = 0 Then
        TheExec.Datalog.WriteComment ("Error!! - Please check input argument DigSrc_Sample_Size")
        Exit Function
    End If

    TheHdw.Digital.Patgen.Halt
    TheHdw.Digital.ApplyLevelsTiming True, True, True, tlPowered
    TheHdw.Patterns(patset).Load

    Dim PattArray() As String
    Dim PatCount As Long

    Call PATT_GetPatListFromPatternSet(patset.value, PattArray, PatCount)
    Call TheHdw.Digital.Patgen.Continue(0, cpuA + cpuB + cpuC + cpuD)
    TheHdw.Digital.Patgen.HaltMode = tlHaltOnOpcode

    Do While b_TestDone = False
        For Each site In TheExec.sites.Active

            ''  theexec.Datalog.WriteComment ("======== Start Dig Src setup =======")
            If SourceNum = 0 Then
                Call Create_DigSrc_Data(DigSrc_pin, DigSrc_DataWidth, DigSrc_Sample_Size, DigSrc_Equation, digsrc_assignment, InDSPWave, site)
            End If

            If SourceNum > 0 And SourceNum < 9 Then

                If MeasureVoltage.Pins(0).value(site) > TargetValue_Volt Then
                    For x = 0 To 3
                        InDSPWave(site).Element(4 * (8 - SourceNum) + x) = 0
                    Next x
                    'InDSPWave(site).Element(4 * (8 - SourceNum) + 1) = 0
                    'InDSPWave(site).Element(4 * (8 - SourceNum) + 2) = 0
                    'InDSPWave(site).Element(4 * (8 - SourceNum) + 3) = 0
                End If

                If SourceNum < 8 Then
                    For x = 0 To 3
                        InDSPWave(site).Element(4 * (7 - SourceNum) + x) = 1
                    Next x
                    'InDSPWave(site).Element(4 * (7 - SourceNum) + 1) = 1
                    'InDSPWave(site).Element(4 * (7 - SourceNum) + 2) = 1
                    'InDSPWave(site).Element(4 * (7 - SourceNum) + 3) = 1
                End If

                If gl_Disable_HIP_debug_log = False Then TheExec.Datalog.WriteComment "Site:" & site & ", Measured Voltage: " & MeasureVoltage.Pins(0).value(site)
            End If

        Next site

        Call SetupDigSrcDspWave(PattArray(0), DigSrc_pin, "Meas_Src", DigSrc_Sample_Size, InDSPWave)

        If SourceNum = 8 Then
            If gl_Disable_HIP_debug_log = False Then TheExec.Datalog.WriteComment "==========SOURCING TRIMMED DATA BITS==============="
        End If

        For Each site In TheExec.sites.Active
            If SourceNum > 0 Then
                SrcOut = vbNullString
                For i = 0 To DigSrc_Sample_Size - 1
                    SrcOut = SrcOut & InDSPWave(site).Element(i)
                    If i Mod DigSrc_DataWidth = DigSrc_DataWidth - 1 Then
                        SrcOut = SrcOut & ", "
                    ElseIf i Mod 4 = 3 Then
                        SrcOut = SrcOut & " "
                    End If
                Next i
                If gl_Disable_HIP_debug_log = False Then TheExec.Datalog.WriteComment "Site:" & site & ", Source Data=" & SrcOut
            End If
        Next site

        Call TheHdw.Patterns(PattArray(0)).start

        If (CPUA_Flag_In_Pat) Then
            Call TheHdw.Digital.Patgen.FlagWait(cpuA, 0)    '' Meas during CPUA loop
        Else
            Call TheHdw.Digital.Patgen.HaltWait '' Pattern without CPUA loop
        End If

''        Call HardIP_SetupAndMeasureVolt_UVI80(MeasureV_PinS, MeasureVoltage, True)
        ''20170621
        Dim MV_TestCond_UVI80(0) As DUTConditions
        MV_TestCond_UVI80(0).PinName = MeasureV_PinS
        Call HardIP_SetupAndMeasureVolt_UVI80_old(MV_TestCond_UVI80, MeasureVoltage)

        SourceNum = SourceNum + 1

        If (CPUA_Flag_In_Pat) Then
            Call TheHdw.Digital.Patgen.Continue(0, cpuA)    '' Jump out CPUA loop
        End If

        TheHdw.Digital.Patgen.HaltWait ' haltwait at patten end

        PatCount = PatCount + 1

        If SourceNum = 9 Then
            b_TestDone = True
        End If
    Loop

    TheExec.Flow.TestLimit resultVal:=MeasureVoltage, unit:=unitVolt, Tname:="Volt_meas_ADC_Trim", ForceResults:=tlForceNone 'eng_forceflow_transfer

    Call HardIP_WriteFuncResult

    For Each site In TheExec.sites.Active
        For i = 0 To DigSrc_Sample_Size - 1
            Src_DSPWave(site).Element(i) = InDSPWave(site).Element(i)
        Next i
    Next site
    
    '20210406 Add for new Efuse
    Dim Efuse_Value As New SiteDouble
    If CurrentJobName_U Like "*FT*" Then
        Efuse_Value = GetEfuseHipValue("UDR", "ADC_vTRIM")
    End If

    For Each site In TheExec.sites.Active
        SrcOut = vbNullString
        For i = 0 To DigSrc_Sample_Size - 1
            SrcOut = SrcOut & Src_DSPWave(site).Element(i)
        Next i
        If gl_Disable_HIP_debug_log = False Then TheExec.Datalog.WriteComment "Site:" & site & ", Stored Data=" & SrcOut

        '''''''''''''''''''''eFUSE
        data = 0
        For i = 0 To DigSrc_DataWidth / 4 - 1
            data = data + InDSPWave(site).Element(4 * i) * (2 ^ i)
        Next i
        If CurrentJobName_U Like "*FT*" Then
            If gl_Disable_HIP_debug_log = False Then TheExec.Datalog.WriteComment vbNullString
            data = CLng(Efuse_Value(site))
            'data = auto_eFuse_GetReadDecimal("UDR", "ADC_vTRIM", True)
            If gl_Disable_HIP_debug_log = False Then TheExec.Datalog.WriteComment vbNullString
            For i = 0 To DigSrc_Sample_Size / 4 - 1
                For x = 0 To 3
                    Src_DSPWave(site).Element(4 * i + x) = data Mod 2
                Next x
                'Src_DSPWave(site).Element(4 * i + 1) = Data Mod 2
                'Src_DSPWave(site).Element(4 * i + 2) = Data Mod 2
                'Src_DSPWave(site).Element(4 * i + 3) = Data Mod 2
                data = data \ 2
            Next i
        Else
            If TheHdw.Digital.Patgen.PatternBurstPassed(site) = False Then 'Pattern Fail
                PassFlag_ADC(site) = False
            Else
                PassFlag_ADC(site) = True
            End If
'            If CUS_Str = "ADC_VTRIM" Then
'                Call auto_eFuse_SetPatTestPass_Flag("UDR", "ADC_vTRIM", PassFlag_ADC(site))
'                Call auto_eFuse_SetWriteDecimal("UDR", "ADC_vTRIM", data)
'            End If
        End If
    Next site
    
    If (Not CurrentJobName_U Like "*FT*") And CUS_Str = "ADC_VTRIM" Then
        Set opbank = GetBdfBank("UDR")
        Set field = opbank.Fields("ADC_vTRIM")
        opbank.SetEfuse field.name, data, PassFlag_ADC, , , , True
    End If
    
    Exit Function

errHandler:
    TheExec.Datalog.WriteComment "error in ADC_Trim function"
    If AbortTest Then Exit Function Else Resume Next

End Function
    



'No Used in Sicily, Oscar, 20200423
Public Function MTR_Sense_Calibration_Coeff_Verification(SensorArray As String, Temperature As String, FusedCoeffDicName_1 As String, FusedCoeffDicName_2 As String, _
Optional MTRMatricesSheet As String, Optional SensorCalculate As String, Optional SweepVArryDic As String, Optional Validating_ As Boolean) As Long
'MTR Record
If Validating_ Then
    
    Exit Function    ' Exit after validation
End If
    
    If TheHdw.DSP.ExecutionMode = tlDSPModeHostDebug Then
        TheHdw.DSP.ExecutionMode = tlDSPModeAutomatic
    End If
    
    HIP_Init_Datalog_Setup


    Dim site As Variant
    Dim PowerPinsGroup() As String
    Dim PowerPinsCounter As Long
    Dim PowerPinLevelsGroup() As String
    Dim PowerPinLevelsCounter As Long
    Dim SensorsGroup() As String
    Dim SensorsCounter As Long
    Dim ScanSeson() As String
    Dim ScanSesonCounter As Long
    Dim MatrixStr As String
    Dim VoltageScaner As New SiteLong
    Dim VoltageScanerStr As String
    Dim FullGroup() As String
    Dim FullLevels() As String
    
    Dim SetInformation() As New DSPWave
    Dim Aininformation() As New DSPWave
    Dim Aixinformation() As New DSPWave
    Dim PiUInformation() As New DSPWave
    
    Dim Fused_ROT_Decimal_Vector As New DSPWave
    Dim Fused_ROV_Decimal_Vector As New DSPWave
    Fused_ROT_Decimal_Vector.CreateConstant 0, 4, DspDouble
    Fused_ROV_Decimal_Vector.CreateConstant 0, 3, DspDouble
            
    Dim Output_ROT_Freq_Vector As New DSPWave
    Dim Output_ROV_Freq_Vector As New DSPWave
    Output_ROT_Freq_Vector.CreateConstant 0, 8, DspDouble
    Output_ROV_Freq_Vector.CreateConstant 0, 8, DspDouble
        
    Dim Difference_ROT_Freq_Vector As New DSPWave
    Dim Difference_ROV_Freq_Vector As New DSPWave
    Difference_ROT_Freq_Vector.CreateConstant 0, 8, DspDouble
    Difference_ROV_Freq_Vector.CreateConstant 0, 8, DspDouble
    
    
    FullGroup = Split(SweepVArryDic, ";")
    PowerPinsGroup = Split(SensorCalculate, ";")
    For PowerPinsCounter = 0 To UBound(PowerPinsGroup)
        FullLevels = Split(FullGroup(PowerPinsCounter), ",")
        PowerPinLevelsGroup = Split(PowerPinsGroup(PowerPinsCounter), ",")
        ReDim SetInformation(UBound(PowerPinLevelsGroup))
        ReDim Aininformation(UBound(PowerPinLevelsGroup))
        ReDim Aixinformation(UBound(PowerPinLevelsGroup))
        ReDim PiUInformation(UBound(PowerPinLevelsGroup))
        For PowerPinLevelsCounter = 0 To UBound(PowerPinLevelsGroup)
            SensorsGroup = Split(SensorArray, ";")
            For SensorsCounter = 0 To UBound(SensorsGroup)
                ScanSeson = Split(SensorsGroup(SensorsCounter), ",")
                For ScanSesonCounter = 0 To UBound(ScanSeson)
                    If CStr(PowerPinLevelsGroup(PowerPinLevelsCounter)) = "VDD_GPU" Then
                        If ScanSeson(ScanSesonCounter) Like "GPU*" Then
                            VoltageScanerStr = CStr(PowerPinLevelsGroup(PowerPinLevelsCounter)) + "_" + "VoltageLevelCNT"
                            VoltageScaner = GetStoredMeasurement(VoltageScanerStr)
                            For Each site In TheExec.sites
                                SetInformation(PowerPinLevelsCounter).CreateConstant 0, 300, DspLong
                                Aininformation(PowerPinLevelsCounter).CreateConstant 0, 300, DspLong
                                Aixinformation(PowerPinLevelsCounter).CreateConstant 0, 300, DspLong
                                PiUInformation(PowerPinLevelsCounter).CreateConstant 0, 300, DspLong
                                If VoltageScaner = UBound(FullLevels) Then
                                    MatrixStr = "MetrologyMatrix_1150_450"
                                ElseIf VoltageScaner = UBound(FullLevels) - 1 Then
                                    MatrixStr = "MetrologyMatrix_1150_475"
                                ElseIf VoltageScaner = UBound(FullLevels) - 2 Then
                                    MatrixStr = "MetrologyMatrix_1150_500"
                                ElseIf VoltageScaner = UBound(FullLevels) - 3 Then
                                    MatrixStr = "MetrologyMatrix_1150_525"
                                ElseIf VoltageScaner <= UBound(FullLevels) - 4 Then
                                    MatrixStr = "MetrologyMatrix_1150_550"
                                End If
                                SetInformation(PowerPinLevelsCounter)(site) = Public_GetStoredCaptureData(MatrixStr + "_" + "size")
                                Aixinformation(PowerPinLevelsCounter)(site) = Public_GetStoredCaptureData(MatrixStr + "_" + "Aix")
                                Aininformation(PowerPinLevelsCounter)(site) = Public_GetStoredCaptureData(MatrixStr + "_" + "Ain")
                                PiUInformation(PowerPinLevelsCounter)(site) = Public_GetStoredCaptureData(MatrixStr + "_" + "PiUInfo")
                            Next site
                            Call MTR_Verification_Calculate(ScanSeson(ScanSesonCounter), Temperature, FusedCoeffDicName_1, FusedCoeffDicName_2, SetInformation(PowerPinLevelsCounter), _
                            Aininformation(PowerPinLevelsCounter), Aixinformation(PowerPinLevelsCounter), PiUInformation(PowerPinLevelsCounter), FullGroup(PowerPinsCounter))
                        End If
                    ElseIf CStr(PowerPinLevelsGroup(PowerPinLevelsCounter)) = "VDD_ECPU" Then
                        If ScanSeson(ScanSesonCounter) Like "ECPU*" Then
                            VoltageScanerStr = CStr(PowerPinLevelsGroup(PowerPinLevelsCounter)) + "_" + "VoltageLevelCNT"
                            VoltageScaner = GetStoredMeasurement(VoltageScanerStr)
                            For Each site In TheExec.sites
                                SetInformation(PowerPinLevelsCounter).CreateConstant 0, 300, DspLong
                                Aininformation(PowerPinLevelsCounter).CreateConstant 0, 300, DspLong
                                Aixinformation(PowerPinLevelsCounter).CreateConstant 0, 300, DspLong
                                PiUInformation(PowerPinLevelsCounter).CreateConstant 0, 300, DspLong
                                If VoltageScaner = UBound(FullLevels) Then
                                    MatrixStr = "MetrologyMatrix_1150_450"
                                ElseIf VoltageScaner = UBound(FullLevels) - 1 Then
                                    MatrixStr = "MetrologyMatrix_1150_475"
                                ElseIf VoltageScaner = UBound(FullLevels) - 2 Then
                                    MatrixStr = "MetrologyMatrix_1150_500"
                                ElseIf VoltageScaner = UBound(FullLevels) - 3 Then
                                    MatrixStr = "MetrologyMatrix_1150_525"
                                ElseIf VoltageScaner <= UBound(FullLevels) - 4 Then
                                    MatrixStr = "MetrologyMatrix_1150_550"
                                End If
                                SetInformation(PowerPinLevelsCounter)(site) = Public_GetStoredCaptureData(MatrixStr + "_" + "size")
                                Aixinformation(PowerPinLevelsCounter)(site) = Public_GetStoredCaptureData(MatrixStr + "_" + "Aix")
                                Aininformation(PowerPinLevelsCounter)(site) = Public_GetStoredCaptureData(MatrixStr + "_" + "Ain")
                                PiUInformation(PowerPinLevelsCounter)(site) = Public_GetStoredCaptureData(MatrixStr + "_" + "PiUInfo")
                            Next site
                            Call MTR_Verification_Calculate(ScanSeson(ScanSesonCounter), Temperature, FusedCoeffDicName_1, FusedCoeffDicName_2, SetInformation(PowerPinLevelsCounter), _
                            Aininformation(PowerPinLevelsCounter), Aixinformation(PowerPinLevelsCounter), PiUInformation(PowerPinLevelsCounter), FullGroup(PowerPinsCounter))
                        End If
                    ElseIf CStr(PowerPinLevelsGroup(PowerPinLevelsCounter)) = "VDD_PCPU" Then
                        If ScanSeson(ScanSesonCounter) Like "ANE*" Or ScanSeson(ScanSesonCounter) Like "PCPU*" Then
                        VoltageScanerStr = CStr(PowerPinLevelsGroup(PowerPinLevelsCounter)) + "_" + "VoltageLevelCNT"
                            VoltageScaner = GetStoredMeasurement(VoltageScanerStr)
                            For Each site In TheExec.sites
                                SetInformation(PowerPinLevelsCounter).CreateConstant 0, 300, DspLong
                                Aininformation(PowerPinLevelsCounter).CreateConstant 0, 300, DspLong
                                Aixinformation(PowerPinLevelsCounter).CreateConstant 0, 300, DspLong
                                PiUInformation(PowerPinLevelsCounter).CreateConstant 0, 300, DspLong
                                If VoltageScaner = UBound(FullLevels) Then
                                    MatrixStr = "MetrologyMatrix_1150_450"
                                ElseIf VoltageScaner = UBound(FullLevels) - 1 Then
                                    MatrixStr = "MetrologyMatrix_1150_475"
                                ElseIf VoltageScaner = UBound(FullLevels) - 2 Then
                                    MatrixStr = "MetrologyMatrix_1150_500"
                                ElseIf VoltageScaner = UBound(FullLevels) - 3 Then
                                    MatrixStr = "MetrologyMatrix_1150_525"
                                ElseIf VoltageScaner <= UBound(FullLevels) - 4 Then
                                    MatrixStr = "MetrologyMatrix_1150_550"
                                End If
                                SetInformation(PowerPinLevelsCounter)(site) = Public_GetStoredCaptureData(MatrixStr + "_" + "size")
                                Aixinformation(PowerPinLevelsCounter)(site) = Public_GetStoredCaptureData(MatrixStr + "_" + "Aix")
                                Aininformation(PowerPinLevelsCounter)(site) = Public_GetStoredCaptureData(MatrixStr + "_" + "Ain")
                                PiUInformation(PowerPinLevelsCounter)(site) = Public_GetStoredCaptureData(MatrixStr + "_" + "PiUInfo")
                            Next site
                            Call MTR_Verification_Calculate(ScanSeson(ScanSesonCounter), Temperature, FusedCoeffDicName_1, FusedCoeffDicName_2, SetInformation(PowerPinLevelsCounter), _
                            Aininformation(PowerPinLevelsCounter), Aixinformation(PowerPinLevelsCounter), PiUInformation(PowerPinLevelsCounter), FullGroup(PowerPinsCounter))
                        End If
                    ElseIf CStr(PowerPinLevelsGroup(PowerPinLevelsCounter)) = "VDD_AVE" Then
                        If ScanSeson(ScanSesonCounter) Like "AVE*" Then
                            VoltageScanerStr = CStr(PowerPinLevelsGroup(PowerPinLevelsCounter)) + "_" + "VoltageLevelCNT"
                            VoltageScaner = GetStoredMeasurement(VoltageScanerStr)
                            For Each site In TheExec.sites
                                SetInformation(PowerPinLevelsCounter).CreateConstant 0, 300, DspLong
                                Aininformation(PowerPinLevelsCounter).CreateConstant 0, 300, DspLong
                                Aixinformation(PowerPinLevelsCounter).CreateConstant 0, 300, DspLong
                                PiUInformation(PowerPinLevelsCounter).CreateConstant 0, 300, DspLong
                                If VoltageScaner = UBound(FullLevels) Then
                                    MatrixStr = "MetrologyMatrix_950_450"
                                ElseIf VoltageScaner = UBound(FullLevels) - 1 Then
                                    MatrixStr = "MetrologyMatrix_950_475"
                                ElseIf VoltageScaner = UBound(FullLevels) - 2 Then
                                    MatrixStr = "MetrologyMatrix_950_500"
                                ElseIf VoltageScaner = UBound(FullLevels) - 3 Then
                                    MatrixStr = "MetrologyMatrix_950_525"
                                ElseIf VoltageScaner <= UBound(FullLevels) - 4 Then
                                    MatrixStr = "MetrologyMatrix_950_550"
                                End If
                                SetInformation(PowerPinLevelsCounter)(site) = Public_GetStoredCaptureData(MatrixStr + "_" + "size")
                                Aixinformation(PowerPinLevelsCounter)(site) = Public_GetStoredCaptureData(MatrixStr + "_" + "Aix")
                                Aininformation(PowerPinLevelsCounter)(site) = Public_GetStoredCaptureData(MatrixStr + "_" + "Ain")
                                PiUInformation(PowerPinLevelsCounter)(site) = Public_GetStoredCaptureData(MatrixStr + "_" + "PiUInfo")
                            Next site
                            Call MTR_Verification_Calculate(ScanSeson(ScanSesonCounter), Temperature, FusedCoeffDicName_1, FusedCoeffDicName_2, SetInformation(PowerPinLevelsCounter), _
                            Aininformation(PowerPinLevelsCounter), Aixinformation(PowerPinLevelsCounter), PiUInformation(PowerPinLevelsCounter), FullGroup(PowerPinsCounter))
                        End If
                    ElseIf CStr(PowerPinLevelsGroup(PowerPinLevelsCounter)) = "VDD_SOC" Then
                        If ScanSeson(ScanSesonCounter) Like "SOC*" Then
                            VoltageScanerStr = CStr(PowerPinLevelsGroup(PowerPinLevelsCounter)) + "_" + "VoltageLevelCNT"
                            VoltageScaner = GetStoredMeasurement(VoltageScanerStr)
                            For Each site In TheExec.sites
                                SetInformation(PowerPinLevelsCounter).CreateConstant 0, 300, DspLong
                                Aininformation(PowerPinLevelsCounter).CreateConstant 0, 300, DspLong
                                Aixinformation(PowerPinLevelsCounter).CreateConstant 0, 300, DspLong
                                PiUInformation(PowerPinLevelsCounter).CreateConstant 0, 300, DspLong
                                If VoltageScaner = UBound(FullLevels) Then
                                    MatrixStr = "MetrologyMatrix_950_450"
                                ElseIf VoltageScaner = UBound(FullLevels) - 1 Then
                                    MatrixStr = "MetrologyMatrix_950_475"
                                ElseIf VoltageScaner = UBound(FullLevels) - 2 Then
                                    MatrixStr = "MetrologyMatrix_950_500"
                                ElseIf VoltageScaner = UBound(FullLevels) - 3 Then
                                    MatrixStr = "MetrologyMatrix_950_525"
                                ElseIf VoltageScaner <= UBound(FullLevels) - 4 Then
                                    MatrixStr = "MetrologyMatrix_950_550"
                                End If
                                SetInformation(PowerPinLevelsCounter)(site) = Public_GetStoredCaptureData(MatrixStr + "_" + "size")
                                Aixinformation(PowerPinLevelsCounter)(site) = Public_GetStoredCaptureData(MatrixStr + "_" + "Aix")
                                Aininformation(PowerPinLevelsCounter)(site) = Public_GetStoredCaptureData(MatrixStr + "_" + "Ain")
                                PiUInformation(PowerPinLevelsCounter)(site) = Public_GetStoredCaptureData(MatrixStr + "_" + "PiUInfo")
                            Next site
                            Call MTR_Verification_Calculate(ScanSeson(ScanSesonCounter), Temperature, FusedCoeffDicName_1, FusedCoeffDicName_2, SetInformation(PowerPinLevelsCounter), _
                            Aininformation(PowerPinLevelsCounter), Aixinformation(PowerPinLevelsCounter), PiUInformation(PowerPinLevelsCounter), FullGroup(PowerPinsCounter))
                        End If
                    End If
                Next ScanSesonCounter
            Next SensorsCounter
        Next PowerPinLevelsCounter
    Next PowerPinsCounter
'    thehdw.DSP.ExecutionMode = tlDSPModeHostDebug
End Function


'No Used in Sicily, Oscar, 20200423
Public Function Metrology_CAL_eFuse_Write(FuseType As String, SensorArray As String, m_catename As String, Dict_Store_Code_Name As String, Flag_Name As String, Hex_BitSize As String, Optional Temperature As String, _
                                          Optional Calculate_Group As String, Optional Efuse_Hex_Write_Flag As Boolean = True) As Long

    ' Parameter : eFuse Block , eFuse Variable , data
    ' Call auto_eFuse_SetPatTestPass_Flag("CFG", "LPDP_C_RX", TheHdw.Digital.Patgen.PatternBurstPassed(Site))
    ' Call auto_eFuse_SetWriteDecimal("CFG", "LPDP_C_RX", BestCode(Site))

    Dim site As Variant
    Dim DSPWave_Dict As New DSPWave
    Dim DSPWave_DictTemp As DSPWave
    Dim Data_Temp As String
    Dim m_value As New SiteVariant
    Dim i, j, k As Integer
    Dim Pass_Fail_Flag As New SiteBoolean
    On Error GoTo errHandler

    Dim SizeCounter As Long
    Dim SizeCounterTemp As Long
    Dim CalculateSize As String
    Dim CalculateArray() As String
    Dim CalculateSplit() As String
    Dim m_catenameTemp() As String
    Dim m_catenameCombination() As String
    Dim SensorArrayTemp() As String
    Dim Dict_Store_Code_NameTemp As String

    CalculateArray = Split(Calculate_Group, ",")
    m_catenameTemp = Split(m_catename, ",")
    m_catenameCombination = m_catenameTemp
    SensorArrayTemp = Split(SensorArray, ",")

    For i = 0 To UBound(SensorArrayTemp)
        SizeCounter = 1
        Dict_Store_Code_NameTemp = Dict_Store_Code_Name + "_" + SensorArrayTemp(i) + "_" + Temperature + "c"
        DSPWave_Dict = GetStoredCaptureData(Dict_Store_Code_NameTemp)

        For j = 0 To UBound(CalculateArray)
            CalculateSplit = Split(CalculateArray(j), ":")
            CalculateSize = CLng(CalculateSplit(1))

            For Each site In TheExec.sites
                If Efuse_Hex_Write_Flag Then
                    For k = 0 To (DSPWave_Dict(site).SampleSize - 1)
                        Data_Temp = Data_Temp & (DSPWave_Dict(site).Element(k))
                    Next k
                    Data_Temp = StrReverse(Data_Temp)
                    Data_Temp = mid(Data_Temp, SizeCounter, CalculateSize)
                    m_value(site) = "0x" + Calc_MTR_BinStr2HexStr(Data_Temp, CLng(Hex_BitSize))
                    Data_Temp = vbNullString
                Else
                    m_value(site) = DSPWave_Dict(site).Element(0)
                End If
            Next site
            SizeCounter = CalculateSize + SizeCounter

            m_catenameCombination(i) = m_catenameTemp(i) + "_" + "t" + Temperature + "_" + CalculateSplit(0)
            For Each site In TheExec.sites
                If TheExec.Flow.SiteFlag(site, Flag_Name) = 1 Then
                    Pass_Fail_Flag(site) = False
                    'm_value(site) = 0   'Cebu MTRGSNS fuse 0 when MTRGSNS fail
                ElseIf TheExec.Flow.SiteFlag(site, Flag_Name) = 0 Then
                    Pass_Fail_Flag(site) = True
                Else
                    Pass_Fail_Flag(site) = False
                    TheExec.Datalog.WriteComment ("Error! " & Flag_Name & "(" & site & ")" & " status is Clear !")
                End If
                    'Call auto_eFuse_SetPatTestPass_Flag(FuseType, m_catenameCombination(i), Pass_Fail_Flag(site), True)
                    'Call auto_eFuse_SetWriteDecimal(FuseType, m_catenameCombination(i), m_value(site), True)
            Next site
            
            '20210406 Modify for new Efuse
            Dim opbank As eFuseBdfBank
            Dim field As eFuseBdfField
            Set opbank = GetBdfBank(FuseType)
            Set field = opbank.Fields(m_catenameCombination(i))
            opbank.SetEfuse field.name, m_value, Pass_Fail_Flag, , , , True

''''''''''            For Each site In TheExec.sites
''''''''''                If Efuse_Hex_Write_Flag Then
''''''''''                    For k = 0 To (DSPWave_Dict(site).SampleSize - 1)
''''''''''                        Data_Temp = Data_Temp & (DSPWave_Dict(site).Element(k))
''''''''''                    Next k
''''''''''                    Data_Temp = StrReverse(Data_Temp)
''''''''''                    m_value(site) = "0x" + Calc_MTR_BinStr2HexStr(Data_Temp, CLng(Hex_BitSize))
''''''''''                    Data_Temp = ""
''''''''''                Else
''''''''''                    m_value(site) = DSPWave_Dict(site).Element(0)
''''''''''                End If
''''''''''            Next site
''''''''''            For Each site In TheExec.sites
''''''''''                If TheExec.Flow.SiteFlag(site, flag_name) = 1 Then
''''''''''                    Pass_Fail_Flag(site) = False
''''''''''                ElseIf TheExec.Flow.SiteFlag(site, flag_name) = 0 Then
''''''''''                    Pass_Fail_Flag(site) = True
''''''''''                Else
''''''''''                    Pass_Fail_Flag(site) = False
''''''''''                    TheExec.DataLog.WriteComment ("Error! " & flag_name & "(" & site & ")" & " status is Clear !")
''''''''''                End If
''''''''''                    Call auto_eFuse_SetPatTestPass_Flag(FuseType, m_catenameTemp(i), Pass_Fail_Flag(site), True)
''''''''''                    Call auto_eFuse_SetWriteDecimal(FuseType, m_catenameTemp(i), m_value(site), True)
''''''''''            Next site

        Next j
    Next i
Exit Function

errHandler:
    TheExec.Datalog.WriteComment "Error in HIP_eFuse_Write"
    If AbortTest Then Exit Function Else Resume Next

End Function

'''''
'''''Public Function pll_read() As Long
'''''
'''''    Call HIP_eFuse_Read(A, b, c)
'''''    Call HIP_eFuse_Read
'''''
'''''End Function

Public Function HIP_eFuse_Write_by_MTRGSNS(FuseType As String, m_catename As String, Dict_Store_Code_Name As String, Flag_Name As String, Optional Efuse_Binary_Write_Flag As Boolean = False, _
                                Optional Calc_code As String) As Long

    ' Parameter : eFuse Block , eFuse Variable , data
    ' Call auto_eFuse_SetPatTestPass_Flag("CFG", "LPDP_C_RX", TheHdw.Digital.Patgen.PatternBurstPassed(Site))
    ' Call auto_eFuse_SetWriteDecimal("CFG", "LPDP_C_RX", BestCode(Site))

    Dim site As Variant
    Dim DSPWave_Dict As New DSPWave
    Dim Data_Temp As String
    Dim m_value As New SiteDouble
    Dim j As Integer
    Dim Pass_Fail_Flag As New SiteBoolean
    On Error GoTo errHandler

    DSPWave_Dict = GetStoredCaptureData(Dict_Store_Code_Name)

        For Each site In TheExec.sites

                If Efuse_Binary_Write_Flag Then
                                For j = 0 To (DSPWave_Dict(site).SampleSize - 1)
                                        Data_Temp = Data_Temp & (DSPWave_Dict(site).Element(j))
                                Next j
                                m_value(site) = Bin2Dec_rev(Data_Temp)
                                Data_Temp = vbNullString
                Else
                        m_value(site) = DSPWave_Dict(site).Element(0)
                End If
        Next site
'''----------cal write fused code
    If Calc_code <> "" Then
    'Calc_code = "add,100"
        If Split(Calc_code, ",")(0) = "add" Then
            m_value = m_value.Add(Split(Calc_code, ",")(1))
        End If
    End If
'''----------cal write fused code
    For Each site In TheExec.sites
        If TheExec.Flow.SiteFlag(site, Flag_Name) = 1 Then
            Pass_Fail_Flag(site) = True
            m_value(site) = 0   'Cebu MTRGSNS fuse 0 when MTRGSNS fail
        ElseIf TheExec.Flow.SiteFlag(site, Flag_Name) = 0 Then
            Pass_Fail_Flag(site) = True
        Else
            Pass_Fail_Flag(site) = False
            TheExec.Datalog.WriteComment ("Error! " & Flag_Name & "(" & site & ")" & " status is Clear !")
        End If
        'Call auto_eFuse_SetPatTestPass_Flag(FuseType, m_catename, Pass_Fail_Flag(site), True)
        'Call auto_eFuse_SetWriteDecimal(FuseType, m_catename, m_value(site), True)
    Next site
    
    '20210406 Modify for new Efuse
    Dim opbank As eFuseBdfBank
    Dim field As eFuseBdfField
    Set opbank = GetBdfBank(FuseType)
    Set field = opbank.Fields(m_catename)
    opbank.SetEfuse field.name, m_value, Pass_Fail_Flag, , , , True

    Exit Function

errHandler:
    TheExec.Datalog.WriteComment "Error in HIP_eFuse_Write"
    If AbortTest Then Exit Function Else Resume Next

End Function


Public Function Check_MTRGSNS_25C_Fuse()
    On Error GoTo errHandler
    Dim site As Variant
    
    '20210406 Modify for new Efuse
    Dim opbank As eFuseBdfBank
    Dim field As eFuseBdfField
    Set opbank = GetBdfBank("CFG")
    Set field = opbank.Fields("mtr_sense_vt_ts3i_t25_a2_3")
    
    For Each site In TheExec.sites
        If field.DsscDecValue <> 0 Then
        'If CFGFuse.category(CFGIndex("mtr_sense_vt_ts3i_t25_a2_3")).Read.Decimal(site) <> 0 Then
            TheExec.Flow.TestLimit 1, lowVal:=1, hiVal:=1, Tname:="Check_MTRGSNS_25C_Fuse"
        Else
            TheExec.Flow.TestLimit 0, lowVal:=1, hiVal:=1, Tname:="Check_MTRGSNS_25C_Fuse"
        End If
    Next site

    Exit Function

errHandler:
    TheExec.Datalog.WriteComment "Error in Check_MTRGSNS_25C_Fuse"
    If AbortTest Then Exit Function Else Resume Next
End Function


Public Function MTR_REL_Fuse_Calc_Verification(Optional Adc_Offset As String, Optional Adc_Gain As String, Optional Cap_Code_Temp1 As String, Optional Encoded_Temp1 As String, Optional Cap_Code_Temp2 As String, Optional Temp2 As String, Optional Fuse_Calc_Temp1 As String, Optional Fuse_Calc_Temp2 As String, Optional Store_Coeff_c0 As String, Optional c0_size As String, Optional Store_Coeff_c1 As String, Optional c1_size As String, Optional Store_Coeff_c2 As String, Optional c2_size As String, Optional Store_Coeff_c3 As String, Optional c3_size As String, Optional Validating_ As Boolean) As Long

    
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


Dim Store_C0_Coeff_Dic As String
 Dim Store_C1_Coeff_Dic As String
 Dim Store_C2_Coeff_Dic As String
 Dim Store_C3_Coeff_Dic As String
 
 Dim C0_Reg_Size As Long
 Dim C1_Reg_Size As Long
 Dim C2_Reg_Size As Long
 Dim C3_Reg_Size As Long
 
 
 
    
    fuse_read_tfe_vol_0 = Adc_Offset
    fuse_read_tfe_vol_1 = Adc_Gain

    fuse_read_tfe_x0 = Cap_Code_Temp1
    Dict_name_tfe_vol_x1 = Cap_Code_Temp2

    fuse_read_tfe_y0 = Encoded_Temp1
    cal_tfe_vol_y1 = Temp2

    fuse_write_tfe_temp_0 = Fuse_Calc_Temp1
    fuse_write_tfe_temp_1 = Fuse_Calc_Temp2

    Store_C0_Coeff_Dic = Store_Coeff_c0
Store_C1_Coeff_Dic = Store_Coeff_c1
Store_C2_Coeff_Dic = Store_Coeff_c2
Store_C3_Coeff_Dic = Store_Coeff_c3

 C0_Reg_Size = CLng(c0_size)
 C1_Reg_Size = CLng(c1_size)
 C2_Reg_Size = CLng(c2_size)
 C3_Reg_Size = CLng(c3_size)


    'Get Cap data for t5p2 at 85C and Fuse Data for offset,gain and x0 at 25C
    Dim DSP_tfe_vol_x1_binary As New DSPWave
    Dim DSP_fuse_read_tfe_vol_0_2S_binary As New DSPWave
    Dim DSP_fuse_read_tfe_vol_1_binary As New DSPWave
    Dim DSP_fuse_read_tfe_x0_binary As New DSPWave

    Dim Dsp_tfe_temp0_in_binary As New DSPWave
    Dim Dsp_tfe_temp1_in_binary As New DSPWave




    DSP_tfe_vol_x1_binary = GetStoredCaptureData(Dict_name_tfe_vol_x1)
    DSP_fuse_read_tfe_vol_0_2S_binary = GetStoredCaptureData(fuse_read_tfe_vol_0)
    DSP_fuse_read_tfe_vol_1_binary = GetStoredCaptureData(fuse_read_tfe_vol_1)
    DSP_fuse_read_tfe_x0_binary = GetStoredCaptureData(fuse_read_tfe_x0)
    Dsp_tfe_temp0_in_binary = GetStoredCaptureData(fuse_write_tfe_temp_0)
    Dsp_tfe_temp1_in_binary = GetStoredCaptureData(fuse_write_tfe_temp_1)




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
  '  If cal_tfe_vol_y1 Like "CP2" Then

        actual_Temp_CP2 = CDbl(cal_tfe_vol_y1)

  '  End If

    Dim DSP_tfe_y1_in_double As New DSPWave

    DSP_tfe_y1_in_double.CreateConstant 0, 1, DspDouble

    For Each site In TheExec.sites

    DSP_tfe_y1_in_double(site).Element(0) = actual_Temp_CP2

    Next site

    'Start the algo


    'Define Constants

    Dim A0 As Double
    Dim a1 As Double
    Dim a2 As Double
    Dim a3 As Double

    'Values for Constants

    A0 = CDbl("-21.5822184999726")
    a1 = CDbl("428.0092266096283") 'truncated one digit
    a2 = CDbl("-133.4543109228228") 'truncated one digit
    a3 = CDbl("19.0485545665615")




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




    Dim Dsp_tfe_temp0_in_decimal As New DSPWave
    Dim Dsp_tfe_temp1_in_decimal As New DSPWave

    Dsp_tfe_temp0_in_decimal.CreateConstant 0, 1, DspDouble
     Dsp_tfe_temp1_in_decimal.CreateConstant 0, 1, DspDouble

     Set SL_BitWidth = Nothing

     For Each site In TheExec.sites
            SL_BitWidth(site) = 28

    Next site

     Call rundsp.DSP_2S_Complement_To_SignDec(Dsp_tfe_temp0_in_binary, SL_BitWidth, Dsp_tfe_temp0_in_decimal)
    Set SL_BitWidth = Nothing

     For Each site In TheExec.sites
            SL_BitWidth(site) = 28

    Next site
 Call rundsp.DSP_2S_Complement_To_SignDec(Dsp_tfe_temp1_in_binary, SL_BitWidth, Dsp_tfe_temp1_in_decimal)


    Dim detA0 As New SiteDouble
    Dim detA1 As New SiteDouble
    Dim offset As New SiteDouble
    Dim Gain As New SiteDouble

    Dim C0 As New SiteDouble
    Dim c1 As New SiteDouble
    Dim C2 As New SiteDouble
    Dim C3 As New SiteDouble

    Dim C0_coeff_Val As New SiteDouble
    Dim C1_coeff_Val As New SiteDouble
    Dim C2_coeff_Val As New SiteDouble
    Dim C3_coeff_Val As New SiteDouble

    Dim Coeff_C0_in_decimal As New DSPWave
    Dim Coeff_C1_in_decimal As New DSPWave
    Dim Coeff_C2_in_decimal As New DSPWave
    Dim Coeff_C3_in_decimal As New DSPWave
    
     Coeff_C0_in_decimal.CreateConstant 0, 1, DspDouble
     Coeff_C1_in_decimal.CreateConstant 0, 1, DspDouble
    Coeff_C2_in_decimal.CreateConstant 0, 1, DspDouble
     Coeff_C3_in_decimal.CreateConstant 0, 1, DspDouble
    

    Dim Coeff_C0_in_binary As New DSPWave
    Dim Coeff_C1_in_binary As New DSPWave
    Dim Coeff_C2_in_binary As New DSPWave
    Dim Coeff_C3_in_binary As New DSPWave


    For Each site In TheExec.sites

                TheExec.Flow.TestLimit resultVal:=DSP_fuse_read_tfe_vol_0_in_decimal(site).Element(0), Tname:="tfe_vol_0", ForceResults:=tlForceNone 'eng_forceflow_transfer

                TheExec.Flow.TestLimit resultVal:=DSP_fuse_read_tfe_vol_1_in_decimal(site).Element(0), Tname:="tfe_vol_1", ForceResults:=tlForceNone 'eng_forceflow_transfer


                TheExec.Flow.TestLimit resultVal:=Dsp_tfe_temp0_in_decimal(site).Element(0), Tname:="temp_0", ForceResults:=tlForceNone 'eng_forceflow_transfer

                TheExec.Flow.TestLimit resultVal:=Dsp_tfe_temp1_in_decimal(site).Element(0), Tname:="temp_1", ForceResults:=tlForceNone 'eng_forceflow_transfer



                If (DSP_fuse_read_tfe_vol_1_in_decimal(site).Element(0) = 0) Then

                         C0_coeff_Val(site) = 178956970

                        C1_coeff_Val(site) = 178956970


                         C2_coeff_Val(site) = 178956970

                        C3_coeff_Val(site) = 178956970


                            Coeff_C0_in_decimal(site).Element(0) = C0_coeff_Val(site)

                        Coeff_C1_in_decimal(site).Element(0) = C1_coeff_Val(site)

                        Coeff_C2_in_decimal(site).Element(0) = C2_coeff_Val(site)

                        Coeff_C3_in_decimal(site).Element(0) = C3_coeff_Val(site)


                        TheExec.Flow.TestLimit resultVal:=Coeff_C0_in_decimal(site).Element(0), Tname:="Error_coeff_c0", ForceResults:=tlForceNone 'eng_forceflow_transfer

                    TheExec.Flow.TestLimit resultVal:=Coeff_C1_in_decimal(site).Element(0), Tname:="Error_coeff_c1", ForceResults:=tlForceNone 'eng_forceflow_transfer

                    TheExec.Flow.TestLimit resultVal:=Coeff_C2_in_decimal(site).Element(0), Tname:="Error_coeff_c2", ForceResults:=tlForceNone 'eng_forceflow_transfer

                    TheExec.Flow.TestLimit resultVal:=Coeff_C3_in_decimal(site).Element(0), Tname:="Error_coeff_c3", ForceResults:=tlForceNone 'eng_forceflow_transfer

                Else




                    detA0(site) = Dsp_tfe_temp0_in_decimal(site).Element(0) / (2 ^ 13)
                    detA1(site) = Dsp_tfe_temp1_in_decimal(site).Element(0) / (2 ^ 13)
                    offset(site) = DSP_fuse_read_tfe_vol_0_in_decimal(site).Element(0) / (2 ^ 13)
                    Gain(site) = DSP_fuse_read_tfe_vol_1_in_decimal(site).Element(0) / (2 ^ 17)

                    C0(site) = A0 + detA0(site) - (a1 + detA1(site)) * (offset(site) / Gain(site))
                    c1(site) = ((a1 + detA1(site)) / Gain(site)) - 2 * a2 * (offset(site) / (Gain(site) ^ 2))
                    C2(site) = (a2 / (Gain(site) ^ 2)) - 3 * a3 * (offset(site) / (Gain(site) ^ 3))
                    C3(site) = a3 / (Gain(site) ^ 3)


                   TheExec.Flow.TestLimit resultVal:=detA0(site), Tname:="detA0", ForceResults:=tlForceNone 'eng_forceflow_transfer

                   TheExec.Flow.TestLimit resultVal:=detA1(site), Tname:="detA1", ForceResults:=tlForceNone 'eng_forceflow_transfer



                   TheExec.Flow.TestLimit resultVal:=offset(site), Tname:="Offset", ForceResults:=tlForceNone 'eng_forceflow_transfer

                   TheExec.Flow.TestLimit resultVal:=Gain(site), Tname:="Gain", ForceResults:=tlForceNone 'eng_forceflow_transfer

                    TheExec.Flow.TestLimit resultVal:=C0(site), Tname:="C0", ForceResults:=tlForceNone 'eng_forceflow_transfer

                   TheExec.Flow.TestLimit resultVal:=c1(site), Tname:="C1", ForceResults:=tlForceNone 'eng_forceflow_transfer



                   TheExec.Flow.TestLimit resultVal:=C2(site), Tname:="C2", ForceResults:=tlForceNone 'eng_forceflow_transfer

                   TheExec.Flow.TestLimit resultVal:=C3(site), Tname:="C3", ForceResults:=tlForceNone 'eng_forceflow_transfer

            C0_coeff_Val(site) = FormatNumber(C0(site) * 2 ^ 13)
            C1_coeff_Val(site) = FormatNumber(c1(site) * 2 ^ 13)
            C2_coeff_Val(site) = FormatNumber(C2(site) * 2 ^ 13)
            C3_coeff_Val(site) = FormatNumber(C3(site) * 2 ^ 13)








                    If (C0_coeff_Val(site) > 42949672945#) Then


                         TheExec.Flow.TestLimit resultVal:=C0_coeff_Val(site), Tname:="UpperLimit_Reached_c0_coeff", ForceResults:=tlForceNone 'eng_forceflow_transfer


                        C0_coeff_Val(site) = 4294967295#



                    End If
                     If (C1_coeff_Val(site) > 4294967295#) Then


                         TheExec.Flow.TestLimit resultVal:=C1_coeff_Val(site), Tname:="UpperLimit_Reached_c1_coeff", ForceResults:=tlForceNone 'eng_forceflow_transfer


                        C1_coeff_Val(site) = 4294967295#



                    End If
                     If (C2_coeff_Val(site) > 42949672945#) Then


                         TheExec.Flow.TestLimit resultVal:=C2_coeff_Val(site), Tname:="UpperLimit_Reached_c2_coeff", ForceResults:=tlForceNone 'eng_forceflow_transfer


                        C2_coeff_Val(site) = 42949672945#



                    End If
                     If (C3_coeff_Val(site) > 42949672945#) Then


                         TheExec.Flow.TestLimit resultVal:=C3_coeff_Val(site), Tname:="UpperLimit_Reached_c3_coeff", ForceResults:=tlForceNone 'eng_forceflow_transfer


                        C3_coeff_Val(site) = 42949672945#



                    End If

                Coeff_C0_in_decimal(site).Element(0) = C0_coeff_Val(site)

                Coeff_C1_in_decimal(site).Element(0) = C1_coeff_Val(site)
                Coeff_C2_in_decimal(site).Element(0) = C2_coeff_Val(site)

                Coeff_C3_in_decimal(site).Element(0) = C3_coeff_Val(site)



                    TheExec.Flow.TestLimit resultVal:=Coeff_C0_in_decimal(site).Element(0), Tname:="C0_coeff_Val", ForceResults:=tlForceNone 'eng_forceflow_transfer

                   TheExec.Flow.TestLimit resultVal:=Coeff_C1_in_decimal(site).Element(0), Tname:="C1_coeff_Val", ForceResults:=tlForceNone 'eng_forceflow_transfer



                   TheExec.Flow.TestLimit resultVal:=Coeff_C2_in_decimal(site).Element(0), Tname:="C2_coeff_Val", ForceResults:=tlForceNone 'eng_forceflow_transfer

                   TheExec.Flow.TestLimit resultVal:=Coeff_C3_in_decimal(site).Element(0), Tname:="C3_coeff_Val", ForceResults:=tlForceNone 'eng_forceflow_transfer



            End If

    Next site





    Call rundsp.DSPWf_Dec2Binary(Coeff_C0_in_decimal, C0_Reg_Size, Coeff_C0_in_binary)

    Call rundsp.DSPWf_Dec2Binary(Coeff_C1_in_decimal, C1_Reg_Size, Coeff_C1_in_binary)

        Call rundsp.DSPWf_Dec2Binary(Coeff_C2_in_decimal, C2_Reg_Size, Coeff_C2_in_binary)

    Call rundsp.DSPWf_Dec2Binary(Coeff_C3_in_decimal, C3_Reg_Size, Coeff_C3_in_binary)

    'Algo end


Store_C0_Coeff_Dic = Store_Coeff_c0
Store_C1_Coeff_Dic = Store_Coeff_c1
Store_C2_Coeff_Dic = Store_Coeff_c2
Store_C3_Coeff_Dic = Store_Coeff_c3


    Call AddStoredCaptureData(Store_C0_Coeff_Dic, Coeff_C0_in_binary)
    Call AddStoredCaptureData(Store_C1_Coeff_Dic, Coeff_C1_in_binary)

    Call AddStoredCaptureData(Store_C2_Coeff_Dic, Coeff_C2_in_binary)
    Call AddStoredCaptureData(Store_C3_Coeff_Dic, Coeff_C3_in_binary)


End Function

Public Function MTR_Sense_Alignment_Calc(Optional FreqSensorArray As String, Optional SensorArray As String, Optional SweepVArray As String, Optional Temperature As String) As Long
                                                                                                                                                                                                                                                             
    If TheHdw.DSP.ExecutionMode = tlDSPModeHostDebug Then
        TheHdw.DSP.ExecutionMode = tlDSPModeAutomatic
    End If

Dim i, j, k As Long
Dim site As Variant
Dim SweepTempName As String
Dim SensorTempName As String
Dim SweepOriginTempName As String
Dim SensorOriginTempName As String


Dim SplitSensorArray() As String
Dim SplitSweepVArray() As String
Dim SplitSensorGroup() As String
Dim SplitSweepVGroup() As String
Dim SplitOriginSensorArray() As String
Dim SplitOriginSensorGroup() As String
SplitSensorGroup = Split(FreqSensorArray, ";")
SplitSweepVGroup = Split(SweepVArray, ";")
SplitOriginSensorGroup = Split(SensorArray, ";")
Dim TempDSPWave As New DSPWave
Dim RotRovMatrix As New DSPWave
Dim RotRovOriginMatrix As New DSPWave
Dim DicCounterbyGroup As Long
Dim DicCounterbyGroup_temp As Long
RotRovMatrix.CreateConstant 0, 16, DspDouble
RotRovOriginMatrix.CreateConstant 0, 16, DspDouble
DicCounterbyGroup = -1
For i = 0 To UBound(SplitSensorGroup)
    SplitSensorArray = Split(SplitSensorGroup(i), ",")
    For j = 0 To UBound(SplitSensorArray)
        DicCounterbyGroup = DicCounterbyGroup + 1
    Next j
Next i

DicCounterbyGroup_temp = 0
For i = 0 To UBound(SplitSensorGroup)
    SplitSensorArray = Split(SplitSensorGroup(i), ",")
    SplitSweepVArray = Split(SplitSweepVGroup(i), ",")
    SplitOriginSensorArray = Split(SplitOriginSensorGroup(i), ",")
    
    For j = 0 To UBound(SplitSensorArray)
        Dim RotRovMatrix_temp() As New DSPWave
        ReDim RotRovMatrix_temp(DicCounterbyGroup)
        Dim RotRovOriginMatrix_temp() As New DSPWave
        ReDim RotRovOriginMatrix_temp(DicCounterbyGroup)
        
        For Each site In TheExec.sites
            RotRovMatrix(site).CreateConstant 0, UBound(SplitSweepVArray) + 1, DspDouble
            RotRovOriginMatrix(site).CreateConstant 0, UBound(SplitSweepVArray) + 1, DspDouble
        Next site
        
        SensorTempName = SplitSensorArray(j) + "_" + Temperature + "C"
        SensorOriginTempName = SplitOriginSensorArray(j) + "_" + Temperature + "C"
        
        For k = 0 To UBound(SplitSweepVArray)
            
            
            Set TempDSPWave = Nothing
            Dim DPSWaveConvert As New SiteDouble
            SweepTempName = SplitSensorArray(j) + "_" + SplitSweepVArray(k)
            SweepOriginTempName = SplitOriginSensorArray(j) + "_" + SplitSweepVArray(k)
            
            TempDSPWave = GetStoredCaptureData(SweepTempName)
            DPSWaveConvert = GetStoredData(SweepOriginTempName & "_para")
            
            For Each site In TheExec.sites
                RotRovMatrix(site).Element(k) = TempDSPWave(site).Element(0)
                RotRovOriginMatrix(site).Element(k) = CDbl(DPSWaveConvert)
            Next site
            
        Next k
    
        RotRovMatrix_temp(DicCounterbyGroup_temp) = RotRovMatrix
        RotRovOriginMatrix_temp(DicCounterbyGroup_temp) = RotRovOriginMatrix
        Call AddStoredCaptureData(SensorTempName, RotRovMatrix_temp(DicCounterbyGroup_temp))
        Call AddStoredCaptureData(SensorOriginTempName, RotRovOriginMatrix_temp(DicCounterbyGroup_temp))
        DicCounterbyGroup_temp = DicCounterbyGroup_temp + 1
        
    Next j
Next i
End Function


Public Function MTR_Sense_Calibration_Coeff_Calc(Optional FreqSensorAlignment As String, Optional SensorAlignment As String, _
Optional SweepVArrayDic As String, Optional SweepVArrayValue As String, Optional Temperature As String, Optional FuseSize_1 As String, _
Optional StoreFuseDicName_1 As String, Optional FuseSize_2 As String, Optional StoreFuseDicName_2 As String, Optional SensorCalculate As String, _
Optional MTR_CAL_Sheet As String, Optional Validating_ As Boolean) As Long '
'MTR Record
    
    If Validating_ Then
        Exit Function    ' Exit after validation
    End If
    
    


    If TheHdw.DSP.ExecutionMode = tlDSPModeHostDebug Then
        TheHdw.DSP.ExecutionMode = tlDSPModeAutomatic
    End If
    
    Call MTR_Sense_Alignment_Calc(FreqSensorAlignment, SensorAlignment, SweepVArrayDic, Temperature)

'===================================================================================================================================================='
'Read bincut voltage from eFuse
'===================================================================================================================================================='
'    If LCase(TheExec.CurrentJob) <> "cp1" Then Call Read_DVFM_To_GradeVDD
    Dim PrintStr As String
    Dim i, j, k As Integer
    Dim CP_GB_Record As Double
    Dim SensorSplitStr() As String
    Dim eFuseValueOnly() As New SiteDouble
    Dim eFuseCPGBOnly() As New SiteDouble
    Dim eFuseValueLowest() As New SiteDouble
    Dim eFuseValueHighest() As New SiteDouble
    Dim CurrentPassBinCutNum_MTR As New SiteLong
    SensorSplitStr = Split(SensorCalculate, ",")
    
    ReDim eFuseValueOnly(UBound(SensorSplitStr))
    ReDim eFuseCPGBOnly(UBound(SensorSplitStr))
    ReDim eFuseValueLowest(UBound(SensorSplitStr))
    ReDim eFuseValueHighest(UBound(SensorSplitStr))
    
    '20210406 Modify for new Efuse
    Dim opbank As eFuseBdfBank
    Dim field As eFuseBdfField
    Set opbank = GetBdfBank("CFG")
    Set field = opbank.Fields("Product_Identifier")

    For Each site In TheExec.sites
        For i = 0 To UBound(SensorSplitStr)
            If LCase(SensorSplitStr(i)) Like "vdd_ecpu" Then
                For j = 0 To UBound(BinCut_Power_Seq(VddBinStr2Enum("vdd_ecpu")).Power_Seq)
                    eFuseValueLowest(i) = VBIN_RESULT(VddBinStr2Enum(BinCut_Power_Seq(VddBinStr2Enum("vdd_ecpu")).Power_Seq(j))).GRADEVDD
                    If eFuseValueLowest(i) <> 0 Then
                        CurrentPassBinCutNum_MTR(site) = field.DsscValue + 1   '20210406 Modify for new Efuse
                        'CurrentPassBinCutNum_MTR(site) = auto_eFuse_GetReadValue("CFG", "Product_Identifier") + 1
                        CP_GB_Record = BinCut(VddBinStr2Enum(BinCut_Power_Seq(VddBinStr2Enum("vdd_ecpu")).Power_Seq(j)), CurrentPassBinCutNum_MTR).CP_GB(0)
                        eFuseCPGBOnly(i) = CP_GB_Record / 1000
                        eFuseValueOnly(i) = eFuseValueLowest(i) / 1000
                        eFuseValueLowest(i) = eFuseValueLowest(i) - CP_GB_Record
                        eFuseValueLowest(i) = (eFuseValueLowest(i)) / 1000
                        Exit For
                    End If
                Next j
            ElseIf LCase(SensorSplitStr(i)) Like "vdd_pcpu" Then
                For j = 0 To UBound(BinCut_Power_Seq(VddBinStr2Enum("vdd_pcpu")).Power_Seq)
                    eFuseValueLowest(i) = VBIN_RESULT(VddBinStr2Enum(BinCut_Power_Seq(VddBinStr2Enum("vdd_pcpu")).Power_Seq(j))).GRADEVDD
                    If eFuseValueLowest(i) <> 0 Then
                        CurrentPassBinCutNum_MTR(site) = field.DsscValue + 1   '20210406 Modify for new Efuse
                        'CurrentPassBinCutNum_MTR(site) = auto_eFuse_GetReadValue("CFG", "Product_Identifier") + 1
                        CP_GB_Record = BinCut(VddBinStr2Enum(BinCut_Power_Seq(VddBinStr2Enum("vdd_pcpu")).Power_Seq(j)), CurrentPassBinCutNum_MTR).CP_GB(0)
                        eFuseCPGBOnly(i) = CP_GB_Record / 1000
                        eFuseValueOnly(i) = eFuseValueLowest(i) / 1000
                        eFuseValueLowest(i) = eFuseValueLowest(i) - CP_GB_Record
                        eFuseValueLowest(i) = (eFuseValueLowest(i)) / 1000
                        Exit For
                    End If
                Next j
            ElseIf LCase(SensorSplitStr(i)) Like "vdd_gpu" Then
                For j = 0 To UBound(BinCut_Power_Seq(VddBinStr2Enum("vdd_gpu")).Power_Seq)
                    eFuseValueLowest(i) = VBIN_RESULT(VddBinStr2Enum(BinCut_Power_Seq(VddBinStr2Enum("vdd_gpu")).Power_Seq(j))).GRADEVDD
                    If eFuseValueLowest(i) <> 0 Then
                        CurrentPassBinCutNum_MTR(site) = field.DsscValue + 1   '20210406 Modify for new Efuse
                        'CurrentPassBinCutNum_MTR(site) = auto_eFuse_GetReadValue("CFG", "Product_Identifier") + 1
                        CP_GB_Record = BinCut(VddBinStr2Enum(BinCut_Power_Seq(VddBinStr2Enum("vdd_gpu")).Power_Seq(j)), CurrentPassBinCutNum_MTR).CP_GB(0)
                        eFuseCPGBOnly(i) = CP_GB_Record / 1000
                        eFuseValueOnly(i) = eFuseValueLowest(i) / 1000
                        eFuseValueLowest(i) = eFuseValueLowest(i) - CP_GB_Record
                        eFuseValueLowest(i) = (eFuseValueLowest(i)) / 1000
                        Exit For
                    End If
                Next j
            ElseIf LCase(SensorSplitStr(i)) Like "vdd_soc" Then
                For j = 0 To UBound(BinCut_Power_Seq(VddBinStr2Enum("vdd_soc")).Power_Seq)
                    eFuseValueLowest(i) = VBIN_RESULT(VddBinStr2Enum(BinCut_Power_Seq(VddBinStr2Enum("vdd_soc")).Power_Seq(j))).GRADEVDD
                    If eFuseValueLowest(i) <> 0 Then
                        CurrentPassBinCutNum_MTR(site) = field.DsscValue + 1   '20210406 Modify for new Efuse
                        'CurrentPassBinCutNum_MTR(site) = auto_eFuse_GetReadValue("CFG", "Product_Identifier") + 1
                        CP_GB_Record = BinCut(VddBinStr2Enum(BinCut_Power_Seq(VddBinStr2Enum("vdd_soc")).Power_Seq(j)), CurrentPassBinCutNum_MTR).CP_GB(0)
                        eFuseCPGBOnly(i) = CP_GB_Record / 1000
                        eFuseValueOnly(i) = eFuseValueLowest(i) / 1000
                        eFuseValueLowest(i) = eFuseValueLowest(i) - CP_GB_Record
                        eFuseValueLowest(i) = (eFuseValueLowest(i)) / 1000
                        Exit For
                    End If
                Next j
            ElseIf LCase(SensorSplitStr(i)) Like "vdd_ave" Then
                For j = 0 To UBound(BinCut_Power_Seq(VddBinStr2Enum("vdd_ave")).Power_Seq)
                    eFuseValueLowest(i) = VBIN_RESULT(VddBinStr2Enum(BinCut_Power_Seq(VddBinStr2Enum("vdd_ave")).Power_Seq(j))).GRADEVDD
                    If eFuseValueLowest(i) <> 0 Then
                        CurrentPassBinCutNum_MTR(site) = field.DsscValue + 1   '20210406 Modify for new Efuse
                        'CurrentPassBinCutNum_MTR(site) = auto_eFuse_GetReadValue("CFG", "Product_Identifier") + 1
                        CP_GB_Record = BinCut(VddBinStr2Enum(BinCut_Power_Seq(VddBinStr2Enum("vdd_ave")).Power_Seq(j)), CurrentPassBinCutNum_MTR).CP_GB(0)
                        eFuseCPGBOnly(i) = CP_GB_Record / 1000
                        eFuseValueOnly(i) = eFuseValueLowest(i) / 1000
                        eFuseValueLowest(i) = eFuseValueLowest(i) - CP_GB_Record
                        eFuseValueLowest(i) = (eFuseValueLowest(i)) / 1000
                        Exit For
                    End If
                Next j
            End If
        Next i
    Next site

    For Each site In TheExec.sites
        For i = 0 To UBound(SensorSplitStr)
            If LCase(SensorSplitStr(i)) Like "vdd_ecpu" Then
                For j = UBound(BinCut_Power_Seq(VddBinStr2Enum("vdd_ecpu")).Power_Seq) To 0 Step -1
                    eFuseValueHighest(i) = VBIN_RESULT(VddBinStr2Enum(BinCut_Power_Seq(VddBinStr2Enum("vdd_ecpu")).Power_Seq(j))).GRADEVDD
                    If eFuseValueHighest(i) <> 0 Then
                        eFuseValueHighest(i) = (eFuseValueHighest(i)) / 1000
                        Exit For
                    End If
                Next j
            ElseIf LCase(SensorSplitStr(i)) Like "vdd_pcpu" Then
                For j = UBound(BinCut_Power_Seq(VddBinStr2Enum("vdd_pcpu")).Power_Seq) To 0 Step -1
                    eFuseValueHighest(i) = VBIN_RESULT(VddBinStr2Enum(BinCut_Power_Seq(VddBinStr2Enum("vdd_pcpu")).Power_Seq(j))).GRADEVDD
                    If eFuseValueHighest(i) <> 0 Then
                        eFuseValueHighest(i) = (eFuseValueHighest(i)) / 1000
                        Exit For
                    End If
                Next j
            ElseIf LCase(SensorSplitStr(i)) Like "vdd_gpu" Then
                For j = UBound(BinCut_Power_Seq(VddBinStr2Enum("vdd_gpu")).Power_Seq) To 0 Step -1
                    eFuseValueHighest(i) = VBIN_RESULT(VddBinStr2Enum(BinCut_Power_Seq(VddBinStr2Enum("vdd_gpu")).Power_Seq(j))).GRADEVDD
                    If eFuseValueHighest(i) <> 0 Then
                        eFuseValueHighest(i) = (eFuseValueHighest(i)) / 1000
                        Exit For
                    End If
                Next j
            ElseIf LCase(SensorSplitStr(i)) Like "vdd_soc" Then
                For j = UBound(BinCut_Power_Seq(VddBinStr2Enum("vdd_soc")).Power_Seq) To 0 Step -1
                    eFuseValueHighest(i) = VBIN_RESULT(VddBinStr2Enum(BinCut_Power_Seq(VddBinStr2Enum("vdd_soc")).Power_Seq(j))).GRADEVDD
                    If eFuseValueHighest(i) <> 0 Then
                        eFuseValueHighest(i) = (eFuseValueHighest(i)) / 1000
                        Exit For
                    End If
                Next j
            ElseIf LCase(SensorSplitStr(i)) Like "vdd_ave" Then
                For j = UBound(BinCut_Power_Seq(VddBinStr2Enum("vdd_ave")).Power_Seq) To 0 Step -1
                    eFuseValueHighest(i) = VBIN_RESULT(VddBinStr2Enum(BinCut_Power_Seq(VddBinStr2Enum("vdd_ave")).Power_Seq(j))).GRADEVDD
                    If eFuseValueHighest(i) <> 0 Then
                        eFuseValueHighest(i) = (eFuseValueHighest(i)) / 1000
                        Exit For
                    End If
                Next j
            End If
        Next i
    Next site

''===================================================================================================================================================='
''Calculate measurement point is greater or not
''===================================================================================================================================================='
    Dim Catchdone() As Boolean
    Dim SensorVoltage() As String
    Dim VoltageMark() As New SiteDouble
    Dim VoltageLevelCNT() As New SiteLong
    Dim VoltageLevelOffset() As New SiteLong
    ReDim Catchdone(UBound(SensorSplitStr))
    ReDim VoltageMark(UBound(SensorSplitStr))
    ReDim VoltageLevelCNT(UBound(SensorSplitStr))
    ReDim VoltageLevelOffset(UBound(SensorSplitStr))
    
    SensorVoltage = Split(SweepVArrayValue, ",")
    For Each site In TheExec.sites
        For i = 0 To UBound(SensorSplitStr)
            VoltageLevelCNT(i) = UBound(SensorVoltage)
        Next i
    Next site
    
    For Each site In TheExec.sites
        For i = 0 To UBound(SensorSplitStr)
            Catchdone(i) = True
            VoltageLevelOffset(i) = 0
        Next i
        For i = 0 To UBound(SensorSplitStr)
            For j = 0 To UBound(SensorVoltage)
                If Catchdone(i) = True Then
                    If CDbl(SensorVoltage(j)) < eFuseValueLowest(i) Then
                        VoltageLevelOffset(i) = VoltageLevelOffset(i) + 1
                        VoltageMark(i) = CDbl(SensorVoltage(j))
                    ElseIf (CDbl(SensorVoltage(j)) >= eFuseValueLowest(i)) And Catchdone(i) = True And VoltageLevelOffset(i) <> 0 Then
                        If VoltageMark(i) < eFuseValueLowest(i) Then
                            VoltageLevelOffset(i) = VoltageLevelOffset(i) - 1
'                            PrintStr = "Site" + CStr(site) + "_" + CStr(SensorSplitStr(i)) + "_" + "LowestMode(eFuse-CPBG) : " + CStr(eFuseValueLowest(i))
'                            TheExec.Datalog.WriteComment (PrintStr)
                            Catchdone(i) = False
                        End If
                    End If
                End If
            Next j
            If VoltageLevelOffset(i) > 4 Then
                VoltageLevelOffset(i) = 4
                VoltageLevelCNT(i) = VoltageLevelCNT(i) - VoltageLevelOffset(i)
            Else
                VoltageLevelCNT(i) = VoltageLevelCNT(i) - VoltageLevelOffset(i)
            End If
        Next i
    Next site

    For i = 0 To UBound(SensorSplitStr)
        TheExec.Flow.TestLimit resultVal:=eFuseValueOnly(i), Tname:=CStr(SensorSplitStr(i)) + "_Lowest_eFuse", ForceResults:=tlForceNone, scaletype:=scaleNone 'eng_forceflow_transfer
        TheExec.Flow.TestLimit resultVal:=eFuseCPGBOnly(i), Tname:=CStr(SensorSplitStr(i)) + "_Lowest_CPGB", ForceResults:=tlForceNone, scaletype:=scaleNone 'eng_forceflow_transfer
        TheExec.Flow.TestLimit resultVal:=eFuseValueLowest(i), Tname:=CStr(SensorSplitStr(i)) + "_Lowest(eFuse - CPGB)", ForceResults:=tlForceNone, scaletype:=scaleNone 'eng_forceflow_transfer
    Next i

'    For i = 0 To UBound(SensorSplitStr)
'        PrintStr = CStr(SensorSplitStr(i)) + "_" + "VoltageLevelCNT"
'        Call AddStoredMeasurement(PrintStr, VoltageLevelCNT(i))
'    Next i
    Dim DicEfuseName As String
    Dim DicEfuseRecord() As New DSPWave
    ReDim DicEfuseRecord(UBound(SensorSplitStr))
    For i = 0 To UBound(SensorSplitStr)
        DicEfuseRecord(i).CreateConstant 0, 1, DspLong
        PrintStr = CStr(SensorSplitStr(i)) + "_" + "VoltageLevelCNT"
        Call AddStoredMeasurement(PrintStr, VoltageLevelCNT(i))
        If SensorSplitStr(i) = "VDD_ECPU" Then
            DicEfuseName = "dic_mtr_comp_matrix_vdd_ecpu"
        ElseIf SensorSplitStr(i) = "VDD_PCPU" Then
            DicEfuseName = "dic_mtr_comp_matrix_vdd_pcpu"
        ElseIf SensorSplitStr(i) = "VDD_GPU" Then
            DicEfuseName = "dic_mtr_comp_matrix_vdd_gpu"
        ElseIf SensorSplitStr(i) = "VDD_SOC" Then
            DicEfuseName = "dic_mtr_comp_matrix_vdd_soc"
        ElseIf SensorSplitStr(i) = "VDD_AVE" Then
            DicEfuseName = "dic_mtr_comp_matrix_vdd_ave"
        End If
        For Each site In TheExec.sites
            If UBound(SensorVoltage) > 13 Then
                If VoltageLevelCNT(i) = UBound(SensorVoltage) Then
                    DicEfuseRecord(i).Element(0) = 5
                ElseIf VoltageLevelCNT(i) = UBound(SensorVoltage) - 1 Then
                    DicEfuseRecord(i).Element(0) = 4
                ElseIf VoltageLevelCNT(i) = UBound(SensorVoltage) - 2 Then
                    DicEfuseRecord(i).Element(0) = 3
                ElseIf VoltageLevelCNT(i) = UBound(SensorVoltage) - 3 Then
                    DicEfuseRecord(i).Element(0) = 2
                ElseIf VoltageLevelCNT(i) = UBound(SensorVoltage) - 4 Then
                    DicEfuseRecord(i).Element(0) = 1
                End If
            Else
                If VoltageLevelCNT(i) = UBound(SensorVoltage) Then
                    DicEfuseRecord(i).Element(0) = 13
                ElseIf VoltageLevelCNT(i) = UBound(SensorVoltage) - 1 Then
                    DicEfuseRecord(i).Element(0) = 12
                ElseIf VoltageLevelCNT(i) = UBound(SensorVoltage) - 2 Then
                    DicEfuseRecord(i).Element(0) = 11
                ElseIf VoltageLevelCNT(i) = UBound(SensorVoltage) - 3 Then
                    DicEfuseRecord(i).Element(0) = 10
                ElseIf VoltageLevelCNT(i) = UBound(SensorVoltage) - 4 Then
                    DicEfuseRecord(i).Element(0) = 9
                End If
            End If
        Next site
        Call AddStoredCaptureData(DicEfuseName, DicEfuseRecord(i))
    Next i
    
'===================================================================================================================================================='
'DataProcess for Matrix calculrate
'===================================================================================================================================================='
    Dim SenName As String
    Dim MatrixMaxValue As New SiteLong
    Dim MatrixMinValue As New SiteLong
    Dim ThisPinType As String
    Dim SetInformation() As New DSPWave
    Dim Aininformation() As New DSPWave
    Dim Aixinformation() As New DSPWave
    Dim PiUInformation() As New DSPWave
    ReDim SetInformation(UBound(SensorSplitStr))
    ReDim Aininformation(UBound(SensorSplitStr))
    ReDim Aixinformation(UBound(SensorSplitStr))
    ReDim PiUInformation(UBound(SensorSplitStr))
    Dim ProcessLock As Boolean
    Dim RuleMonitor() As New SiteLong
'    Dim MaxEndPointMonitor_First() As New SiteLong
'    Dim MaxEndPointMonitor_Second() As New SiteLong
'    Dim MinEndPointMonitor_First() As New SiteLong
'    Dim MinEndPointMonitor_Second() As New SiteLong
    Dim FinalCheck As New SiteLong
    Dim Coeff_DspWave_a1 As New DSPWave
    Dim Coeff_DspWave_a2 As New DSPWave
    Dim SensorTempName_rot As String
    Dim SensorTempName_rov As String
    Dim sizeOfFuse_1 As Long
    sizeOfFuse_1 = CLng(FuseSize_1)
    Dim OutFuseDspWave_1 As New DSPWave
    OutFuseDspWave_1.CreateConstant 0, sizeOfFuse_1, DspLong
    Dim sizeOfFuse_2 As Long
    sizeOfFuse_2 = CLng(FuseSize_2)
    Dim OutFuseDspWave_2 As New DSPWave
    OutFuseDspWave_2.CreateConstant 0, sizeOfFuse_2, DspLong
    Dim DicCounterbyMain As Long
    Dim DicCounterbyMain_temp As Long
    Dim SensorGroup_main() As String
    Dim SensorArray_main() As String
    Dim StoreFuseDicNameTemp_1 As String
    Dim StoreFuseDicNameTemp_2 As String
    DicCounterbyMain = -1
    DicCounterbyMain_temp = 0
    SensorGroup_main = Split(FreqSensorAlignment, ";")
        
    For i = 0 To UBound(SensorGroup_main)
        SensorArray_main = Split(SensorGroup_main(i), ",")
        For j = 0 To UBound(SensorArray_main)
            DicCounterbyMain = DicCounterbyMain + 1
        Next j
    Next i
    DicCounterbyMain = (DicCounterbyMain + 1) / 2
        
    For i = 0 To UBound(SensorGroup_main)
        SensorArray_main = Split(SensorGroup_main(i), ",")
        ReDim RuleMonitor(UBound(SensorArray_main))
'        ReDim MaxEndPointMonitor_First(UBound(SensorArray_main))
'        ReDim MinEndPointMonitor_First(UBound(SensorArray_main))
'        ReDim MaxEndPointMonitor_Second(UBound(SensorArray_main))
'        ReDim MinEndPointMonitor_Second(UBound(SensorArray_main))
        For j = 0 To UBound(SensorSplitStr)
            For k = 0 To UBound(SensorArray_main)
'''''                If (k + 1) Mod 2 = 0 Then
'''''                    If CStr(SensorSplitStr(j)) = "VDD_GPU" Then
'''''                        If SensorArray_main(k) Like "*GPU*" Then
'''''                            ProcessLock = True
'''''                        Else
'''''                            ProcessLock = False
'''''                        End If
'''''                    ElseIf CStr(SensorSplitStr(j)) = "VDD_ECPU" Then
'''''                        If SensorArray_main(k) Like "*ECPU*" Then
'''''                            ProcessLock = True
'''''                        Else
'''''                            ProcessLock = False
'''''                        End If
'''''                    ElseIf CStr(SensorSplitStr(j)) = "VDD_PCPU" Then
'''''                        If SensorArray_main(k) Like "*ANE*" Or SensorArray_main(k) Like "*PCPU*" Then
'''''                            ProcessLock = True
'''''                        Else
'''''                            ProcessLock = False
'''''                        End If
'''''                    ElseIf CStr(SensorSplitStr(j)) = "VDD_AVE" Then
'''''                        If SensorArray_main(k) Like "*AVE*" Then
'''''                            ProcessLock = True
'''''                        Else
'''''                            ProcessLock = False
'''''                        End If
'''''                    ElseIf CStr(SensorSplitStr(j)) = "VDD_SOC" Then
'''''                        If SensorArray_main(k) Like "*SOC*" Then
'''''                            ProcessLock = True
'''''                        Else
'''''                            ProcessLock = False
'''''                        End If
'''''                    Else
'''''                        ProcessLock = False
'''''                    End If
                
'''''                    If ProcessLock = True Then
'''''                        For Each site In TheExec.sites
'''''                            SetInformation(j).CreateConstant 0, 300, DspLong
'''''                            Aininformation(j).CreateConstant 0, 300, DspLong
'''''                            Aixinformation(j).CreateConstant 0, 300, DspLong
'''''                            PiUInformation(j).CreateConstant 0, 300, DspLong
'''''                            If MTR_CAL_Sheet Like "*1150*" Then
'''''                                MatrixMaxValue = 1150
'''''                            Else
'''''                                MatrixMaxValue = 950
'''''                            End If
'''''                            If VoltageLevelCNT(j) = UBound(SensorVoltage) Then
'''''                                ThisPinType = MTR_CAL_Sheet + "450"
'''''                                MatrixMinValue = 450
'''''                            ElseIf VoltageLevelCNT(j) = UBound(SensorVoltage) - 1 Then
'''''                                ThisPinType = MTR_CAL_Sheet + "475"
'''''                                MatrixMinValue = 475
'''''                            ElseIf VoltageLevelCNT(j) = UBound(SensorVoltage) - 2 Then
'''''                                ThisPinType = MTR_CAL_Sheet + "500"
'''''                                MatrixMinValue = 500
'''''                            ElseIf VoltageLevelCNT(j) = UBound(SensorVoltage) - 3 Then
'''''                                ThisPinType = MTR_CAL_Sheet + "525"
'''''                                MatrixMinValue = 525
'''''                            ElseIf VoltageLevelCNT(j) = UBound(SensorVoltage) - 4 Then
'''''                                ThisPinType = MTR_CAL_Sheet + "550"
'''''                                MatrixMinValue = 550
'''''                            End If
'''''                            SetInformation(j)(site) = Public_GetStoredCaptureData(ThisPinType + "_" + "size")
'''''                            Aixinformation(j)(site) = Public_GetStoredCaptureData(ThisPinType + "_" + "Aix")
'''''                            Aininformation(j)(site) = Public_GetStoredCaptureData(ThisPinType + "_" + "Ain")
'''''                            PiUInformation(j)(site) = Public_GetStoredCaptureData(ThisPinType + "_" + "PiUInfo")
'''''                            SenName = Left(SensorArray_main(k), (InStr(SensorArray_main(k), "_") - 1))
''''''                            PrintStr = "Site" + CStr(site) + "_" + CStr(SensorSplitStr(j)) + "_" + SenName + "_" + ThisPinType
''''''                            TheExec.Datalog.WriteComment (PrintStr)
'''''                        Next site
'''''
'''''                        TheExec.Flow.TestLimit resultVal:=MatrixMinValue, TName:=SenName + "_" + "Matrix_StartPoint", ForceResults:=tlForceFlow
'''''                        TheExec.Flow.TestLimit resultVal:=MatrixMaxValue, TName:=SenName + "_" + "Matrix_EndPoint", ForceResults:=tlForceFlow
'''''
''''''    Call MTR_LinearRegression(SensorArray_main(k - 1), SensorArray_main(k), SweepVArrayValue, SweepVArrayDic, Temperature, VoltageLevelCNT(j))
'''''
'''''                        Dim OutFuseDspWave_1_temp() As New DSPWave
'''''                        Dim OutFuseDspWave_2_temp() As New DSPWave
'''''                        ReDim OutFuseDspWave_1_temp(DicCounterbyMain)
'''''                        ReDim OutFuseDspWave_2_temp(DicCounterbyMain)
'''''
'''''                        SensorTempName_rot = SensorArray_main(k - 1) + "_" + Temperature + "C"
'''''                        SensorTempName_rov = SensorArray_main(k) + "_" + Temperature + "C"
'''''
'''''                        Call Calc_FromLoad_MTR_SE_CAL_Coeff(SensorTempName_rot, SensorTempName_rov, Temperature, sizeOfFuse_1, sizeOfFuse_2, SetInformation(j), PiUInformation(j), _
'''''                        Aixinformation(j), Aininformation(j), OutFuseDspWave_1, OutFuseDspWave_2, CLng(UBound(SensorVoltage)), VoltageLevelCNT(j), RuleMonitor(k - 1), RuleMonitor(k))
'''''
'''''                        StoreFuseDicNameTemp_1 = StoreFuseDicName_1
'''''                        StoreFuseDicNameTemp_2 = StoreFuseDicName_2
'''''
'''''                        OutFuseDspWave_1_temp(DicCounterbyMain_temp) = OutFuseDspWave_1
'''''                        OutFuseDspWave_2_temp(DicCounterbyMain_temp) = OutFuseDspWave_2
'''''
'''''                        StoreFuseDicNameTemp_1 = StoreFuseDicNameTemp_1 + "_" + SensorTempName_rot
'''''                        Call AddStoredCaptureData(StoreFuseDicNameTemp_1, OutFuseDspWave_1_temp(DicCounterbyMain_temp))
'''''                        StoreFuseDicNameTemp_2 = StoreFuseDicNameTemp_2 + "_" + SensorTempName_rov
'''''                        Call AddStoredCaptureData(StoreFuseDicNameTemp_2, OutFuseDspWave_2_temp(DicCounterbyMain_temp))
'''''                        DicCounterbyMain_temp = DicCounterbyMain_temp + 1
'''''                    End If
'''''                End If
            Next k
        Next j
    Next i
'''''
'''''    For i = 0 To UBound(SensorArray_main)
'''''        TheExec.Flow.TestLimit resultVal:=RuleMonitor(i), ForceResults:=tlForceFlow
'''''    Next i
End Function



Public Function MTRG_t5p2a_DigSrc_Coefficient_PreCalc(Optional v0 As String, Optional v1 As String, Optional x0a As String, Optional x1a As String, Optional c0_DictName As String, Optional c1_DictName As String, Optional c2_DictName As String, Optional c3_DictName As String, Optional Validating_ As Boolean) As Long
    If Validating_ Then
        Exit Function    ' Exit after validation
    End If
     Dim site As Variant
     
    On Error GoTo errHandler
    
    Dim DSPWave_v0 As New DSPWave
    DSPWave_v0.CreateConstant 0, 1
    Dim DSPWave_v1 As New DSPWave
    DSPWave_v1.CreateConstant 0, 1
    Dim DSPWave_x0a As New DSPWave
    DSPWave_x0a.CreateConstant 0, 1
    Dim DSPWave_x1a As New DSPWave
    DSPWave_x1a.CreateConstant 0, 1
    
    Dim DSPWave_Binary_v0 As New DSPWave
    DSPWave_v0.CreateConstant 0, 18
    Dim DSPWave_Binary_v1 As New DSPWave
    DSPWave_v1.CreateConstant 0, 18
    Dim DSPWave_Binary_x0a As New DSPWave
    DSPWave_x0a.CreateConstant 0, 18
    Dim DSPWave_Binary_x1a As New DSPWave
    DSPWave_x1a.CreateConstant 0, 18
    
    Dim DSPWave_x0 As New DSPWave
    DSPWave_x0.CreateConstant 0, 1
    Dim DSPWave_x1 As New DSPWave
    DSPWave_x1.CreateConstant 0, 1
    Dim DSPWave_y0 As New DSPWave
    DSPWave_y0.CreateConstant 0, 1
    Dim DSPWave_y1 As New DSPWave
    DSPWave_y1.CreateConstant 0, 1
    
    Dim DSPWave_a0cal As New DSPWave
    DSPWave_a0cal.CreateConstant 0, 1
    Dim DSPWave_a1cal As New DSPWave
    DSPWave_a1cal.CreateConstant 0, 1
    
    Dim DSPWave_c0 As New DSPWave
    DSPWave_c0.CreateConstant 0, 1
    Dim DSPWave_c1 As New DSPWave
    DSPWave_c1.CreateConstant 0, 1
    Dim DSPWave_c2 As New DSPWave
    DSPWave_c2.CreateConstant 0, 1
    Dim DSPWave_c3 As New DSPWave
    DSPWave_c3.CreateConstant 0, 1
    
    Dim DSPWave_Binary_c0 As New DSPWave
    DSPWave_Binary_c0.CreateConstant 0, 32
    Dim DSPWave_Binary_c1 As New DSPWave
    DSPWave_Binary_c1.CreateConstant 0, 32
    Dim DSPWave_Binary_c2 As New DSPWave
    DSPWave_Binary_c2.CreateConstant 0, 32
    Dim DSPWave_Binary_c3 As New DSPWave
    DSPWave_Binary_c3.CreateConstant 0, 32
    Dim C_BitWidth As New SiteLong
    For Each site In TheExec.sites
        C_BitWidth(site) = DSPWave_Binary_c0(site).SampleSize
    Next site
    
    Dim A0 As Double: A0 = -21.5822184999726
    Dim a1 As Double: a1 = 428.009226609628
    Dim a2 As Double: a2 = -133.454310922823
    Dim a3 As Double: a3 = 19.0485545665615
    
    
    DSPWave_Binary_v0 = GetStoredCaptureData(v0)
    DSPWave_Binary_v1 = GetStoredCaptureData(v1)
    DSPWave_Binary_x0a = GetStoredCaptureData(x0a)
    DSPWave_Binary_x1a = GetStoredCaptureData(x1a)     '25C and 85C have same dictionary name.
    
    Dim i As Integer
    If currentJobName Like "cp1" Then   'work around 25C and 85C have same dictionary name.
        For Each site In TheExec.sites
            For i = 0 To DSPWave_Binary_x1a(site).SampleSize - 1
                DSPWave_Binary_x1a(site).Element(i) = 0
            Next i
        Next site
    End If
    
    Dim SL_BitWidth As New SiteLong
    For Each site In TheExec.sites
        SL_BitWidth(site) = DSPWave_Binary_v0(site).SampleSize
    Next site
    
    Call rundsp.DSP_2S_Complement_To_SignDec(DSPWave_Binary_v0, SL_BitWidth, DSPWave_v0)
    Call rundsp.DSP_DivideConstant(DSPWave_v0, 2 ^ 13)
    Call rundsp.BinToDec(DSPWave_Binary_v1, DSPWave_v1)
    Call rundsp.DSP_DivideConstant(DSPWave_v1, 2 ^ 17)
    Call rundsp.BinToDec(DSPWave_Binary_x0a, DSPWave_x0a)
    Call rundsp.DSP_DivideConstant(DSPWave_x0a, 2 ^ 13)
    Call rundsp.BinToDec(DSPWave_Binary_x1a, DSPWave_x1a)
    Call rundsp.DSP_DivideConstant(DSPWave_x1a, 2 ^ 13)
    
'    DSPWave_x0 = DSPWave_x0a
'    Call rundsp.DSP_Subtract(DSPWave_x0, DSPWave_v0)
'    Call rundsp.DSP_Divide(DSPWave_x0, DSPWave_v1)
'
'    DSPWave_x1 = DSPWave_x1a
'    Call rundsp.DSP_Subtract(DSPWave_x1, DSPWave_v0)
'    Call rundsp.DSP_Divide(DSPWave_x1, DSPWave_v1)
    
    For Each site In TheExec.sites
        If DSPWave_v1(site).Element(0) = 0 Then DSPWave_v1(site).Element(0) = 0.0000000001
        DSPWave_x0(site).Element(0) = (DSPWave_x0a(site).Element(0) - DSPWave_v0(site).Element(0)) / DSPWave_v1(site).Element(0)
        DSPWave_x1(site).Element(0) = (DSPWave_x1a(site).Element(0) - DSPWave_v0(site).Element(0)) / DSPWave_v1(site).Element(0)
        DSPWave_y0(site).Element(0) = 25 + 273.15 - a2 * DSPWave_x0(site).Element(0) ^ 2 - a3 * DSPWave_x0(site).Element(0) ^ 3
        DSPWave_y1(site).Element(0) = 85 + 273.15 - a2 * DSPWave_x1(site).Element(0) ^ 2 - a3 * DSPWave_x1(site).Element(0) ^ 3
        If DSPWave_x1(site).Element(0) = DSPWave_x0(site).Element(0) Then DSPWave_x1(site).Element(0) = DSPWave_x1(site).Element(0) + 0.0000000001
        DSPWave_a0cal(site).Element(0) = (DSPWave_x1(site).Element(0) * DSPWave_y0(site).Element(0) - DSPWave_x0(site).Element(0) * DSPWave_y1(site).Element(0)) / (DSPWave_x1(site).Element(0) - DSPWave_x0(site).Element(0))
        DSPWave_a1cal(site).Element(0) = (DSPWave_y1(site).Element(0) - DSPWave_y0(site).Element(0)) / (DSPWave_x1(site).Element(0) - DSPWave_x0(site).Element(0)) / (DSPWave_x1(site).Element(0) - DSPWave_x0(site).Element(0))
        DSPWave_c0(site).Element(0) = DSPWave_a0cal(site).Element(0) - DSPWave_a1cal(site).Element(0) * DSPWave_v0(site).Element(0) / DSPWave_v1(site).Element(0)
        DSPWave_c1(site).Element(0) = DSPWave_a1cal(site).Element(0) / DSPWave_v1(site).Element(0) - 2 * a2 * DSPWave_v0(site).Element(0) / DSPWave_v1(site).Element(0) ^ 2
        DSPWave_c2(site).Element(0) = a2 / DSPWave_v1(site).Element(0) ^ 2 - 3 * a3 * DSPWave_v0(site).Element(0) / DSPWave_v1(site).Element(0) ^ 3
        DSPWave_c3(site).Element(0) = a3 / DSPWave_v1(site).Element(0) ^ 3
        DSPWave_c0(site).Element(0) = FormatNumber(DSPWave_c0(site).Element(0) * 2 ^ 13, 0)
        DSPWave_c1(site).Element(0) = FormatNumber(DSPWave_c1(site).Element(0) * 2 ^ 13, 0)
        DSPWave_c2(site).Element(0) = FormatNumber(DSPWave_c2(site).Element(0) * 2 ^ 13, 0)
        DSPWave_c3(site).Element(0) = FormatNumber(DSPWave_c3(site).Element(0) * 2 ^ 13, 0)
    Next site
    
    TheExec.Flow.TestLimit resultVal:=DSPWave_x0.Element(0), Tname:="x0", ForceResults:=tlForceNone 'eng_forceflow_transfer
    TheExec.Flow.TestLimit resultVal:=DSPWave_x1.Element(0), Tname:="x1", ForceResults:=tlForceNone 'eng_forceflow_transfer
    TheExec.Flow.TestLimit resultVal:=DSPWave_y0.Element(0), Tname:="y0", ForceResults:=tlForceNone 'eng_forceflow_transfer
    TheExec.Flow.TestLimit resultVal:=DSPWave_y1.Element(0), Tname:="y1", ForceResults:=tlForceNone 'eng_forceflow_transfer
    TheExec.Flow.TestLimit resultVal:=DSPWave_c0.Element(0), Tname:="c0", ForceResults:=tlForceNone 'eng_forceflow_transfer
    TheExec.Flow.TestLimit resultVal:=DSPWave_c1.Element(0), Tname:="c1", ForceResults:=tlForceNone 'eng_forceflow_transfer
    TheExec.Flow.TestLimit resultVal:=DSPWave_c2.Element(0), Tname:="c2", ForceResults:=tlForceNone 'eng_forceflow_transfer
    TheExec.Flow.TestLimit resultVal:=DSPWave_c3.Element(0), Tname:="c3", ForceResults:=tlForceNone 'eng_forceflow_transfer
    
    'Call rundsp.DSP_Convert_2S_Complement(DSPWave_c0, C_BitWidth, DSPWave_Binary_c0)
    Call rundsp.DSPWf_Dec2Binary(DSPWave_c0, C_BitWidth, DSPWave_Binary_c0)
    Call rundsp.DSPWf_Dec2Binary(DSPWave_c1, C_BitWidth, DSPWave_Binary_c1)
    Call rundsp.DSPWf_Dec2Binary(DSPWave_c2, C_BitWidth, DSPWave_Binary_c2)
    Call rundsp.DSPWf_Dec2Binary(DSPWave_c3, C_BitWidth, DSPWave_Binary_c3)
    
    Call AddStoredCaptureData(c0_DictName, DSPWave_Binary_c0)
    Call AddStoredCaptureData(c1_DictName, DSPWave_Binary_c1)
    Call AddStoredCaptureData(c2_DictName, DSPWave_Binary_c2)
    Call AddStoredCaptureData(c3_DictName, DSPWave_Binary_c3)
    
    Exit Function
errHandler:
    TheExec.Datalog.WriteComment "Error in MTRG_t5p2a_PreCalculation"
    If AbortTest Then Exit Function Else Resume Next
End Function


Public Function MTRG_t6p3abc_DigSrc_Coefficient_PreCalc(Optional v0 As String, Optional v1 As String, Optional c0_DictName As String, Optional c1_DictName As String, Optional Validating_ As Boolean) As Long
    If Validating_ Then
        Exit Function    ' Exit after validation
    End If
    On Error GoTo errHandler
    Dim site As Variant
    
    Dim DSPWave_v0 As New DSPWave
    DSPWave_v0.CreateConstant 0, 1
    Dim DSPWave_v1 As New DSPWave
    DSPWave_v1.CreateConstant 0, 1
    
    Dim DSPWave_Binary_v0 As New DSPWave
    DSPWave_v0.CreateConstant 0, 18
    Dim DSPWave_Binary_v1 As New DSPWave
    DSPWave_v1.CreateConstant 0, 18
    
    Dim DSPWave_c0 As New DSPWave
    DSPWave_c0.CreateConstant 0, 1
    Dim DSPWave_c1 As New DSPWave
    DSPWave_c1.CreateConstant 0, 1
    
    Dim DSPWave_Binary_c0 As New DSPWave
    DSPWave_Binary_c0.CreateConstant 0, 32
    Dim DSPWave_Binary_c1 As New DSPWave
    DSPWave_Binary_c1.CreateConstant 0, 32
    
    Dim Vref As Double: Vref = 10
    
    DSPWave_Binary_v0 = GetStoredCaptureData(v0)
    DSPWave_Binary_v1 = GetStoredCaptureData(v1)
    
    Dim SL_BitWidth As New SiteLong
    For Each site In TheExec.sites
        SL_BitWidth(site) = DSPWave_Binary_v0(site).SampleSize
    Next site
    Dim C_BitWidth As New SiteLong
    For Each site In TheExec.sites
        C_BitWidth(site) = DSPWave_Binary_c0(site).SampleSize
    Next site
    
    Call rundsp.DSP_2S_Complement_To_SignDec(DSPWave_Binary_v0, SL_BitWidth, DSPWave_v0)
    Call rundsp.DSP_DivideConstant(DSPWave_v0, 2 ^ 13)
    Call rundsp.BinToDec(DSPWave_Binary_v1, DSPWave_v1)
    Call rundsp.DSP_DivideConstant(DSPWave_v1, 2 ^ 17)
    
    For Each site In TheExec.sites
        If DSPWave_v1(site).Element(0) = 0 Then DSPWave_v1(site).Element(0) = 0.0000000001
        DSPWave_c0(site).Element(0) = 273.15 - Vref * DSPWave_v0(site).Element(0) / DSPWave_v1(site).Element(0)
        DSPWave_c1(site).Element(0) = Vref / DSPWave_v1(site).Element(0)
        DSPWave_c0(site).Element(0) = FormatNumber(DSPWave_c0(site).Element(0) * 2 ^ 13, 0)
        DSPWave_c1(site).Element(0) = FormatNumber(DSPWave_c1(site).Element(0) * 2 ^ 13, 0)
    Next site
    
    TheExec.Flow.TestLimit resultVal:=DSPWave_c0.Element(0), Tname:="c0", ForceResults:=tlForceNone 'eng_forceflow_transfer
    TheExec.Flow.TestLimit resultVal:=DSPWave_c1.Element(0), Tname:="c1", ForceResults:=tlForceNone 'eng_forceflow_transfer
    
    Call rundsp.DSPWf_Dec2Binary(DSPWave_c0, C_BitWidth, DSPWave_Binary_c0)
    Call rundsp.DSPWf_Dec2Binary(DSPWave_c1, C_BitWidth, DSPWave_Binary_c1)
    
    Call AddStoredCaptureData(c0_DictName, DSPWave_Binary_c0)
    Call AddStoredCaptureData(c1_DictName, DSPWave_Binary_c1)
    
    Exit Function
errHandler:
    TheExec.Datalog.WriteComment "Error in MTRG_t5p2a_PreCalculation"
    If AbortTest Then Exit Function Else Resume Next
End Function

Public Function DSSC_Search(Optional Pat As String, Optional MeasureV_pin As PinList, Optional DigSrc_pin As PinList, _
    Optional TrimTarget As Double, Optional TrimStart As Long, Optional TrimFormat As String, Optional TrimRepeat As Long, _
    Optional TrimStoreName As String, Optional TrimFuseName As String, Optional TrimFuseTypeName As String, Optional Validating_ As Boolean)
''    Dim PatCount As Long
    Dim i As Integer
    Dim pats() As String
    Dim code As New SiteLong
    Dim vout As New SiteDouble
    Dim BestCode As New SiteLong, BestVal As New SiteDouble, verr As New SiteDouble, temp As New SiteLong
    Dim First As New SiteBoolean, Done As New SiteBoolean
    Dim trace As Boolean
    Dim site As Variant
    
    Dim PatCount As Long, PattArray() As String
    
    If Validating_ Then
        Call PrLoadPattern(Pat)
        Exit Function    ' Exit after validation
    End If
    
    Call HardIP_InitialSetupForPatgen
    
    On Error GoTo errHandler
    TheHdw.Digital.ApplyLevelsTiming True, True, True, tlPowered
    
    PATT_GetPatListFromPatternSet Pat, pats, PatCount
    If glb_TesterType = "Jaguar" Then
        With TheHdw.DCVI.Pins(MeasureV_pin)
            .Gate = False
            .Disconnect tlDCVIConnectDefault
            .mode = tlDCVIModeHighImpedance
            .Connect tlDCVIConnectHighSense
            .Voltage = 6
            .Current = 0
             TheHdw.Wait 0.5 * ms
            .Gate = True
        End With
    End If
    
    '20210416,Add for ufp
    Dim TempPat As New Pattern
    TempPat.value = Pat
    '---------------------------------------UFP_Corr fix 20200413
    Call ProcessInputToGLB(patset:=TempPat, TestSequence:="V", MeasV_PinS:=MeasureV_pin.value)
    '---------------------------------------UFP_Corr fix 20200413
    code = TrimStart
    
    Dim SplitByEqual() As String, SplitByColon() As String, TrimCodeSize As Long
    Dim TrimCodeValue_Min As Long, TrimCodeValue_Mid As Long, TrimCodeValue_Max As Long
    SplitByEqual = Split(TrimFormat, "=")
    SplitByColon = Split(SplitByEqual(1), ":")
    TrimCodeSize = SplitByColon(0) + 1
    TrimCodeValue_Min = 0
    TrimCodeValue_Mid = (2 ^ TrimCodeSize) / 2
    TrimCodeValue_Max = 2 ^ TrimCodeSize - 1
    
    Call DSSC_Search_par_run(pats(0), DigSrc_pin.value, code, MeasureV_pin.value, vout, TrimCodeSize, TrimRepeat)
    If trace Then TheExec.Flow.TestLimit code
    If trace Then TheExec.Flow.TestLimit vout
    First = vout.compare(LessThan, TrimTarget)
    If trace Then TheExec.Flow.TestLimit First, , , , , , , , "first"
    BestCode = code
    BestVal = vout
    verr = vout.Subtract(TrimTarget).Abs

    For i = 0 To TrimCodeValue_Mid - 1
        If trace Then TheExec.Datalog.WriteComment ("i = " & i)
        temp = vout.compare(LessThan, TrimTarget)
        code = code.Add(temp.Multiply(-2).Subtract(1))  ' If vout < TrimTarget Then code++ Else code--
        
        For Each site In TheExec.sites.Active
            If code(site) > TrimCodeValue_Max Then: code(site) = code(site) - 1: ''GoTo EndTrim
            If code(site) < TrimCodeValue_Min Then: code(site) = code(site) + 1: ''GoTo EndTrim
        Next site
      
        Call DSSC_Search_par_run(pats(0), DigSrc_pin.value, code, MeasureV_pin.value, vout, TrimCodeSize, TrimRepeat)
        If trace Then TheExec.Flow.TestLimit code
        If trace Then TheExec.Flow.TestLimit vout
        For Each site In TheExec.sites
            If Abs(vout - TrimTarget) < verr Then
                BestCode = code
                BestVal = vout
                verr = Abs(vout - TrimTarget)
            End If
        Next site

        Done = Done.LogicalOr(First.LogicalXor(vout.compare(LessThan, TrimTarget)))
        If trace Then TheExec.Flow.TestLimit Done, , , , , , , , "done"
        If Done.All(True) Then Exit For
    Next i
    
EndTrim:
    TheExec.Flow.TestLimit BestVal, , , , , , unitVolt, , "Volt_meas_LPDPRX_LDO", , MeasureV_pin, , , , , tlForceFlow
    TheExec.Flow.TestLimit BestCode, 0, 15, , , , , , "LPDPRX_LDO_Trim", , , , , , , tlForceFlow
    HardIP_WriteFuncResult
    
    Dim TempVal As Integer
    Dim FinalTrimCode As New DSPWave
    
    FinalTrimCode.CreateConstant 0, TrimCodeSize
    
    For Each site In TheExec.sites
        TempVal = BestCode(site)
        For i = 0 To TrimCodeSize - 1
            FinalTrimCode(site).Element(i) = TempVal Mod 2
            TempVal = TempVal \ 2
        Next i
    Next site
    
    If TrimStoreName <> "" Then
        Call Checker_StoreDigCapAllToDictionary(TrimStoreName, FinalTrimCode)
    End If
    
    '' 20170704 - Add write efuse function
''    Dim sl_Fuse_Val As SiteLong
    If TheExec.TesterMode = testModeOffline Then
    Else
''        For Each Site In TheExec.sites
''            sl_Fuse_Val(Site) = BestCode(Site)
''        Next Site
        
        If TrimFuseName <> "" And TrimFuseTypeName <> "" Then
            ''Call HIP_eFuse_Write(TrimFuseTypeName, TrimFuseName, BestCode) ''set fuse information from flow
        End If
    End If
        
    Exit Function
    
errHandler:
    TheExec.Datalog.WriteComment ("ERROR in DSSC_Search: " & err.Description)
    DSSC_Search = TL_ERROR
End Function




Public Function DSSC_Search_LDO(Optional Pat As Pattern, Optional MeasureV_pin As PinList, Optional MeasV_Name As String, Optional MeasV_Name_Trim As String, Optional MeaV_WaitTime As String, Optional DigSrc_pin As PinList, Optional DigSrc_Sample_Size As Long, Optional DigSrc_Equation As String, Optional digsrc_assignment As String, Optional TrimStoreName As String, Optional TrimTarget As Double, Optional TrimStart As Long, Optional TrimCodeSize As Long, Optional TrimMethod As String, Optional TrimStepSize As Double, Optional Validating_ As Boolean)
    Dim site As Variant
    Dim i As Integer
    Dim pats() As String
    Dim code As New SiteLong: code = TrimStart
    Dim vout As New SiteDouble
    Dim NumberOfMeasV As Integer: NumberOfMeasV = UBound(Split(MeasV_Name, "+")) + 1
    Dim BestCode As New SiteLong, BestVal() As New SiteDouble, temp As New SiteLong
    ReDim BestVal(NumberOfMeasV - 1) As New SiteDouble
    Dim First As New SiteBoolean, Done As New SiteBoolean
    Dim PreviousNegative As New SiteBoolean
    Dim PreviousPositive As New SiteBoolean
    Dim step As New SiteBoolean
    Dim DecideTrim As New SiteBoolean
        For Each site In TheExec.sites.Active
            PreviousNegative = False
            PreviousPositive = False
            DecideTrim = False
        Next site

    glb_TestInstance = vbNullString
    glb_TestInstance = UCase(TheExec.DataManager.instancename)
    
    Dim BlockName() As String: BlockName = Split(glb_TestInstance, "_")
    Dim MeasValue() As New SiteDouble: ReDim MeasValue(NumberOfMeasV - 1)
    Dim PreviousMeasValue() As New SiteDouble: ReDim PreviousMeasValue(NumberOfMeasV - 1) As New SiteDouble
    Dim StepCount As Long: StepCount = 0
    Dim OutputTname_format() As String
    Dim TestNameInput As String
    Dim PatCount As Long, PattArray() As String
    Dim MeasV_Name_Array() As String: MeasV_Name_Array = Split(MeasV_Name, "+")
    Dim MeasV_Name_Trim_Array() As String: MeasV_Name_Trim_Array = Split(MeasV_Name_Trim, "+")
    Dim TrimPoint() As Long: ReDim TrimPoint(UBound(Split(MeasV_Name, "+")))
    
    Call ProcessInputToGLB(Pat, "V", False, , , , , MeasureV_pin.value, , , , , , , , , , , , , , , , , , , , , , , , , , , , , , , , , , , , , , , , , , , , , MeaV_WaitTime)
     
    For i = 0 To NumberOfMeasV - 1
        For Each site In TheExec.sites.Active
            PreviousMeasValue(i).value = 0
            MeasValue(i).value = 0
        Next site
        If MeasV_Name_Array(i) = MeasV_Name_Trim_Array(i) Then: TrimPoint(i) = 1
    Next i
    
    Call GetFlowTName

    If Validating_ Then
        Call PrLoadPattern(Pat.value)
        Exit Function    ' Exit after validation
    End If
    
    On Error GoTo errHandler
    
    If TheExec.DevChar.Setups.IsRunning Then
        If TheExec.DevChar.Setups(TheExec.DevChar.Setups.ActiveSetupName).Shmoo.Axes.Contains(tlDevCharShmooAxis_Y) Then
            If gl_Flag_HardIP_Trim_Set_PrePoint And Not (gl_Flag_HardIP_Characterization_1stRun) Then
                TheHdw.Digital.ApplyLevelsTiming True, True, True, tlPowered
                Call TheExec.Overlays.ApplyUniformSpecToHW("XI0_0_Shmoo_Freq_VAR", TheExec.DevChar.Results(TheExec.DevChar.Setups.ActiveSetupName).Shmoo.CurrentPoint.Axes(tlDevCharShmooAxis_Y).value)
                Call TheExec.Overlays.ApplyUniformSpecToHW("XI0_1_Shmoo_Freq_VAR", TheExec.DevChar.Results(TheExec.DevChar.Setups.ActiveSetupName).Shmoo.CurrentPoint.Axes(tlDevCharShmooAxis_Y).value)
                Call TheExec.Overlays.ApplyUniformSpecToHW("PCIE_XI0_0_Shmoo_Freq_VAR", TheExec.DevChar.Results(TheExec.DevChar.Setups.ActiveSetupName).Shmoo.CurrentPoint.Axes(tlDevCharShmooAxis_Y).value)
            ElseIf gl_Flag_HardIP_Trim_Set_PostPoint Then
                TheHdw.Digital.ApplyLevelsTiming True, True, True, tlPowered
                Call TheExec.Overlays.ApplyUniformSpecToHW("XI0_0_Shmoo_Freq_VAR", 24000000#)
                Call TheExec.Overlays.ApplyUniformSpecToHW("XI0_1_Shmoo_Freq_VAR", 24000000#)
                Call TheExec.Overlays.ApplyUniformSpecToHW("PCIE_XI0_0_Shmoo_Freq_VAR", 24000000#)
            Else
                TheHdw.Digital.ApplyLevelsTiming True, True, True, tlPowered
            End If
        End If
    Else
        TheHdw.Digital.ApplyLevelsTiming True, True, True, tlPowered
    End If
    TheHdw.Digital.Patgen.Continue 0, cpuA + cpuB + cpuC + cpuD
    
    PATT_GetPatListFromPatternSet Pat.value, pats, PatCount
        
    Dim TrimCodeValue_Min As Long, TrimCodeValue_Mid As Long, TrimCodeValue_Max As Long
    TrimCodeValue_Min = 0
    TrimCodeValue_Mid = (2 ^ TrimCodeSize) / 2
    TrimCodeValue_Max = 2 ^ TrimCodeSize - 1
  

    If gl_Disable_HIP_debug_log = False Then TheExec.Datalog.WriteComment ("**************** The Measurement at Trim Start Point ****************")
    Call DSSC_Search_par_run_LDO(pats(0), DigSrc_pin, code, MeasureV_pin, vout, TrimCodeSize, NumberOfMeasV, MeasV_Name_Array(), MeasValue(), TrimPoint(), DigSrc_Sample_Size, DigSrc_Equation, digsrc_assignment, TrimStoreName, MeaV_WaitTime)

    First = vout.compare(LessThan, TrimTarget)
    BestCode = code
    For Each site In TheExec.sites.Active
        For i = 0 To NumberOfMeasV - 1
            BestVal(i) = MeasValue(i)
        Next i
    Next site
    
    If TrimMethod = "LinearSearch" Then
        For Each site In TheExec.sites.Active
            If vout.compare(LessThan, TrimTarget) Then
                code = code + 1
            ElseIf vout.compare(GreaterThan, TrimTarget) Then
                code = code - 1
            End If
            step = True
        Next site
    Else
        For Each site In TheExec.sites.Active
            code = code + Fix((TrimTarget - vout) / TrimStepSize)
            If Fix((TrimTarget - vout) / TrimStepSize) <> 0 Then: step = True
        Next site
    End If
    
StartTrim:
        If step.Any(True) Then
            StepCount = StepCount + 1
            
            If gl_Disable_HIP_debug_log = False Then
                If right(CStr(StepCount), 1) = "1" Then
                    TheExec.Datalog.WriteComment ("**************** The " & StepCount & "st Trim Process ****************")
                ElseIf right(CStr(StepCount), 1) = "2" Then
                    TheExec.Datalog.WriteComment ("**************** The " & StepCount & "nd Trim Process ****************")
                ElseIf right(CStr(StepCount), 1) = "3" Then
                    TheExec.Datalog.WriteComment ("**************** The " & StepCount & "rd Trim Process ****************")
                Else
                    TheExec.Datalog.WriteComment ("**************** The " & StepCount & "th Trim Process ****************")
                End If
            End If
            
            
            Call DSSC_Search_par_run_LDO(pats(0), DigSrc_pin, code, MeasureV_pin, vout, TrimCodeSize, NumberOfMeasV, MeasV_Name_Array(), MeasValue(), TrimPoint(), DigSrc_Sample_Size, DigSrc_Equation, digsrc_assignment, TrimStoreName, MeaV_WaitTime)
        End If

        For Each site In TheExec.sites.Active
        If StepCount > (TrimCodeValue_Max - TrimCodeValue_Min) Then
            BestCode = code
            For i = 0 To NumberOfMeasV - 1
                BestVal(i) = MeasValue(i)
            Next i
            DecideTrim = False
        ElseIf code.compare(GreaterThan, TrimCodeValue_Max) Then
            code = TrimCodeValue_Max
            DecideTrim = True

        ElseIf code.compare(LessThan, TrimCodeValue_Min) Then
            code = TrimCodeValue_Min
            DecideTrim = True

        ElseIf code.compare(EqualTo, TrimCodeValue_Max) Or code.compare(EqualTo, TrimCodeValue_Min) Then
            BestCode = code
            For i = 0 To NumberOfMeasV - 1
                BestVal(i) = MeasValue(i)
            Next i
            DecideTrim = False
        ElseIf vout.compare(LessThan, TrimTarget) And PreviousPositive Then
            BestCode = code + 1
            For i = 0 To NumberOfMeasV - 1
                BestVal(i) = PreviousMeasValue(i)
            Next i
            DecideTrim = False
        ElseIf vout.compare(LessThan, TrimTarget) And Not (PreviousPositive) Then
            code(site) = code(site) + 1
            PreviousNegative = True
            For i = 0 To NumberOfMeasV - 1
                PreviousMeasValue(i) = MeasValue(i)
            Next i
            step = True
            DecideTrim = True

        ElseIf vout.compare(GreaterThan, TrimTarget) And PreviousNegative Then
            BestCode = code
            For i = 0 To NumberOfMeasV - 1
                BestVal(i) = MeasValue(i)
            Next i
            DecideTrim = False
        ElseIf vout.compare(GreaterThan, TrimTarget) And Not (PreviousNegative) Then
            code(site) = code(site) - 1
            PreviousPositive = True
            For i = 0 To NumberOfMeasV - 1
                PreviousMeasValue(i) = MeasValue(i)
            Next i
            step = True
            DecideTrim = True

        End If
        Next site
        
    If DecideTrim.Any(True) Then GoTo StartTrim
        
    For i = 0 To NumberOfMeasV - 1
        TestNameInput = Report_TName_From_Instance("V", MeasureV_pin.value, vbNullString, i, 0)
        If Not ByPassTestLimit Then: TheExec.Flow.TestLimit BestVal(i), , , , , , unitVolt, , TestNameInput, , MeasureV_pin, , , , , tlForceFlow
    Next i
    
    TestNameInput = Report_TName_From_Instance("C", vbNullString, , 0, 0)
    If Not ByPassTestLimit Then: TheExec.Flow.TestLimit resultVal:=BestCode, Tname:=TestNameInput, ForceResults:=tlForceFlow
    

    
    Dim TempVal As Integer
    Dim FinalTrimCode As New DSPWave
    
    FinalTrimCode.CreateConstant 0, TrimCodeSize
    
    For Each site In TheExec.sites
        TempVal = BestCode(site)
        For i = 0 To TrimCodeSize - 1
            FinalTrimCode(site).Element(i) = TempVal Mod 2
            TempVal = TempVal \ 2
        Next i
    Next site
    
    If TrimStoreName <> "" Then
        Call AddStoredCaptureData(TrimStoreName, FinalTrimCode)
    End If
    DebugPrintFunc Pat.value
           
    If InStr(UCase(glb_TestInstance), UCase("VREGDT")) <> 0 Or InStr(UCase(glb_TestInstance), UCase("DTVREG")) Then
        Dim FinalTrimCode_DCVREG As New DSPWave
        Dim BestCode_DCVREG As New SiteLong
        FinalTrimCode_DCVREG.CreateConstant 0, 3
        
        
        For Each site In TheExec.sites
            If BestCode(site) = 0 Then BestCode_DCVREG(site) = 0
            If BestCode(site) <= 6 And BestCode(site) >= 1 Then BestCode_DCVREG(site) = 1
            If BestCode(site) <= 12 And BestCode(site) >= 7 Then BestCode_DCVREG(site) = 2
            If BestCode(site) <= 17 And BestCode(site) >= 13 Then BestCode_DCVREG(site) = 3
            If BestCode(site) <= 23 And BestCode(site) >= 18 Then BestCode_DCVREG(site) = 4
            If BestCode(site) <= 29 And BestCode(site) >= 24 Then BestCode_DCVREG(site) = 5
            If BestCode(site) <= 31 And BestCode(site) >= 30 Then BestCode_DCVREG(site) = 6
            
            TempVal = BestCode_DCVREG(site)
            For i = 0 To 2
                FinalTrimCode_DCVREG(site).Element(i) = TempVal Mod 2
                TempVal = TempVal \ 2
            Next i
        Next site
        TestNameInput = Report_TName_From_Instance("C", vbNullString, , 0, 0)
        If Not ByPassTestLimit Then: TheExec.Flow.TestLimit resultVal:=BestCode_DCVREG, Tname:=TestNameInput, ForceResults:=tlForceFlow
        If UCase(glb_TestInstance) Like "CIO*NV" Then: Call AddStoredCaptureData("CIO_LDO_TRIM_VREGDC", FinalTrimCode_DCVREG)
        If UCase(glb_TestInstance) Like "PCIE*NV" Then: Call AddStoredCaptureData("PCIE_LDO_TRIM_VREGDC", FinalTrimCode_DCVREG)
        If UCase(glb_TestInstance) Like "LPDPRX*NV" Then: Call AddStoredCaptureData("LPDPRX_LDO_TRIM_DCVREG", FinalTrimCode_DCVREG)
    End If
    
    Call HardIP_WriteFuncResult(, , glb_TestInstance)
    Exit Function
    
errHandler:
    TheExec.Datalog.WriteComment ("ERROR in DSSC_Search_LDO: " & err.Description)
    DSSC_Search = TL_ERROR
End Function

Public Function TrimUVI80Code_VFI_2sComplement(TwoS_Complement As Boolean, Optional Pat As String, Optional TestSequence As String, Optional MeasV_PinS As String, Optional MeasI_pinS As String, Optional MeasI_Range As Double, _
    Optional MeasF_PinS_SingleEnd As PinList, Optional MeasF_Interval As String, Optional MeasF_EventSourceWithTerminationMode As EventSourceWithTerminationMode, _
    Optional TrimTarget As Double, Optional TrimStart As Long, Optional TrimFormat As String, _
    Optional DigSrc_pin As PinList, Optional DigSrc_DataWidth As Long, Optional DigSrc_Sample_Size As Long, Optional DigSrc_Equation As String, Optional digsrc_assignment As String, _
    Optional DigCap_Pin As PinList, Optional DigCap_DataWidth As Long, Optional DigCap_Sample_Size As Long, Optional CUS_Str_DigCapData As String = vbNullString, _
    Optional TrimStoreName As String, Optional Diffaccuracy As Double, Optional TrimMethod As Long, _
    Optional Meas_StoreName As String, Optional Calc_Eqn As String, Optional TrimCal_Name As String, Optional Antitrim As Boolean = False, Optional Validating_ As Boolean, Optional Interpose_PrePat As String, Optional CPUA_Flag_In_Pat As Boolean = True, Optional Interpose_PostTest As String, _
    Optional Final_Calc As Boolean = False)
''    Dim PatCount As Long
    Dim i As Integer
    Dim pats() As String
    Dim code As New SiteLong
    Dim MeasValue As New SiteDouble
    Dim BestCode As New SiteLong, BestVal As New SiteDouble, verr As New SiteDouble, temp As New SiteLong
    Dim First As New SiteBoolean, Done As New SiteBoolean
    Dim trace As Boolean
    Dim site As Variant
    Dim Ts As Variant
    Dim ADCOUT As New SiteBoolean
    Dim TrimStep As Long
    Dim OutDSP As New DSPWave
    Dim PatCount As Long, PattArray() As String
    Dim TestSequence_array() As String
    Dim doallFlag As Boolean
    Dim finalflag As Boolean
    Dim OutputTname_format() As String
    Dim TName_Ary() As String
    Dim TestNameInput As String
    Dim ReCalc As New SiteDouble
    Dim DoneEndTrim As New SiteBoolean
    gl_TName_Pat = Pat
    
    
    Dim TempPat As New Pattern
    TempPat.value = Pat
    Call ProcessInputToGLB(TempPat, TestSequence, CPUA_Flag_In_Pat, , , , , MeasV_PinS, MeasF_PinS_SingleEnd.value, MeasF_Interval, MeasF_EventSourceWithTerminationMode, , , , MeasI_pinS, CStr(MeasI_Range), , DigCap_Pin, DigCap_DataWidth, DigCap_Sample_Size, DigSrc_pin, CStr(DigSrc_DataWidth) _
                    , CStr(DigSrc_Sample_Size), DigSrc_Equation, digsrc_assignment, , , , , CUS_Str_DigCapData, , , , , , , , , , , , , , , Interpose_PrePat, , Interpose_PostTest)
    
    Call GetFlowTName
    
    If Validating_ Then
        Call PrLoadPattern(Pat)
        Exit Function    ' Exit after validation
    End If
    
    Call HardIP_InitialSetupForPatgen
    
    TestSequence_array = Split(TestSequence, ",")

    glb_TestInstance = vbNullString
    glb_TestInstance = UCase(TheExec.DataManager.instancename)
    
    On Error GoTo errHandler
    TheHdw.Digital.ApplyLevelsTiming True, True, True, tlPowered
    
    PATT_GetPatListFromPatternSet Pat, pats, PatCount
    
    If Interpose_PrePat <> "" Then
        Call SetForceCondition(Interpose_PrePat & ";STOREPREPAT")
    End If
    
    TName_Ary = Split(gl_Tname_Meas, "+")
    
    Dim SplitByEqual() As String, SplitByColon() As String, TrimCodeSize As Long
    Dim TrimCodeValue_Min As Long, TrimCodeValue_Max As Long
    Dim Trimname As String
    SplitByEqual = Split(TrimFormat, "=")
    SplitByColon = Split(SplitByEqual(1), ":")
    Trimname = SplitByEqual(0)
    TrimCodeSize = SplitByColon(0) + 1
    
    If TwoS_Complement = True Then
        TrimCodeValue_Min = -(2 ^ (TrimCodeSize - 1))
        TrimCodeValue_Max = (2 ^ (TrimCodeSize - 1)) - 1
    Else
        TrimCodeValue_Min = 0
        TrimCodeValue_Max = (2 ^ TrimCodeSize) - 1
    End If
    
    Dim binaryFlag As Boolean
    Dim temp_assignment As String
    temp_assignment = digsrc_assignment
     
    '''''''''''''''''''''''''''''''''''Process Trim Method'''''''''''''''''''''''''''''''''''''
    If TrimMethod = 0 Then
        binaryFlag = False
        doallFlag = False
    ElseIf TrimMethod = 2 Then
        binaryFlag = False
        doallFlag = True
        
    Else
        binaryFlag = True
    End If
    
    '''''''''''''''''''''''''''''''''''Linear Search'''''''''''''''''''''''''''''''''''''''''''
    If binaryFlag = False Then
    
        code = TrimStart
        'If doallFlag = True Then
        TrimStep = TrimCodeValue_Max - TrimCodeValue_Min
        'Else
            'TrimStep = TrimCodeValue_Max - TrimCodeValue_Min
        'End If
        For Each site In TheExec.sites
            TheExec.Datalog.WriteComment "site" & site & " decimal code is " & code
        Next site
        
        Call TrimUVI80_Meas_VFI(pats(0), TestSequence_array, DigSrc_pin, code, CStr(MeasV_PinS), MeasValue, MeasI_pinS, MeasI_Range, MeasF_PinS_SingleEnd, MeasF_Interval, MeasF_EventSourceWithTerminationMode, DigSrc_DataWidth, DigSrc_Sample_Size, DigSrc_Equation, digsrc_assignment, DigCap_Pin, DigCap_DataWidth, DigCap_Sample_Size, CUS_Str_DigCapData, OutDSP, TrimCodeSize, Trimname, Meas_StoreName, Calc_Eqn, TrimCal_Name, CPUA_Flag_In_Pat, Final_Calc)
        
        
        If trace Then TheExec.Flow.TestLimit code
        If trace Then TheExec.Flow.TestLimit MeasValue
        First = MeasValue.compare(LessThan, TrimTarget)
        If trace Then TheExec.Flow.TestLimit First, , , , , , , , "first"
        BestCode = code
        BestVal = MeasValue
        verr = MeasValue.Subtract(TrimTarget).Abs
    
        For i = 0 To TrimStep - 1
            'DoneEndTrim = False
            digsrc_assignment = temp_assignment
            If trace Then TheExec.Datalog.WriteComment ("i = " & i)
            If doallFlag = True Then
            
            Else
                ADCOUT = verr.compare(LessThan, Diffaccuracy)
                If ADCOUT.All(True) Then Exit For
            End If
            
            
            If doallFlag = True Then
                code = code.Add(1) 'do all should set start value to smallest value
            Else
                temp = MeasValue.compare(LessThan, TrimTarget)
                If Antitrim = True Then
                    code = code.Add(temp.Multiply(2).Add(1))
                Else
                    code = code.Add(temp.Multiply(-2).Subtract(1))  ' If MeasValue < TrimTarget Then code++ Else code--
                End If
                
            End If
            
            For Each site In TheExec.sites.Active
                If code(site) > TrimCodeValue_Max Then
                    code(site) = code(site) - 1
                    DoneEndTrim = True
                End If ''GoTo EndTrim
                If code(site) < TrimCodeValue_Min Then
                    code(site) = code(site) + 1
                    DoneEndTrim = True
                End If ''GoTo EndTrim
            Next site
          
          
          
            For Each site In TheExec.sites
                TheExec.Datalog.WriteComment "site" & site & " decimal code is " & code
            Next site
            Call TrimUVI80_Meas_VFI(pats(0), TestSequence_array, DigSrc_pin, code, CStr(MeasV_PinS), MeasValue, MeasI_pinS, MeasI_Range, MeasF_PinS_SingleEnd, MeasF_Interval, MeasF_EventSourceWithTerminationMode, DigSrc_DataWidth, DigSrc_Sample_Size, DigSrc_Equation, digsrc_assignment, DigCap_Pin, DigCap_DataWidth, DigCap_Sample_Size, CUS_Str_DigCapData, OutDSP, TrimCodeSize, Trimname, Meas_StoreName, Calc_Eqn, TrimCal_Name, CPUA_Flag_In_Pat, Final_Calc)
            
            If Diffaccuracy <> 0 And TheExec.TesterMode = testModeOffline And i = 3 Then
                For Each site In TheExec.sites
                    MeasValue(site) = TrimTarget - Diffaccuracy + 0.0003
                    If gl_Disable_HIP_debug_log = False Then TheExec.Datalog.WriteComment "Site " & site & ",Code " & code(site) & ", Voltage = " & MeasValue(site)
                Next site
            End If
    
            If trace Then TheExec.Flow.TestLimit code
            If trace Then TheExec.Flow.TestLimit MeasValue
            For Each site In TheExec.sites
                If Abs(MeasValue - TrimTarget) < verr Then
                    BestCode = code
                    BestVal = MeasValue
                    verr = Abs(MeasValue - TrimTarget)
                End If
            Next site
            
            If doallFlag = True Then
            Else
                Done = Done.LogicalOr(First.LogicalXor(MeasValue.compare(LessThan, TrimTarget)))
                If trace Then TheExec.Flow.TestLimit Done, , , , , , , , "done"
                DoneEndTrim = DoneEndTrim.LogicalOr(Done)
                'If Done.All(True) Then Exit For
                If DoneEndTrim.All(True) Then Exit For
            End If
            
        Next i
    
    '''''''''''''''''''''''''''''''''''Binary Search'''''''''''''''''''''''''''''''''''''''''''
    Else
        Dim counter As Long
        Dim trimmax As New SiteLong
        Dim trimmin As New SiteLong
        
        trimmax = TrimCodeValue_Max
        trimmin = TrimCodeValue_Min
        
        counter = 0
        code = (trimmax + trimmin) / 2
        
        Do While counter < TrimCodeSize
            digsrc_assignment = temp_assignment
            
            For Each site In TheExec.sites
                TheExec.Datalog.WriteComment "site" & site & " decimal code is " & code
            Next site
            Call TrimUVI80_Meas_VFI(pats(0), TestSequence_array, DigSrc_pin, code, CStr(MeasV_PinS), MeasValue, MeasI_pinS, MeasI_Range, MeasF_PinS_SingleEnd, MeasF_Interval, MeasF_EventSourceWithTerminationMode, DigSrc_DataWidth, DigSrc_Sample_Size, DigSrc_Equation, digsrc_assignment, DigCap_Pin, DigCap_DataWidth, DigCap_Sample_Size, CUS_Str_DigCapData, OutDSP, TrimCodeSize, Trimname, Meas_StoreName, Calc_Eqn, TrimCal_Name, CPUA_Flag_In_Pat, Final_Calc)
            If counter = 0 Then
            verr = MeasValue.Subtract(TrimTarget).Abs
            BestCode = code
            BestVal = MeasValue
            End If

            
            For Each site In TheExec.sites
                If MeasValue.Subtract(TrimTarget).Abs < verr Then
                    BestCode = code
                    BestVal = MeasValue
                    verr = MeasValue.Subtract(TrimTarget).Abs
                End If
            Next site
            
            For Each site In TheExec.sites
                If Antitrim = True Then
                    If MeasValue(site) < TrimTarget Then
                        If counter = TrimCodeSize - 1 Then
                            code(site) = code(site) - 1
                        Else
                            trimmax(site) = code(site)
                        End If
                    Else
                        If counter = TrimCodeSize - 1 Then
                            'code(Site) = code(Site)
                            code(site) = code(site) + 1
                        Else
                            trimmin(site) = code(site)
                        End If
                    End If
                Else
                    If MeasValue(site) < TrimTarget Then
                        If counter = TrimCodeSize - 1 Then
                            'code(Site) = code(Site)
                            code(site) = code(site) + 1
                        Else
                            trimmin(site) = code(site)
                        End If
                    Else
                        If counter = TrimCodeSize - 1 Then
                            code(site) = code(site) - 1
                        Else
                            trimmax(site) = code(site)
                        End If
                    End If
                End If
            Next site
            
            If counter = TrimCodeSize - 1 Then
            Else
                code = trimmax.Add(trimmin).divide(2)
            End If
            
        
            counter = counter + 1
        Loop
        digsrc_assignment = temp_assignment
        For Each site In TheExec.sites
            TheExec.Datalog.WriteComment "site" & site & " decimal code is " & code
        Next site
        Call TrimUVI80_Meas_VFI(pats(0), TestSequence_array, DigSrc_pin, code, CStr(MeasV_PinS), MeasValue, MeasI_pinS, MeasI_Range, MeasF_PinS_SingleEnd, MeasF_Interval, MeasF_EventSourceWithTerminationMode, DigSrc_DataWidth, DigSrc_Sample_Size, DigSrc_Equation, digsrc_assignment, DigCap_Pin, DigCap_DataWidth, DigCap_Sample_Size, CUS_Str_DigCapData, OutDSP, TrimCodeSize, Trimname, Meas_StoreName, Calc_Eqn, TrimCal_Name, CPUA_Flag_In_Pat, Final_Calc)
        
            For Each site In TheExec.sites
                If MeasValue.Subtract(TrimTarget).Abs < verr Then
                    BestCode = code
                    BestVal = MeasValue
                    verr = MeasValue.Subtract(TrimTarget).Abs
                End If
            Next site

    End If
    
    
    finalflag = True
    
    
EndTrim:

    If Interpose_PostTest <> "" Then
        Call SetForceCondition(Interpose_PostTest & ";STOREPREPAT")
    End If
    
    If MeasV_PinS <> "" Then
        TestNameInput = Report_TName_From_Instance("V", MeasV_PinS, "TrimmedVoltage", 0, 0)
        If InStr(glb_TestInstance, "T4P2") <> 0 Then
            ReCalc = BestVal.Add(1).Multiply(0.7975).Add(0.4)
            OutputTname_format(6) = gl_Tname_Meas_FromFlow(TheExec.Flow.TestLimitIndex)
            TestNameInput = Merge_TName(OutputTname_format)
            TheExec.Flow.TestLimit ReCalc, , , , , , unitVolt, , TestNameInput, , MeasV_PinS, , , , , tlForceFlow
            OutputTname_format(6) = gl_Tname_Meas_FromFlow(TheExec.Flow.TestLimitIndex)
            TestNameInput = Merge_TName(OutputTname_format)
            TheExec.Flow.TestLimit BestVal, , , , , , unitVolt, , TestNameInput, , , , , , , tlForceFlow
        Else
            TheExec.Flow.TestLimit BestVal, , , , , , unitVolt, , TestNameInput, , MeasV_PinS, , , , , tlForceFlow
        End If
    ElseIf MeasI_pinS <> "" Then
        TestNameInput = Report_TName_From_Instance("I", MeasI_pinS, "TrimmedVoltage", 0, 0)
        TheExec.Flow.TestLimit BestVal, , , , , , unitAmp, , TestNameInput, , MeasI_pinS, , , , , tlForceFlow
    ElseIf MeasF_PinS_SingleEnd <> "" Then
        TestNameInput = Report_TName_From_Instance("F", MeasF_PinS_SingleEnd.value, "TrimmedFrequency", 0, 0)
        TheExec.Flow.TestLimit BestVal, , , , , , unitHz, , TestNameInput, , MeasF_PinS_SingleEnd.value, , , , , tlForceFlow
    Else
        TestNameInput = Report_TName_From_Instance("C", DigCap_Pin.value, "TrimmedCode(Decimal)", 0, 0)
        TheExec.Flow.TestLimit BestVal, , , , , , unitNone, , TestNameInput, , DigCap_Pin, , , , , tlForceFlow
    End If
    TestNameInput = Report_TName_From_Instance("C", "x", , 0, 0)
    TheExec.Flow.TestLimit BestCode, TrimCodeValue_Min, TrimCodeValue_Max, , , , , , TestNameInput, , , , , , , tlForceNone 'eng_forceflow_transfer
    
    
    
    Dim TempVal As Integer
    Dim FinalTrimCode As New DSPWave
    Dim Binary_FinalTrimCode As String
    FinalTrimCode.CreateConstant 0, TrimCodeSize
    
    For Each site In TheExec.sites
'        TempVal = BestCode(Site)
        TheExec.Datalog.WriteComment "site" & site & " decimal best code is " & BestCode
        Binary_FinalTrimCode = vbNullString
        For i = 0 To TrimCodeSize - 1
            FinalTrimCode(site).Element(i) = ((BestCode(site) And (2 ^ i)) \ (2 ^ i))
            Binary_FinalTrimCode = Binary_FinalTrimCode & FinalTrimCode(site).Element(i)
'            If i = 0 Then
'                code_bin(Site) = CStr(code(Site) And 1)
'            Else
'                code_bin(Site) = code_bin(Site) & CStr((code(Site) And (2 ^ i)) \ (2 ^ i))
'            End If
'            FinalTrimCode(Site).Element(i) = TempVal Mod 2
'            TempVal = TempVal \ 2
            
        Next i
        TheExec.Datalog.WriteComment "site" & site & " Binary best code is " & Binary_FinalTrimCode
    Next site
    
    If TrimStoreName <> "" Then
        Call Checker_StoreDigCapAllToDictionary(TrimStoreName, FinalTrimCode)
    End If
    
    If TrimCal_Name <> "" Then
        If Final_Calc = True Then
            Call ProcessCalcEquation(Calc_Eqn)
        End If
    Else
        If Calc_Eqn <> "" Then
            Call ProcessCalcEquation(Calc_Eqn)
        End If
    End If
    
    HardIP_WriteFuncResult
    
    DebugPrintFunc Pat
    

    If TheExec.TesterMode = testModeOffline Then
    Else
    End If
    TheHdw.Alarms.Check
    Exit Function
    
errHandler:
'       ByPassTestLimit = False
    TheExec.Datalog.WriteComment ("ERROR in DSSC_Search: " & err.Description)
    'TrimUVI80Code_VFI = TL_ERROR
    If AbortTest Then Exit Function Else Resume Next
End Function



Public Function TrimUVI80Code_VFI(Optional Pat As String, Optional TestSequence As String, Optional MeasV_PinS As String, Optional MeasI_pinS As String, Optional MeasI_Range As Double, _
    Optional MeasF_PinS_SingleEnd As PinList, Optional MeasF_Interval As String, Optional MeasF_EventSourceWithTerminationMode As EventSourceWithTerminationMode, _
    Optional TrimTarget As Double, Optional TrimStart As Long, Optional TrimFormat As String, _
    Optional DigSrc_pin As PinList, Optional DigSrc_DataWidth As Long, Optional DigSrc_Sample_Size As Long, Optional DigSrc_Equation As String, Optional digsrc_assignment As String, _
    Optional DigCap_Pin As PinList, Optional DigCap_DataWidth As Long, Optional DigCap_Sample_Size As Long, Optional CUS_Str_DigCapData As String = vbNullString, _
    Optional TrimStoreName As String, Optional Diffaccuracy As Double, Optional TrimMethod As Long, _
    Optional Meas_StoreName As String, Optional Calc_Eqn As String, Optional TrimCal_Name As String, Optional Antitrim As Boolean = False, Optional Interpose_PrePat As String, Optional CPUA_Flag_In_Pat As Boolean = True, Optional Interpose_PostTest As String, _
    Optional Final_Calc As Boolean = False, Optional BestMeasVal_StoreName As String, Optional Final_Calc_Eqn As String, Optional MSB_First_Flag As Boolean = False, Optional Validating_ As Boolean)
''    Dim PatCount As Long
    Dim i As Integer
    Dim pats() As String
    Dim code As New SiteLong
    Dim MeasValue As New SiteDouble
    Dim BestCode As New SiteLong, BestVal As New SiteDouble, verr As New SiteDouble, temp As New SiteLong
    Dim First As New SiteBoolean, Done As New SiteBoolean
    Dim trace As Boolean
    Dim site As Variant
    Dim Ts As Variant
    Dim ADCOUT As New SiteBoolean
    Dim TrimStep As Long
    Dim OutDSP As New DSPWave
    Dim PatCount As Long, PattArray() As String
    Dim TestSequence_array() As String
    Dim doallFlag As Boolean
    Dim finalflag As Boolean
    Dim OutputTname_format() As String
    Dim TName_Ary() As String
    Dim TestNameInput As String
    Dim ReCalc As New SiteDouble
    Dim TempPat As New Pattern
    TempPat.value = Pat
    Call ProcessInputToGLB(TempPat, TestSequence, CPUA_Flag_In_Pat, , , , , MeasV_PinS, MeasF_PinS_SingleEnd.value, MeasF_Interval, MeasF_EventSourceWithTerminationMode, , , , MeasI_pinS, CStr(MeasI_Range), , DigCap_Pin, DigCap_DataWidth, DigCap_Sample_Size, DigSrc_pin, CStr(DigSrc_DataWidth) _
                    , CStr(DigSrc_Sample_Size), DigSrc_Equation, digsrc_assignment, , , , , CUS_Str_DigCapData, , , , , , , , , , , , , , , Interpose_PrePat, , Interpose_PostTest)
    gl_TName_Pat = Pat

    glb_TestInstance = vbNullString
    glb_TestInstance = UCase(TheExec.DataManager.instancename)
    
    Call GetFlowTName
    
'    If InStr(TheExec.DataManager.InstanceName, "MTRGR_T4P2") <> 0 Then
'        Cal_Eqn = "ALG::Calc_Metrology_GainError(GainError,V1)"
'        Meas_StoreName = "V1+"
'        TrimCal_Name = "GainError"
'    End If
    'ByPassTestLimit = True
    If Validating_ Then
        Call PrLoadPattern(Pat)
        Exit Function    ' Exit after validation
    End If
    
    
    Call HardIP_InitialSetupForPatgen

    TestSequence_array = Split(TestSequence, ",")
    
    
    On Error GoTo errHandler
    TheHdw.Digital.ApplyLevelsTiming True, True, True, tlPowered
    
    PATT_GetPatListFromPatternSet Pat, pats, PatCount
    
    If Interpose_PrePat <> "" Then
        Call SetForceCondition(Interpose_PrePat & ";STOREPREPAT")
    End If
    
    TName_Ary = Split(gl_Tname_Meas, "+")
    
  '  If (UBound(TestSequence_array) > UBound(TName_Ary)) Then
  '      ReDim Preserve TName_Ary(UBound(TestSequence_array)) As String
  '
  '  End If
    
    
'    '''''''''''''''setup UVI80 for meas V''''''''''''''''''
'    With thehdw.DCVI.Pins(MeasV_PinS)
'            .Gate = False
'            .Disconnect tlDCVIConnectDefault
'            .mode = tlDCVIModeHighImpedance
'            .Connect tlDCVIConnectHighSense
'            .Voltage = 6
'            .Current = 0
'             thehdw.Wait 0.5 * ms
'            .Gate = True
'    End With
    ''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    '''''''''''''''Setup UVI180 for meas I'''''''''''''''''
    
    Dim SplitByEqual() As String, SplitByColon() As String, TrimCodeSize As Long
    Dim TrimCodeValue_Min As Long, TrimCodeValue_Mid As Long, TrimCodeValue_Max As Long
    Dim Trimname As String
    SplitByEqual = Split(TrimFormat, "=")
    SplitByColon = Split(SplitByEqual(1), ":")
    Trimname = SplitByEqual(0)
    TrimCodeSize = SplitByColon(0) + 1
    TrimCodeValue_Min = 0
    TrimCodeValue_Mid = (2 ^ TrimCodeSize) / 2
    If SplitByColon(1) = 0 Then
        TrimCodeValue_Max = 2 ^ TrimCodeSize - 1
    Else
        TrimCodeValue_Max = SplitByColon(1)
    End If
    Dim binaryFlag As Boolean
    Dim temp_assignment As String
    temp_assignment = digsrc_assignment
     
    '''''''''''''''''''''''''''''''''''Process Trim Method'''''''''''''''''''''''''''''''''''''
    If TrimMethod = 0 Then
        binaryFlag = False
        doallFlag = False
    ElseIf TrimMethod = 2 Then
        binaryFlag = False
        doallFlag = True
        
    Else
        binaryFlag = True
    End If
    
    '''''''''''''''''''''''''''''''''''Linear Search'''''''''''''''''''''''''''''''''''''''''''
    If binaryFlag = False Then
    
        code = TrimStart
        If doallFlag = True Then
            TrimStep = TrimCodeValue_Max - TrimStart
        Else
            TrimStep = TrimCodeValue_Max
        End If
        
        
        Call TrimUVI80_Meas_VFI(pats(0), TestSequence_array, DigSrc_pin, code, MeasV_PinS, MeasValue, MeasI_pinS, MeasI_Range, MeasF_PinS_SingleEnd, MeasF_Interval, MeasF_EventSourceWithTerminationMode, DigSrc_DataWidth, DigSrc_Sample_Size, DigSrc_Equation, digsrc_assignment, DigCap_Pin, DigCap_DataWidth, DigCap_Sample_Size, CUS_Str_DigCapData, OutDSP, TrimCodeSize, Trimname, Meas_StoreName, Calc_Eqn, TrimCal_Name, CPUA_Flag_In_Pat, Final_Calc, MSB_First_Flag:=MSB_First_Flag)
        
        
    '    If Diffaccuracy <> 0 And TheExec.TesterMode = testModeOffline Then
    '        For Each Site In TheExec.sites
    '            MeasValue(Site) = TrimTarget + Diffaccuracy - 0.0003
    '            TheExec.Datalog.WriteComment "Site " & Site & ",Code " & code(Site) & ", Voltage = " & MeasValue(Site)
    '        Next Site
    '    End If
        If trace Then TheExec.Flow.TestLimit code
        If trace Then TheExec.Flow.TestLimit MeasValue
        First = MeasValue.compare(LessThan, TrimTarget)
        If trace Then TheExec.Flow.TestLimit First, , , , , , , , "first"
        BestCode = code
        BestVal = MeasValue
        verr = MeasValue.Subtract(TrimTarget).Abs
    
        For i = 0 To TrimStep - 1
            digsrc_assignment = temp_assignment
            If trace Then TheExec.Datalog.WriteComment ("i = " & i)
            If doallFlag = True Then
            
            Else
                ADCOUT = verr.compare(LessThan, Diffaccuracy)
                If ADCOUT.All(True) Then Exit For
            End If
            
            If doallFlag = True Then
                code = code.Add(1)
            Else
                temp = MeasValue.compare(LessThan, TrimTarget)
                If Antitrim = True Then
                    code = code.Add(temp.Multiply(2).Add(1))
                Else
                    code = code.Add(temp.Multiply(-2).Subtract(1))  ' If MeasValue < TrimTarget Then code++ Else code--
                End If
                
            End If
            
            For Each site In TheExec.sites.Active
                If code(site) > TrimCodeValue_Max Then: code(site) = code(site) - 1: ''GoTo EndTrim
                If code(site) < TrimCodeValue_Min Then: code(site) = code(site) + 1: ''GoTo EndTrim
            Next site
          
            Call TrimUVI80_Meas_VFI(pats(0), TestSequence_array, DigSrc_pin, code, MeasV_PinS, MeasValue, MeasI_pinS, MeasI_Range, MeasF_PinS_SingleEnd, MeasF_Interval, MeasF_EventSourceWithTerminationMode, DigSrc_DataWidth, DigSrc_Sample_Size, DigSrc_Equation, digsrc_assignment, DigCap_Pin, DigCap_DataWidth, DigCap_Sample_Size, CUS_Str_DigCapData, OutDSP, TrimCodeSize, Trimname, Meas_StoreName, Calc_Eqn, TrimCal_Name, CPUA_Flag_In_Pat, Final_Calc, MSB_First_Flag:=MSB_First_Flag)
            
            If Diffaccuracy <> 0 And TheExec.TesterMode = testModeOffline And i = 3 Then
                For Each site In TheExec.sites
                    MeasValue(site) = TrimTarget - Diffaccuracy + 0.0003
                    If gl_Disable_HIP_debug_log = False Then TheExec.Datalog.WriteComment "Site " & site & ",Code " & code(site) & ", Voltage = " & MeasValue(site)
                Next site
            End If
    
            If trace Then TheExec.Flow.TestLimit code
            If trace Then TheExec.Flow.TestLimit MeasValue
            For Each site In TheExec.sites
                If Abs(MeasValue - TrimTarget) < verr Then
                    BestCode = code
                    BestVal = MeasValue
                    verr = Abs(MeasValue - TrimTarget)
                End If
            Next site
            
            If doallFlag = True Then
            Else
                Done = Done.LogicalOr(First.LogicalXor(MeasValue.compare(LessThan, TrimTarget)))
                If trace Then TheExec.Flow.TestLimit Done, , , , , , , , "done"
                If Done.All(True) Then Exit For
            End If
        Next i
    
    '''''''''''''''''''''''''''''''''''Binary Search'''''''''''''''''''''''''''''''''''''''''''
    Else
        Dim counter As Long
        Dim trimmax As New SiteLong
        Dim trimmin As New SiteLong
        
        trimmax = TrimCodeValue_Max
        trimmin = TrimCodeValue_Min
        
        counter = 0
        code = (trimmax + trimmin) / 2
        
        Do While counter < TrimCodeSize
            digsrc_assignment = temp_assignment
            
            Call TrimUVI80_Meas_VFI(pats(0), TestSequence_array, DigSrc_pin, code, MeasV_PinS, MeasValue, MeasI_pinS, MeasI_Range, MeasF_PinS_SingleEnd, MeasF_Interval, MeasF_EventSourceWithTerminationMode, DigSrc_DataWidth, DigSrc_Sample_Size, DigSrc_Equation, digsrc_assignment, DigCap_Pin, DigCap_DataWidth, DigCap_Sample_Size, CUS_Str_DigCapData, OutDSP, TrimCodeSize, Trimname, Meas_StoreName, Calc_Eqn, TrimCal_Name, CPUA_Flag_In_Pat, Final_Calc, MSB_First_Flag:=MSB_First_Flag)
            If counter = 0 Then
            verr = MeasValue.Subtract(TrimTarget).Abs
            BestCode = code
            BestVal = MeasValue
            End If

            
            
            For Each site In TheExec.sites
                If MeasValue.Subtract(TrimTarget).Abs < verr Then
                    BestCode = code
                    BestVal = MeasValue
                    verr = MeasValue.Subtract(TrimTarget).Abs
                End If
            Next site
            
            For Each site In TheExec.sites
                If Antitrim = True Then
                    If MeasValue(site) < TrimTarget Then
                        If counter = TrimCodeSize - 1 Then
                            code(site) = code(site) - 1
                        Else
                            trimmax(site) = code(site)
                        End If
                    Else
                        If counter = TrimCodeSize - 1 Then
                            code(site) = code(site)
                        Else
                            trimmin(site) = code(site)
                        End If
                    End If
                Else
                    If MeasValue(site) < TrimTarget Then
                        If counter = TrimCodeSize - 1 Then
                            code(site) = code(site)
                        Else
                            trimmin(site) = code(site)
                        End If
                    Else
                        If counter = TrimCodeSize - 1 Then
                            code(site) = code(site) - 1
                        Else
                            trimmax(site) = code(site)
                        End If
                    End If
                End If
            Next site
            
            If counter = TrimCodeSize - 1 Then
            Else
                code = trimmax.Add(trimmin).divide(2)
            End If
            
        
            counter = counter + 1
        Loop
        digsrc_assignment = temp_assignment
        Call TrimUVI80_Meas_VFI(pats(0), TestSequence_array, DigSrc_pin, code, MeasV_PinS, MeasValue, MeasI_pinS, MeasI_Range, MeasF_PinS_SingleEnd, MeasF_Interval, MeasF_EventSourceWithTerminationMode, DigSrc_DataWidth, DigSrc_Sample_Size, DigSrc_Equation, digsrc_assignment, DigCap_Pin, DigCap_DataWidth, DigCap_Sample_Size, CUS_Str_DigCapData, OutDSP, TrimCodeSize, Trimname, Meas_StoreName, Calc_Eqn, TrimCal_Name, CPUA_Flag_In_Pat, Final_Calc, MSB_First_Flag:=MSB_First_Flag)
        
            For Each site In TheExec.sites
                If MeasValue.Subtract(TrimTarget).Abs < verr Then
                    BestCode = code
                    BestVal = MeasValue
                    verr = MeasValue.Subtract(TrimTarget).Abs
                End If
            Next site

    End If
    
    
    finalflag = True
    
    'Call TrimUVI80_Meas_VFI(pats(0), TestSequence_array, DigSrc_pin, BestCode, MeasV_PinS, MeasValue, MeasI_pinS, MeasI_Range, MeasF_PinS_SingleEnd, MeasF_Interval, MeasF_EventSourceWithTerminationMode, DigSrc_DataWidth, DigSrc_Sample_Size, DigSrc_Equation, DigSrc_Assignment, DigCap_Pin, DigCap_DataWidth, DigCap_Sample_Size, CUS_Str_DigCapData, OutDSP, TrimCodesize, Trimname, Meas_StoreName, Calc_Eqn, TrimCal_Name, CPUA_Flag_In_Pat, finalflag)
    
    
EndTrim:

    If Interpose_PostTest <> "" Then
        Call SetForceCondition(Interpose_PostTest & ";STOREPREPAT")
    End If
    
    Dim MeasV_Split() As String
    MeasV_Split = Split(MeasV_PinS, "+")
    
    If MeasV_PinS <> "" Then
        TestNameInput = Report_TName_From_Instance("V", vbNullString, "TrimmedVoltage", 0, 0)
        If InStr(glb_TestInstance, "T4P2") <> 0 Then
            ReCalc = BestVal.Add(1).Multiply(0.7975).Add(0.4)
            'OutputTname_format(6) = gl_Tname_Meas_FromFlow(TheExec.Flow.TestLimitIndex)
            'TestNameInput = Merge_TName(OutputTname_format)
            TheExec.Flow.TestLimit ReCalc, , , , , , unitVolt, , TestNameInput, , MeasV_Split(0), , , , , tlForceFlow
            'OutputTname_format(6) = gl_Tname_Meas_FromFlow(TheExec.Flow.TestLimitIndex)
            'TestNameInput = Merge_TName(OutputTname_format)
            TheExec.Flow.TestLimit BestVal, , , , , , unitVolt, , TestNameInput, , , , , , , tlForceFlow
        Else
            TheExec.Flow.TestLimit BestVal, , , , , , unitVolt, , TestNameInput, , MeasV_Split(0), , , , , tlForceFlow
        End If
    ElseIf MeasI_pinS <> "" Then
        TestNameInput = Report_TName_From_Instance("I", vbNullString, "TrimmedVoltage", 0, 0)
        TheExec.Flow.TestLimit BestVal, , , , , , unitAmp, , TestNameInput, , MeasI_pinS, , , , , tlForceFlow
    ElseIf MeasF_PinS_SingleEnd <> "" Then
        TestNameInput = Report_TName_From_Instance("F", vbNullString, "TrimmedFrequency", 0, 0)
        TheExec.Flow.TestLimit BestVal, , , , , , unitHz, , TestNameInput, , MeasF_PinS_SingleEnd.value, , , , , tlForceFlow
    Else
        TestNameInput = Report_TName_From_Instance("C", vbNullString, "TrimmedCode(Decimal)", 0, 0)
        TheExec.Flow.TestLimit BestVal, , , , , , unitNone, , TestNameInput, , DigCap_Pin, , , , , tlForceFlow
    End If
        
    If BestMeasVal_StoreName <> "" Then         'Cebu MTRG GR t1p1 store best meas value 20180806
        Call AddStoredMeasurement(BestMeasVal_StoreName, BestVal)
    End If
    
    TestNameInput = Report_TName_From_Instance("C", vbNullString, "TrimmedCode", 0, 0)
    TheExec.Flow.TestLimit BestCode, 0, 2 ^ TrimCodeSize, , , , , , TestNameInput, , , , , , , tlForceNone 'eng_forceflow_transfer
    'ByPassTestLimit = False
    If DigCap_Sample_Size <> 0 Then
        Dim DigCapPinAry() As String, NumberPins As Long
        
        'Call TrimUVI80_Meas_VFI(pats(0), TestSequence_array, DigSrc_pin, BestCode, MeasV_PinS, MeasValue, MeasI_pinS, MeasI_Range, MeasF_PinS_SingleEnd, MeasF_Interval, MeasF_EventSourceWithTerminationMode, DigSrc_DataWidth, DigSrc_Sample_Size, DigSrc_Equation, DigSrc_Assignment, DigCap_Pin, DigCap_DataWidth, DigCap_Sample_Size, CUS_Str_DigCapData, OutDSP, TrimCodesize, Trimname, Meas_StoreName, Cal_Eqn, TrimCal_Name)
        Call TheExec.DataManager.DecomposePinList(DigCap_Pin, DigCapPinAry(), NumberPins)
        
        If NumberPins > 1 Then
            'Call CreateSimulateDataDSPWave_Parallel(OutDSP, DigCap_Sample_Size)
            Call Checker_StoreDigCapAllToDictionary(CUS_Str_DigCapData, OutDSP, NumberPins)
            Call DigCapDataProcessByDSP_Parallel(CUS_Str_DigCapData, OutDSP, DigCap_Sample_Size, NumberPins)

        ElseIf NumberPins = 1 Then
            'Call CreateSimulateDataDSPWave(OutDSP, DigCap_Sample_Size, DigCap_DataWidth)
            Call Checker_StoreDigCapAllToDictionary(CUS_Str_DigCapData, OutDSP, NumberPins)
            Call DigCapDataProcessByDSP(CUS_Str_DigCapData, OutDSP, DigCap_Sample_Size, DigCap_DataWidth)
        End If
    End If
    
    
    Dim TempVal As Integer
    Dim FinalTrimCode As New DSPWave
    Dim TrimValue As String
    
    FinalTrimCode.CreateConstant 0, TrimCodeSize
    
    For Each site In TheExec.sites
        TempVal = BestCode(site)
        TrimValue = vbNullString
        For i = 0 To TrimCodeSize - 1
            FinalTrimCode(site).Element(i) = TempVal Mod 2
            TempVal = TempVal \ 2
            TrimValue = TrimValue & CStr(FinalTrimCode(site).Element(i))
            
        Next i
        'theexec.Datalog.WriteComment "Site==> " & Site & ", Trimed code==>" & trimvalue
        
        If UCase(glb_TestInstance) Like "*MTRGRT1P1*" And FinalTrimCode(site).Element(TrimCodeSize - 1) = 0 Then
                FinalTrimCode(site).Element(TrimCodeSize - 1) = 1
        ElseIf UCase(glb_TestInstance) Like "*MTRGRT1P1*" And FinalTrimCode(site).Element(TrimCodeSize - 1) = 1 Then
                FinalTrimCode(site).Element(TrimCodeSize - 1) = 0
        End If
    Next site
    
    If TrimStoreName <> "" Then
        
        Call Checker_StoreDigCapAllToDictionary(TrimStoreName, FinalTrimCode)
        
        Dim negate_FinalTrimCode As New DSPWave
        negate_FinalTrimCode = FinalTrimCode
        For Each site In TheExec.sites
            If negate_FinalTrimCode.Element(0) = 1 Then
                negate_FinalTrimCode.Element(0) = 0
            Else
                negate_FinalTrimCode.Element(0) = 1
            End If
        Next site
        Call Checker_StoreDigCapAllToDictionary(TrimStoreName + "_negate", FinalTrimCode)

    End If
    
    If TrimCal_Name <> "" Then
        If Final_Calc = True Then
            Call ProcessCalcEquation(Calc_Eqn)
        End If
    Else
        If Calc_Eqn <> "" Then
            Call ProcessCalcEquation(Calc_Eqn)
        End If
    End If
    
    If Final_Calc_Eqn <> "" Then
        Call ProcessCalcEquation(Final_Calc_Eqn)
    End If
    
    Call HardIP_WriteFuncResult(, , glb_TestInstance)
    
    DebugPrintFunc Pat
    
    '' 20170704 - Add write efuse function
''    Dim sl_Fuse_Val As SiteLong
    If TheExec.TesterMode = testModeOffline Then
    Else
''        For Each Site In TheExec.sites
''            sl_Fuse_Val(Site) = BestCode(Site)
''        Next Site
        
'        If TrimFuseName <> "" And TrimFuseTypeName <> "" Then
'            Call HIP_eFuse_Write(TrimFuseTypeName, TrimFuseName, BestCode)
'        End If
    End If
        
    Exit Function
    
errHandler:
'       ByPassTestLimit = False
    TheExec.Datalog.WriteComment ("ERROR in DSSC_Search: " & err.Description)
    'TrimUVI80Code_VFI = TL_ERROR
    If AbortTest Then Exit Function Else Resume Next
End Function

Public Function TrimCodeImpedence(Optional patset As Pattern, Optional CPUA_Flag_In_Pat As Boolean, _
    Optional MeasR_Pins_SingleEnd As PinList, Optional MeasR_Pins_Differential As PinList, Optional StrForceVolt As String, Optional DigSrc_pin As PinList, Optional DigSrc_Sample_Size As Long, _
    Optional TrimPrcocessAll As Boolean = False, Optional UseMinimumTrimCode As Boolean = False, Optional PreCheckMinMaxTrimCode As Boolean = False, _
    Optional TrimTarget As Double = 50, Optional TrimTargetTolerance As Double = 0, Optional TrimStart As String, Optional TrimFormat As String, _
    Optional TrimStoreName As String, Optional TrimFuseName As String, Optional TrimFuseTypeName As String, _
    Optional Fixed_DigSrc_DataWidth As Long, Optional Fixed_DigSrc_Sample_Size As Long, Optional Fixed_DigSrc_Equation As String, Optional Fixed_DigSrc_Assignment As String, _
    Optional Calc_Eqn As String, Optional GetbitNumber As Long, Optional b_PD_Mode As Boolean = True, Optional Validating_ As Boolean, Optional Interpose_PrePat As String) As Long

    Dim PatCount As Long, PattArray() As String
    Dim TrimmedImpedance() As New SiteDouble
    Dim TrimCode() As New SiteLong
    Dim x As Long
    Dim TestNameInput As String
    Dim OutputTname_format() As String
    Dim TName_Ary() As String
    
    If Validating_ Then
        Call PrLoadPattern(patset.value)
        Exit Function
    End If
    
    Call HardIP_InitialSetupForPatgen
    
    Dim InitialDSPWave As New DSPWave, PastDSPWave As New DSPWave, InDSPWave As New DSPWave

    Dim site As Variant
    Dim Pat As String
    Dim i As Long, j As Long, k As Long
    
    Dim MeasureImped As New PinListData, MeasureImped_F1 As New PinListData, MeasureImped_F2 As New PinListData

    On Error GoTo ErrorHandler

    TheHdw.Digital.Patgen.Halt
    
    TheHdw.Digital.ApplyLevelsTiming True, True, True, tlPowered
    
    If Calc_Eqn <> "" Then
        Call ProcessCalcEquation(Calc_Eqn)
    End If
    TName_Ary = Split(gl_Tname_Meas, "+")
    
    Call HardIP_InitialSetupForPatgen
    
    TheHdw.Patterns(patset).Load

    gl_TName_Pat = patset.value

    Call PATT_GetPatListFromPatternSet(patset.value, PattArray, PatCount)
    
    '' 20160425 - Check format from TrimFormat
    Dim StrSeparatebyComma() As String
    Dim ExecutionMax As Long
    StrSeparatebyComma = Split(TrimFormat, ";")
    
    ExecutionMax = UBound(StrSeparatebyComma)
    Dim StrSeparatebyEqual() As String, StrSeparatebyColon() As String '' Get Src bit
    Dim SrcStartBit As Long, SrcEndBit As Long
    
    '' 20161230 - To check "Fixed" key word to decide trim code process
    Dim b_ExecuteTrimCode() As Boolean
    ReDim b_ExecuteTrimCode(ExecutionMax) As Boolean
    Dim b_HasFIXED_InTrimFormat As Boolean
    
    For i = 0 To ExecutionMax
        If InStr(UCase(StrSeparatebyComma(i)), "FIXED") <> 0 Then
            b_ExecuteTrimCode(i) = False
            b_HasFIXED_InTrimFormat = True
        Else
            b_ExecuteTrimCode(i) = True

        End If
    Next i
    
    Dim Pin_Ary() As String, Pin_Cnt As Long, pin As Variant
    Dim StorePerPinFinalTrimCode() As New DSPWave
    Dim StorePerPinFinalTrimCode_Dec() As New DSPWave
    Dim b_IsDifferential As Boolean
    Dim TempPins As String
    If MeasR_Pins_SingleEnd <> "" Then
        TheExec.DataManager.DecomposePinList MeasR_Pins_SingleEnd, Pin_Ary, Pin_Cnt
        b_IsDifferential = False
    ElseIf MeasR_Pins_Differential <> "" Then
        TheExec.DataManager.DecomposePinList MeasR_Pins_Differential, Pin_Ary, Pin_Cnt
        
        For i = 0 To Pin_Cnt - 1
            If InStr(UCase(Pin_Ary(i)), "_P") <> 0 Then
                If i = 0 Then
                    TempPins = Pin_Ary(i)
                Else
                    TempPins = TempPins & "," & Pin_Ary(i)
                End If
            End If
        Next i
        Pin_Cnt = Pin_Cnt / 2
        ReDim Pin_Ary(Pin_Cnt) As String
        Pin_Ary = Split(TempPins, ",")
        b_IsDifferential = True
    End If
    
    ReDim StorePerPinFinalTrimCode(Pin_Cnt - 1) As New DSPWave
    ReDim StorePerPinFinalTrimCode_Dec(Pin_Cnt - 1) As New DSPWave
        
    For i = 0 To Pin_Cnt - 1
        StorePerPinFinalTrimCode(i).CreateConstant 0, DigSrc_Sample_Size, DspLong
        StorePerPinFinalTrimCode_Dec(i).CreateConstant 0, DigSrc_Sample_Size, DspLong
    Next i
    
    Dim b_HighThanTargetImped As New SiteBoolean
    b_HighThanTargetImped = False
    
    Dim OutputTrimCode As String
    Dim TestLimitIndex As Long, LastSectionF1F2_Index As Long
    LastSectionF1F2_Index = 0
    
    Dim Dec_TrimStart_1st As Long
    
    '' 20160706 Create value for final Impeduency
    Dim b_DefineFinalImped As New SiteBoolean, FinalImped As New PinListData
    
    ''20160712 - If match taget Imped just store the trim code
    Dim b_MatchTagetImped As New SiteBoolean, b_DisplayImped As New SiteBoolean, StoredTargetTrimCode As New DSPWave
    
    b_MatchTagetImped = False
    b_DisplayImped = False
    StoredTargetTrimCode.CreateConstant 0, DigSrc_Sample_Size, DspLong
    
    Dim StoreEachTrimImped() As New PinListData
    Dim StoreEachTrimCode() As New DSPWave
    ReDim StoreEachTrimImped(DigSrc_Sample_Size + 1) As New PinListData
    ReDim StoreEachTrimCode(DigSrc_Sample_Size + 1) As New DSPWave
    Dim StoreEachIndex As Long
    
    ''20161128-Stop trim code process
    Dim b_StopTrimCodeProcess As New SiteBoolean
    b_StopTrimCodeProcess = False
    
    For i = 0 To UBound(StoreEachTrimCode)
        StoreEachTrimCode(i).CreateConstant 0, DigSrc_Sample_Size, DspLong
    Next i
    
    ''20161230-Add loop to trim code for each pin
    Dim PerPinIndex As Long
    PerPinIndex = 0
    
    ''20170117 - For long length fixed code
    Dim InitialTrimDSPWave As New DSPWave
    Dim InitialFixedDSPWave As New DSPWave
    Dim Trim_DigSrc_Sample_Size As Long
    
    '' 20170117-Evaluate for ForceVolt
    Dim SplitForceVolt() As String
    SplitForceVolt = Split(StrForceVolt, ",")
    Dim ForceVolt As String
    Call HIP_Evaluate_ForceVal(SplitForceVolt)
    For i = 0 To UBound(SplitForceVolt)
        If i = 0 Then
            ForceVolt = SplitForceVolt(i)
        Else
            ForceVolt = ForceVolt & "," & SplitForceVolt(i)
        End If
    Next i
    
    TrimStart = Replace(TrimStart, "&", vbNullString)
    ReDim TrimCode(Pin_Cnt)
    ReDim TrimmedImpedance(Pin_Cnt)
    Dim PinNumber As Long
    ''20170210-Specified item to store trim code to fuse by 1 pin
    Trim_DigSrc_Sample_Size = DigSrc_Sample_Size - Fixed_DigSrc_Sample_Size
    Dim wkds_StoreTrimCodeToDict_DEC As New DSPWave
    Dim wkds_StoreTrimCodeToDict_BIN As New DSPWave
    wkds_StoreTrimCodeToDict_DEC.CreateConstant 0, 1, DspLong
    
    If Fixed_DigSrc_Equation <> "" Then
        wkds_StoreTrimCodeToDict_BIN.CreateConstant 0, Trim_DigSrc_Sample_Size, DspLong
    Else
        wkds_StoreTrimCodeToDict_BIN.CreateConstant 0, DigSrc_Sample_Size, DspLong
    End If
    
    Dim b_wkds_Store_Flag As Boolean
    b_wkds_Store_Flag = False

    For Each pin In Pin_Ary

        If LCase(TrimStoreName) = LCase("wkds_1") And LCase(pin) = LCase("DDR0_ADDR_SOP_P0") Then
            b_wkds_Store_Flag = True
        ElseIf LCase(TrimStoreName) = LCase("wkds_2") And LCase(pin) = LCase("DDR0_ADDR_SOP_P1") Then
            b_wkds_Store_Flag = True
        ElseIf LCase(TrimStoreName) = LCase("wkds_3") And LCase(pin) = LCase("DDR1_ADDR_SOP_P0") Then
            b_wkds_Store_Flag = True
        ElseIf LCase(TrimStoreName) = LCase("wkds_4") And LCase(pin) = LCase("DDR1_ADDR_SOP_P1") Then
            b_wkds_Store_Flag = True
        End If
        
        If Fixed_DigSrc_Equation <> "" Then
            For Each site In TheExec.sites.Active
                Call Create_DigSrc_Data(DigSrc_pin, Fixed_DigSrc_DataWidth, Fixed_DigSrc_Sample_Size, Fixed_DigSrc_Equation, Fixed_DigSrc_Assignment, InitialFixedDSPWave, site)
            Next site
            Dec_TrimStart_1st = Bin2Dec(TrimStart)
            
            InitialDSPWave.CreateConstant Dec_TrimStart_1st, 1, DspLong
''            Trim_DigSrc_Sample_Size = DigSrc_Sample_Size - Fixed_DigSrc_Sample_Size
            
            Call rundsp.CreateFlexibleDSPWave(InitialDSPWave, Trim_DigSrc_Sample_Size, InitialTrimDSPWave)
           
           Call rundsp.CombineDSPWave(InitialFixedDSPWave, InitialTrimDSPWave, Fixed_DigSrc_Sample_Size, Trim_DigSrc_Sample_Size, InDSPWave)
        Else
            Dec_TrimStart_1st = Bin2Dec(TrimStart)
            
            InitialDSPWave.CreateConstant Dec_TrimStart_1st, 1, DspLong
        
            Call rundsp.CreateFlexibleDSPWave(InitialDSPWave, DigSrc_Sample_Size, InDSPWave)
        End If
        
        Call SetupDigSrcDspWave(PattArray(0), DigSrc_pin, "TrimCodeImped", DigSrc_Sample_Size, InDSPWave)

        If gl_Disable_HIP_debug_log = False Then TheExec.Datalog.WriteComment ("First Time Setup")
        '' Debug use
        For Each site In TheExec.sites
            OutputTrimCode = vbNullString
            For k = 0 To InDSPWave(site).SampleSize - 1
                OutputTrimCode = OutputTrimCode & CStr(InDSPWave(site).Element(k))
            Next k
            If gl_Disable_HIP_debug_log = False Then TheExec.Datalog.WriteComment ("Site_" & site & " Initial Output Trim Code = " & OutputTrimCode)
        Next site
        
        For Each site In TheExec.sites
            StoreEachTrimCode(0)(site) = InDSPWave(site).Copy
        Next site
        
        Call TheHdw.Patterns(PattArray(0)).start
        
        Call SubMeasR(CPUA_Flag_In_Pat, CStr(pin), ForceVolt, MeasureImped, b_IsDifferential, b_PD_Mode)
        
        StoreEachTrimImped(0) = MeasureImped
        
        b_HighThanTargetImped = MeasureImped.Math.Subtract(TrimTarget).compare(LessThan, 0)
        PastDSPWave = InDSPWave
        
        TestNameInput = "Imped"
        TestLimitIndex = 0
        
        '' 20160712 - Modify to use WriteComment to display output Impeduency.
        For Each site In TheExec.sites
            If gl_Disable_HIP_debug_log = False Then TheExec.Datalog.WriteComment ("Site " & site & " Pin " & CStr(pin) & " Impedence = " & FormatNumber((MeasureImped.Pins(0).value(site)), 3) & " Ohm")
            
            TrimmedImpedance(PinNumber) = MeasureImped.Pins(0).value(site)
            TrimCode(PinNumber) = 0
             For x = 0 To GetbitNumber - 1
                TrimCode(PinNumber) = InDSPWave.Element(InDSPWave.SampleSize - x - 1) * 2 ^ (GetbitNumber - 1 - x) + TrimCode(PinNumber)
            Next x
            
        Next site
        
        '' 20160712 - Compare Measure Impeduency whether match target Imped
        b_MatchTagetImped = MeasureImped.Math.Subtract(TrimTarget).Abs.compare(LessThanOrEqualTo, TrimTargetTolerance)
        
        b_DisplayImped = b_DisplayImped.LogicalOr(b_MatchTagetImped)
        For Each site In TheExec.sites
            If b_MatchTagetImped(site) = True And StoredTargetTrimCode(site).CalcSum = 0 Then
                StoredTargetTrimCode(site) = InDSPWave(site).Copy
                b_StopTrimCodeProcess(site) = True
            End If
        Next site
       If gl_Disable_HIP_debug_log = False Then TheExec.Datalog.WriteComment ("======================================================================================")
        
        ''========================================================================================
        ''20161128 Pre check Min/Max trim code process
        Dim b_KeepGoing As New SiteBoolean
        Dim PreviousImped As New PinListData
        If PreCheckMinMaxTrimCode = True Then
            PreviousImped = MeasureImped
            Call rundsp.PreCheckMinMaxTrimCode(b_HighThanTargetImped, InDSPWave)
            Call SetupDigSrcDspWave(PattArray(0), DigSrc_pin, "TrimCodeImped", DigSrc_Sample_Size, InDSPWave)
    
            Call TheHdw.Patterns(PattArray(0)).start
        
            Call SubMeasR(CPUA_Flag_In_Pat, CStr(pin), ForceVolt, MeasureImped, b_IsDifferential, b_PD_Mode)

            If TheExec.TesterMode = testModeOffline Then
                MeasureImped.Pins(CStr(pin)).value(0) = MeasureImped.Pins(CStr(pin)).value(0) + 10
                MeasureImped.Pins(CStr(pin)).value(1) = MeasureImped.Pins(CStr(pin)).value(1) - 10
            End If
            
            For Each site In TheExec.sites
                OutputTrimCode = vbNullString
                For k = 0 To InDSPWave(site).SampleSize - 1
                    OutputTrimCode = OutputTrimCode & CStr(InDSPWave(site).Element(k))
                Next k
                If gl_Disable_HIP_debug_log = False Then TheExec.Datalog.WriteComment ("Pre Check Min and Max Trim Code, Site_" & site & " Initial Output Trim Code = " & OutputTrimCode)
            Next site
            
            For Each site In TheExec.sites
                TheExec.Datalog.WriteComment ("Pre Check Min and Max Trim Code, Site " & site & " Pin " & CStr(pin) & " Impedence = " & FormatNumber((MeasureImped.Pins(0).value(site)), 3) & " Ohm")
                If Abs(TrimmedImpedance(PinNumber) - TrimTarget) > Abs(MeasureImped.Pins(0).value(site) - TrimTarget) Then
                    TrimmedImpedance(PinNumber) = MeasureImped.Pins(0).value(site)
                    
                     TrimCode(PinNumber) = 0
                     For x = 0 To GetbitNumber - 1
                        TrimCode(PinNumber) = InDSPWave.Element(InDSPWave.SampleSize - x - 1) * 2 ^ (GetbitNumber - 1 - x) + TrimCode(PinNumber)
                    Next x
                End If
            
            Next site
            
            For Each site In TheExec.sites
                If b_HighThanTargetImped(site) = True Then
                    b_KeepGoing(site) = MeasureImped.Math.Subtract(PreviousImped).compare(LessThan, 0)
                Else
                    b_KeepGoing(site) = MeasureImped.Math.Subtract(PreviousImped).compare(GreaterThan, 0)
                End If
            Next site
    
            Dim PreCheckBinStr As String, PreCheckDecVal As Double
            For Each site In TheExec.sites
                If b_KeepGoing(site) = False Then
                    b_StopTrimCodeProcess(site) = True
                    PreCheckBinStr = vbNullString
                    StoredTargetTrimCode(site) = InDSPWave(site).Copy
                    For i = 0 To StoredTargetTrimCode(site).SampleSize - 1
                        PreCheckBinStr = PreCheckBinStr & StoredTargetTrimCode.Element(i)
                    Next i
                    PreCheckDecVal = Bin2Dec_rev_Double(PreCheckBinStr)
    
                End If
            Next site
        End If
        
        ''========================================================================================
        Dim b_ControlNextBit As Boolean
        b_ControlNextBit = False
        Dim b_FirstExecution As Boolean
        b_FirstExecution = False
        StoreEachIndex = 1

        ''20170103-Setup b_KeepGoing to true if PreCheckMinMaxTrimCode=false
        If PreCheckMinMaxTrimCode = False Then
            b_KeepGoing = True
        End If
        
        If b_KeepGoing.All(False) Then
        Else
    
            For i = 0 To ExecutionMax
                If TrimPrcocessAll = False Then
                    If b_StopTrimCodeProcess.All(True) Then
                        Exit For
                    End If
                End If
                StrSeparatebyEqual = Split(StrSeparatebyComma(i), "=")
                StrSeparatebyColon = Split(StrSeparatebyEqual(1), ":")
                SrcStartBit = StrSeparatebyColon(0)
                SrcEndBit = StrSeparatebyColon(1)
                If b_HasFIXED_InTrimFormat Then
                    If UBound(StrSeparatebyColon) = 1 Then
                        If i = 0 Then
                            b_FirstExecution = True
                        Else
                            b_FirstExecution = False
                            SrcStartBit = SrcStartBit
                        End If
                    End If
                
                Else
                    If i = 0 Then
                        b_FirstExecution = True
                    Else
                        b_FirstExecution = False
                        SrcStartBit = SrcStartBit + 1
                    End If
                End If
                If b_ExecuteTrimCode(i) = True Then
                    For j = SrcStartBit To SrcEndBit Step -1
                    
                        If b_FirstExecution = True Then
                            b_ControlNextBit = True
                            If j = SrcEndBit Then
                                b_ControlNextBit = False
                            End If
                        Else
                        ''20160716-Control next bit to 1 no matter first or last progress
                            b_ControlNextBit = True
            ''                b_ControlNextBit = False
                            If j = SrcEndBit Then
                                b_ControlNextBit = False
                            End If
                        End If
            
''                        If b_FirstExecution = True And j = SrcEndBit Then
                        If j = SrcEndBit Then
                            Call rundsp.SetupTrimCodeBit(PastDSPWave, True, j, b_ControlNextBit, InDSPWave)
            ''            ElseIf b_FirstExecution = False And j = SrcStartBit Then
            ''                j = SrcStartBit + 1
                        Else
                            Call rundsp.SetupTrimCodeBit(PastDSPWave, b_HighThanTargetImped, j, b_ControlNextBit, InDSPWave)
                        End If
                        
                        Call SetupDigSrcDspWave(PattArray(0), DigSrc_pin, "TrimCodeImped", DigSrc_Sample_Size, InDSPWave)
                        
                        For Each site In TheExec.sites
                            StoreEachTrimCode(StoreEachIndex)(site) = InDSPWave(site).Copy
                        Next site
                    
                        '' Debug use
                        '' ==============================================================================================
                        '' 20160716 - Modify trim code rule
                        
                        If b_FirstExecution = True And gl_Disable_HIP_debug_log = False Then
                            If j = SrcEndBit Then
                                TheExec.Datalog.WriteComment ("Setup Bit " & j & " to 0")
                            Else
                                TheExec.Datalog.WriteComment ("Setup Bit " & j & ", Trim Code Bit " & j - 1)
                            End If
                        ElseIf gl_Disable_HIP_debug_log = False Then
                            If j = SrcEndBit Then
                                TheExec.Datalog.WriteComment ("Setup Bit " & j)
                            Else
                                TheExec.Datalog.WriteComment ("Setup Bit " & j & ", Trim Code Bit " & j - 1)
                            End If
                        End If
                        
                        For Each site In TheExec.sites
        
                            If b_KeepGoing(site) = True Then
                                OutputTrimCode = vbNullString
                                For k = 0 To InDSPWave(site).SampleSize - 1
                                    OutputTrimCode = OutputTrimCode & CStr(InDSPWave(site).Element(k))
                                Next k
                                
                                If gl_Disable_HIP_debug_log = False Then TheExec.Datalog.WriteComment ("Site_" & site & " Output Trim Code = " & OutputTrimCode)
                            End If
        
                        Next site
                        '' ==============================================================================================
                        
                        Call TheHdw.Patterns(PattArray(0)).start
        
                        Call SubMeasR(CPUA_Flag_In_Pat, CStr(pin), ForceVolt, MeasureImped, b_IsDifferential, b_PD_Mode)
                        
                        If TheExec.TesterMode = testModeOffline Then
                            Dim SimuIndex As Long
                            SimuIndex = TestLimitIndex
                            If SimuIndex >= 3 Then
                                SimuIndex = 3
                            End If
    
                            MeasureImped.Pins(CStr(pin)).value(0) = MeasureImped.Pins(CStr(pin)).value(0) + (SimuIndex * 1) + ((PerPinIndex + 1) * 1 / (PerPinIndex + 1))
                            MeasureImped.Pins(CStr(pin)).value(1) = MeasureImped.Pins(CStr(pin)).value(1) - (SimuIndex * 1) + ((PerPinIndex + 1) * 1 / (PerPinIndex + 1))
    
                        End If
                        
                        If j = SrcEndBit + 1 Then
                            MeasureImped_F1 = MeasureImped
                        ElseIf j = SrcEndBit Then
                            MeasureImped_F2 = MeasureImped
                        End If
                        
                        StoreEachTrimImped(StoreEachIndex) = MeasureImped
                        StoreEachIndex = StoreEachIndex + 1
                        
                        If j = SrcEndBit Then
                            b_HighThanTargetImped = False
                            b_HighThanTargetImped = MeasureImped_F1.Math.Subtract(TrimTarget).Abs.compare(GreaterThan, MeasureImped_F2.Math.Subtract(TrimTarget).Abs)
''                            Call rundsp.SetupTrimCodeBit(PastDSPWave, b_HighThanTargetImped, j, b_ControlNextBit, InDSPWave)
                            PastDSPWave = InDSPWave
                        Else
                            b_HighThanTargetImped = False
                            b_HighThanTargetImped = MeasureImped.Math.Subtract(TrimTarget).compare(LessThan, 0)
                            PastDSPWave = InDSPWave
                        End If
            
                        TestLimitIndex = TestLimitIndex + 1
                        
                        '' 20160712 - Modify to use WriteComment to display output Impeduency.
                        For Each site In TheExec.sites
                            If b_KeepGoing(site) = True Then
                                If gl_Disable_HIP_debug_log = False Then TheExec.Datalog.WriteComment ("Site " & site & " Pin " & CStr(pin) & " Impedence = " & FormatNumber((MeasureImped.Pins(0).value(site)), 3) & " Ohm")
                                If Abs(TrimmedImpedance(PinNumber) - TrimTarget) > Abs(MeasureImped.Pins(0).value(site) - TrimTarget) Then
                                    
                                    TrimCode(PinNumber) = 0
                                    TrimmedImpedance(PinNumber) = MeasureImped.Pins(0).value(site)
                                     For x = 0 To GetbitNumber - 1
                                        TrimCode(PinNumber) = InDSPWave.Element(InDSPWave.SampleSize - x - 1) * 2 ^ (GetbitNumber - 1 - x) + TrimCode(PinNumber)
                                    Next x
                                End If
                            End If
                        Next site
                        
                        ''20160716 - Modify display info sequence when source bit in the section end
                        If j = SrcEndBit Then
                            For Each site In TheExec.sites
        
                                If b_KeepGoing(site) = True And gl_Disable_HIP_debug_log = False Then
                                    TheExec.Datalog.WriteComment ("Site " & site & " Pin " & CStr(pin) & " R" & LastSectionF1F2_Index + 1 & " Impedence = " & FormatNumber((MeasureImped_F1.Pins(0).value(site)), 3) & " Ohm")
                                    TheExec.Datalog.WriteComment ("Site " & site & " Pin " & CStr(pin) & " R" & LastSectionF1F2_Index + 2 & " Impedence = " & FormatNumber((MeasureImped_F2.Pins(0).value(site)), 3) & " Ohm")
                                End If
        
                            Next site
                            LastSectionF1F2_Index = LastSectionF1F2_Index + 2
                        End If
                        
                        '' 20160712 - Compare Measure Impeduency whether match target Imped
                        b_MatchTagetImped = MeasureImped.Math.Subtract(TrimTarget).Abs.compare(LessThanOrEqualTo, TrimTargetTolerance)
                        b_DisplayImped = b_DisplayImped.LogicalOr(b_MatchTagetImped)
                        For Each site In TheExec.sites
                            If b_KeepGoing(site) = True Then
                                If b_MatchTagetImped(site) = True And StoredTargetTrimCode(site).CalcSum = 0 Then
                                    StoredTargetTrimCode(site) = InDSPWave(site).Copy
                                    b_StopTrimCodeProcess(site) = True
                                End If
                            End If
                        Next site
                        ''20161128-Stop trim code process if found out match code of all site
                        If TrimPrcocessAll = False Then
                            If b_StopTrimCodeProcess.All(True) Then
                                Exit For
                            End If
                        End If
                        If gl_Disable_HIP_debug_log = False Then TheExec.Datalog.WriteComment ("======================================================================================")
                    Next j
                End If
            Next i
        End If
    
        ''============================================================================
        ''20161128 Findout mimiumn trim code
        Dim CloseTargetImped As New PinListData
        Dim DiffValue As New SiteDouble, PreviousDiffValue As New SiteDouble, CloseIndex As New SiteLong
        
        Dim b_UseMinTrim As New SiteBoolean, MinDiffVal As New SiteDouble
        Dim binstr As String
        Dim CloseTargetTrimCode As New DSPWave
        Dim DecVal As Double, PreviousDecVal As Double, MinDecVal As Double
        Dim b_FirstTimeSwitch As Boolean
        
        
        If b_KeepGoing.All(False) Then
        Else
''            If PerPinIndex = 0 Then
                Set CloseTargetTrimCode = Nothing
''                If Fixed_DigSrc_Equation <> "" Then
''                    CloseTargetTrimCode.CreateConstant 0, Trim_DigSrc_Sample_Size, DspLong
''                Else
                    CloseTargetTrimCode.CreateConstant 0, DigSrc_Sample_Size, DspLong
''                End If
''            Else
''                For Each Site In TheExec.Sites
''                    CloseTargetTrimCode.Clear
''                Next Site
''            End If
            
            For Each site In TheExec.sites
                If b_KeepGoing(site) = True Then
                    If StoredTargetTrimCode(site).CalcSum = 0 Then
                        b_UseMinTrim(site) = True
                    End If
                End If
            Next site
                
            If UseMinimumTrimCode = True Then
                b_UseMinTrim = True
            End If
                
            For Each site In TheExec.sites
                If b_KeepGoing(site) = True Then
                    If b_UseMinTrim(site) = True Then
                        '' Findout minimum difference value
    
                        For i = 0 To StoreEachIndex - 1
                            DiffValue(site) = Abs(StoreEachTrimImped(i).Pins(0).value(site) - TrimTarget)
                            If DiffValue(site) <= PreviousDiffValue(site) Then
                                CloseIndex(site) = i
                                PreviousDiffValue(site) = DiffValue(site)
                                MinDiffVal(site) = DiffValue(site)
                            End If
                            If i = 0 Then
                                PreviousDiffValue(site) = DiffValue(site)
                                MinDiffVal(site) = DiffValue(site)
                            End If
                        Next i
                        '' Transfer to decimal value to findout minimum code
                        PreviousDecVal = 0
                        DecVal = 0
                        b_FirstTimeSwitch = False
    
                        For i = 0 To StoreEachIndex - 1
                            binstr = vbNullString
                            If Abs(StoreEachTrimImped(i).Pins(0).value(site) - TrimTarget) <= MinDiffVal(site) Then
                                For j = 0 To StoreEachTrimCode(i)(site).SampleSize - 1
                                    binstr = binstr & StoreEachTrimCode(i)(site).Element(j)
                                Next j
                                DecVal = Bin2Dec_rev_Double(binstr)
                               
                                If DecVal < PreviousDecVal Then
                                    MinDecVal = DecVal
                                    CloseTargetTrimCode(site) = StoreEachTrimCode(i)(site).Copy
                                End If
                                PreviousDecVal = DecVal
                                If b_FirstTimeSwitch = False Then
                                    CloseTargetTrimCode(site) = StoreEachTrimCode(i)(site).Copy
                                    b_FirstTimeSwitch = True
                                End If
                            End If
                        Next i
                    End If
                End If
            Next site
        End If
        
        For Each site In TheExec.sites
            If b_KeepGoing(site) = True Then
                If b_UseMinTrim(site) = True Then
                    StoredTargetTrimCode(site) = CloseTargetTrimCode(site).Copy
                Else
                    StoredTargetTrimCode(site) = StoredTargetTrimCode(site).Copy
                End If
            Else
                StoredTargetTrimCode(site) = StoredTargetTrimCode(site).Copy
            End If
        Next site
        
        For Each site In TheExec.sites
            OutputTrimCode = vbNullString
            For k = 0 To StoredTargetTrimCode(site).SampleSize - 1
                OutputTrimCode = OutputTrimCode & CStr(StoredTargetTrimCode(site).Element(k))
            Next k
            If gl_Disable_HIP_debug_log = False Then TheExec.Datalog.WriteComment ("Site_" & site & " Pin " & CStr(pin) & " Final Output Trim Code = " & OutputTrimCode)
        Next site
        
        ''20170103-Store per pin final trim code
        StorePerPinFinalTrimCode(PerPinIndex) = StoredTargetTrimCode
        PerPinIndex = PerPinIndex + 1
        LastSectionF1F2_Index = 0
        
        PinNumber = PinNumber + 1
        
        ''20170210-Specified item to store trim code to fuse by 1 pin
        If b_wkds_Store_Flag Then
            If Fixed_DigSrc_Equation <> "" Then
                Call rundsp.SelectCertainBitsToDec(StoredTargetTrimCode, Fixed_DigSrc_Sample_Size, Trim_DigSrc_Sample_Size, wkds_StoreTrimCodeToDict_DEC)
                Call rundsp.DSPWaveDecToBinary(wkds_StoreTrimCodeToDict_DEC, Trim_DigSrc_Sample_Size, wkds_StoreTrimCodeToDict_BIN)
                b_wkds_Store_Flag = False
            End If
        End If
    Next pin
    ''============================================================================
    ''20161230 - Average per pin trim code and store it
    Dim TrimCodeForTotalPin_Dec As New DSPWave
    Dim TrimCodeForTotalPin_Bin As New DSPWave
    TrimCodeForTotalPin_Dec.CreateConstant 0, 1, DspLong
    If Fixed_DigSrc_Equation <> "" Then
        TrimCodeForTotalPin_Bin.CreateConstant 0, Trim_DigSrc_Sample_Size, DspLong
    Else
        TrimCodeForTotalPin_Bin.CreateConstant 0, DigSrc_Sample_Size, DspLong
    End If
    For i = 0 To UBound(StorePerPinFinalTrimCode_Dec)
    
        If Fixed_DigSrc_Equation <> "" Then
            '' 20170117-only select trim code,  not select all to convert decimal
            Call rundsp.SelectCertainBitsToDec(StorePerPinFinalTrimCode(i), Fixed_DigSrc_Sample_Size, Trim_DigSrc_Sample_Size, StorePerPinFinalTrimCode_Dec(i))
        Else
            Call rundsp.ConvertToLongAndSerialToParrel(StorePerPinFinalTrimCode(i), DigSrc_Sample_Size, StorePerPinFinalTrimCode_Dec(i))
        End If
        Call rundsp.DSP_Add(TrimCodeForTotalPin_Dec, StorePerPinFinalTrimCode_Dec(i))
    Next i
    Dim DenominatorConstant As Double
    DenominatorConstant = UBound(StorePerPinFinalTrimCode_Dec) + 1
    If DenominatorConstant = 0 Then
        TheExec.Datalog.WriteComment ("Error! Divide 0.")
        Exit Function
    End If
''    Call rundsp.DSP_DivideConstant(TrimCodeForTotalPin_Dec, DenominatorConstant)
    For Each site In TheExec.sites
        TrimCodeForTotalPin_Dec(site).Element(0) = Int(CDbl(TrimCodeForTotalPin_Dec(site).Element(0) / DenominatorConstant) + 0.5)
    Next site
    
''    For Each Site In TheExec.Sites
''        If TrimCodeForTotalPin_Dec(Site).Element(0) > 192 Then
''            TrimCodeForTotalPin_Dec(Site).Element(0) = TrimCodeForTotalPin_Dec(Site).Element(0) - 192
''        End If
''    Next Site
    
    If Fixed_DigSrc_Equation <> "" Then
        Call rundsp.DSPWaveDecToBinary(TrimCodeForTotalPin_Dec, Trim_DigSrc_Sample_Size, TrimCodeForTotalPin_Bin)
        ''20170214 - If trim code size = 8 , set element 7 and 6 to 0.
        For Each site In TheExec.sites
            If TrimCodeForTotalPin_Bin(site).SampleSize = 8 Then
                TrimCodeForTotalPin_Bin(site).Element(6) = 0
                TrimCodeForTotalPin_Bin(site).Element(7) = 0
            End If
        Next site
    Else
        Call rundsp.DSPWaveDecToBinary(TrimCodeForTotalPin_Dec, DigSrc_Sample_Size, TrimCodeForTotalPin_Bin)
    End If
    ''============================================================================
    ''20170104 - TestLimit for found code and measured R for each pin
    i = 0
    Dim TrimCodeDSP_DEC() As New DSPWave
    ReDim TrimCodeDSP_DEC(Pin_Cnt) As New DSPWave
    Dim TrimCodeDSP_BIN() As New DSPWave
    ReDim TrimCodeDSP_BIN(Pin_Cnt) As New DSPWave
    
    For Each pin In Pin_Ary
        Set TrimCodeDSP_DEC(i) = Nothing
        TrimCodeDSP_DEC(i).CreateConstant 0, 1, DspLong
        For Each site In TheExec.sites
''            If TrimCode(i) > 192 Then
''                TrimCode(i) = TrimCode(i) - 192
''            End If

''          ''20170214 - If trim code size = 8 , set element 7 and 6 to 0.
            TrimCodeDSP_DEC(i)(site).Element(0) = TrimCode(i)(site)
        Next site
        Call rundsp.DSPWaveDecToBinary(TrimCodeDSP_DEC(i), GetbitNumber, TrimCodeDSP_BIN(i))
        If GetbitNumber = 8 Then
            For Each site In TheExec.sites
                TrimCodeDSP_BIN(i)(site).Element(6) = 0
                TrimCodeDSP_BIN(i)(site).Element(7) = 0
            Next site
        End If
        Call rundsp.BinToDec(TrimCodeDSP_BIN(i), TrimCodeDSP_DEC(i))
''        TheExec.Flow.TestLimit TrimCode(i), , , , , , , , CStr(pin), , , , , , , tlForceFlow
        
        TestNameInput = Report_TName_From_Instance("C", vbNullString, vbNullString, 0)
        TheExec.Flow.TestLimit TrimCodeDSP_DEC(i).Element(0), , , , , , , , TestNameInput, , , , , , , tlForceNone 'eng_forceflow_transfer

        
        TestNameInput = Report_TName_From_Instance("R", vbNullString, vbNullString, 0)
        TheExec.Flow.TestLimit TrimmedImpedance(i), , , , , , unitCustom, , TestNameInput, , , , , "ohm", , tlForceNone 'eng_forceflow_transfer

        i = i + 1
    Next pin
    ''============================================================================
    If TrimStoreName <> "" Then
        ''20170210-Specified item to store trim code to fuse by 1 pin
        If LCase(TrimStoreName) Like LCase("wkds_*") Then

            Call Checker_StoreDigCapAllToDictionary(TrimStoreName, wkds_StoreTrimCodeToDict_BIN)
        Else
            Call Checker_StoreDigCapAllToDictionary(TrimStoreName, TrimCodeForTotalPin_Bin)
        End If
    End If
    
    Call HardIP_WriteFuncResult



''    For Each site In TheExec.Sites
''        If TrimCodeForTotalPin_Dec(site).Element(0) > 192 Then
''            TrimCodeForTotalPin_Dec(site).Element(0) = TrimCodeForTotalPin_Dec(site).Element(0) - 192
''        End If
''    Next site
    ''20170214 - If trim code size = 8 , set element 7 and 6 to 0.
    Dim TEMP_TrimCodeForTotalPin_Bin As New DSPWave
    Call rundsp.DSPWaveDecToBinary(TrimCodeForTotalPin_Dec, Trim_DigSrc_Sample_Size, TEMP_TrimCodeForTotalPin_Bin)
    For Each site In TheExec.sites
        If TEMP_TrimCodeForTotalPin_Bin(site).SampleSize = 8 Then
            TEMP_TrimCodeForTotalPin_Bin(site).Element(6) = 0
            TEMP_TrimCodeForTotalPin_Bin(site).Element(7) = 0
        End If
    Next site
    Call rundsp.BinToDec(TEMP_TrimCodeForTotalPin_Bin, TrimCodeForTotalPin_Dec)
    
    TestNameInput = Report_TName_From_Instance("C", vbNullString, "AverageTrimCode", 0)
    TheExec.Flow.TestLimit TrimCodeForTotalPin_Dec.Element(0), 0, 2 ^ Trim_DigSrc_Sample_Size - 1, Tname:=TestNameInput, ForceResults:=tlForceFlow
    
    ''20170213-Source average trim code to measure impedence again
    If Fixed_DigSrc_Equation <> "" Then
        Call rundsp.CombineDSPWave(InitialFixedDSPWave, InitialTrimDSPWave, Fixed_DigSrc_Sample_Size, Trim_DigSrc_Sample_Size, InDSPWave)
    End If

    '' 20170210 - Source final code to do re-measurement
''    If Fixed_DigSrc_Equation <> "" Then
''        Call rundsp.CombineDSPWave(InitialFixedDSPWave, TrimCodeForTotalPin_Bin, Fixed_DigSrc_Sample_Size, Trim_DigSrc_Sample_Size, InDSPWave)
''    End If
''    For Each pin In pin_ary
''        Call SetupDigSrcDspWave(PattArray(0), DigSrc_Pin, "TrimCodeImped", DigSrc_Sample_Size, InDSPWave)
''
''        Call thehdw.Patterns(PattArray(0)).start
''        Call SubMeasR(CPUA_Flag_In_Pat, CStr(pin), ForceVolt, MeasureImped, b_IsDifferential, b_PD_Mode)
''        TheExec.Flow.TestLimit resultVal:=MeasureImped, unit:=unitCustom, customUnit:="ohm", Tname:="SourceAverCode" & "_Pin_" & pin, ForceResults:=tlForceFlow
''
''    Next pin
    
    Dim SplitByComma() As String
    Dim DictName_FUSE As String
    Dim sl_FUSE_Val As New SiteLong
    '' 20170119 - Process calculate equation by dictionary.
    If Calc_Eqn <> "" Then
        Call ProcessCalcEquation(Calc_Eqn)
        If UCase(Calc_Eqn) Like UCase("*ADDRIO_TrimCodeAverage*") Then
            SplitByComma() = Split(Calc_Eqn, ",")
            DictName_FUSE = SplitByComma(UBound(SplitByComma))
            DictName_FUSE = Replace(DictName_FUSE, ")", vbNullString)
            Call DictDSPToSiteLong(DictName_FUSE, sl_FUSE_Val, TrimFuseName)
        End If
    End If
    '' 20170704- Comment this
''    If TrimFuseName <> "" And TrimFuseName <> "addr-wkds_u" Then
''        Call HIP_eFuse_Write("ECID", TrimFuseName, sl_FUSE_Val)
''    End If
    If TheExec.TesterMode = testModeOffline Then
    Else
        If TrimFuseName <> "" And TrimFuseTypeName <> "" Then
            ''Call HIP_eFuse_Write(TrimFuseTypeName, TrimFuseName, sl_FUSE_Val) ''set fuse information from flow
        End If
    End If
    
    If Interpose_PrePat <> "" Then
        Call SetForceCondition("RESTOREPREPAT")
    End If
    
    '' NOTE : Efuse write TrimCodeForTotalPin_Dec ( DSPWave )
    
''    Dim ConvertedDataWf As New DSPWave
''    rundsp.ConvertToLongAndSerialToParrel StoredTargetTrimCode, DigSrc_Sample_Size, ConvertedDataWf

''    Call SetupDigSrcDspWave(PattArray(0), DigSrc_Pin, "TrimCodeImped", DigSrc_Sample_Size, TrimCodeForTotalPin_Bin)
''
''    Call TheHdw.Patterns(PattArray(0)).start
''
''    Call SubMeasR(CPUA_Flag_In_Pat, MeasR_Pins_SingleEnd.Value, ForceVolt, MeasureImped, b_IsDifferential)
''
''    TheExec.Flow.TestLimit resultVal:=MeasureImped, unit:=unitCustom, customUnit:="ohm", Tname:=TestNameInput & "Final", ForceResults:=tlForceFlow
    
    
''    If TheExec.TesterMode = testModeOffline Then
''    Else
''        '' eFUSE
''        For Each Site In TheExec.Sites.Active
''            Dim PassFlag_LPRO As New SiteBoolean
''
''            If CurrentJobName_U Like "*FT*" Then
''                TheExec.Datalog.WriteComment ""
''                ConvertedDataWf(Site).Element(0) = auto_eFuse_GetReadDecimal("ECID", "OSC", True)
''                TheExec.Datalog.WriteComment ""
''                For i = 0 To DigSrc_Sample_Size - 1
''                    '' 20161110 - Hint! Need to check Src_DSPWave
''                    Src_DSPWave(Site).Element(i) = ConvertedDataWf(Site).Element(0) Mod 2
''                    ConvertedDataWf(Site).Element(0) = ConvertedDataWf(Site).Element(0) \ 2
''                Next i
''            Else
''                If TheHdw.Digital.Patgen.PatternBurstPassed(Site) = False Then 'Pattern Fail
''                    PassFlag_LPRO(Site) = False
''                Else
''                    PassFlag_LPRO(Site) = True
''                End If
''
''                If UCase(TrimFuseName) = "OSC" Then
''                    Call auto_eFuse_SetPatTestPass_Flag("ECID", "OSC", PassFlag_LPRO(Site))
''                    Call auto_eFuse_SetWriteDecimal("ECID", "OSC", ConvertedDataWf(Site).Element(0))
''                End If
''            End If
''        Next Site
''    End If
    
    Exit Function
    
ErrorHandler:
    TheExec.Datalog.WriteComment "error in TrimCodeImped function"
    If AbortTest Then Exit Function Else Resume Next
    
    
End Function


Public Function MTR_Verification_Calculate(SensorName As String, Temperature As String, FusedCoeffDicName_1 As String, FusedCoeffDicName_2 As String, SetInformation As DSPWave, _
                                        Aininformation As DSPWave, Aixinformation As DSPWave, PiUInformation As DSPWave, LevelsRecord As String) As Long
 
        Dim i As Long
        Dim Row As Long
        Dim Col As Long
        Dim bitIter As Long
        Dim tmpCount As Long
        Dim TestName As String
        Dim ScanOffset As Integer
        Dim ElementCNT As Integer
        Dim PiUCntTempA As Integer
        Dim PiUCntTempB As Integer
        Dim VoltageName() As String
        Dim tmpValue_ROT  As Double
        Dim tmpValue_ROV  As Double
        Dim InheritanceCnt As Integer
        Dim temp_rowVal_f As New SiteDouble
        Dim decimalPlaces As Long: decimalPlaces = 4
        VoltageName = Split(LevelsRecord, ",")
        
        Dim actualROTMatrix As New DSPWave
        Dim actualROVMatrix As New DSPWave
        Dim RotMatrixDicName As String
        Dim RovMatrixDicName As String
    
        Dim FusedCoeffDicNameTemp_1 As String
        Dim FusedCoeffDicNameTemp_2 As String
        Dim readFusedCoeffDspWave1 As New DSPWave
        Dim readFusedCoeffDspWave2 As New DSPWave
        
        Dim Fused_ROT_Decimal_Vector As New DSPWave
        Dim Fused_ROV_Decimal_Vector As New DSPWave
        Fused_ROT_Decimal_Vector.CreateConstant 0, 4, DspDouble
        Fused_ROV_Decimal_Vector.CreateConstant 0, 3, DspDouble
                
        Dim Output_ROT_Freq_Vector As New DSPWave
        Dim Output_ROV_Freq_Vector As New DSPWave
        Output_ROT_Freq_Vector.CreateConstant 0, 20, DspDouble
        Output_ROV_Freq_Vector.CreateConstant 0, 20, DspDouble
            
        Dim Difference_ROT_Freq_Vector As New DSPWave
        Dim Difference_ROV_Freq_Vector As New DSPWave
        Difference_ROT_Freq_Vector.CreateConstant 0, 20, DspDouble
        Difference_ROV_Freq_Vector.CreateConstant 0, 20, DspDouble
    
        RotMatrixDicName = "Freq" + "_" + SensorName + "_" + "rot" + "_" + Temperature + "c"
        RovMatrixDicName = "Freq" + "_" + SensorName + "_" + "rov" + "_" + Temperature + "c"
    
        actualROTMatrix = GetStoredCaptureData(RotMatrixDicName)
        actualROVMatrix = GetStoredCaptureData(RovMatrixDicName)
    
        FusedCoeffDicNameTemp_1 = FusedCoeffDicName_1 + "_" + "Freq" + "_" + SensorName + "_" + "rot" + "_" + Temperature + "c"
        FusedCoeffDicNameTemp_2 = FusedCoeffDicName_2 + "_" + "Freq" + "_" + SensorName + "_" + "rov" + "_" + Temperature + "c"
        readFusedCoeffDspWave1 = GetStoredCaptureData(FusedCoeffDicNameTemp_1)
        readFusedCoeffDspWave2 = GetStoredCaptureData(FusedCoeffDicNameTemp_2)
        
        If Temperature = 25 Then
            For Each site In TheExec.sites
                tmpValue_ROT = 0
                tmpValue_ROV = 0
                tmpCount = 0
                For bitIter = 0 To readFusedCoeffDspWave1.SampleSize - 1
                    If bitIter = 14 Then
                        tmpValue_ROT = tmpValue_ROT + (CDbl(readFusedCoeffDspWave1(site).Element(readFusedCoeffDspWave1.SampleSize - bitIter - 1))) / (2 ^ (tmpCount + 1))
                        Fused_ROT_Decimal_Vector(site).Element(0) = tmpValue_ROT * (Aixinformation.Element(0) - Aininformation.Element(0)) + Aininformation.Element(0)
                        tmpValue_ROT = 0
                        tmpCount = 0
                    ElseIf bitIter = 28 Then
                        tmpValue_ROT = tmpValue_ROT + (CDbl(readFusedCoeffDspWave1(site).Element(readFusedCoeffDspWave1.SampleSize - bitIter - 1))) / (2 ^ (tmpCount + 1))
                        Fused_ROT_Decimal_Vector(site).Element(1) = tmpValue_ROT * (Aixinformation.Element(1) - Aininformation.Element(1)) + Aininformation.Element(1)
                        tmpValue_ROT = 0
                        tmpCount = 0
                    ElseIf bitIter = 42 Then
                        tmpValue_ROT = tmpValue_ROT + (CDbl(readFusedCoeffDspWave1(site).Element(readFusedCoeffDspWave1.SampleSize - bitIter - 1))) / (2 ^ (tmpCount + 1))
                        Fused_ROT_Decimal_Vector(site).Element(2) = tmpValue_ROT * (Aixinformation.Element(2) - Aininformation.Element(2)) + Aininformation.Element(2)
                        tmpValue_ROT = 0
                        tmpCount = 0
                    ElseIf bitIter = 56 Then
                        tmpValue_ROT = tmpValue_ROT + (CDbl(readFusedCoeffDspWave1(site).Element(readFusedCoeffDspWave1.SampleSize - bitIter - 1))) / (2 ^ (tmpCount + 1))
                        Fused_ROT_Decimal_Vector(site).Element(3) = tmpValue_ROT * (Aixinformation.Element(3) - Aininformation.Element(3)) + Aininformation.Element(3)
                        tmpValue_ROT = 0
                        tmpCount = 0
                    Else
                        tmpValue_ROT = tmpValue_ROT + (CDbl(readFusedCoeffDspWave1(site).Element(readFusedCoeffDspWave1.SampleSize - bitIter - 1))) / (2 ^ (tmpCount + 1))
                        tmpCount = tmpCount + 1
                    End If
                Next bitIter
                tmpValue_ROT = 0
                tmpValue_ROV = 0
                tmpCount = 0
                For bitIter = 0 To readFusedCoeffDspWave2.SampleSize - 1
                    If bitIter = 14 Then
                        tmpValue_ROV = tmpValue_ROV + (CDbl(readFusedCoeffDspWave2(site).Element(readFusedCoeffDspWave2.SampleSize - bitIter - 1))) / (2 ^ (tmpCount + 1))
                        Fused_ROV_Decimal_Vector(site).Element(0) = tmpValue_ROV * (Aixinformation.Element(4) - Aininformation.Element(4)) + Aininformation.Element(4)
                        tmpValue_ROV = 0
                        tmpCount = 0
                    ElseIf bitIter = 28 Then
                        tmpValue_ROV = tmpValue_ROV + (CDbl(readFusedCoeffDspWave2(site).Element(readFusedCoeffDspWave2.SampleSize - bitIter - 1))) / (2 ^ (tmpCount + 1))
                        Fused_ROV_Decimal_Vector(site).Element(1) = tmpValue_ROV * (Aixinformation.Element(5) - Aininformation.Element(5)) + Aininformation.Element(5)
                        tmpValue_ROV = 0
                        tmpCount = 0
                    ElseIf bitIter = 42 Then
                        tmpValue_ROV = tmpValue_ROV + (CDbl(readFusedCoeffDspWave2(site).Element(readFusedCoeffDspWave2.SampleSize - bitIter - 1))) / (2 ^ (tmpCount + 1))
                        Fused_ROV_Decimal_Vector(site).Element(2) = tmpValue_ROV * (Aixinformation.Element(6) - Aininformation.Element(6)) + Aininformation.Element(6)
                        tmpValue_ROV = 0
                        tmpCount = 0
                    Else
                        tmpValue_ROV = tmpValue_ROV + (CDbl(readFusedCoeffDspWave2(site).Element(readFusedCoeffDspWave2.SampleSize - bitIter - 1))) / (2 ^ (tmpCount + 1))
                        tmpCount = tmpCount + 1
                    End If
                Next bitIter
                tmpValue_ROT = 0
                tmpValue_ROV = 0
                tmpCount = 0
            Next site
        ElseIf Temperature = 85 Then
            For Each site In TheExec.sites
                tmpValue_ROT = 0
                tmpValue_ROV = 0
                tmpCount = 0
                For bitIter = 0 To readFusedCoeffDspWave1.SampleSize - 1
                    If bitIter = 14 Then
                        tmpValue_ROT = tmpValue_ROT + (CDbl(readFusedCoeffDspWave1(site).Element(readFusedCoeffDspWave1.SampleSize - bitIter - 1))) / (2 ^ (tmpCount + 1))
                        Fused_ROT_Decimal_Vector(site).Element(0) = tmpValue_ROT * (Aixinformation.Element(7) - Aininformation.Element(7)) + Aininformation.Element(7)
                        tmpValue_ROT = 0
                        tmpCount = 0
                    ElseIf bitIter = 28 Then
                        tmpValue_ROT = tmpValue_ROT + (CDbl(readFusedCoeffDspWave1(site).Element(readFusedCoeffDspWave1.SampleSize - bitIter - 1))) / (2 ^ (tmpCount + 1))
                        Fused_ROT_Decimal_Vector(site).Element(1) = tmpValue_ROT * (Aixinformation.Element(8) - Aininformation.Element(8)) + Aininformation.Element(8)
                        tmpValue_ROT = 0
                        tmpCount = 0
                    ElseIf bitIter = 42 Then
                        tmpValue_ROT = tmpValue_ROT + (CDbl(readFusedCoeffDspWave1(site).Element(readFusedCoeffDspWave1.SampleSize - bitIter - 1))) / (2 ^ (tmpCount + 1))
                        Fused_ROT_Decimal_Vector(site).Element(2) = tmpValue_ROT * (Aixinformation.Element(9) - Aininformation.Element(9)) + Aininformation.Element(9)
                        tmpValue_ROT = 0
                        tmpCount = 0
                    ElseIf bitIter = 56 Then
                        tmpValue_ROT = tmpValue_ROT + (CDbl(readFusedCoeffDspWave1(site).Element(readFusedCoeffDspWave1.SampleSize - bitIter - 1))) / (2 ^ (tmpCount + 1))
                        Fused_ROT_Decimal_Vector(site).Element(3) = tmpValue_ROT * (Aixinformation.Element(10) - Aininformation.Element(10)) + Aininformation.Element(10)
                        tmpValue_ROT = 0
                        tmpCount = 0
                    Else
                        tmpValue_ROT = tmpValue_ROT + (CDbl(readFusedCoeffDspWave1(site).Element(readFusedCoeffDspWave1.SampleSize - bitIter - 1))) / (2 ^ (tmpCount + 1))
                        tmpCount = tmpCount + 1
                    End If
                Next bitIter
                tmpValue_ROT = 0
                tmpValue_ROV = 0
                tmpCount = 0
                For bitIter = 0 To readFusedCoeffDspWave2.SampleSize - 1
                    If bitIter = 14 Then
                        tmpValue_ROV = tmpValue_ROV + (CDbl(readFusedCoeffDspWave2(site).Element(readFusedCoeffDspWave2.SampleSize - bitIter - 1))) / (2 ^ (tmpCount + 1))
                        Fused_ROV_Decimal_Vector(site).Element(0) = tmpValue_ROV * (Aixinformation.Element(11) - Aininformation.Element(11)) + Aininformation.Element(11)
                        tmpValue_ROV = 0
                        tmpCount = 0
                    ElseIf bitIter = 28 Then
                        tmpValue_ROV = tmpValue_ROV + (CDbl(readFusedCoeffDspWave2(site).Element(readFusedCoeffDspWave2.SampleSize - bitIter - 1))) / (2 ^ (tmpCount + 1))
                        Fused_ROV_Decimal_Vector(site).Element(1) = tmpValue_ROV * (Aixinformation.Element(12) - Aininformation.Element(12)) + Aininformation.Element(12)
                        tmpValue_ROV = 0
                        tmpCount = 0
                    ElseIf bitIter = 42 Then
                        tmpValue_ROV = tmpValue_ROV + (CDbl(readFusedCoeffDspWave2(site).Element(readFusedCoeffDspWave2.SampleSize - bitIter - 1))) / (2 ^ (tmpCount + 1))
                        Fused_ROV_Decimal_Vector(site).Element(2) = tmpValue_ROV * (Aixinformation.Element(13) - Aininformation.Element(13)) + Aininformation.Element(13)
                        tmpValue_ROV = 0
                        tmpCount = 0
                    Else
                        tmpValue_ROV = tmpValue_ROV + (CDbl(readFusedCoeffDspWave2(site).Element(readFusedCoeffDspWave2.SampleSize - bitIter - 1))) / (2 ^ (tmpCount + 1))
                        tmpCount = tmpCount + 1
                    End If
                Next bitIter
                tmpValue_ROT = 0
                tmpValue_ROV = 0
                tmpCount = 0
            Next site
        End If
        
        If Temperature = 25 Then
            For Each site In TheExec.sites
                temp_rowVal_f(site) = 0
                ElementCNT = SetInformation.Element(0)
                ScanOffset = (UBound(VoltageName) + 1) - ElementCNT
                For Col = 0 To ElementCNT - 1
                    temp_rowVal_f(site) = 0
                    For Row = 0 To 3
                        PiUCntTempA = (Row * ElementCNT) + Col
                        temp_rowVal_f(site) = temp_rowVal_f(site) + PiUInformation.Element(PiUCntTempA) * Fused_ROT_Decimal_Vector(site).Element(Row)
                    Next Row
                    Output_ROT_Freq_Vector(site).Element(Col + CLng(ScanOffset)) = temp_rowVal_f(site)
    '                    TestName = "Compression_Freq_ROT_" + SensorName + "_" + CStr(VoltageName(col + CLng(ScanOffset)))
    '                    TheExec.Flow.TestLimit resultval:=actualROTMatrix(site).Element(col + ScanOffset), Tname:=TestName, ForceResults:=tlForceFlow
    '                    TestName = "Decompression_Freq_ROT_" + SensorName + "_" + CStr(VoltageName(col + CLng(ScanOffset)))
    '                    TheExec.Flow.TestLimit resultval:=Output_ROT_Freq_Vector(site).Element(col + CLng(ScanOffset)), Tname:=TestName, ForceResults:=tlForceFlow
                    If (actualROTMatrix(site).Element(Col + CLng(ScanOffset)) = 0) Then
                        Difference_ROT_Freq_Vector(site).Element(Col + CLng(ScanOffset)) = 100 * (Output_ROT_Freq_Vector(site).Element(Col + CLng(ScanOffset)) - 0.0001) / 0.0001
                    Else
                        Difference_ROT_Freq_Vector(site).Element(Col + CLng(ScanOffset)) = 100 * (Output_ROT_Freq_Vector(site).Element(Col + CLng(ScanOffset)) - actualROTMatrix(site).Element(Col + CLng(ScanOffset))) / actualROTMatrix(site).Element(Col + CLng(ScanOffset))
                    End If
                Next Col
                temp_rowVal_f(site) = 0
                ElementCNT = SetInformation.Element(0)
                InheritanceCnt = SetInformation.Element(4)
                For Col = 0 To ElementCNT - 1
                    temp_rowVal_f(site) = 0
                    For Row = 0 To 2
                        PiUCntTempB = (Row * ElementCNT) + Col
                        temp_rowVal_f(site) = temp_rowVal_f(site) + PiUInformation.Element(InheritanceCnt + PiUCntTempB) * Fused_ROV_Decimal_Vector(site).Element(Row)
                    Next Row
                    Output_ROV_Freq_Vector(site).Element(Col + CLng(ScanOffset)) = temp_rowVal_f(site)
    '                    TestName = "Compression_Freq_ROV_" + SensorName + "_" + CStr(VoltageName(col + CLng(ScanOffset)))
    '                    TheExec.Flow.TestLimit resultval:=actualROVMatrix(site).Element(col + CLng(ScanOffset)), Tname:=TestName, ForceResults:=tlForceFlow
    '                    TestName = "Decompression_Freq_ROV_" + SensorName + "_" + CStr(VoltageName(col + CLng(ScanOffset)))
    '                    TheExec.Flow.TestLimit resultval:=Output_ROV_Freq_Vector(site).Element(col + CLng(ScanOffset)), Tname:=TestName, ForceResults:=tlForceFlow
                    If (actualROVMatrix(site).Element(Col + CLng(ScanOffset)) = 0) Then
                        Difference_ROV_Freq_Vector(site).Element(Col + CLng(ScanOffset)) = 100 * (Output_ROV_Freq_Vector(site).Element(Col + CLng(ScanOffset)) - 0.0001) / 0.0001
                    Else
                        Difference_ROV_Freq_Vector(site).Element(Col + CLng(ScanOffset)) = 100 * (Output_ROV_Freq_Vector(site).Element(Col + CLng(ScanOffset)) - actualROVMatrix(site).Element(Col + CLng(ScanOffset))) / actualROVMatrix(site).Element(Col + CLng(ScanOffset))
                    End If
                Next Col
            Next site
    
    
            For Col = 0 To UBound(VoltageName)
                    TestName = "Compression_Freq_ROT_" + SensorName + "_" + CStr(VoltageName(Col))
                    TheExec.Flow.TestLimit resultVal:=actualROTMatrix.Element(Col), Tname:=TestName, ForceResults:=tlForceNone, scaletype:=scaleNoScaling 'eng_forceflow_transfer
                    TestName = "Decompression_Freq_ROT_" + SensorName + "_" + CStr(VoltageName(Col))
                    TheExec.Flow.TestLimit resultVal:=Output_ROT_Freq_Vector.Element(Col), Tname:=TestName, ForceResults:=tlForceNone, scaletype:=scaleNoScaling 'eng_forceflow_transfer
                    TestName = "Compression_Freq_ROV_" + SensorName + "_" + CStr(VoltageName(Col))
                    TheExec.Flow.TestLimit resultVal:=actualROVMatrix.Element(Col), Tname:=TestName, ForceResults:=tlForceNone, scaletype:=scaleNoScaling 'eng_forceflow_transfer
                    TestName = "Decompression_Freq_ROV_" + SensorName + "_" + CStr(VoltageName(Col))
                    TheExec.Flow.TestLimit resultVal:=Output_ROV_Freq_Vector.Element(Col), Tname:=TestName, ForceResults:=tlForceNone, scaletype:=scaleNoScaling 'eng_forceflow_transfer
            Next Col
    
            For Col = 0 To UBound(VoltageName)
                TestName = "PercentageDiff_Freq_ROT_" + SensorName + "_" + CStr(VoltageName(Col))
                TheExec.Flow.TestLimit resultVal:=Difference_ROT_Freq_Vector.Element(Col), hiVal:=1, lowVal:=-1, Tname:=TestName, ForceResults:=tlForceNone, scaletype:=scaleNoScaling 'eng_forceflow_transfer
            Next Col
            
            For Col = 0 To UBound(VoltageName)
                TestName = "PercentageDiff_Freq_ROV_" + SensorName + "_" + CStr(VoltageName(Col))
                TheExec.Flow.TestLimit resultVal:=Difference_ROV_Freq_Vector.Element(Col), hiVal:=1, lowVal:=-1, Tname:=TestName, ForceResults:=tlForceNone, scaletype:=scaleNoScaling 'eng_forceflow_transfer
            Next Col
                
    '            For Each site In TheExec.sites
    '                For col = 0 To ElementCNT - 1
    '                    TestName = "PercentageDiff_Freq_ROT_" + SensorName + "_" + CStr(VoltageName(col + CLng(ScanOffset)))
    '                    TheExec.Flow.TestLimit resultval:=Difference_ROT_Freq_Vector.Element(col + CLng(ScanOffset)), hival:=1, lowval:=-1, Tname:=TestName, ForceResults:=tlForceFlow
    '                Next col
                
    '                For col = 0 To ElementCNT - 1
    '                    TestName = "PercentageDiff_Freq_ROV_" + SensorName + "_" + CStr(VoltageName(col + CLng(ScanOffset)))
    '                    TheExec.Flow.TestLimit resultval:=Difference_ROV_Freq_Vector.Element(col + CLng(ScanOffset)), hival:=1, lowval:=-1, Tname:=TestName, ForceResults:=tlForceFlow
    '                Next col
    '            Next site
            
        ElseIf Temperature = 85 Then
            For Each site In TheExec.sites
                temp_rowVal_f(site) = 0
                ElementCNT = SetInformation.Element(0)
                ScanOffset = (UBound(VoltageName) + 1) - ElementCNT
                InheritanceCnt = SetInformation.Element(4) + SetInformation.Element(5)
                For Col = 0 To ElementCNT - 1
                    temp_rowVal_f(site) = 0
                    For Row = 0 To 3
                        PiUCntTempA = (Row * ElementCNT) + Col
                        temp_rowVal_f(site) = temp_rowVal_f(site) + PiUInformation.Element(InheritanceCnt + PiUCntTempA) * Fused_ROT_Decimal_Vector(site).Element(Row)
                    Next Row
                    Output_ROT_Freq_Vector(site).Element(Col + CLng(ScanOffset)) = temp_rowVal_f(site)
    '                    TestName = "Compression_Freq_ROT_" + SensorName + "_" + CStr(VoltageName(col + CLng(ScanOffset)))
    '                    TheExec.Flow.TestLimit resultval:=actualROTMatrix(site).Element(col + CLng(ScanOffset)), Tname:=TestName, ForceResults:=tlForceFlow
    '                    TestName = "Decompression_Freq_ROT_" + SensorName + "_" + CStr(VoltageName(col + CLng(ScanOffset)))
    '                    TheExec.Flow.TestLimit resultval:=Output_ROT_Freq_Vector(site).Element(col + CLng(ScanOffset)), Tname:=TestName, ForceResults:=tlForceFlow
                    If (actualROTMatrix(site).Element(Col + CLng(ScanOffset)) = 0) Then
                        Difference_ROT_Freq_Vector(site).Element(Col + CLng(ScanOffset)) = 100 * (Output_ROT_Freq_Vector(site).Element(Col + CLng(ScanOffset)) - 0.0001) / 0.0001
                    Else
                        Difference_ROT_Freq_Vector(site).Element(Col + CLng(ScanOffset)) = 100 * (Output_ROT_Freq_Vector(site).Element(Col + CLng(ScanOffset)) - actualROTMatrix(site).Element(Col + CLng(ScanOffset))) / actualROTMatrix(site).Element(Col + CLng(ScanOffset))
                    End If
                Next Col
                temp_rowVal_f(site) = 0
                ElementCNT = SetInformation.Element(0)
                InheritanceCnt = SetInformation.Element(4) + SetInformation.Element(5) + SetInformation.Element(6)
                For Col = 0 To ElementCNT - 1
                    temp_rowVal_f(site) = 0
                    For Row = 0 To 2
                        PiUCntTempB = (Row * ElementCNT) + Col
                        temp_rowVal_f(site) = temp_rowVal_f(site) + PiUInformation.Element(InheritanceCnt + PiUCntTempB) * Fused_ROV_Decimal_Vector(site).Element(Row)
                    Next Row
                    Output_ROV_Freq_Vector(site).Element(Col + CLng(ScanOffset)) = temp_rowVal_f(site)
    '                    TestName = "Compression_Freq_ROV_" + SensorName + "_" + CStr(VoltageName(col + CLng(ScanOffset)))
    '                    TheExec.Flow.TestLimit resultval:=actualROVMatrix(site).Element(col + CLng(ScanOffset)), Tname:=TestName, ForceResults:=tlForceFlow
    '                    TestName = "Decompression_Freq_ROV_" + SensorName + "_" + CStr(VoltageName(col))
    '                    TheExec.Flow.TestLimit resultval:=Output_ROV_Freq_Vector(site).Element(col + CLng(ScanOffset)), Tname:=TestName, ForceResults:=tlForceFlow
                    If (actualROVMatrix(site).Element(Col + CLng(ScanOffset)) = 0) Then
                        Difference_ROV_Freq_Vector(site).Element(Col + CLng(ScanOffset)) = 100 * (Output_ROV_Freq_Vector(site).Element(Col + CLng(ScanOffset)) - 0.0001) / 0.0001
                    Else
                        Difference_ROV_Freq_Vector(site).Element(Col + CLng(ScanOffset)) = 100 * (Output_ROV_Freq_Vector(site).Element(Col + CLng(ScanOffset)) - actualROVMatrix(site).Element(Col + CLng(ScanOffset))) / actualROVMatrix(site).Element(Col + CLng(ScanOffset))
                    End If
                Next Col
            Next site
            
            For Col = 0 To UBound(VoltageName)
                    TestName = "Compression_Freq_ROT_" + SensorName + "_" + CStr(VoltageName(Col))
                    TheExec.Flow.TestLimit resultVal:=actualROTMatrix.Element(Col), Tname:=TestName, ForceResults:=tlForceNone, scaletype:=scaleNoScaling 'eng_forceflow_transfer
                    TestName = "Decompression_Freq_ROT_" + SensorName + "_" + CStr(VoltageName(Col))
                    TheExec.Flow.TestLimit resultVal:=Output_ROT_Freq_Vector.Element(Col), Tname:=TestName, ForceResults:=tlForceNone, scaletype:=scaleNoScaling 'eng_forceflow_transfer
                    TestName = "Compression_Freq_ROV_" + SensorName + "_" + CStr(VoltageName(Col))
                    TheExec.Flow.TestLimit resultVal:=actualROVMatrix.Element(Col), Tname:=TestName, ForceResults:=tlForceNone, scaletype:=scaleNoScaling 'eng_forceflow_transfer
                    TestName = "Decompression_Freq_ROV_" + SensorName + "_" + CStr(VoltageName(Col))
                    TheExec.Flow.TestLimit resultVal:=Output_ROV_Freq_Vector.Element(Col), Tname:=TestName, ForceResults:=tlForceNone, scaletype:=scaleNoScaling 'eng_forceflow_transfer
            Next Col
            
             For Col = 0 To UBound(VoltageName)
                TestName = "PercentageDiff_Freq_ROT_" + SensorName + "_" + CStr(VoltageName(Col))
                TheExec.Flow.TestLimit resultVal:=Difference_ROT_Freq_Vector.Element(Col), hiVal:=1, lowVal:=-1, Tname:=TestName, ForceResults:=tlForceNone, scaletype:=scaleNoScaling 'eng_forceflow_transfer
            Next Col
            
            For Col = 0 To UBound(VoltageName)
                TestName = "PercentageDiff_Freq_ROV_" + SensorName + "_" + CStr(VoltageName(Col))
                TheExec.Flow.TestLimit resultVal:=Difference_ROV_Freq_Vector.Element(Col), hiVal:=1, lowVal:=-1, Tname:=TestName, ForceResults:=tlForceNone, scaletype:=scaleNoScaling 'eng_forceflow_transfer
            Next Col
            
    
    '            For Each site In TheExec.sites
    '                For col = 0 To ElementCNT - 1
    '                    TestName = "PercentageDiff_Freq_ROT_" + SensorName + "_" + CStr(VoltageName(col + CLng(ScanOffset)))
    '                    TheExec.Flow.TestLimit resultval:=Difference_ROT_Freq_Vector.Element(col + CLng(ScanOffset)), hival:=1, lowval:=-1, Tname:=TestName, ForceResults:=tlForceFlow
    '                Next col
    '                For col = 0 To ElementCNT - 1
    '                    TestName = "PercentageDiff_Freq_ROV_" + SensorName + "_" + CStr(VoltageName(col + CLng(ScanOffset)))
    '                    TheExec.Flow.TestLimit resultval:=Difference_ROV_Freq_Vector.Element(col + CLng(ScanOffset)), hival:=1, lowval:=-1, Tname:=TestName, ForceResults:=tlForceFlow
    '                Next col
    '            Next site
        End If
End Function


'From Sicily,20200423, Oscar
Public Function HardIP_Trim_Set_PrePoint(argc As Integer, argv() As String) As Long

    Dim i As Long
    Dim Previous_Disable_HardIP_Debug_Log_Flag As Boolean: Previous_Disable_HardIP_Debug_Log_Flag = gl_Disable_HIP_debug_log
    Dim Previous_ByPassTestLimit_Flag As Boolean: Previous_ByPassTestLimit_Flag = ByPassTestLimit
    Dim Previous_Disable_CurrRangeSetting_Print_Flag As Boolean: Previous_Disable_CurrRangeSetting_Print_Flag = glb_Disable_CurrRangeSetting_Print
    Dim Previous_HardIP_Disable_Functional_Result_Flag As Boolean: Previous_HardIP_Disable_Functional_Result_Flag = gl_Flag_HardIP_Disable_Functional_Result
    
    gl_Flag_HardIP_Trim_Set_PrePoint = True
    gl_Disable_HIP_debug_log = True
    ByPassTestLimit = True
    glb_Disable_CurrRangeSetting_Print = True
    gl_Flag_HardIP_Disable_Functional_Result = True
        
    For i = 0 To argc - 1
        Call TheExec.Flow.instance(argv(i)).Execute
    Next i
    
    gl_Flag_HardIP_Disable_Functional_Result = Previous_HardIP_Disable_Functional_Result_Flag
    glb_Disable_CurrRangeSetting_Print = Previous_Disable_CurrRangeSetting_Print_Flag
    ByPassTestLimit = Previous_ByPassTestLimit_Flag
    gl_Disable_HIP_debug_log = Previous_Disable_HardIP_Debug_Log_Flag
    gl_Flag_HardIP_Trim_Set_PrePoint = False
    
End Function

'From Sicily,20200423, Oscar
Public Function HardIP_Trim_Set_PostPoint(argc As Integer, argv() As String) As Long

    Dim i As Long
    Dim Previous_Disable_HardIP_Debug_Log_Flag As Boolean: Previous_Disable_HardIP_Debug_Log_Flag = gl_Disable_HIP_debug_log
    Dim Previous_ByPassTestLimit_Flag As Boolean: Previous_ByPassTestLimit_Flag = ByPassTestLimit
    Dim Previous_Disable_CurrRangeSetting_Print_Flag As Boolean: Previous_Disable_CurrRangeSetting_Print_Flag = glb_Disable_CurrRangeSetting_Print
    Dim Previous_HardIP_Disable_Functional_Result_Flag As Boolean: Previous_HardIP_Disable_Functional_Result_Flag = gl_Flag_HardIP_Disable_Functional_Result
    
    gl_Flag_HardIP_Trim_Set_PostPoint = True
    gl_Disable_HIP_debug_log = True
    ByPassTestLimit = True
    glb_Disable_CurrRangeSetting_Print = True
    gl_Flag_HardIP_Disable_Functional_Result = True
    
    For i = 0 To argc - 1
        Call TheExec.Flow.instance(argv(i)).Execute
    Next i
    
    gl_Flag_HardIP_Disable_Functional_Result = Previous_HardIP_Disable_Functional_Result_Flag
    glb_Disable_CurrRangeSetting_Print = Previous_Disable_CurrRangeSetting_Print_Flag
    ByPassTestLimit = Previous_ByPassTestLimit_Flag
    gl_Disable_HIP_debug_log = Previous_Disable_HardIP_Debug_Log_Flag
    gl_Flag_HardIP_Trim_Set_PostPoint = False
    
End Function



Public Function CUS_AMP_SDLL_SWP_Init(Loop_count As Long, Loop_Init As Long, Loop_Idx As Long, CUS_Str_MainProgram As String, Ori_CUST_Str_MainProgram As String) As Long
    
    If Loop_count = Loop_Init Then
        Ori_CUST_Str_MainProgram = CUS_Str_MainProgram
    End If
    
    CUS_Str_MainProgram = Ori_CUST_Str_MainProgram
    CUS_Str_MainProgram = Replace(UCase(CUS_Str_MainProgram), UCase("Loop_Idx"), CStr(Loop_Idx))
    CUS_Str_MainProgram = Replace(UCase(CUS_Str_MainProgram), UCase("HexSrcCode"), CStr(Loop_count))
    ''                                    CUS_Str_MainProgram = Replace(UCase(CUS_Str_MainProgram), UCase("HexSrcStep"), CStr(Loop_Step))
     Loop_Idx = Loop_Idx + 1
     
End Function

Public Function AMP_EYE_VT_Setup(Char_Flag As Boolean)

    If Char_Flag = True Then
        AMP_EYE_VT_CZ_Flag = True
    Else
        AMP_EYE_VT_CZ_Flag = False
    End If
    'Alarm check From Sicily,20200423, Oscar
    ' Check implicit alarms
    TheHdw.Alarms.Check
    
End Function

Public Function ADCLK_Matrix_Loading()
Dim ADCLK_Matrix_Sheet As Worksheet: Set ADCLK_Matrix_Sheet = Sheets("Flow_HARDIP_ADCLK")
Dim Column_Index As Long: Column_Index = 1
Dim Row_Index As Long: Row_Index = 1
Dim Matrix_Index As Long
Dim ADCLK_Matrix_Index As Long: ADCLK_Matrix_Index = 0
Dim ADCLK_Matrix_Range As Variant
Dim Max_Rows_Count As Long
Dim Max_Columns_Count As Long

With ADCLK_Matrix_Sheet
    Max_Rows_Count = .UsedRange.Rows.Count
    Max_Columns_Count = .UsedRange.Columns.Count
    ADCLK_Matrix_Range = .range(.Cells(5, 1), .Cells(Max_Rows_Count, Max_Columns_Count))
End With

Dim add_Matrix_Sheet As Worksheet: Set add_Matrix_Sheet = Sheets("add")
Dim add_Matrix_Index As Long: add_Matrix_Index = 0
Dim add_Matrix_Range As Variant
Dim Max_Rows_Count_A As Long
Dim Max_Columns_Count_A As Long
Dim Column_Index_A As Long: Column_Index_A = 1
Dim Row_Index_A As Long: Row_Index_A = 1


With add_Matrix_Sheet
    Max_Rows_Count_A = .UsedRange.Rows.Count
    Max_Columns_Count_A = .UsedRange.Columns.Count
    add_Matrix_Range = .range(.Cells(1, 1), .Cells(Max_Rows_Count_A, Max_Columns_Count_A))
End With

Dim temp_str() As String

For Row_Index = 1 To Max_Rows_Count - 4
    For Column_Index = 0 To Max_Columns_Count
        If ADCLK_Matrix_Range(Row_Index, 7) = "Use-Limit" And ADCLK_Matrix_Range(Row_Index, 14) = "Hz" Then
            temp_str() = Split(ADCLK_Matrix_Range(Row_Index, 8), "_")
                    For Row_Index_A = 1 To Max_Rows_Count_A
                      'For Column_Index_A = 0 To Max_Columns_Count_A
                        If temp_str(14) = add_Matrix_Range(Row_Index_A, 1) Then
                                 ADCLK_Matrix_Range(Row_Index, 11) = add_Matrix_Range(Row_Index_A, 2)
                                 ADCLK_Matrix_Range(Row_Index, 12) = add_Matrix_Range(Row_Index_A, 3)
                         End If
                      'Next Column_Index_A
                  Next Row_Index_A
        End If
    Next Column_Index
Next Row_Index

With ADCLK_Matrix_Sheet
.range(.Cells(5, 1), .Cells(Max_Rows_Count, Max_Columns_Count)) = ADCLK_Matrix_Range
End With

End Function

Public Function Time_Measure_kit_UP1600(pat_name As Pattern, pin_name As PinList, jitter_meas As Boolean, eye_meas As Boolean, _
    Optional CPUA_Flag_In_Pat As Boolean, Optional TestSequence As String, _
    Optional DigCap_Pin As PinList, Optional DigCap_DataWidth As Long, Optional DigCap_Sample_Size As Long, _
    Optional DigSrc_pin As PinList, Optional DigSrc_DataWidth As Long, Optional DigSrc_Sample_Size As Long, Optional DigSrc_Equation As String, Optional digsrc_assignment As String, Optional DigSrc_FlowForLoopIntegerName As String = vbNullString, _
    Optional CUS_Str_MainProgram As String = vbNullString, Optional CUS_Str_DigCapData As String = vbNullString, Optional CUS_Str_DigSrcData As String = vbNullString, _
    Optional MeasPin_Differential As PinList, _
    Optional MeasF_WalkingStrobe_Flag As Boolean, Optional MeasF_WalkingStrobe_StartV As Double, Optional MeasF_WalkingStrobe_EndV As Double, Optional MeasF_WalkingStrobe_StepVoltage As Double, _
    Optional MeasF_WalkingStrobe_BothVohVolDiffV As Double, Optional MeasF_WalkingStrobe_interval As Double, Optional MeasF_WalkingStrobe_miniFreq As Double) As Long

''duty_freq_meas As Boolean, log_on As Boolean,

    Dim DSPCapture As New PinListData
    Dim RJ As New PinListData, DDJ As New PinListData, Tj As New PinListData
    Dim measUI As New PinListData
    Dim Tr As New PinListData, Tf As New PinListData
    Dim Eye20 As New PinListData, Eye50 As New PinListData, Eye80 As New PinListData
    Dim dspStatus As New PinListData
    
    Dim RJ_J As New PinListData, DDJ_J As New PinListData
    Dim MeasUI_J As New PinListData
    Dim dspStatus_J As New PinListData
    
    Dim dutycycle As New PinListData
    Dim freq As New PinListData
    
    Dim site As Variant
    Dim pin As Variant
    Dim InDSPWave As New DSPWave, OutDspWave As New DSPWave
    
    Dim Pat As String
''    Dim ShowDec As String, ShowOut As String
    Dim PattArray() As String
    Dim patt As Variant
    Dim PatCount As Long
    
    Dim index As Long
    
    On Error GoTo errHandler
    
    Dim TestSequenceArray() As String
    TestSequenceArray = Split(TestSequence, ",")
    Dim Ts As Variant, TestOption As Variant
    Dim TestOptLen As Integer
    Dim i As Long, j As Long, k As Long
    Dim TestLimitPinName As String

   'setup and run pattern
    TheHdw.Digital.ApplyLevelsTiming True, True, True, tlPowered
    
    Call HardIP_InitialSetupForPatgen
    TheHdw.Patterns(pat_name).Load
    Call PATT_GetPatListFromPatternSet(pat_name.value, PattArray, PatCount)
    
    ''20161107-Return sweep test name
    Dim Rtn_SweepTestName As String
    Rtn_SweepTestName = vbNullString
    
    For Each patt In PattArray
    
        Pat = CStr(patt)
    
        TheHdw.Patterns(Pat).Load
        
        Call GeneralDigSrcSetting(Pat, DigSrc_pin, DigSrc_Sample_Size, DigSrc_DataWidth, DigSrc_Equation, digsrc_assignment, _
                                               DigSrc_FlowForLoopIntegerName, CUS_Str_DigSrcData, InDSPWave, Rtn_SweepTestName)
        
        Call GeneralDigCapSetting(Pat, DigCap_Pin, DigCap_Sample_Size, OutDspWave)
        
        Call PrintDigCapSetting(DigCap_Pin, DigCap_Sample_Size, CUS_Str_DigCapData)
        
        '' 20160713 - If no cpuflags in the test item modify the code to run pattern by using .test
        If (CPUA_Flag_In_Pat) Then
            Call TheHdw.Patterns(Pat).start
        Else
            Call TheHdw.Patterns(Pat).test(pfAlways, 0)
        End If

        TheHdw.Wait 0.5
       
        For Each Ts In TestSequenceArray
        
            If (CPUA_Flag_In_Pat) Then
                Call TheHdw.Digital.Patgen.FlagWait(cpuA, 0) 'Meas during CPUA loop
            Else
                Call TheHdw.Digital.Patgen.HaltWait 'Pattern without CPUA loop
            End If
            
            ''20170105-Add freq walking strobe function to fine turn Voh/Vol(single end) Vt(differential) before jitter measurement
            If pin_name <> "" Then
                If MeasF_WalkingStrobe_Flag = True Then
                    Call Freq_WalkingStrobe_Meas_VOHVOL(pin_name, MeasF_WalkingStrobe_StartV, MeasF_WalkingStrobe_EndV, MeasF_WalkingStrobe_StepVoltage, MeasF_WalkingStrobe_BothVohVolDiffV, MeasF_WalkingStrobe_interval, MeasF_WalkingStrobe_miniFreq)
                End If
            ElseIf MeasPin_Differential <> "" Then
                If MeasF_WalkingStrobe_Flag = True Then
                    Call Freq_WalkingStrobe_Meas_VOD_Diff(MeasPin_Differential, MeasF_WalkingStrobe_StartV, MeasF_WalkingStrobe_EndV, MeasF_WalkingStrobe_StepVoltage, MeasF_WalkingStrobe_BothVohVolDiffV, MeasF_WalkingStrobe_interval, MeasF_WalkingStrobe_miniFreq)
                End If
            End If
            
            DSPCapture = TheHdw.Digital.Jitter.SingleDSPWaves
            TheHdw.Digital.Jitter.TimeoutEnable = True          'enable time out
            TheHdw.Digital.Jitter.TimeOut = 10                      'setup time out value
       
            TheHdw.Wait 0.1
    
            'start time measure
            TheHdw.Digital.Jitter.start     'start jitter
            TheHdw.Digital.Jitter.Wait
        
            Call TheHdw.Digital.Jitter.UpdateSingleDSPWaves(DSPCapture)     'update the dspwave
        
            'measure jitter block
            If jitter_meas = True Then
                Call rundsp.LoopJitterMeas(DSPCapture, RJ_J, DDJ_J, MeasUI_J, dspStatus_J, dutycycle, freq)     'DSP
            End If
        
            'measure eye block
            If eye_meas = True Then
                Call rundsp.LoopEyeMeas(DSPCapture, RJ, DDJ, Tj, measUI, Tr, Tf, Eye20, Eye50, Eye80, dspStatus)        'DSP
            End If
        
''            'measure Duty Freq block
''            If duty_freq_meas = True Then
''                'dutycycle = TheHdw.PPMU.Pins(pin_name).Read(tlPPMUReadMeasurements, 1)   ' for assign pin information to dutycycle
''                'freq = TheHdw.PPMU.Pins(pin_name).Read(tlPPMUReadMeasurements, 1)        ' for assign pin information to dutycycle
''                'Call rundsp.duty_freq_meas(DSPCapture, dutycycle, freq)    'DSP
''            End If
            
''            ''generate raw data
''            '' 20161116 pinlist data seems not workable
''            If log_on = True Then
''                For Each Pin In DSPCapture.Pins
''                    For Each Site In TheExec.Sites
''                        TheHdw.Digital.Jitter.FileExport DSPCapture.Pins(Pin).Value(Site), TheExec.DataManager.InstanceName & "_raw_data_" & DSPCapture.Pins(Pin) & "_site" & Site & ".txt"
''                    Next Site
''                Next Pin
''            End If
            
            'judgment
            
            If pin_name <> "" Then
                TestLimitPinName = pin_name
            ElseIf MeasPin_Differential <> "" Then
                TestLimitPinName = MeasPin_Differential
            End If
            
            Dim TestName As String
            TestName = vbNullString
            
            If jitter_meas = True Then
                For index = 0 To (RJ_J.Pins.Count - 1)
                    Call TheExec.Flow.TestLimit(resultVal:=RJ_J.Pins(index), scaletype:=scalePico, unit:=unitTime, Tname:=TestName & "Jitter_RJ", PinName:=TestLimitPinName, ForceResults:=tlForceFlow)
                    Call TheExec.Flow.TestLimit(resultVal:=DDJ_J.Pins(index), scaletype:=scalePico, unit:=unitTime, Tname:=TestName & "Jitter_DDJ", PinName:=TestLimitPinName, ForceResults:=tlForceFlow)
                    Call TheExec.Flow.TestLimit(resultVal:=MeasUI_J.Pins(index), scaletype:=scalePico, unit:=unitTime, Tname:=TestName & "Jitter_UI", PinName:=TestLimitPinName, ForceResults:=tlForceFlow)
                    Call TheExec.Flow.TestLimit(resultVal:=dutycycle.Pins(index), unit:=unitCustom, Tname:=TestName & "Duty_cycle", PinName:=TestLimitPinName, ForceResults:=tlForceFlow)
                    Call TheExec.Flow.TestLimit(resultVal:=freq.Pins(index), scaletype:=scaleMega, unit:=unitHz, Tname:=TestName & "Freq", PinName:=TestLimitPinName, ForceResults:=tlForceFlow)
                Next index
            End If
            
            
            If eye_meas = True Then
                Call TheExec.Flow.TestLimit(resultVal:=RJ, scaletype:=scalePico, unit:=unitTime, Tname:="Eye_RJ", PinName:=TestLimitPinName, ForceResults:=tlForceNone) 'eng_forceflow_transfer
                Call TheExec.Flow.TestLimit(resultVal:=DDJ, scaletype:=scalePico, unit:=unitTime, Tname:="Eye_DDJ", PinName:=TestLimitPinName, ForceResults:=tlForceNone) 'eng_forceflow_transfer
                Call TheExec.Flow.TestLimit(resultVal:=measUI, scaletype:=scalePico, unit:=unitTime, Tname:="Eye_Width", PinName:=TestLimitPinName, ForceResults:=tlForceNone) 'eng_forceflow_transfer
                Call TheExec.Flow.TestLimit(resultVal:=Eye20, scaletype:=scalePico, unit:=unitTime, Tname:="Eye_Widthhigh", PinName:=TestLimitPinName, ForceResults:=tlForceNone) 'eng_forceflow_transfer
                Call TheExec.Flow.TestLimit(resultVal:=Eye50, scaletype:=scalePico, unit:=unitTime, Tname:="Eye_Widthmid", PinName:=TestLimitPinName, ForceResults:=tlForceNone) 'eng_forceflow_transfer
                Call TheExec.Flow.TestLimit(resultVal:=Eye80, scaletype:=scalePico, unit:=unitTime, Tname:="Eye_Widthlow", PinName:=TestLimitPinName, ForceResults:=tlForceNone) 'eng_forceflow_transfer
            End If
        
''            If duty_freq_meas = True Then
''                Call TheExec.Flow.TestLimit(resultVal:=dutycycle, unit:=unitCustom, Tname:="Duty_cycle", PinName:=pin_name, ForceResults:=tlForceFlow)
''                Call TheExec.Flow.TestLimit(resultVal:=freq, ScaleType:=scaleMega, unit:=unitHz, Tname:="Freq", PinName:=pin_name, ForceResults:=tlForceFlow)
''            End If
           
''            TestSeqNum = TestSeqNum + 1
                
            If (CPUA_Flag_In_Pat) Then Call TheHdw.Digital.Patgen.Continue(0, cpuA)
        Next Ts
        
        TheHdw.Digital.Patgen.HaltWait ' Haltwait at patten end

        If DigCap_Sample_Size <> 0 Then
            Dim DigCapPinAry() As String, NumberPins As Long
            Call TheExec.DataManager.DecomposePinList(DigCap_Pin, DigCapPinAry(), NumberPins)
            
            If NumberPins > 1 Then
                Call CreateSimulateDataDSPWave_Parallel(OutDspWave, DigCap_Sample_Size)
                Call DigCapDataProcessByDSP_Parallel(CUS_Str_DigCapData, OutDspWave, DigCap_Sample_Size, NumberPins)
            ElseIf NumberPins = 1 Then
                Call CreateSimulateDataDSPWave(OutDspWave, DigCap_Sample_Size, DigCap_DataWidth)
                Call Checker_StoreDigCapAllToDictionary(CUS_Str_DigCapData, OutDspWave)
                Call DigCapDataProcessByDSP(CUS_Str_DigCapData, OutDspWave, DigCap_Sample_Size, DigCap_DataWidth)
            End If
        End If
    
    Next patt
    
    '' 20160713 - Call write functional result if cpu flag in pattern
    If (CPUA_Flag_In_Pat) Then
        Call HardIP_WriteFuncResult
    End If
  
    Exit Function
  
errHandler:
    TheExec.Datalog.WriteComment "error in Time_Measure_kit_UP1600"
    If AbortTest Then Exit Function Else Resume Next
End Function

Public Function MetrologySense_Calibration(Optional patset As Pattern, Optional DigCap_Pin As PinList, Optional DigCap_DataWidth As Long, Optional DigCap_Sample_Size As Long, Optional CUS_Str_MainProgram As String = vbNullString, Optional CUS_Str_DigCapData As String = vbNullString, Optional Interpose_PrePat_Sweep As String, Optional AC_Category_Sweep As String, Optional SweepVArrayValue As String, Optional Validating_ As Boolean) As Long

    Dim PatCount As Long
    Dim PattArray() As String
    Dim i, j, k As Long
    If Validating_ Then
        Call PrLoadPattern(patset.value)
        Exit Function    ' Exit after validation
    End If
    
    If TheHdw.DSP.ExecutionMode = tlDSPModeHostDebug Then
        TheHdw.DSP.ExecutionMode = tlDSPModeAutomatic
    End If

    glb_TestInstance = vbNullString
    glb_TestInstance = UCase(TheExec.DataManager.instancename)
    
    Call HardIP_InitialSetupForPatgen
    
    Dim TestOptLen As Integer
    Dim Ts As Variant, TestOption As Variant ', Site As Variant
    Dim TestSeqNum As Integer
    Dim testnum As Long
    Dim Interpose_PrePat() As String: Interpose_PrePat = Split(Interpose_PrePat_Sweep, "&")
    Dim AC_Category() As String: AC_Category = Split(AC_Category_Sweep, "&")
    Dim OutDspWave() As New DSPWave
    ReDim OutDspWave(UBound(Interpose_PrePat)) As New DSPWave
    Dim patt As Variant
    Dim Pat As String
     
    On Error GoTo errHandler
    Dim CheckDSPWave As New DSPWave
    Dim OutputTname() As String
    Call tl_PinListDataSort(True)
    
    '----------------------------20180523
    'Roger New,20180510 TName
    '--------------------------------------------------------------------
    Call GetFlowTName
    
    ''20161130-Get test name from flow table
    Dim FlowTestNme() As String
    Dim OutputTname_format() As String
    Dim TestNameInput As String
    Dim Temp_index As Long
    Dim SensorVoltage() As String: SensorVoltage = Split(SweepVArrayValue, ",")
    Dim DSSC_Out_DecompseByComma() As String: DSSC_Out_DecompseByComma = Split(CUS_Str_DigCapData, ",")
    Dim DSSC_Out_DecompseByColon() As String
    Dim ParseStringByBits As String: ParseStringByBits = vbNullString
    Dim ParseStringForTestName As String: ParseStringForTestName = vbNullString
    Dim DecomposeTestName() As String
    Dim DecomposeParseDigCapBit() As String
    Dim TestLimitWithTestName As New PinListData
    For i = 0 To UBound(DSSC_Out_DecompseByComma)
        DSSC_Out_DecompseByColon = Split(DSSC_Out_DecompseByComma(i), ":")
        If UBound(DSSC_Out_DecompseByColon) > 0 Then
            If ParseStringByBits = "" And ParseStringForTestName = "" Then
                ParseStringByBits = DSSC_Out_DecompseByColon(0)
                ParseStringForTestName = DSSC_Out_DecompseByColon(1)
            Else
                ParseStringByBits = ParseStringByBits & "," & DSSC_Out_DecompseByColon(0)
                ParseStringForTestName = ParseStringForTestName & "," & DSSC_Out_DecompseByColon(1)
            End If
        End If
    Next i
    DecomposeTestName = Split(ParseStringForTestName, ",")
    DecomposeParseDigCapBit = Split(ParseStringByBits, ",")

    If patset.value <> "" Then
        Shmoo_Pattern = patset.value '' 20170808 add for shmoo pattern name print
        TheHdw.Patterns(patset).Load
        Call PATT_GetPatListFromPatternSet(patset.value, PattArray, PatCount)
    Else
        ReDim PattArray(0)
        PattArray(0) = vbNullString
    End If

    For i = 0 To UBound(Interpose_PrePat)
        If i = 0 Then
            TheHdw.Digital.ApplyLevelsTiming True, True, True, tlPowered, , , , "Levels_HardIp", "HardIP", "Typ", "TIMESET_SKUA0_S_AN_SI_2", AC_Category(i), "Typ"
        ElseIf StrComp(AC_Category(i), AC_Category(i - 1)) Then
            TheHdw.Digital.ApplyLevelsTiming True, True, True, tlPowered, , , , "Levels_HardIp", "HardIP", "Typ", "TIMESET_SKUA0_S_AN_SI_2", AC_Category(i), "Typ"
        End If
         '' 20160923 - Add Interpose_PrePat entry point
        If Interpose_PrePat(i) <> "" Then: Call SetForceCondition(Interpose_PrePat(i) & ";STOREPREPAT")
        
        For Each patt In PattArray
                If patt <> "" Then
                    Pat = CStr(patt)
                    TheHdw.Patterns(Pat).Load
                    Set OutDspWave(i) = Nothing

                    Call GeneralDigCapSetting(Pat, DigCap_Pin, DigCap_Sample_Size, OutDspWave(i))
                    Call PrintDigCapSetting(DigCap_Pin, DigCap_Sample_Size, CUS_Str_DigCapData)
                    Call TheHdw.Patterns(Pat).test(pfAlways, 0)
                    Call CreateSimulateDataDSPWave(OutDspWave(i), DigCap_Sample_Size, DigCap_DataWidth)
                End If
        Next patt
             If Interpose_PrePat(i) <> "" Then: Call SetForceCondition("RESTOREPREPAT")
    Next i
    
    Dim Freq_DSP As New DSPWave
    Dim DigCap_DataWidth_SiteLong As New SiteLong
    For Each site In TheExec.sites.Active
        DigCap_DataWidth_SiteLong = DigCap_DataWidth
    Next site
    If UCase(patset.value) Like "*ASGMTR*" Then
        rundsp.MTR_ASGMTR_Freq_Calculation OutDspWave(0), OutDspWave(1), OutDspWave(2), OutDspWave(3), OutDspWave(4), OutDspWave(5), OutDspWave(6), OutDspWave(7), OutDspWave(8), OutDspWave(9), OutDspWave(10), OutDspWave(11), OutDspWave(12), OutDspWave(13), OutDspWave(14), OutDspWave(15), OutDspWave(16), OutDspWave(17), OutDspWave(18), OutDspWave(19), OutDspWave(20), OutDspWave(21), DigCap_DataWidth, Freq_DSP
   ElseIf UCase(patset.value) Like "*DSGMTR*" Then
        rundsp.MTR_DSGMTR_Freq_Calculation OutDspWave(0), OutDspWave(1), OutDspWave(2), OutDspWave(3), OutDspWave(4), OutDspWave(5), OutDspWave(6), OutDspWave(7), OutDspWave(8), OutDspWave(9), OutDspWave(10), OutDspWave(11), OutDspWave(12), OutDspWave(13), OutDspWave(14), OutDspWave(15), OutDspWave(16), OutDspWave(17), OutDspWave(18), OutDspWave(19), OutDspWave(20), OutDspWave(21), DigCap_DataWidth, Freq_DSP
   End If
   
    TheHdw.Digital.Patgen.HaltWait ' Haltwait at patten end

    '' 20160211 - Process DigCapData by using DSP
        If DigCap_Sample_Size <> 0 Then
            Temp_index = TheExec.Flow.TestLimitIndex
            For i = 0 To UBound(Interpose_PrePat)
                TheExec.Flow.TestLimitIndex = Temp_index
                For k = 0 To UBound(DecomposeTestName)
                    TestLimitWithTestName.AddPin (DecomposeTestName(k) & "_" & k)
                    TestLimitWithTestName.Pins(DecomposeTestName(k) & "_" & k).value = OutDspWave(i).Element(k)
                    TestNameInput = Report_TName_From_Instance("C", DigCap_Pin.value, , CInt(k), , , SensorVoltage(i))
                    TheExec.Flow.TestLimit TestLimitWithTestName.Pins(DecomposeTestName(k) & "_" & k), 0, 2 ^ DecomposeParseDigCapBit(k) - 1, Tname:=TestNameInput, ForceResults:=tlForceFlow, scaletype:=scaleNoScaling, formatStr:="%.0f"
                Next k
                For j = 0 To DigCap_Sample_Size / DigCap_DataWidth - 1
                    TestNameInput = Report_TName_From_Instance("Calc", vbNullString, , , , , SensorVoltage(i))
                    TheExec.Flow.TestLimit resultVal:=Freq_DSP.Element((DigCap_Sample_Size / DigCap_DataWidth) * i + j), Tname:=TestNameInput, ForceResults:=tlForceFlow
                Next j
                Set TestLimitWithTestName = Nothing
            Next i
        End If

     DebugPrintFunc patset.value  ' print all debug information

    Exit Function

errHandler:
    TheExec.Datalog.WriteComment "error in Metrology_Sense_Calibration"
    If AbortTest Then Exit Function Else Resume Next
  
End Function

Public Function MTRTMPS_OffSet_AVG(StoreName_OffSet As String, StoreName_OffSetMean As String, Integer_Bit As Long, StoreName_OffSetMean_Fuse As String) As Long
Dim InWf As New DSPWave
Dim StoreName_OffSet_Split() As String: StoreName_OffSet_Split = Split(StoreName_OffSet, ",")
Dim i As Long
Dim DSP_OffSet_Mean As New DSPWave
Dim DSP_OffSet_Mean_Array(0) As Double
Dim DSP_OffSet_Mean_Fuse As New DSPWave
Dim DSP_OffSet_Mean_Fuse_Array(0) As Double
Dim TestNameInput As String
Dim High_limit As Double: High_limit = Bin2Dec_rev(String(Integer_Bit - 1, "1"))
Dim Low_limit As Double: Low_limit = -2 ^ (Integer_Bit - 1)

    For i = 0 To UBound(StoreName_OffSet_Split)
        If i = 0 Then
            InWf = GetStoredCaptureData(StoreName_OffSet_Split(i))
        Else
            For Each site In TheExec.sites.Active
                InWf = InWf.Concatenate(GetStoredCaptureData(StoreName_OffSet_Split(i)))
            Next site
        End If
    Next i
    For Each site In TheExec.sites.Active
        DSP_OffSet_Mean_Array(0) = FormatNumber(InWf.CalcMean, 0)
        DSP_OffSet_Mean.data = DSP_OffSet_Mean_Array
        If DSP_OffSet_Mean_Array(0) < Low_limit Then
            DSP_OffSet_Mean_Fuse_Array(0) = 2 ^ (Integer_Bit) + FormatNumber(Low_limit, 0)
        ElseIf DSP_OffSet_Mean_Array(0) >= Low_limit And DSP_OffSet_Mean_Array(0) < 0 Then
            DSP_OffSet_Mean_Fuse_Array(0) = 2 ^ (Integer_Bit) + FormatNumber(DSP_OffSet_Mean_Array(0), 0)
        ElseIf DSP_OffSet_Mean_Array(0) < High_limit And DSP_OffSet_Mean_Array(0) >= 0 Then
            DSP_OffSet_Mean_Fuse_Array(0) = FormatNumber(DSP_OffSet_Mean_Array(0), 0)
        Else
            DSP_OffSet_Mean_Fuse_Array(0) = FormatNumber(High_limit, 0)
        End If
        DSP_OffSet_Mean_Fuse.data = DSP_OffSet_Mean_Fuse_Array
    Next site
    
    TestNameInput = Report_TName_From_Instance("C", "X", , 0, 0)
    TheExec.Flow.TestLimit resultVal:=DSP_OffSet_Mean.Element(0), Tname:=TestNameInput, ForceResults:=tlForceFlow, scaletype:=scaleNoScaling, formatStr:="%.0f"
    
    Call AddStoredCaptureData(StoreName_OffSetMean, DSP_OffSet_Mean)
    If StoreName_OffSetMean_Fuse <> "" Then: Call AddStoredCaptureData(StoreName_OffSetMean_Fuse, DSP_OffSet_Mean_Fuse)
    
End Function

Public Function MTRTMPS_Gain_Cal(StoreName_Gain As String, StoreName_Gain_Mean As String, StoreName_OffSetMean As String, StoreName_GainMean_Fuse As String) As Long
Dim InWf As New DSPWave
Dim StoreName_Gain_Split() As String: StoreName_Gain_Split = Split(StoreName_Gain, ",")
Dim DSP_OffSet_Mean As New DSPWave: DSP_OffSet_Mean = GetStoredCaptureData(StoreName_OffSetMean)
Dim DSP_OffSet_Mean_Array() As Double
Dim MTRTMPS_Gain As New DSPWave
Dim MTRTMPS_Gain_Array(0) As Double
Dim TestNameInput As String

    For i = 0 To UBound(StoreName_Gain_Split)
        If i = 0 Then
            InWf = GetStoredCaptureData(StoreName_Gain_Split(i))
        Else
            For Each site In TheExec.sites.Active
                InWf = InWf.Concatenate(GetStoredCaptureData(StoreName_Gain_Split(i)))
            Next site
        End If
    Next i

    For Each site In TheExec.sites.Active
        DSP_OffSet_Mean_Array = DSP_OffSet_Mean.data
        MTRTMPS_Gain_Array(0) = FormatNumber(InWf.CalcMean - 8 * DSP_OffSet_Mean_Array(0), 0)
        If MTRTMPS_Gain_Array(0) < 0 Then: MTRTMPS_Gain_Array(0) = 0
        MTRTMPS_Gain.data = MTRTMPS_Gain_Array
    Next site
    
    TestNameInput = Report_TName_From_Instance("C", "X", , 0, 0)
    TheExec.Flow.TestLimit resultVal:=MTRTMPS_Gain.Element(0), Tname:=TestNameInput, ForceResults:=tlForceFlow, scaletype:=scaleNoScaling, formatStr:="%.0f"
    
    If StoreName_GainMean_Fuse <> "" Then: Call AddStoredCaptureData(StoreName_GainMean_Fuse, MTRTMPS_Gain)
    
End Function

Public Function PCIE_PI_TEST(patset1 As Pattern, patset2 As Pattern, patset3 As Pattern, patset4 As Pattern, PhAdjMax As Integer, Total_Lane_Num As Long, Tested_Lane_Num As Long) As Long
    
    Dim pat1_count As Long
    Dim pat2_count As Long
    Dim pat3_count As Long
    Dim pat4_count As Long
    Dim PatTrimArray_pat1() As String
    Dim PatTrimArray_pat2() As String
    Dim PatTrimArray_pat3() As String
    Dim PatTrimArray_pat4() As String
    Dim site As Variant
    Dim Current_test_number As Long
    TheHdw.Digital.ApplyLevelsTiming True, True, True, tlPowered
    
    Call GetPatFromPatternSet(patset1.value, PatTrimArray_pat1, pat1_count)
    Call GetPatFromPatternSet(patset2.value, PatTrimArray_pat2, pat2_count)
    Call GetPatFromPatternSet(patset3.value, PatTrimArray_pat3, pat3_count)
    Call GetPatFromPatternSet(patset4.value, PatTrimArray_pat4, pat4_count)
    
    offset_angle = 5.625
    
    For Each site In TheExec.sites
        
       Current_test_number = TheExec.sites(site).TestNumber
        
    Next site
    
    Call PCIE_PI_Pat1(PatTrimArray_pat1(0), Total_Lane_Num)
    
    For Each site In TheExec.sites
        
        pat4_check(site) = 0
        
    Next site
        
        For lane = 0 To Tested_Lane_Num - 1
                 
            Call PCIE_PI_Pat2(PatTrimArray_pat2(lane), PhAdjMax)
            Call PCIE_PI_Pat3(PatTrimArray_pat3(lane), "JTAG_TDI", PatTrimArray_pat2(lane), PhAdjMax)

        Next lane
        
     
    For Each site In TheExec.sites
     
     TheExec.sites(site).TestNumber = Current_test_number + 50000
     
    Next site
     
     Call PCIE_PI_Pat4(PatTrimArray_pat4(0), Total_Lane_Num, Tested_Lane_Num)
     
     DebugPrintFunc patset1 & "," & patset2 & "," & patset3 & "," & patset4 ' print all debug information
     
Exit Function

errHandler:
    TheExec.Datalog.WriteComment "error in PCIE_PI_TEST"
    If AbortTest Then Exit Function Else Resume Next
  
End Function
Public Function PCIE_PI_Pat1(pat1 As String, Total_Lane_Num As Long)

    Dim pat1_dspwave As New DSPWave
    Dim abc As Integer
    Dim output_pin As New PinList
    Dim merge_bit As String
    output_pin = "JTAG_TDO"
    Dim site As Variant
    Dim total_bit As Long
    
    total_bit = Total_Lane_Num * 7
    
    Call DigCapSetup(pat1, output_pin, "test", total_bit, pat1_dspwave)

     
    TheHdw.Patterns(pat1).Load
    Call TheHdw.Patterns(pat1).test(pfAlways, 0, tlResultModeDomain)
    
    Call TheHdw.Digital.Patgen.HaltWait
    
    For Each site In TheExec.sites
    
    pat1_dspwave = TheHdw.DSSC.Pins("JTAG_TDO").Pattern(pat1).Capture.Signals("test").DSPWave
    
        For abc = UBound(pat1_dspwave(site).data) To 0 Step -1
        merge_bit = merge_bit & pat1_dspwave(site).Element(abc)
        Next abc
        
        TheExec.Datalog.WriteComment "site : " & site & " " & merge_bit & " <=LSB"
        merge_bit = vbNullString
        
    Next site
   
End Function

Public Function PCIE_PI_Pat2(pat2 As String, PhAdjMax As Integer)

    Dim pat2_dspwave As New DSPWave
    Dim abc As Integer
    Dim output_pin As New PinList
    output_pin = "JTAG_TDO"
    Dim merge_bit As String
    Dim site As Variant
    Call DigCapSetup(pat2, output_pin, "test1", 8, pat2_dspwave)
    
    TheHdw.Patterns(pat2).Load
    Call TheHdw.Patterns(pat2).test(pfAlways, 0, tlResultModeDomain)
    Call TheHdw.Digital.Patgen.HaltWait
    
    

    
        
    pat2_dspwave = TheHdw.DSSC.Pins("JTAG_TDO").Pattern(pat2).Capture.Signals("test1").DSPWave
        
    For Each site In TheExec.sites
    
        For abc = UBound(pat2_dspwave(site).data) To 0 Step -1
            merge_bit = merge_bit & pat2_dspwave(site).Element(abc)
        Next abc
            
        'TheExec.Datalog.WriteComment "site : " & Site & " " & merge_bit & " <=LSB"
    
        If merge_bit Like "00000000" Or merge_bit Like "00100000" Then
            PhAdj(site) = PhAdjMax + 1
            pat4_check(site) = 1
        End If
        
        
        Qd(site) = (pat2_dspwave(site).Element(6) + pat2_dspwave(site).Element(7) * 2)
        Ph0Rel(site) = pat2_dspwave(site).Element(0) + pat2_dspwave(site).Element(1) * 2 + pat2_dspwave(site).Element(2) * 4 + pat2_dspwave(site).Element(3) * 8 + pat2_dspwave(site).Element(4) * 16


        Select Case Qd(site)
            Case 0:
               Ph0(site) = 0 + (Ph0Rel(site) * offset_angle)
               Adj_steps(site) = Ph0(site) / offset_angle
            Case 1:
               Ph0(site) = 360 - (Ph0Rel(site) * offset_angle)
               Adj_steps(site) = (360 - Ph0(site)) / offset_angle
            Case 2:
               Ph0(site) = 180 - (Ph0Rel(site) * offset_angle)
               Adj_steps(site) = Ph0(site) / offset_angle
            Case 3:
               Ph0(site) = 180 + (Ph0Rel(site) * offset_angle)
               Adj_steps(site) = (360 - Ph0(site)) / offset_angle
        End Select
      
        TheExec.Datalog.WriteComment "Lane" & lane & "  site " & site & ":" & merge_bit & "<=LSB" & "  Angle: " & Format(Ph0(site), "##0.000") & "  Steps: " & Adj_steps(site)

        merge_bit = vbNullString
    Next site

End Function

Public Function PCIE_PI_Pat3(Adjust_pat As String, DigSrcPin As String, PCIE_Pat2 As String, PCIE_PhAdjMax As Integer) As Long

  Dim Data_Out As New DSPWave
  Dim pat_count As Long
  Dim Adj_patArray() As String
  Dim Adj_common_steps As Integer
  Dim i As Integer
  Dim i_loop As Integer
  Dim j_loop As Integer
  Dim DigSrcPin1 As New PinList
  Dim site As Variant
  Dim ChecPhaseResult As Boolean
  Dim InWave As New DSPWave
  Dim Current_test_number1 As Long
  InWave.CreateConstant 0, 3
  TheHdw.Patterns(Adjust_pat).Load
  
  For Each site In TheExec.sites
       Current_test_number1 = TheExec.sites(site).TestNumber
  Next site
  
  
  For i_loop = 0 To PCIE_PhAdjMax
  
      For Each site In TheExec.sites
    
'      Select Case Qd(Site)
'      Case 0:
'         Ph0(Site) = 0 + (Ph0Rel(Site) * offset_angle)
'         Adj_steps(Site) = Ph0(Site) / offset_angle
'      Case 1:
'         Ph0(Site) = 360 - (Ph0Rel(Site) * offset_angle)
'         Adj_steps(Site) = (360 - Ph0(Site)) / offset_angle
'      Case 2:
'         Ph0(Site) = 180 - (Ph0Rel(Site) * offset_angle)
'         Adj_steps(Site) = Ph0(Site) / offset_angle
'      Case 3:
'         Ph0(Site) = 180 + (Ph0Rel(Site) * offset_angle)
'         Adj_steps(Site) = (360 - Ph0(Site)) / offset_angle
'      End Select
         
      If Qd(site) = 0 Or Qd(site) = 2 Then
         InWave.Element(0) = 1
         InWave.Element(1) = 0
         InWave.Element(2) = 1
      Else
         InWave.Element(0) = 1
         InWave.Element(1) = 1
         InWave.Element(2) = 0
      End If
      DigSrcPin1 = DigSrcPin
    
    Next site
    
    Call SetupDigSrcDspWave(Adjust_pat, DigSrcPin1, "phase_adj", 3, InWave)
     
  
    Adj_common_steps = 999
    
    For Each site In TheExec.sites
        
        If Adj_common_steps > Adj_steps(site) Then
            Adj_common_steps = Adj_steps(site)
        End If
        
    Next site
  
    For j_loop = 0 To Adj_common_steps - 1
'        Call TheHdw.Patterns(Adjust_pat).start
        Call TheHdw.Patterns(Adjust_pat).test(pfAlways, 0, tlResultModeDomain)
        TheHdw.Digital.Patgen.HaltWait
    Next j_loop
  
  
     For Each site In TheExec.sites
            For i = 0 To Adj_steps(site) - 1 - Adj_common_steps
'                   If i < Adj_steps(Site) - 1 - Adj_common_steps Then
'                    Call TheHdw.Patterns(Adjust_pat).start
'                   Else
                 Call TheHdw.Patterns(Adjust_pat).test(pfAlways, 0, tlResultModeDomain)
'                   End If
                 TheHdw.Digital.Patgen.HaltWait
            Next i
     Next site


    For Each site In TheExec.sites
        
        TheExec.sites(site).TestNumber = Current_test_number1 + 2000 * (i_loop + 1)
        
    Next site
     
    Call PCIE_PI_Pat2(PCIE_Pat2, PCIE_PhAdjMax)
     
    ChecPhaseResult = True
    For Each site In TheExec.sites
    
        If pat4_check(site) = 0 Then
           ChecPhaseResult = False
           If i_loop = PCIE_PhAdjMax Then TheExec.sites.item(site).testResult = siteFail
        End If
    
    Next site
    
    If ChecPhaseResult = True Then i_loop = PCIE_PhAdjMax + 1
                                 
  Next i_loop
  
  
End Function

Public Function PCIE_PI_Pat4(pat4 As String, Total_Lane_Num As Long, Tested_Lane_Num As Long)

    Dim pat4_dspwave As New DSPWave
    Dim i As Integer
    Dim output_pin As New PinList
    Dim site As Variant
    Dim show_out As String
    Dim lane_status_str() As String
    Dim lane_phase_result() As String
    Dim phase_result As String
    Dim phase_num As Integer
    Dim Lane_Num As Integer
    Dim total_bit As Long
    Dim x As Integer
    Dim phase_Array() As String
    Dim eye_width_site() As New SiteDouble
    Dim eye_count_site() As New SiteDouble
    Dim con_fail_count_site() As New SiteDouble
    Dim pat4_output_bit_num As Long
    
    ReDim lane_phase_result(0 To Total_Lane_Num - 1)
    ReDim eye_width_site(0 To Total_Lane_Num - 1)
    ReDim eye_count_site(0 To Total_Lane_Num - 1)
    ReDim con_fail_count_site(0 To Total_Lane_Num - 1)

    glb_TestInstance = vbNullString
    glb_TestInstance = UCase(TheExec.DataManager.instancename)
    
    output_pin = "JTAG_TDO"
    total_bit = Total_Lane_Num * 16 * 17

    TheHdw.Patterns(pat4).Load
    Call DigCapSetup(pat4, output_pin, "test1", total_bit, pat4_dspwave)

    Call TheHdw.Patterns(pat4).test(pfAlways, 0, tlResultModeDomain)
 
    Call TheHdw.Digital.Patgen.HaltWait
    
    pat4_dspwave = TheHdw.DSSC.Pins("JTAG_TDO").Pattern(pat4).Capture.Signals("test1").DSPWave
    

    For Each site In TheExec.sites
    
        show_out = vbNullString
        
        For x = 0 To total_bit - 1
            show_out = show_out & pat4_dspwave.Element(x)
        Next x
     
        For x = 0 To Total_Lane_Num - 1
            lane_phase_result(x) = vbNullString
        Next x

        For phase_num = 0 To 15
            For x = 0 To Total_Lane_Num - 1
                phase_result = mid(show_out, (x + 1) + (phase_num * (17 * Total_Lane_Num)), 1) & mid(show_out, (Total_Lane_Num + 1 + (x * 16)) + (phase_num * (17 * Total_Lane_Num)), 16)
                If phase_result = "10000000000000000" Then
                    lane_phase_result(x) = lane_phase_result(x) & "1,"
                Else
                    lane_phase_result(x) = lane_phase_result(x) & "0,"
                End If
            Next x
        Next phase_num
        
 
       For i = 0 To Tested_Lane_Num - 1
            
                phase_Array = Split(lane_phase_result(i), ",")
                eye_count_site(i) = EyeCount(phase_Array)
                
                phase_Array = Split(lane_phase_result(i), ",")
                eye_width_site(i) = EyeWidth(phase_Array)
                
                phase_Array = Split(lane_phase_result(i), ",")
                con_fail_count_site(i) = PI_continuous_fai_count(phase_Array)

                TheExec.Datalog.WriteComment "Site: " & site & ", Lane" & i & " phase result " & " = " & lane_phase_result(i) & vbNullString
        Next i
        
        TheExec.Datalog.WriteComment "Site: " & site & ", Capture bits " & total_bit & " = " & show_out & " "
        
    Next site
    
   For i = 0 To Tested_Lane_Num - 1
                TheExec.Flow.TestLimit eye_count_site(i), 1, 2, Tname:=glb_TestInstance & "con_fail_count", PinName:="Lane_" & i & "_eye_count", ForceResults:=tlForceFlow
                TheExec.Flow.TestLimit eye_width_site(i), 6, 16, Tname:=glb_TestInstance & "con_fail_count", PinName:="Lane_" & i & "_eye_width", ForceResults:=tlForceFlow
                TheExec.Flow.TestLimit con_fail_count_site(i), 0, 2, Tname:=glb_TestInstance & "con_fail_count", PinName:="Lane_" & i & "_max_con_fail_count", ForceResults:=tlForceFlow
  Next i
End Function

Public Function PI_continuous_fai_count(phase_Array() As String) As Integer
'================continuous fail count============================================
        Dim temp_count As Integer
        Dim con_fail_count As Integer
        Dim loop_2nd As Integer
        Dim i As Integer
        temp_count = 0
        con_fail_count = 0
        loop_2nd = 0
        For i = 0 To 15
            If temp_count = 16 Then
                con_fail_count = 16
                i = 16
            End If
            If phase_Array(i) = "0" Then
                temp_count = temp_count + 1
                If i = 15 Then
                    i = 0
                    loop_2nd = 1
                End If
                    
            ElseIf phase_Array(i) = "1" Then
                If con_fail_count <= temp_count Then con_fail_count = temp_count
                temp_count = 0
                If loop_2nd = 1 Then i = i + 15
            Else
            End If
        Next i
        
    PI_continuous_fai_count = con_fail_count

End Function

'Public Check_Eye(17) As String
Public Function EyeCount(EyeStrArray() As String) As Integer
'Public Function EyeCount() As Integer
'Dim EyeStrArray(17) As String
Dim EyeArrayCnt As Integer
Dim loop_i As Integer
Dim Record_bit As String
Dim Current_bit As String
Dim transitionCnt As Integer
Dim CycleCnt As Integer
Dim First_bit_Test As Boolean
Dim check_result As Integer

'EyeStrArray(0) = "1"
'EyeStrArray(1) = "1"
'EyeStrArray(2) = "1"
'EyeStrArray(3) = "1"
'EyeStrArray(4) = "1"
'EyeStrArray(5) = "1"
'EyeStrArray(6) = "1"
'EyeStrArray(7) = "1"
'EyeStrArray(8) = "1"
'EyeStrArray(9) = "1"
'EyeStrArray(10) = "1"
'EyeStrArray(11) = "1"
'EyeStrArray(12) = "1"
'EyeStrArray(13) = "1"
'EyeStrArray(14) = "1"
'EyeStrArray(15) = "1"
'EyeStrArray(16) = ""

    First_bit_Test = True
    CycleCnt = 0

    EyeArrayCnt = UBound(EyeStrArray)
    
    transitionCnt = 0
    
    For loop_i = 0 To EyeArrayCnt
    
        Current_bit = EyeStrArray(loop_i) ' current bit

        If First_bit_Test = False Then
        
            If Record_bit <> Current_bit Then transitionCnt = transitionCnt + 1 'cal the transition
            
            If CycleCnt = 1 Then Exit For
            
        End If

        Record_bit = EyeStrArray(loop_i) ' recored bit
    
        If loop_i = (EyeArrayCnt - 1) And CycleCnt = 0 Then
        
            CycleCnt = 1 'set start another cycle
            loop_i = -1
        End If
    
    First_bit_Test = False
    
    Next

    check_result = transitionCnt Mod 2
    transitionCnt = transitionCnt - check_result
    If transitionCnt = 0 Then
        If Current_bit Like "1" Then
            EyeCount = 1
        Else
            EyeCount = 0
        End If
    Else
        EyeCount = transitionCnt / 2 'Mod 2
    End If
    
End Function

Public Function EyeWidth(EyeWidthStrArray() As String) As Integer
'
'End Function
'Public Function EyeWidth() As Integer
'Dim EyeWidthStrArray(17) As String
Dim EyeArrayCnt As Integer
Dim loop_i As Integer
Dim Record_bit As String
Dim Current_bit As String
Dim transitionCnt As Integer
Dim CycleCnt As Integer
Dim First_bit_Test As Boolean
Dim check_result As Integer
Dim MaxEyeWidth As Integer
Dim CurrentEyeWidth As Integer
Dim Start_record As Boolean
'EyeWidthStrArray(0) = "0"
'EyeWidthStrArray(1) = "0"
'EyeWidthStrArray(2) = "0"
'EyeWidthStrArray(3) = "0"
'EyeWidthStrArray(4) = "0"
'EyeWidthStrArray(5) = "0"
'EyeWidthStrArray(6) = "0"
'EyeWidthStrArray(7) = "0"
'EyeWidthStrArray(8) = "0"
'EyeWidthStrArray(9) = "0"
'EyeWidthStrArray(10) = "0"
'EyeWidthStrArray(11) = "0"
'EyeWidthStrArray(12) = "0"
'EyeWidthStrArray(13) = "0"
'EyeWidthStrArray(14) = "0"
'EyeWidthStrArray(15) = "0"
'EyeWidthStrArray(16) = ""

    MaxEyeWidth = 16
    CurrentEyeWidth = 16
    Start_record = False
    First_bit_Test = True
    CycleCnt = 0

    EyeArrayCnt = UBound(EyeWidthStrArray) '- 1
    transitionCnt = 0
    
    For loop_i = 0 To EyeArrayCnt - 1
    
        Current_bit = EyeWidthStrArray(loop_i) ' current bit

        If First_bit_Test = False Then
        
            If Record_bit <> Current_bit Then
            
                If Current_bit Like "1" Then
                    Start_record = True
                End If
                
                If Current_bit Like "0" Then
                    If CurrentEyeWidth < MaxEyeWidth Then MaxEyeWidth = CurrentEyeWidth
                    CurrentEyeWidth = 0
                End If
                
                
                transitionCnt = transitionCnt + 1 'cal the transition
                
            End If
            
            If Current_bit Like "1" And Start_record = True Then
                CurrentEyeWidth = CurrentEyeWidth + 1
            End If
            
                
            
        End If

        Record_bit = EyeWidthStrArray(loop_i) ' recored bit
    
        If loop_i = (EyeArrayCnt - 1) And CycleCnt = 0 Then
        
            CycleCnt = 1 'set start another cycle
            loop_i = -1
        End If
    
    First_bit_Test = False
    
    Next

'    check_result = transitionCnt Mod 2
'    transitionCnt = transitionCnt - check_result
'    If transitionCnt = 0 Then
'        EyeWidth = 1
'    Else
'        EyeWidth = transitionCnt / 2 'Mod 2
'    End If
    
    If transitionCnt = 0 And Record_bit Like "0" Then
    
        EyeWidth = 0
    
    Else
    
        EyeWidth = MaxEyeWidth
        
    End If
    
    
    
End Function

Public Function ReadEfuseDecimal_from_Dictionary(keyname As String, EfuseDeimal As SiteLong) As Integer
    keyname = UCase(Trim(keyname))
    If Not EfuseDecimalDictionary.Exists(keyname) Then
        'Stop    ' Please Check, no data in Dictionary
        ReadEfuseDecimal_from_Dictionary = -1
    Else
        EfuseDeimal = EfuseDecimalDictionary(keyname)
        ReadEfuseDecimal_from_Dictionary = 1
    End If
End Function


Public Function SetWriteDecimalPassFlagToEFuse(eFuseDefNameList As String, eFuseType As String, ActionFailFlag As String) As Long

    Dim i As Integer
    Dim site As Variant
    
    Dim PassFailFlag As New SiteBoolean
    Dim DecimalData As New SiteLong
    
    
    Dim eFuseDefNameArray() As String
    Dim eFuseDefName As String
    
    Dim GetEfuseDataNum As Integer
    
    '20210406 Modify for new Efuse
    Dim opbank As eFuseBdfBank
    Dim field As eFuseBdfField
    
    
    
    'Debug.Print "eFuseDefNameList :", eFuseDefNameList
    'Debug.Print "eFuseDefName :",
    
    
    '/ *** --------------------- ***/
    '/ *** --- PassFailFlag  --- ***/
    '/ *** --------------------- ***/
    
    For Each site In TheExec.sites
        If TheExec.Flow.SiteFlag(site, ActionFailFlag) = 1 Then
            PassFailFlag = False
        Else
            PassFailFlag = True
        End If
    Next site
    
    TheExec.Flow.TestLimit PassFailFlag, Tname:=ActionFailFlag, ForceResults:=tlForceNone 'eng_forceflow_transfer
    
    '/ *** --------------------- ***/
    '/ *** --- DecimalData   --- ***/
    '/ *** --------------------- ***/
    
    
    eFuseDefNameArray = Split(eFuseDefNameList, ",")
    
    For i = 0 To UBound(eFuseDefNameArray)
        eFuseDefName = eFuseDefNameArray(i)
        
        'Debug.Print eFuseDefName,
        
        GetEfuseDataNum = ReadEfuseDecimal_from_Dictionary(eFuseDefName, DecimalData)
        
        If GetEfuseDataNum < 0 Then
            TheExec.Datalog.WriteComment " *** EFUSE data not found! *** " & eFuseDefName
        Else
            TheExec.Flow.TestLimit DecimalData, Tname:=eFuseDefName, scaletype:=scaleNoScaling, formatStr:="%d", ForceResults:=tlForceFlow
        End If
        
        
    
        For Each site In TheExec.sites
        
            If GetEfuseDataNum < 0 Then
                PassFailFlag(site) = False
            Else
            
            End If
            
            'Call auto_eFuse_SetPatTestPass_Flag(eFuseType, eFuseDefName, PassFailFlag(site))
            'Call auto_eFuse_SetWriteDecimal(eFuseType, eFuseDefName, DecimalData(site), True, False)
        
        Next site
        
        '20210406 Modify for new Efuse
        Set opbank = GetBdfBank(eFuseType)
        Set field = opbank.Fields(eFuseDefName)
        opbank.SetEfuse field.name, DecimalData, PassFailFlag, , , , True

    Next i
    
    'Debug.Print
    'Debug.Print

End Function




Public Function TrimImpedance_M9(Optional patset As Pattern, Optional CPUA_Flag_In_Pat As Boolean, _
    Optional MeasR_Pins_SingleEnd As PinList, Optional MeasR_Pins_Differential As PinList, _
    Optional MeasR_TrimTarget As Double, _
    Optional ForceCondition As String, _
    Optional DigSrc_pin As PinList, Optional DigSrc_Sample_Size As Long, _
    Optional TrimStart_Dec As Long, Optional TrimEnd_Dec As Long, Optional TrimStep As Long, _
    Optional TrimResultByPins As PinList, _
    Optional eFuseName As String, Optional eFuseType As String, _
    Optional storename As String, _
    Optional PullDown_Mode As Boolean, _
    Optional IRange_mA As Double, _
    Optional Validating_ As Boolean) As Long
    

    Dim Pin_Ary() As String
    Dim Pin_Cnt As Long
    Dim TrimCode As Long
    Dim site As Variant
    Dim pin As Variant
    
    Dim TrimDSPwave As New DSPWave
    Dim InDSPWave As New DSPWave
    Dim MeasureValue As New PinListData
    Dim ColseToTargetValue As New PinListData
    Dim CloseToTargetTrimCode As New PinListData
    Dim CompareResult As New PinListData
    Dim CompareResult_BelowTarget As New PinListData
    Dim CompareResult_OverTarget As New PinListData
    
    
    
    Dim MeasR_Pins As String
    
    Dim TrimResult As New SiteLong
    
    
    '/* ----------------------------------------- */
    '/* --- added by Kaino for CZ style Tname --- */
    '/* ----------------------------------------- */
    Dim TNameUV7 As String
    Dim TNameUV6 As String
    Dim TNameUV5 As String
    Dim TNamePin As String
    
    Call CZ_Style_TName_InstanceInfo_Reg(PatSetName:=patset.value)
    Dim TrimCode_AllPins_M9 As PinListData
    '/* ----------------------------------------- */
    Dim upperbound As Double
    Dim lowerbound As Double
    
    Dim ForceValueStr As String
    Dim ForceValueArray() As String
    Dim patt_ary() As String
    Dim pat_count As Long
    
    
    Dim MeasR_Pins_P As String
    Dim MeasR_Pins_N As String
    


    Call Randomize
    upperbound = MeasR_TrimTarget + MeasR_TrimTarget * 20 / 100
    lowerbound = MeasR_TrimTarget - MeasR_TrimTarget * 20 / 100

    
    If Validating_ Then
        Call PrLoadPattern(patset.value)
        Exit Function
    End If

    If IRange_mA = 0 Then
        IRange_mA = 50
    End If

    
    
    Call PATT_GetPatListFromPatternSet(patset.value, patt_ary, pat_count)
    

    
    '/* ----------------------------------------- */
    '/*    SingleEnd    /    Differential
    '/* ----------------------------------------- */

    If MeasR_Pins_SingleEnd <> "" Then
        TheExec.DataManager.DecomposePinList MeasR_Pins_SingleEnd, Pin_Ary, Pin_Cnt
        
        For Each pin In Pin_Ary
    
            MeasR_Pins = MeasR_Pins & pin & ","
        Next pin
    
        MeasR_Pins = left(MeasR_Pins, Len(MeasR_Pins) - 1)
        
        
    Else
        'TheExec.DataManager.DecomposePinList MeasR_Pins_Differential, Pin_Ary, Pin_Cnt
        Pin_Cnt = MeasurePins_to_Differential_Pair(MeasR_Pins_Differential, MeasR_Pins_P, MeasR_Pins_N)
        
        If Pin_Cnt < 0 Then
            Stop
        Else
            TheExec.Datalog.WriteComment " *** vvv *** Differential Pair P pis: " & MeasR_Pins_P
            TheExec.Datalog.WriteComment " *** vvv *** Differential Pair N pis: " & MeasR_Pins_N
        
        End If
    End If
    
    

    
    
    '/* ----------------------------------------- */
    Call ForceConditionToFocreValue(ForceCondition, ForceValueStr, ForceValueArray)
    
    
    
    
    
    TheHdw.Digital.ApplyLevelsTiming True, True, True, tlPowered
    TheHdw.Digital.Patgen.Halt
    TheHdw.Digital.Patgen.Continue 0, cpuA + cpuB + cpuC + cpuD
    TheHdw.Digital.Patgen.HaltMode = tlHaltOnOpcode
    TheHdw.Patterns(patset).Load
    
    
    
    'Call PATT_GetPatListFromPatternSet(PatSet.Value, PattArray, PatCount)
    
    
    TrimDSPwave.CreateConstant 0, 1, DspLong
    InDSPWave.CreateConstant 0, DigSrc_Sample_Size, DspLong
    
   

    
    
    For TrimCode = TrimStart_Dec To TrimEnd_Dec Step TrimStep
    
        For Each site In TheExec.sites
            TrimDSPwave.Element(0) = TrimCode
            InDSPWave = TrimDSPwave.ConvertStreamTo(tldspSerial, DigSrc_Sample_Size, 0, Bit0IsMsb)
        Next site
        
        '/* ----------------------------------------- */
        '/* --- Setup Trim Code
        '/* --- Pattern start
        '/* --- Measure R
        '/* ----------------------------------------- */
        Call SetupDigSrcDspWave(patt_ary(0), DigSrc_pin, "TrimImpedance", DigSrc_Sample_Size, InDSPWave)
        
        Call TheHdw.Patterns(patt_ary(0)).start
        
        If MeasR_Pins_SingleEnd <> "" Then
            TrimImpedance_MeasR CPUA_Flag_In_Pat, MeasR_Pins, MeasR_Pins_P, MeasR_Pins_N, ForceValueStr, MeasureValue, False, PullDown_Mode, IRange_mA
        Else
            TrimImpedance_MeasR CPUA_Flag_In_Pat, MeasR_Pins, MeasR_Pins_P, MeasR_Pins_N, ForceValueStr, MeasureValue, True, PullDown_Mode, IRange_mA
        End If
        
        
        
        If TheExec.TesterMode = testModeOffline Then
            If MeasR_Pins_SingleEnd <> "" Then
                MeasureValue = TheHdw.PPMU.Pins(MeasR_Pins).Read(tlPPMUReadMeasurements, pc_Def_PPMU_ReadPoint)
            Else
                MeasureValue = TheHdw.PPMU.Pins(MeasR_Pins_P).Read(tlPPMUReadMeasurements, pc_Def_PPMU_ReadPoint)
            End If
            
            For Each pin In MeasureValue.Pins
                For Each site In TheExec.sites
                    MeasureValue.Pins(pin) = (upperbound - lowerbound + 1) * Rnd() + lowerbound + 120
                Next site
            Next pin
        End If
        '/* ----------------------------------------- */
        '/* --- Show Trim Log
        '/* ----------------------------------------- */
        
        Report_TestLimit_by_CZ_Format MeasureValue, unit:=unitOhm, ForceVal:=ForceValueArray(0), ForceUnit:=unitVolt, UserVar7:="Trim", TestSeqNum:=TrimCode
        TheExec.Datalog.WriteComment " *** ^^^ Trim Code = " & TrimCode & " ^^^ *** "

        '/* ----------------------------------------- */
        '/* --- ColseToTargetValue / CloseToTargetTrimCode
        '/* ----------------------------------------- */
       
        'If ColseToTargetValue.Pins.Count = 0 Then
        If TrimCode = TrimStart_Dec Then
            '/* ----------------------------------------- */
            '/* --- Create Pins in pinlistdata
            '/* ----------------------------------------- */
            ColseToTargetValue = MeasureValue.Copy
            CloseToTargetTrimCode = MeasureValue.Copy   '0 ' MeasureValue.Math.Multiply(0).Add(TrimCode)
            'CloseToTargetTrimCode = TrimCode
            
            '/* ----------------------------------------- */
            '/* --- Reset the value
            '/* ----------------------------------------- */
        
            ColseToTargetValue = 0
            CloseToTargetTrimCode = 0
        End If
        'Else
        
        
        
        '/* --- Closest To Target --- */
        CompareResult = MeasureValue.Math.Subtract(MeasR_TrimTarget).Abs.compare(LessThan, ColseToTargetValue.Math.Subtract(MeasR_TrimTarget).Abs)
        
        
        '/* --- Below To Target --- */
        CompareResult_BelowTarget = MeasureValue.Math.compare(LessThanOrEqualTo, MeasR_TrimTarget)


        '/* --- Over To Target for ColseToTargetValue --- */ 2017-11-01
        CompareResult_OverTarget = ColseToTargetValue.Math.compare(GreaterThan, MeasR_TrimTarget)
        
              
        
        
        For Each pin In MeasureValue.Pins
            For Each site In TheExec.sites
                If CompareResult.Pins(pin).value And CompareResult_BelowTarget.Pins(pin).value Then     '/* --- Closest To Target --- */ && '/* --- Below To Target --- */
                
                    CloseToTargetTrimCode.Pins(pin).value = TrimCode
                    ColseToTargetValue.Pins(pin).value = MeasureValue.Pins(pin).value
                    
                'ElseIf CompareResult_OverTarget.Pins(Pin).Value And (CompareResult.Pins(Pin).Value Or CompareResult_BelowTarget.Pins(Pin).Value) Then    '/* --- Closest To Target --- */ && '/* --- Over To Target for All CompareResult_OverTarget --- */ 2017-11-01
                    ' for the impedance for all trim code > target
                    
                '    CloseToTargetTrimCode.Pins(Pin).Value = TrimCode
                '    ColseToTargetValue.Pins(Pin).Value = MeasureValue.Pins(Pin).Value
                    
                Else
                     If TrimCode = TrimEnd_Dec Then
                    
                        If ColseToTargetValue.Pins(pin).value = 0 Then
                            CloseToTargetTrimCode.Pins(pin).value = TrimCode
                            ColseToTargetValue.Pins(pin).value = MeasureValue.Pins(pin).value
                        End If
                
                    End If
                
                End If
            Next site
        Next pin
            'CloseToTargetTrimCode = MeasureValue.Math.Multiply(0).Add(TrimCode)
        'End If
        
        
        
        
        
        
        
        
        
    
    
        TheHdw.DSSC.Pins(DigSrc_pin).Pattern(patt_ary(0)).Source.Signals.Unload
    Next TrimCode
    
    
    
    '/* ----------------------------------------- */
    '/* --- Show Select Impedance & Trim Code
    '/* ----------------------------------------- */
    
    Report_TestLimit_by_CZ_Format ColseToTargetValue, unit:=unitOhm, ForceVal:=ForceValueArray(0), ForceUnit:=unitVolt, UserVar6:="Imp", UserVar5:="Closest"
   
    TheExec.Datalog.WriteComment " *** ^^^ Closest & Below to Target Impedance ( " & MeasR_TrimTarget & " ohm ) ^^^ *** Impedance"
    
    Report_TestLimit_by_CZ_Format CloseToTargetTrimCode, MeasType:="c", UserVar6:="Code", UserVar5:="Closest"
    
    TheExec.Datalog.WriteComment " *** ^^^ Closest & Below to Target Impedance ( " & MeasR_TrimTarget & " ohm ) ^^^ *** Trim Code"
    
    'TheExec.Datalog.WriteComment " *** ^^^ Closest & Below to Target Impedance ( " & MeasR_TrimTarget & " ohm )  Trim Code  ^^^ *** "
    
    
    
    If InStr(1, TrimResultByPins, "grp", vbTextCompare) > 0 Then
    
        CloseToTargetTrimCode = CloseToTargetTrimCode.Analyze.mean
        
        
        For Each site In TheExec.sites
        
            TrimDSPwave.Element(0) = Round(CloseToTargetTrimCode.Pins(0).value, 0)
            
            
            TheExec.Datalog.WriteComment " *** ^^^ Trim Code : site (" & site & "): " & TrimDSPwave.Element(0) & " , ( " & TrimResultByPins & " ) ^^^ *** "
            InDSPWave = TrimDSPwave.ConvertStreamTo(tldspSerial, DigSrc_Sample_Size, 0, Bit0IsMsb)
            
            TrimResult = TrimDSPwave.Element(0)
        Next site
        
        
        TNamePin = "x"
        'Stop
    ElseIf TrimResultByPins <> "" Then
        
        For Each site In TheExec.sites
        
            TrimDSPwave.Element(0) = CloseToTargetTrimCode.Pins(TrimResultByPins).value
            
            TheExec.Datalog.WriteComment " *** ^^^ Trim Code : site (" & site & "): " & TrimDSPwave.Element(0) & " , ( " & TrimResultByPins & " ) ^^^ *** "
            
            InDSPWave = TrimDSPwave.ConvertStreamTo(tldspSerial, DigSrc_Sample_Size, 0, Bit0IsMsb)
            
            TrimResult = TrimDSPwave.Element(0)
        Next site
        
        TNamePin = TrimResultByPins
    
    Else
        Stop
    
    End If
    
    
    
    '/* ----------------------------------------- */
    '/* --- Fianl Code test
    '/* ----------------------------------------- */
    
    '/* ----------------------------------------- */
    '/* --- Setup Trim Code
    '/* --- Pattern start
    '/* --- Measure R
    '/* ----------------------------------------- */
    
    
    Call SetupDigSrcDspWave(patt_ary(0), DigSrc_pin, "TrimImpedance", DigSrc_Sample_Size, InDSPWave)
    
    Call TheHdw.Patterns(patt_ary(0)).start
    
    If MeasR_Pins_SingleEnd <> "" Then
        TrimImpedance_MeasR CPUA_Flag_In_Pat, MeasR_Pins, MeasR_Pins_P, MeasR_Pins_N, ForceValueStr, MeasureValue, False, PullDown_Mode, IRange_mA
    Else
        TrimImpedance_MeasR CPUA_Flag_In_Pat, MeasR_Pins, MeasR_Pins_P, MeasR_Pins_N, ForceValueStr, MeasureValue, True, PullDown_Mode, IRange_mA
    End If
    
    If TheExec.TesterMode = testModeOffline Then
        If MeasR_Pins_SingleEnd <> "" Then
                MeasureValue = TheHdw.PPMU.Pins(MeasR_Pins).Read(tlPPMUReadMeasurements, pc_Def_PPMU_ReadPoint)
            Else
                MeasureValue = TheHdw.PPMU.Pins(MeasR_Pins_P).Read(tlPPMUReadMeasurements, pc_Def_PPMU_ReadPoint)
            End If
        
        For Each pin In MeasureValue.Pins
            For Each site In TheExec.sites
                MeasureValue.Pins(pin) = (upperbound - lowerbound + 1) * Rnd() + lowerbound
            Next site
        Next pin
    End If
    '/* ----------------------------------------- */
    '/* --- Show Trim Log
    '/* ----------------------------------------- */
    
    TheExec.Datalog.WriteComment " *** Final Trim Impedance ***"
    
    For Each pin In MeasureValue.Pins
        Report_TestLimit_by_CZ_Format MeasureValue.Pins(pin), unit:=unitOhm, ForceVal:=ForceValueArray(0), ForceUnit:=unitVolt, ForceResults:=tlForceFlow, UserVar5:="FINAL", UserVar6:="IMP"
    Next pin
    TheExec.Datalog.WriteComment " *** Final Trim Code ***"
    Report_TestLimit_by_CZ_Format TrimDSPwave.Element(0), ForceResults:=tlForceFlow, MeasType:="c", PinName:=TNamePin, UserVar5:="FINAL", UserVar6:="CODE"
    
    
    
    
    
    
    TheHdw.DSSC.Pins(DigSrc_pin).Pattern(patt_ary(0)).Source.Signals.Unload
    
    
    
    
    Call HardIP_WriteFuncResult(m_testName:=TheExec.DataManager.instancename)
    
    
    If eFuseName <> "" Then
         TheExec.Datalog.WriteComment " *** save efuse Trim Code ***"
         
        Call SaveEfuseDecimal_to_Dictionary(eFuseName, TrimResult)
        'Call WriteDecToEFuse(eFuseName, TrimResult)
    End If
    
    
    
    If storename <> "" Then
        Call AddStoredCaptureData(storename, InDSPWave)
    End If
    
    
    
    DebugPrintFunc patset.value
    
    Call CZ_Style_TName_InstanceInfo_Clear   'added by Kaino for CZ style Tname
    

End Function




Public Function Show_Fuse_Check(ActionPassFlagName As String) As Long
    Dim site As Variant
    Dim Fuse_Check_Result As New SiteLong
    
    For Each site In TheExec.sites
        Fuse_Check_Result = TheExec.Flow.SiteFlag(site, ActionPassFlagName)
    Next site
    
    TheExec.Flow.TestLimit Fuse_Check_Result, Tname:=UCase("ADDRIO_GOT_TRIM_CODE_IN_EFUSE_AND_DO_NOT_Trim"), ForceResults:=tlForceFlow

End Function



Public Function HIP_eFuse_Read_and_Set_SiteVar(FuseType As String, m_catename As String, Dict_Store_Code_Name As String, dspwavesize As Long, _
                    Optional Get_Efuse_Dec_Flag As Boolean = False, Optional Dict_Store_Dec_Name As String = vbNullString, _
                    Optional SiteVarName As String) As Long

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
        
    On Error GoTo errHandler
    
    Read_Code.CreateConstant 0, dspwavesize

    If Get_Efuse_Dec_Flag = True Then
        Read_Value.CreateConstant 0, 1
    End If
    
    '20210406 Add for New Efuse
    Efuse_Value = GetEfuseHipValue(FuseType, m_catename)

    For Each site In TheExec.sites

        'Efuse_Value(site) = auto_eFuse_GetReadDecimal(FuseType, m_catename, True)
'''''        Efuse_Value(Site) = CLng(Site) + 8

        If Get_Efuse_Dec_Flag = True Then
            Read_Value.Element(0) = Efuse_Value(site)
        End If

        TempVal = Efuse_Value(site)
        For i = 0 To dspwavesize - 1
            Read_Code.Element(i) = TempVal Mod 2
            TempVal = TempVal \ 2
        Next i
        
        If Efuse_Value(site) = 0 Then                                                'If Read out value = 0 then bin out
            Efuse_Value_Chk(site) = 0
        Else
            Efuse_Value_Chk(site) = 1
        End If
        
        If SiteVarName <> "" Then
            TheExec.sites(site).SiteVariableValue(SiteVarName) = TheExec.sites(site).SiteVariableValue(SiteVarName) + Efuse_Value_Chk(site)
        End If
        
    Next site

    If SiteVarName = "" Then
        TheExec.Flow.TestLimit resultVal:=Efuse_Value_Chk, lowVal:=1, hiVal:=1, Tname:="NonZero_Val_Chk", ForceResults:=tlForceFlow
    Else
        'do not show fail
        TheExec.Flow.TestLimit resultVal:=Efuse_Value_Chk, Tname:="NonZero_Val_Chk", ForceResults:=tlForceFlow
        
        'TheExec.Flow.TestLimit resultVal:=Efuse_Value_Chk, ForceResults:=tlForceFlow
    End If
        
    Call AddStoredCaptureData(Dict_Store_Code_Name, Read_Code)
    
    If Get_Efuse_Dec_Flag = True Then
        Call AddStoredCaptureData(Dict_Store_Dec_Name, Read_Value)
    End If
    
    Exit Function

errHandler:
    TheExec.Datalog.WriteComment "Error in HIP_eFuse_Read_and_Set_Flag"
    If AbortTest Then Exit Function Else Resume Next
    
End Function


Public Function Show_Num_Of_Data_In_Fuse(SiteVarName As String) As Long
    Dim site As Variant
    Dim Fuse_Check_Result As New SiteLong
        
    For Each site In TheExec.sites
        Fuse_Check_Result = TheExec.sites(site).SiteVariableValue(SiteVarName)
    Next site
    
    TheExec.Flow.TestLimit Fuse_Check_Result, Tname:=UCase("ADDRIO_GOT_TRIM_CODE_IN_EFUSE_AND_DO_NOT_Trim"), ForceResults:=tlForceFlow

End Function


Public Function CombineTrimCodeInEFuseAndTrimII(StoreNameList As String, StoreNameList_in_fuse As String, StoreNameList_Trim As String, SiteVarName As String) As Long

    Dim StoredData As DSPWave
    Dim StoredDataDec As New DSPWave
    Dim StoreDataSiteVariant As New SiteLong
    
    
    Dim StoredDataTrim As New DSPWave
    Dim StoredDataInFuse As New DSPWave
    
    Dim DSPWaveDataSize As Long
    
    Dim site As Variant
    Dim i As Integer
    
    Dim StoreName_array() As String
    Dim StoreNameTrim_array() As String
    Dim StoreNameInFuse_array() As String
    
    Dim storename As String
    Dim StoreNameTrim As String
    Dim StoreNameInfuse As String
    
    Dim PrintOutString As String
    
    Dim TrimSitesNum As Integer
    
    StoreName_array = Split(StoreNameList, ",")
    StoreNameTrim_array = Split(StoreNameList_Trim, ",")
    StoreNameInFuse_array = Split(StoreNameList_in_fuse, ",")
    
    
    If UBound(StoreNameTrim_array) <> UBound(StoreNameInFuse_array) Then
        Stop
    Else
        For i = 0 To UBound(StoreNameTrim_array)
            storename = StoreName_array(i)
            StoreNameTrim = StoreNameTrim_array(i)
            StoreNameInfuse = StoreNameInFuse_array(i)

            '/* --- get DSPWaveDataSize --- */
            'fuse
            If IsExists_StoredCaptureData(StoreNameInfuse) Then
                StoredDataInFuse = GetStoredCaptureData(StoreNameInfuse)
            End If
            'trim
            If IsExists_StoredCaptureData(StoreNameTrim) Then
                StoredDataTrim = GetStoredCaptureData(StoreNameTrim)
            End If
            
                        
            TrimSitesNum = 0
            For Each site In TheExec.sites
                If TheExec.sites(site).SiteVariableValue(SiteVarName) > 0 Then
                    'fuse
                    DSPWaveDataSize = StoredDataInFuse(site).DataSize
                Else
                    'trim
                    DSPWaveDataSize = StoredDataTrim(site).DataSize
                    TrimSitesNum = TrimSitesNum + 1
                End If
            Next site
            

            '/* --- create StoredData dspwave --- */
            Set StoredData = New DSPWave
            StoredData.CreateConstant 0, DSPWaveDataSize, DspLong
            
            '/* --- Combine TrimCode In EFuse And Trim --- */
            PrintOutString = " *** Data Source: site "
            For Each site In TheExec.sites
                If TheExec.sites(site).SiteVariableValue(SiteVarName) > 0 Then
                    'fuse
                    PrintOutString = PrintOutString & "(" & Trim(str(site)) & "): *FuseReaded*; "
                    StoredData(site).data = StoredDataInFuse(site).data
                Else
                    'trim
                    PrintOutString = PrintOutString & "(" & Trim(str(site)) & "): *TrimResult*; "
                    StoredData(site).data = StoredDataTrim(site).data
                End If
            Next site
            
            
            
            
            Call rundsp.ConvertToLongAndSerialToParrel(StoredData, DSPWaveDataSize, StoredDataDec)
            
            '/*--- Report the DigCap Data ---*/
            StoreDataSiteVariant = StoredDataDec.Element(0)
            
            TheExec.Flow.TestLimit StoreDataSiteVariant, Tname:=storename, formatStr:="%d", ForceResults:=tlForceFlow, scaletype:=scaleNoScaling
            
            
            
            PrintOutString = PrintOutString & " ( " & storename & " )"
            TheExec.Datalog.WriteComment PrintOutString & Chr(13)
            
            
            
            Call AddStoredCaptureData(storename, StoredData)
            
            'Debug.Print "--------------read store-------------------------------------"
            'StoredData = GetStoredCaptureData(StoreName)
            
            'For Each Site In TheExec.sites
            '    Debug.Print Site, StoreName, StoredData(Site).DataSize
            'Next Site
            
            'Debug.Print "------------------------------------------------------------"
            
            Set StoredData = Nothing
        Next i

    
    End If
    
    
    

End Function


Public Function RemoveAllEfuseDecimal_to_Dictionary()
    EfuseDecimalDictionary.RemoveAll
End Function




Public Function TrimImpedance_M9_Check_TrimCode_V1(Optional patset As Pattern, Optional CPUA_Flag_In_Pat As Boolean, _
    Optional MeasR_Pins_SingleEnd As PinList, Optional MeasR_Pins_Differential As PinList, _
    Optional MeasR_TrimTarget As Double, _
    Optional ForceCondition As String, _
    Optional DigSrc_pin As PinList, Optional DigSrc_Sample_Size As Long, _
    Optional TrimCodeStoreName As String, Optional TrimEnd_Dec As Long, Optional TrimStep As Long, _
    Optional TrimResultByPins As PinList, _
    Optional eFuseName As String, Optional eFuseType As String, _
    Optional storename As String, _
    Optional PullDown_Mode As Boolean, _
    Optional IRange_mA As Double, _
    Optional FuseDefault_check As String, _
    Optional StoreName_check As String, _
    Optional ReadCodeWithoutChecking As Boolean, _
    Optional Validating_ As Boolean) As Long
    '/* 2017-12-04 update by Kaino  for IRange_mA */

    Dim Pin_Ary() As String
    Dim Pin_Cnt As Long
    Dim TrimCode As Long
    Dim site As Variant
    Dim pin As Variant
    
    Dim TrimDSPwave As New DSPWave
    Dim InDSPWave As New DSPWave
    Dim MeasureValue As New PinListData
    Dim ColseToTargetValue As New PinListData
    Dim CloseToTargetTrimCode As New PinListData
    Dim CompareResult As New PinListData
    Dim CompareResult_BelowTarget As New PinListData
    
    Dim MeasR_Pins As String
    
    Dim TrimResult As New SiteLong
    
    
    '/* ----------------------------------------- */
    '/* --- added by Kaino for CZ style Tname --- */
    '/* ----------------------------------------- */
    Dim TNameUV7 As String
    Dim TNameUV6 As String
    Dim TNameUV5 As String
    Dim TNamePin As String
    
    Call CZ_Style_TName_InstanceInfo_Reg(PatSetName:=patset.value)
    Dim TrimCode_AllPins_M9 As PinListData
    '/* ----------------------------------------- */
    Dim upperbound As Double
    Dim lowerbound As Double
    
    Dim ForceValueStr As String
    Dim ForceValueArray() As String
    Dim patt_ary() As String
    Dim pat_count As Long
    
    
    Dim MeasR_Pins_P As String
    Dim MeasR_Pins_N As String
    
    Dim i As Integer
    
    Dim CodeForCheck As New SiteLong
    Dim UserVar5 As String

    Call Randomize
    upperbound = MeasR_TrimTarget + MeasR_TrimTarget * 20 / 100
    lowerbound = MeasR_TrimTarget - MeasR_TrimTarget * 20 / 100

    
    If Validating_ Then
        Call PrLoadPattern(patset.value)
        Exit Function
    End If
        '/* 2017-12-04 update by Kaino */
    If IRange_mA = 0 Then
        IRange_mA = 50
    End If
    

    
    
    Call PATT_GetPatListFromPatternSet(patset.value, patt_ary, pat_count)
    

    
    '/* ----------------------------------------- */
    '/*    SingleEnd    /    Differential
    '/* ----------------------------------------- */

    If MeasR_Pins_SingleEnd <> "" Then
        TheExec.DataManager.DecomposePinList MeasR_Pins_SingleEnd, Pin_Ary, Pin_Cnt
        
        For Each pin In Pin_Ary
    
            MeasR_Pins = MeasR_Pins & pin & ","
        Next pin
    
        MeasR_Pins = left(MeasR_Pins, Len(MeasR_Pins) - 1)
        
        
    Else
        'TheExec.DataManager.DecomposePinList MeasR_Pins_Differential, Pin_Ary, Pin_Cnt
        Pin_Cnt = MeasurePins_to_Differential_Pair(MeasR_Pins_Differential, MeasR_Pins_P, MeasR_Pins_N)
        
        If Pin_Cnt < 0 Then
            Stop
        Else
            TheExec.Datalog.WriteComment " *** vvv *** Differential Pair P pis: " & MeasR_Pins_P
            TheExec.Datalog.WriteComment " *** vvv *** Differential Pair N pis: " & MeasR_Pins_N
        
        End If
    End If
    
    

    
    
    '/* ----------------------------------------- */
    Call ForceConditionToFocreValue(ForceCondition, ForceValueStr, ForceValueArray)
    
    
    
    
    
    TheHdw.Digital.ApplyLevelsTiming True, True, True, tlPowered
    TheHdw.Digital.Patgen.Halt
    TheHdw.Digital.Patgen.Continue 0, cpuA + cpuB + cpuC + cpuD
    TheHdw.Digital.Patgen.HaltMode = tlHaltOnOpcode
    TheHdw.Patterns(patset).Load
    
    

    
    
    
    
    
    TrimDSPwave.CreateConstant 0, 1, DspLong
    InDSPWave.CreateConstant 0, DigSrc_Sample_Size, DspLong
    
    '20210406 Modify for new Efuse
    Dim opbank As New eFuseBdfBank
    Dim field As New eFuseBdfField
    Dim fieldStr As Variant
    Set opbank = GetBdfBank("CFG")

    If FuseDefault_check <> "" Then
        TheExec.Datalog.WriteComment " *** Trim Code from Fuse Default Value  *** " & FuseDefault_check
        
        For Each fieldStr In opbank.Fields  '20210406 Modify for new Efuse
        'For i = LBound(CFGFuse.category) To UBound(CFGFuse.category)
            '20210406 Modify for new Efuse
            Set field = opbank.Fields(fieldStr)
            If UCase(field.name) = UCase(FuseDefault_check) Then
            'If UCase(CFGFuse.category(i).name) = UCase(FuseDefault_check) Then
            
                'Stop
    
                
                For Each site In TheExec.sites
                    'TrimDSPwave.Element(0) = Site_TrimCode
                    ''20180325 judge default code is hex or decimal
                    '20210406 Modify for new Efuse
                    If LCase(field.Default) Like "*0x*" Then
                        TrimDSPwave.Element(0) = auto_HexStr2Value(field.Default)
                    Else
                        TrimDSPwave.Element(0) = val(field.Default)
                    End If
'                    If LCase(CFGFuse.category(i).DefaultValue) Like "*0x*" Then
'                        TrimDSPwave.Element(0) = auto_HexStr2Value(CFGFuse.category(i).DefaultValue)
'                    Else
'                        TrimDSPwave.Element(0) = val(CFGFuse.category(i).DefaultValue)
'                    End If
                    
                    InDSPWave = TrimDSPwave.ConvertStreamTo(tldspSerial, DigSrc_Sample_Size, 0, Bit0IsMsb)
                Next site
                
                
                Exit For
            End If
        Next
        'Next i
        
    ElseIf StoreName_check <> "" Then
        TheExec.Datalog.WriteComment " *** Trim Code from Store *** " & StoreName_check
        
        If IsExists_StoredCaptureData(StoreName_check) Then
            InDSPWave = GetStoredCaptureData(StoreName_check)
            
            Call rundsp.ConvertToLongAndSerialToParrel(InDSPWave, DigSrc_Sample_Size, TrimDSPwave)
        End If
    Else
        Stop
        
    End If
    
    'CodeForCheck = TrimDSPwave.Element(0)
    
    
    
    
    If InStr(1, TrimResultByPins, "grp", vbTextCompare) > 0 Then
        TNamePin = "x"
        'Stop
    ElseIf TrimResultByPins <> "" Then
       
        TNamePin = TrimResultByPins
   
    Else
        Stop
    
    End If
    
    
    
    If ReadCodeWithoutChecking Then
        'Stop
        'If FuseDefault_check <> "" Then
        '    UserVar5 = "FuseDefault"
        'ElseIf StoreName_check <> "" Then
        '    UserVar5 = "Store"
        'Else
        '
        'End If
        
        'Report_TestLimit_by_CZ_Format TrimDSPwave.Element(0), ForceResults:=tlForceFlow, MeasType:="c", PinName:=TNamePin, UserVar5:=UserVar5, UserVar6:="CODE"
        Report_TestLimit_by_CZ_Format TrimDSPwave.Element(0), ForceResults:=tlForceFlow, MeasType:="c", PinName:=TNamePin, UserVar6:="CODE"
    Else
    
        
        '/* ----------------------------------------- */
        '/* --- Fianl Code test
        '/* ----------------------------------------- */
        
        '/* ----------------------------------------- */
        '/* --- Setup Trim Code
        '/* --- Pattern start
        '/* --- Measure R
        '/* ----------------------------------------- */
        
        
        Call SetupDigSrcDspWave(patt_ary(0), DigSrc_pin, "TrimImpedance", DigSrc_Sample_Size, InDSPWave)
        
        Call TheHdw.Patterns(patt_ary(0)).start
        
        If MeasR_Pins_SingleEnd <> "" Then
                    '/* 2017-12-04 update by Kaino */
            TrimImpedance_MeasR CPUA_Flag_In_Pat, MeasR_Pins, MeasR_Pins_P, MeasR_Pins_N, ForceValueStr, MeasureValue, False, PullDown_Mode, IRange_mA
        Else
                    '/* 2017-12-04 update by Kaino */
            TrimImpedance_MeasR CPUA_Flag_In_Pat, MeasR_Pins, MeasR_Pins_P, MeasR_Pins_N, ForceValueStr, MeasureValue, True, PullDown_Mode, IRange_mA
        End If
        
        If TheExec.TesterMode = testModeOffline Then
            If MeasR_Pins_SingleEnd <> "" Then
                    MeasureValue = TheHdw.PPMU.Pins(MeasR_Pins).Read(tlPPMUReadMeasurements, pc_Def_PPMU_ReadPoint)
                Else
                    MeasureValue = TheHdw.PPMU.Pins(MeasR_Pins_P).Read(tlPPMUReadMeasurements, pc_Def_PPMU_ReadPoint)
                End If
            
            For Each pin In MeasureValue.Pins
                For Each site In TheExec.sites
                    MeasureValue.Pins(pin) = (upperbound - lowerbound + 1) * Rnd() + lowerbound
                Next site
            Next pin
        End If
        '/* ----------------------------------------- */
        '/* --- Show Trim Log
        '/* ----------------------------------------- */
        
        TheExec.Datalog.WriteComment " *** Final Trim Impedance ***"
        
        For Each pin In MeasureValue.Pins
            Report_TestLimit_by_CZ_Format MeasureValue.Pins(pin), unit:=unitOhm, ForceVal:=ForceValueArray(0), ForceUnit:=unitVolt, ForceResults:=tlForceFlow, UserVar6:="IMP"
        Next pin
        TheExec.Datalog.WriteComment " *** Final Trim Code ***"
        Report_TestLimit_by_CZ_Format TrimDSPwave.Element(0), ForceResults:=tlForceFlow, MeasType:="c", PinName:=TNamePin, UserVar6:="CODE"
        
        
        
        
        
        
        TheHdw.DSSC.Pins(DigSrc_pin).Pattern(patt_ary(0)).Source.Signals.Unload
        
        
        
        
        Call HardIP_WriteFuncResult(m_testName:=TheExec.DataManager.instancename)
    
    
    End If
    
    'Stop
    
    If eFuseName <> "" Then
         TheExec.Datalog.WriteComment " *** save efuse Trim Code *** " & UCase(eFuseName)
    
        Call SaveEfuseDecimal_to_Dictionary(eFuseName, TrimResult)
        'Call WriteDecToEFuse(eFuseName, TrimResult)
    End If
    
    
    If storename <> "" Then
        TheExec.Datalog.WriteComment " *** save Trim Code *** " & storename
        Call AddStoredCaptureData(storename, InDSPWave)
    End If
   
    
    
    
    DebugPrintFunc patset.value
    
    Call CZ_Style_TName_InstanceInfo_Clear   'added by Kaino for CZ style Tname
    

End Function

Public Function K_ADDRIO_FUSE_WRITE(eFuseName As String, eFuseType As String, ActionFailFlag As String) As Long
        'auto_eFuse_SetPatTestPass_Flag
        'call auto_eFuse_SetWriteDecimal
    Dim i As Integer
    Dim site As Variant
    
    Dim PassFailFlag As New SiteBoolean
    Dim DecimalData As New SiteLong
    
    
    Dim eFuseDefNameArray() As String
    Dim eFuseDefName As String
    
    Dim GetEfuseDataNum As Integer
    
    
    
    'Debug.Print "eFuseDefNameList :", eFuseDefNameList
    'Debug.Print "eFuseDefName :",
    
    
    '/ *** --------------------- ***/
    '/ *** --- PassFailFlag  --- ***/
    '/ *** --------------------- ***/
    
    For Each site In TheExec.sites
        If TheExec.Flow.SiteFlag(site, ActionFailFlag) = 1 Or ActionFailFlag = "" Then
            PassFailFlag = False
        Else
            PassFailFlag = True
        End If
    Next site
    
    TheExec.Flow.TestLimit PassFailFlag, Tname:=ActionFailFlag, ForceResults:=tlForceFlow
    
    '/ *** --------------------- ***/
    '/ *** --- DecimalData   --- ***/
    '/ *** --------------------- ***/
    
    
    eFuseDefNameArray = Split(eFuseName, ",")
    
    For i = 0 To UBound(eFuseDefNameArray)
        eFuseDefName = eFuseDefNameArray(i)
        
        'Debug.Print eFuseDefName,
        
        GetEfuseDataNum = ReadEfuseDecimal_from_Dictionary(eFuseDefName, DecimalData)
        
        If GetEfuseDataNum < 0 Then
            TheExec.Datalog.WriteComment " *** EFUSE data not found! *** " & eFuseDefName
        Else
            TheExec.Flow.TestLimit DecimalData, Tname:=eFuseDefName, scaletype:=scaleNoScaling, formatStr:="%d", ForceResults:=tlForceFlow
        End If
        
        
    
        For Each site In TheExec.sites
        
            If GetEfuseDataNum < 0 Then
                PassFailFlag(site) = False
            Else
            
            End If
            
            'Call auto_eFuse_SetPatTestPass_Flag(eFuseType, eFuseDefName, PassFailFlag(site))
            'Call auto_eFuse_SetWriteDecimal(eFuseType, eFuseDefName, DecimalData(site), True, False)
        
        Next site
        
        '20210406 Modify for new Efuse
        Dim opbank As eFuseBdfBank
        Dim field As eFuseBdfField
        Set opbank = GetBdfBank(eFuseType)
        Set field = opbank.Fields(eFuseDefName)
        opbank.SetEfuse field.name, DecimalData, PassFailFlag, , , , True

    Next i
    
    'Debug.Print
    'Debug.Print

End Function


Public Function K_ADDRIO_FUSE_READ(eFuseName As String, eFuseType As String, _
                                    storename As String, DspWave_SampleSize As Long, _
                                    Optional FuseDefRead As Boolean, _
                                    Optional SiteVarName As String, _
                                    Optional Validating_ As Boolean) As Long
'Public Function HIP_eFuse_Read_and_Set_SiteVar(FuseType As String, eFuseType As String, Dict_Store_Code_Name As String, dspwavesize As Long, _
'                    Optional Get_Dec_from_Fuse As Boolean = False, Optional Dict_Store_Dec_Name As String = "", _
'                    Optional SiteVarName As String) As Long

    ' Parameter : eFuse Block , eFuse Variable , data , Data Width
    ' Create dictionary , if exist then remove and re-create
    ' MUST :  if necessary , we can set limit if read out value = 0 then bin out .
    
    Dim site As Variant
    Dim Read_Code_BIN As New DSPWave
    Dim Read_Code_DEC As New DSPWave
    Dim Efuse_Value As New SiteLong
    Dim TempVal As Long
    Dim Efuse_Value_Chk As New SiteVariant
    Dim i As Long
    
    Dim Get_Dec_from_Fuse As Boolean
    
    If Validating_ Then
        Exit Function
    End If

    'On Error GoTo errHandler
    
    
    Get_Dec_from_Fuse = True
    '/* ---------------------------- */
    '/* Added by Kaino on 2019/06/05 */
    '/* Read Fuse Default Value      */
    '/* ---------------------------- */
    '/* Efuse_Value (Site)
    '/* SiteVariableValue(SiteVarName)
    '/* ---------------------------- */
        
    Dim FuseDefValue As Variant
    '20210406 Modify for new Efuse
    Dim opbank As New eFuseBdfBank
    Dim field As New eFuseBdfField
    Dim fieldStr As Variant
    Set opbank = GetBdfBank("CFG")
    
    If FuseDefRead Then
        '20210406 Modify for new Efuse
        i = 1
        For Each fieldStr In opbank.Fields  '20210406 Modify for new Efuse
        'For i = LBound(CFGFuse.category) To UBound(CFGFuse.category)
            Set field = opbank.Fields(fieldStr)
            If UCase(field.name) = UCase(eFuseName) Then
            'If UCase(CFGFuse.category(i).name) = UCase(eFuseName) Then
                FuseDefValue = field.Default
'                FuseDefValue = CFGFuse.category(i).DefaultValue
                FuseDefValue = CLng(Replace(FuseDefValue, "0x", "&H", , , vbTextCompare))
                TheExec.Datalog.WriteComment " *** Read Fuse Default Value  *** " & eFuseName & " : " & FuseDefValue & " ( 0x" & Hex(FuseDefValue) & " )"
    
                Efuse_Value = FuseDefValue
    
                Exit For
            End If
            i = i + 1
        Next
        'Next i
        
        If i > opbank.Fields.Count Then
        'If i > UBound(CFGFuse.category) Then
            TheExec.Datalog.WriteComment " *** Read Fuse Default Value  *** " & eFuseName & " : Not Found!!!"
            Stop
        End If
    Else
        '20210406 Add for New Efuse
        Efuse_Value = GetEfuseHipValue(eFuseType, eFuseName)
'        For Each site In TheExec.sites
'            Efuse_Value(site) = auto_eFuse_GetReadDecimal(eFuseType, eFuseName, True)
'
'        Next site
    End If
    '/* ---------------------------- */
    
    

    Read_Code_BIN.CreateConstant 0, DspWave_SampleSize, DspLong
    Read_Code_DEC.CreateConstant 0, 1, DspLong


    'If Get_Dec_from_Fuse = True Then
    '    Read_Code_DEC.CreateConstant 0, 1
    'End If


    
    For Each site In TheExec.sites
        Read_Code_DEC.Element(0) = Efuse_Value
        Read_Code_BIN = Read_Code_DEC.ConvertStreamTo(tldspSerial, DspWave_SampleSize, 0, Bit0IsMsb)
        
        
        If Efuse_Value(site) = 0 Then                                                'If Read out value = 0 then bin out
            Efuse_Value_Chk(site) = 0
        Else
            Efuse_Value_Chk(site) = 1
        End If
        
        If SiteVarName <> "" Then
            TheExec.sites(site).SiteVariableValue(SiteVarName) = TheExec.sites(site).SiteVariableValue(SiteVarName) + Efuse_Value_Chk(site)
        End If
        
    Next site

    
        


    If SiteVarName = "" Then
        TheExec.Flow.TestLimit resultVal:=Efuse_Value_Chk, lowVal:=1, hiVal:=1, Tname:="NonZero_Val_Chk", ForceResults:=tlForceFlow
    Else
        'do not show fail
        TheExec.Flow.TestLimit resultVal:=Efuse_Value_Chk, Tname:="NonZero_Val_Chk", ForceResults:=tlForceFlow
        
        'TheExec.Flow.TestLimit resultVal:=Efuse_Value_Chk, ForceResults:=tlForceFlow
    End If
        
    Call AddStoredCaptureData(storename, Read_Code_BIN)
    
    If Get_Dec_from_Fuse = True Then
        Dim Dict_Store_Dec_Name As String
        'Stop
        Dict_Store_Dec_Name = UCase(storename & "_DEC")
        Call AddStoredCaptureData(Dict_Store_Dec_Name, Read_Code_DEC)
    End If
    
    Exit Function

errHandler:
    TheExec.Datalog.WriteComment "Error in HIP_eFuse_Read_and_Set_Flag"
    If AbortTest Then Exit Function Else Resume Next
    
End Function

Public Function K_Show_Data_Count_In_Fuse(SiteVarName As String, Optional Validating_ As Boolean) As Long
    Dim site As Variant
    Dim Fuse_Check_Result As New SiteLong
            
    If Validating_ Then
        Exit Function
    End If
        
    For Each site In TheExec.sites
        Fuse_Check_Result = TheExec.sites(site).SiteVariableValue(SiteVarName)
    Next site
    
    TheExec.Flow.TestLimit Fuse_Check_Result, Tname:=UCase("ADDRIO_TRIM_CODE_IN_EFUSE_Count"), ForceResults:=tlForceFlow
    TheExec.Datalog.WriteComment vbNullString
    TheExec.Datalog.WriteComment "   *** Except for reading from Fuse Default value *** "
    TheExec.Datalog.WriteComment vbNullString
    
    If TheExec.Flow.enableWord("ADDRIO_TRIM_ON") Then
        TheExec.Datalog.WriteComment "   *** Trimming is Turned ON by Enable Word *** "
    End If
    If TheExec.Flow.enableWord("ADDRIO_TRIM_OFF") Then
        TheExec.Datalog.WriteComment "   *** Trimming is Turned OFF by Enable Word *** "
    End If

End Function
''''' Modify by TY
'Public Function IsExists_StoredCaptureData(keyname As String) As Boolean
'    keyname = LCase(keyname)
'    IsExists_StoredCaptureData = gDictDSPWaves.Exists(keyname)
'End Function

Public Function K_Merge_TrimCode_From_Fuse_n_Trim(StoreNameList As String, StoreNameList_in_fuse As String, StoreNameList_Trim As String, FuseDefValueFlagList As String, SiteVarName As String) As Long

    Dim StoredData As DSPWave
    Dim StoredDataDec As New DSPWave
    Dim StoreDataSiteVariant As New SiteLong
    
    
    Dim StoredDataTrim As New DSPWave
    Dim StoredDataInFuse As New DSPWave
    
    Dim DSPWaveDataSize As Long
    
    Dim site As Variant
    Dim i As Integer
    
    Dim StoreName_array() As String
    Dim StoreNameTrim_array() As String
    Dim StoreNameInFuse_array() As String
    Dim FuseDefValueFlag_array() As String
    
    Dim storename As String
    Dim StoreNameTrim As String
    Dim StoreNameInfuse As String
    Dim FuseDefValueFlag As Boolean
    
    Dim PrintOutString As String
    
    Dim TrimSitesNum As Integer
    
    StoreName_array = Split(StoreNameList, ",")
    StoreNameTrim_array = Split(StoreNameList_Trim, ",")
    StoreNameInFuse_array = Split(StoreNameList_in_fuse, ",")
    FuseDefValueFlag_array = Split(FuseDefValueFlagList, ",")
    
    If UBound(StoreNameTrim_array) <> UBound(StoreNameInFuse_array) Then
        Stop
    Else
        For i = 0 To UBound(StoreNameTrim_array)
            storename = StoreName_array(i)
            StoreNameTrim = StoreNameTrim_array(i)
            StoreNameInfuse = StoreNameInFuse_array(i)
            FuseDefValueFlag = FuseDefValueFlag_array(i)

            '/* --- get DSPWaveDataSize --- */
            'fuse
            If IsExists_StoredCaptureData(StoreNameInfuse) Then
                StoredDataInFuse = GetStoredCaptureData(StoreNameInfuse)
            Else
                Set StoredDataInFuse = Nothing
            End If
            'trim
            If IsExists_StoredCaptureData(StoreNameTrim) Then
                StoredDataTrim = GetStoredCaptureData(StoreNameTrim)
            Else
                Set StoredDataTrim = Nothing
            End If
            
                        
            TrimSitesNum = 0
            For Each site In TheExec.sites
                If TheExec.sites(site).SiteVariableValue(SiteVarName) > 0 Or FuseDefValueFlag Then
                    'fuse
                    DSPWaveDataSize = StoredDataInFuse(site).DataSize
                Else
                    'trim
                    DSPWaveDataSize = StoredDataTrim(site).DataSize
                    TrimSitesNum = TrimSitesNum + 1
                End If
            Next site
            

            '/* --- create StoredData dspwave --- */
            Set StoredData = New DSPWave
            StoredData.CreateConstant 0, DSPWaveDataSize, DspLong
            
            '/* --- Combine TrimCode In EFuse And Trim --- */
            PrintOutString = " *** Data Source: site "
            For Each site In TheExec.sites
                If FuseDefValueFlag Then
                    'fuse default
                    PrintOutString = PrintOutString & "(" & Trim(str(site)) & "): *Default*; "
                    StoredData(site).data = StoredDataInFuse(site).data
                
                Else
                
                    If TheExec.sites(site).SiteVariableValue(SiteVarName) > 0 Then
                        'fuse
                        PrintOutString = PrintOutString & "(" & Trim(str(site)) & "): *Fuse*;    "
                        StoredData(site).data = StoredDataInFuse(site).data
                    Else
                        'trim
                        PrintOutString = PrintOutString & "(" & Trim(str(site)) & "): *Trim*;    "
                        StoredData(site).data = StoredDataTrim(site).data
                    End If
                End If
            Next site
            
            
            
            
            Call rundsp.ConvertToLongAndSerialToParrel(StoredData, DSPWaveDataSize, StoredDataDec)
            
            '/*--- Report the DigCap Data ---*/
            StoreDataSiteVariant = StoredDataDec.Element(0)
            
            TheExec.Flow.TestLimit StoreDataSiteVariant, Tname:=storename, formatStr:="%d", ForceResults:=tlForceFlow, scaletype:=scaleNoScaling
            
            
            
            PrintOutString = PrintOutString & " ( " & storename & " )"
            TheExec.Datalog.WriteComment PrintOutString & Chr(13)
            
            
            
            Call AddStoredCaptureData(storename, StoredData)
            
            'Debug.Print "--------------read store-------------------------------------"
            'StoredData = GetStoredCaptureData(StoreName)
            
            'For Each Site In TheExec.sites
            '    Debug.Print Site, StoreName, StoredData(Site).DataSize
            'Next Site
            
            'Debug.Print "------------------------------------------------------------"
            
            Set StoredData = Nothing
        Next i

    
    End If
    
    
    

End Function




Public Function TMPS(patset As Pattern, CPUA_Flag_In_Pat As Boolean, DigSrc_pin As PinList, DigSrc_DataWidth As Long, DigSrc_Sample_Size As Long, DigSrc_Equation As String, digsrc_assignment As String, _
                           Optional DigCap_Pin As PinList, Optional DigCap_DataWidth As Long, Optional DigCap_Sample_Size As Long, Optional DigCap_DSPWaveSetting As CalculateMethodSetup = 0, _
                            Optional CUS_Str_DigSrcData As String, Optional CUS_Str_DigCapData As String = vbNullString, Optional TMPS_Fuse_string As String, Optional Validating_ As Boolean) As Long

'' Step 1 : trim code is 8 bit, show out measured volt and trimed code, target volt is 0.9v
'' Step 2 : start from 0x8 and add algorithm to decide +/- direction
'' while decimal < 2 ^ DigSrc_Sample_Size
'' convert decimal to binary reverse
'' input the binary reverse data to digSrc_assignment

    If Validating_ Then
        Call PrLoadPattern(patset.value)
        Exit Function    ' Exit after validation
    End If
    Dim x As Long
    Dim InDSPWave As New DSPWave
    Dim OutDspWave As New DSPWave
    Dim CapOut As String
    Dim SrcOut As String
    Dim site As Variant
    Dim Pat As String
    Dim i As Integer
    Dim ShowDec As String
    Dim ShowOut As String
    Dim TrimBits As String
    Dim b_TestDone As Boolean
    Dim code(7) As Integer
    Dim SourceNum As Integer
    Dim k As Integer
    Dim CtrlBits As New DSPWave
    Dim MinValue As New DSPWave
    Dim Data_Array(7) As New SiteLong
    Dim PassFlag_TMPS As New SiteBoolean
    On Error GoTo errHandler

    b_TestDone = False
    SourceNum = 0
    CtrlBits.CreateConstant 0, 2 * DigSrc_Sample_Size / DigSrc_DataWidth
    MinValue.CreateConstant 0, DigSrc_Sample_Size / DigSrc_DataWidth

    If DigSrc_Sample_Size = 0 Then
        TheExec.Datalog.WriteComment ("Error!! - Please check input argument DigSrc_Sample_Size")
        Exit Function
    End If

    TheHdw.Digital.Patgen.Halt
    TheHdw.Digital.ApplyLevelsTiming True, True, True, tlPowered
    TheHdw.Patterns(patset).Load

    Dim PattArray() As String
    Dim PatCount As Long

    Call PATT_GetPatListFromPatternSet(patset.value, PattArray, PatCount)
    Call TheHdw.Digital.Patgen.Continue(0, cpuA + cpuB + cpuC + cpuD)
    TheHdw.Digital.Patgen.HaltMode = tlHaltOnOpcode

    Do While b_TestDone = False
        For Each site In TheExec.sites.Active

          ''  theexec.Datalog.WriteComment ("======== Start Dig Src setup =======")
            If SourceNum = 0 Then
                Call Create_DigSrc_Data(DigSrc_pin, DigSrc_DataWidth, DigSrc_Sample_Size, DigSrc_Equation, digsrc_assignment, InDSPWave, site)
            End If

            '' Processing Captured Code and Printing the Values
            If SourceNum > 0 Then
                For k = 0 To DigCap_Sample_Size / DigCap_DataWidth - 1
                code(k) = 0
                    For i = 0 To DigCap_DataWidth - 1
                        code(k) = code(k) + OutDspWave(site).Element(k * DigCap_DataWidth + i) * (2 ^ i)
                    Next i
                ''   Code(k) = 3770 + Int(Rnd * 20)    '' Random Codes for Offline testing
                If gl_Disable_HIP_debug_log = False Then TheExec.Datalog.WriteComment "Site:" & site & ", Capture data_ " & k & "_Value =" & code(k)
                Next k
            End If

            If SourceNum > 0 And SourceNum < 6 Then
                For k = 0 To DigSrc_Sample_Size / DigSrc_DataWidth - 1

                    If code(k) > 3780 Then
                        For x = 0 To 3
                            InDSPWave(site).Element(k * 28 + 4 * (7 - SourceNum) + x) = 0
                        Next x
                        'InDSPWave(site).Element(k * 28 + 4 * (7 - SourceNum) + 1) = 0
                        'InDSPWave(site).Element(k * 28 + 4 * (7 - SourceNum) + 2) = 0
                        'InDSPWave(site).Element(k * 28 + 4 * (7 - SourceNum) + 3) = 0
                    Else
                        For x = 0 To 3
                            InDSPWave(site).Element(k * 28 + 4 * (7 - SourceNum) + x) = 1
                        Next x
                        'InDSPWave(site).Element(k * 28 + 4 * (7 - SourceNum) + 1) = 1
                        'InDSPWave(site).Element(k * 28 + 4 * (7 - SourceNum) + 2) = 1
                        'InDSPWave(site).Element(k * 28 + 4 * (7 - SourceNum) + 3) = 1
                    End If

                    If SourceNum < 5 Then
                        For x = 0 To 3
                            InDSPWave(site).Element(k * 28 + 4 * (6 - SourceNum) + x) = 1
                        Next x
                        'InDSPWave(site).Element(k * 28 + 4 * (6 - SourceNum) + 1) = 1
                        'InDSPWave(site).Element(k * 28 + 4 * (6 - SourceNum) + 2) = 1
                        'InDSPWave(site).Element(k * 28 + 4 * (6 - SourceNum) + 3) = 1
                    End If
                Next k

            End If

            ''' Trim Control(first 2) Bits

            If SourceNum > 4 And SourceNum < 9 Then
                Select Case SourceNum - 5

                    Case 0
                        For k = 0 To DigSrc_Sample_Size / DigSrc_DataWidth - 1

                            For x = 0 To 7
                                InDSPWave(site).Element(k * 28 + x) = 0
                            Next x
                            'InDSPWave(site).Element(k * 28 + 1) = 0
                            'InDSPWave(site).Element(k * 28 + 2) = 0
                            'InDSPWave(site).Element(k * 28 + 3) = 0
                            'InDSPWave(site).Element(k * 28 + 4) = 0
                            'InDSPWave(site).Element(k * 28 + 5) = 0
                            'InDSPWave(site).Element(k * 28 + 6) = 0
                            'InDSPWave(site).Element(k * 28 + 7) = 0
                        Next k


                    Case 1
                        For k = 0 To DigSrc_Sample_Size / DigSrc_DataWidth - 1

                            MinValue(site).Element(k) = code(k)
                            For i = 0 To 15
                                CtrlBits(site).Element(i) = 0
                            Next i

                            For x = 0 To 3
                                InDSPWave(site).Element(k * 28 + x) = 0
                            Next x
                            'InDSPWave(site).Element(k * 28 + 1) = 0
                            'InDSPWave(site).Element(k * 28 + 2) = 0
                            'InDSPWave(site).Element(k * 28 + 3) = 0
                            For x = 4 To 7
                                InDSPWave(site).Element(k * 28 + x) = 1
                            Next x
                            'InDSPWave(site).Element(k * 28 + 5) = 1
                            'InDSPWave(site).Element(k * 28 + 6) = 1
                            'InDSPWave(site).Element(k * 28 + 7) = 1
                        Next k


                    Case 2
                        For k = 0 To DigSrc_Sample_Size / DigSrc_DataWidth - 1

                            If code(k) < MinValue(site).Element(k) Then
                                MinValue(site).Element(k) = code(k)
                                CtrlBits(site).Element(2 * k) = 0
                                CtrlBits(site).Element(2 * k + 1) = 1
                            End If

                            For x = 0 To 3

                                InDSPWave(site).Element(k * 28 + x) = 1
                            Next
                            'InDSPWave(site).Element(k * 28 + 1) = 1
                            'InDSPWave(site).Element(k * 28 + 2) = 1
                            'InDSPWave(site).Element(k * 28 + 3) = 1
                            For x = 4 To 7
                                InDSPWave(site).Element(k * 28 + x) = 0
                            Next x
                            'InDSPWave(site).Element(k * 28 + 5) = 0
                            'InDSPWave(site).Element(k * 28 + 6) = 0
                            'InDSPWave(site).Element(k * 28 + 7) = 0
                        Next k


                    Case 3
                        For k = 0 To DigSrc_Sample_Size / DigSrc_DataWidth - 1

                            If code(k) < MinValue(site).Element(k) Then
                                MinValue(site).Element(k) = code(k)
                                CtrlBits(site).Element(2 * k) = 1
                                CtrlBits(site).Element(2 * k + 1) = 0
                            End If

                            For x = 0 To 7
                                InDSPWave(site).Element(k * 28 + x) = 1
                            Next x
                            'InDSPWave(site).Element(k * 28 + 1) = 1
                            'InDSPWave(site).Element(k * 28 + 2) = 1
                            'InDSPWave(site).Element(k * 28 + 3) = 1
                            'InDSPWave(site).Element(k * 28 + 4) = 1
                            'InDSPWave(site).Element(k * 28 + 5) = 1
                            'InDSPWave(site).Element(k * 28 + 6) = 1
                            'InDSPWave(site).Element(k * 28 + 7) = 1
                        Next k

                End Select
            End If


            ''''' Sourcing Trimmed Data Bits

            If SourceNum = 9 Then

                For k = 0 To DigSrc_Sample_Size / DigSrc_DataWidth - 1

                    If code(k) < MinValue(site).Element(k) Then
                        MinValue(site).Element(k) = code(k)
                        CtrlBits(site).Element(2 * k) = 1
                        CtrlBits(site).Element(2 * k + 1) = 1
                    End If

                Next k

                For k = 0 To DigSrc_Sample_Size / DigSrc_DataWidth - 1
                    For x = 0 To 3
                        InDSPWave(site).Element(k * 28 + x) = CtrlBits(site).Element(2 * k)
                    Next x

                    'InDSPWave(site).Element(k * 28 + 1) = CtrlBits(site).Element(2 * k)
                    'InDSPWave(site).Element(k * 28 + 2) = CtrlBits(site).Element(2 * k)
                    'InDSPWave(site).Element(k * 28 + 3) = CtrlBits(site).Element(2 * k)
                    For x = 4 To 7
                        InDSPWave(site).Element(k * 28 + x) = CtrlBits(site).Element(2 * k + 1)
                    Next x
                    'InDSPWave(site).Element(k * 28 + 5) = CtrlBits(site).Element(2 * k + 1)
                    'InDSPWave(site).Element(k * 28 + 6) = CtrlBits(site).Element(2 * k + 1)
                    'InDSPWave(site).Element(k * 28 + 7) = CtrlBits(site).Element(2 * k + 1)
                Next k

                For k = 0 To DigSrc_Sample_Size / DigSrc_DataWidth - 1
                    Data_Array(k) = 0
                    For i = 0 To DigSrc_DataWidth / 4 - 1
                        Data_Array(k) = Data_Array(k) + InDSPWave(site).Element(k * DigSrc_DataWidth + 4 * i) * (2 ^ i)
                    Next i
                Next k

            End If



        Next site


        Call SetupDigSrcDspWave(PattArray(0), DigSrc_pin, "Meas_Src", DigSrc_Sample_Size, InDSPWave)

        If SourceNum = 9 Then
            If gl_Disable_HIP_debug_log = False Then TheExec.Datalog.WriteComment "==========SOURCING TRIMMED DATA BITS==============="
        End If

        For Each site In TheExec.sites.Active
            If SourceNum > 0 Then
                SrcOut = vbNullString
                For i = 0 To DigSrc_Sample_Size - 1
                    SrcOut = SrcOut & InDSPWave(site).Element(i)
                    If i Mod DigSrc_DataWidth = DigSrc_DataWidth - 1 Then
                        SrcOut = SrcOut & ", "
                    ElseIf i Mod 4 = 3 Then
                        SrcOut = SrcOut & " "
                    End If

                Next i
                If gl_Disable_HIP_debug_log = False Then TheExec.Datalog.WriteComment "Site:" & site & ", Source Data=" & SrcOut
            End If
        Next site

        If DigCap_Sample_Size <> 0 Then

           TheExec.Datalog.WriteComment ("======== Setup Dig Cap Test Start ========")
           Call DigCapSetup(PattArray(0), DigCap_Pin, "Meas_cap", DigCap_Sample_Size, OutDspWave)

        End If

        SourceNum = SourceNum + 1

        Call TheHdw.Patterns(PattArray(0)).start

        TheHdw.Digital.Patgen.HaltWait ' haltwait at patten end

        PatCount = PatCount + 1
        CapOut = vbNullString


        TheHdw.Digital.Patgen.HaltWait   '' Haltwait at patten end

        If SourceNum = 10 Then
            b_TestDone = True
        End If

    Loop

    For Each site In TheExec.sites.Active
        For k = 0 To DigCap_Sample_Size / DigCap_DataWidth - 1
        code(k) = 0
            For i = 0 To DigCap_DataWidth - 1
                code(k) = code(k) + OutDspWave(site).Element(k * 12 + i) * (2 ^ i)
            Next i
         If CurrentJobName_L Like "*char*" Then
         Disable_Inst_pinname_in_PTR
            TheExec.Flow.TestLimit resultVal:=code(k), unit:=unitNone, ForceResults:=tlForceFlow
         Enable_Inst_pinname_in_PTR
         Else
            TheExec.Flow.TestLimit resultVal:=code(k), unit:=unitNone, Tname:="Capture Data_" & k, ForceResults:=tlForceFlow
            End If
        Next k

'' Add for TMPS fusing 20160526
       If TheHdw.Digital.Patgen.PatternBurstPassed(site) = False Then 'Pattern Fail
           PassFlag_TMPS(site) = False
       Else
           PassFlag_TMPS(site) = True
        End If

        '20210406 Modify for new Efuse
'        If LCase(TMPS_Fuse_string) Like "*t3_fuse_25c*" Then
'            For x = 1 To 7
'                Call auto_eFuse_SetPatTestPass_Flag("ECID", "Temp" & x, PassFlag_TMPS(site))
'            Next x
'            'Call auto_eFuse_SetPatTestPass_Flag("ECID", "Temp2", PassFlag_TMPS(site))
'            'Call auto_eFuse_SetPatTestPass_Flag("ECID", "Temp3", PassFlag_TMPS(site))
'            'Call auto_eFuse_SetPatTestPass_Flag("ECID", "Temp4", PassFlag_TMPS(site))
'            'Call auto_eFuse_SetPatTestPass_Flag("ECID", "Temp5", PassFlag_TMPS(site))
'            'Call auto_eFuse_SetPatTestPass_Flag("ECID", "Temp6", PassFlag_TMPS(site))
'            'Call auto_eFuse_SetPatTestPass_Flag("ECID", "Temp7", PassFlag_TMPS(site))
'            'Call auto_eFuse_SetPatTestPass_Flag("ECID", "Temp8", PassFlag_TMPS(site))
'
'            For x = 0 To 7
'                Call auto_eFuse_SetWriteDecimal("ECID", "Temp" & x + 1, Data_Array(x))
'            Next x
'            'Call auto_eFuse_SetWriteDecimal("ECID", "Temp2", Data_Array(1))
'            'Call auto_eFuse_SetWriteDecimal("ECID", "Temp3", Data_Array(2))
'            'Call auto_eFuse_SetWriteDecimal("ECID", "Temp4", Data_Array(3))
'            'Call auto_eFuse_SetWriteDecimal("ECID", "Temp5", Data_Array(4))
'            'Call auto_eFuse_SetWriteDecimal("ECID", "Temp6", Data_Array(5))
'            'Call auto_eFuse_SetWriteDecimal("ECID", "Temp7", Data_Array(6))
'            'Call auto_eFuse_SetWriteDecimal("ECID", "Temp8", Data_Array(7))
'
'        ElseIf LCase(TMPS_Fuse_string) Like "*t5_fuse_85c*" Then
'
'            Call auto_eFuse_SetPatTestPass_Flag("CFG", "TS_TEMP_REF_CTRL0", True)
'            Call auto_eFuse_SetPatTestPass_Flag("MON", "THERMAL_PARAM_TS_REFERENCE_CTRL", True)
'            Call auto_eFuse_SetPatTestPass_Flag("CFG", "TS_TEMP_REF_CTRL1", True)
'            Call auto_eFuse_SetPatTestPass_Flag("CFG", "TS_TEMP_REF_CTRL2", True)
'            Call auto_eFuse_SetPatTestPass_Flag("CFG", "TS_TEMP_REF_CTRL3", True)
'            Call auto_eFuse_SetPatTestPass_Flag("UDR", "Temp_sensor3_tTRIM", True)
'            Call auto_eFuse_SetPatTestPass_Flag("UDR", "Temp_sensor2_tTRIM", True)
'            Call auto_eFuse_SetPatTestPass_Flag("UDR", "Temp_sensor1_tTRIM", True)
'            Call auto_eFuse_SetPatTestPass_Flag("UDR", "Temp_sensor0_tTRIM", True)
'
'            If PassFlag_TMPS(site) = True Then
'                Call auto_eFuse_SetWriteDecimal("CFG", "TS_TEMP_REF_CTRL0", Data_Array(0))
'                Call auto_eFuse_SetWriteDecimal("MON", "THERMAL_PARAM_TS_REFERENCE_CTRL", Data_Array(0))
'                Call auto_eFuse_SetWriteDecimal("CFG", "TS_TEMP_REF_CTRL1", Data_Array(1))
'                Call auto_eFuse_SetWriteDecimal("CFG", "TS_TEMP_REF_CTRL2", Data_Array(2))
'                Call auto_eFuse_SetWriteDecimal("CFG", "TS_TEMP_REF_CTRL3", Data_Array(3))
'                Call auto_eFuse_SetWriteDecimal("UDR", "Temp_sensor3_tTRIM", Data_Array(4))
'                Call auto_eFuse_SetWriteDecimal("UDR", "Temp_sensor2_tTRIM", Data_Array(5))
'                Call auto_eFuse_SetWriteDecimal("UDR", "Temp_sensor1_tTRIM", Data_Array(6))
'                Call auto_eFuse_SetWriteDecimal("UDR", "Temp_sensor0_tTRIM", Data_Array(7))
'
'            Else  ' if fail then burn the default code
'
'                Call auto_eFuse_SetWriteDecimal("CFG", "TS_TEMP_REF_CTRL0", 64)
'                Call auto_eFuse_SetWriteDecimal("MON", "THERMAL_PARAM_TS_REFERENCE_CTRL", 64)
'                Call auto_eFuse_SetWriteDecimal("CFG", "TS_TEMP_REF_CTRL1", 64)
'                Call auto_eFuse_SetWriteDecimal("CFG", "TS_TEMP_REF_CTRL2", 64)
'                Call auto_eFuse_SetWriteDecimal("CFG", "TS_TEMP_REF_CTRL3", 64)
'                Call auto_eFuse_SetWriteDecimal("UDR", "Temp_sensor3_tTRIM", 64)
'                Call auto_eFuse_SetWriteDecimal("UDR", "Temp_sensor2_tTRIM", 64)
'                Call auto_eFuse_SetWriteDecimal("UDR", "Temp_sensor1_tTRIM", 64)
'                Call auto_eFuse_SetWriteDecimal("UDR", "Temp_sensor0_tTRIM", 64)
'
'            End If
'        End If

    Next site
    
    ''20210406 Modify for newEfuse
    Dim opbank As eFuseBdfBank
    Dim opbank_cfg As eFuseBdfBank, opbank_mon As eFuseBdfBank, opbank_udr As eFuseBdfBank
    Dim Pass_Fail_Flag As New SiteBoolean: Pass_Fail_Flag = True
    Dim Data_Array_Default As New SiteLong: Data_Array_Default = 64
    
    If LCase(TMPS_Fuse_string) Like "*t3_fuse_25c*" Then
        For x = 1 To 7
            Set opbank = GetBdfBank("ECID")
            opbank.SetEfuse "Temp" & x, Data_Array(x), PassFlag_TMPS, , , , True
        Next x
    ElseIf LCase(TMPS_Fuse_string) Like "*t5_fuse_85c*" Then
        Set opbank_cfg = GetBdfBank("CFG")
        Set opbank_mon = GetBdfBank("MON")
        Set opbank_udr = GetBdfBank("UDR")
        
        If PassFlag_TMPS(site) = True Then
            opbank_cfg.SetEfuse "TS_TEMP_REF_CTRL0", Data_Array(0), Pass_Fail_Flag, , , , True
            opbank_mon.SetEfuse "THERMAL_PARAM_TS_REFERENCE_CTRL", Data_Array(0), Pass_Fail_Flag, , , , True
            opbank_cfg.SetEfuse "TS_TEMP_REF_CTRL1", Data_Array(1), Pass_Fail_Flag, , , , True
            opbank_cfg.SetEfuse "TS_TEMP_REF_CTRL2", Data_Array(2), Pass_Fail_Flag, , , , True
            opbank_cfg.SetEfuse "TS_TEMP_REF_CTRL3", Data_Array(3), Pass_Fail_Flag, , , , True
            opbank_udr.SetEfuse "Temp_sensor3_tTRIM", Data_Array(4), Pass_Fail_Flag, , , , True
            opbank_udr.SetEfuse "Temp_sensor2_tTRIM", Data_Array(5), Pass_Fail_Flag, , , , True
            opbank_udr.SetEfuse "Temp_sensor1_tTRIM", Data_Array(6), Pass_Fail_Flag, , , , True
            opbank_udr.SetEfuse "Temp_sensor0_tTRIM", Data_Array(7), Pass_Fail_Flag, , , , True
        Else
            opbank_cfg.SetEfuse "TS_TEMP_REF_CTRL0", Data_Array_Default, Pass_Fail_Flag, , , , True
            opbank_mon.SetEfuse "THERMAL_PARAM_TS_REFERENCE_CTRL", Data_Array(0), Pass_Fail_Flag, , , , True
            opbank_cfg.SetEfuse "TS_TEMP_REF_CTRL1", Data_Array_Default, Pass_Fail_Flag, , , , True
            opbank_cfg.SetEfuse "TS_TEMP_REF_CTRL2", Data_Array_Default, Pass_Fail_Flag, , , , True
            opbank_cfg.SetEfuse "TS_TEMP_REF_CTRL3", Data_Array_Default, Pass_Fail_Flag, , , , True
            opbank_udr.SetEfuse "Temp_sensor3_tTRIM", Data_Array_Default, Pass_Fail_Flag, , , , True
            opbank_udr.SetEfuse "Temp_sensor2_tTRIM", Data_Array_Default, Pass_Fail_Flag, , , , True
            opbank_udr.SetEfuse "Temp_sensor1_tTRIM", Data_Array_Default, Pass_Fail_Flag, , , , True
            opbank_udr.SetEfuse "Temp_sensor0_tTRIM", Data_Array_Default, Pass_Fail_Flag, , , , True
        End If
    End If

    Call HardIP_WriteFuncResult

    Exit Function

errHandler:
    TheExec.Datalog.WriteComment "error in TMPS function"
    If AbortTest Then Exit Function Else Resume Next

End Function







