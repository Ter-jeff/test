Attribute VB_Name = "VBT_LIB_K_ADDRIO_AP"
Option Explicit

Private EfuseDecimalDictionary As New Dictionary
Public Function SaveEfuseDecimal_to_Dictionary(keyname As String, ByRef obj As Variant)
On Error GoTo errHandler 'Add ErrHandler 2023/05/29


    keyname = UCase(Trim(keyname))
    
    If EfuseDecimalDictionary.Exists(keyname) Then
        'Stop 'Please check data will be orverwrite
        'TheExec.Datalog.WriteComment ""
        'TheExec.Datalog.WriteComment " *** Saved data will be overwritten *** " & KeyName
        'TheExec.Datalog.WriteComment ""
        EfuseDecimalDictionary.Remove keyname
        EfuseDecimalDictionary.Add keyname, obj
    Else
        EfuseDecimalDictionary.Add keyname, obj
    End If
    
Exit Function 'Add ErrHandler 2023/05/29
errHandler: 'Add ErrHandler 2023/05/29
    Call Print_Error_Message(Error_Info, "VBT_LIB_K_ADDRIO_AP", "SaveEfuseDecimal_to_Dictionary") 'Add ErrHandler 2023/05/29
    If AbortTest Then Exit Function Else Resume Next 'Add ErrHandler 2023/05/29
End Function
Public Function TrimImpedance_MeasR(CPUA_Flag_In_Pat As Boolean, MeasPin As String, MeasPin_P As String, MeasPin_N As String, ForceVolt As String, ByRef RTN_Imped_Val As PinListData, b_IsDifferential As Boolean, _
                                               b_PD_Mode As Boolean, IRange_mA As Double)
On Error GoTo errHandler 'Add ErrHandler 2023/05/29
'/* 2017-12-04 update by Kaino*/
'Public Function TrimImpedance_MeasR(CPUA_Flag_In_Pat As Boolean, MeasPin As String, MeasPin_P As String, MeasPin_N As String, ForceVolt As String, ByRef RTN_Imped_Val As PinListData, Optional b_IsDifferential As Boolean, _
'                                              Optional b_PD_Mode As Boolean = True, Optional IRange_mA As Double

    If (CPUA_Flag_In_Pat) Then
        Call TheHdw.Digital.Patgen.FlagWait(cpuA, 0)
    Else
        Call TheHdw.Digital.Patgen.HaltWait
    End If
    
    Dim Diff_P_Pin As String, Diff_N_Pin As String
    Dim Diff_P_Pin_total As String, Diff_N_Pin_total As String
    Dim P_Pin_ForceV As Double, N_Pin_ForceV As Double
    Dim SplitForceVolt() As String
    Dim TempStringArray() As String ''20170722 unwant string _P transfer to _N
    Dim i As Integer
    Dim Pin_Ary() As String, Pin_Cnt As Long
    Dim k As Integer
    
    
    'TheExec.DataManager.DecomposePinList Pin, Pin_Ary, Pin_Cnt
    
    If MeasPin <> "" Then
        Call UP1600_PPMU_Measure_R_SingleEnd(MeasPin, ForceVolt, IRange_mA * mA, R_PathWithContact, RTN_Imped_Val, b_PD_Mode)
    Else
        Call UP1600_PPMU_Measure_R_Differential(MeasPin_P, MeasPin_N, ForceVolt, IRange_mA * mA, R_PathWithContact, RTN_Imped_Val)
    End If
    
    If (CPUA_Flag_In_Pat) Then
        Call TheHdw.Digital.Patgen.Continue(0, cpuA)
    Else
        TheHdw.Digital.Patgen.HaltWait
    End If
    
    TheHdw.Digital.Patgen.HaltWait

Exit Function 'Add ErrHandler 2023/05/29
errHandler: 'Add ErrHandler 2023/05/29
    Call Print_Error_Message(Error_Info, "VBT_LIB_K_ADDRIO_AP", "TrimImpedance_MeasR") 'Add ErrHandler 2023/05/29
    If AbortTest Then Exit Function Else Resume Next 'Add ErrHandler 2023/05/29
End Function

Public Function UP1600_PPMU_Measure_R_SingleEnd(MeasurePin As String, ForceVoltStr As String, MeasureCurrRange As Double, Optional RAK_Flag As Enum_RAK = 0, _
                                                                                 Optional ByRef RTN_Imped_Val As PinListData, Optional b_PD_Mode As Boolean = True, Optional IRange_mA As Double) As Long
On Error GoTo errHandler 'Add ErrHandler 2023/05/29

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
        .ForceI pc_Def_PPMU_InitialValue_FI, pc_Def_PPMU_InitialValue_FI_Range
        .Connect
        .Gate = tlOn
        .ForceV ForceVoltVal, MeasureCurrRange
    End With
    TheExec.Datalog.WriteComment " *** set current range *** pins: " & MeasurePin & " = " & MeasureCurrRange & " A"
    TheExec.Datalog.WriteComment " *** set voltage for R *** pins: " & MeasurePin & " = " & ForceVoltVal & " V"

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
            TheExec.Datalog.WriteComment (" *** Site(" & site & "), Pin : " & pin & ", Measure Current = " & MeasureValue.Pins(pin).value(site))
        Next site
    Next pin

    '' Print force condition
    Call Print_Force_Condition("r", MeasureValue)

    Dim PowerVal As Double
    '' Impedance measurement
    If b_PD_Mode Then
        Imped = MeasureValue.Math.Invert.Multiply(ForceVoltVal).Abs
    Else
        PowerVal = TheHdw.DCVS.Pins("VDD_ALLI").Voltage.value
        Imped = MeasureValue.Math.Invert.Multiply(PowerVal - ForceVoltVal).Abs
    End If

    If 0 = 1 And TheExec.TesterMode = testModeOffline Then
        'Call SimulateOutputImped(MeasurePin, Imped)
    Else
        If RAK_Flag = R_PathWithContact Then
            '' Compensate resistance after Kelvin for path resistance considerations
            For Each pin In Imped.Pins
                For Each site In TheExec.sites
                    Imped.Pins.item(pin).value(site) = Imped.Pins.item(pin).value(site) - R_Path_PLD.Pins.item(pin).value(site)
                    'TheExec.Datalog.WriteComment " *** Site(" & Site & "), Pin : " & Pin & ", R-Path = " & R_Path_PLD.Pins.Item(Pin).Value(Site)
                    TheExec.Datalog.WriteComment " *** Site(" & site & "), R-Path : " & pin & " = " & R_Path_PLD.Pins.item(pin).value(site)
                Next site
            Next pin
        End If
    End If

    TheHdw.PPMU.Pins(MeasurePin).Disconnect
    TheHdw.Digital.Pins(MeasurePin).Connect

    RTN_Imped_Val = Imped

Exit Function 'Add ErrHandler 2023/05/29
errHandler: 'Add ErrHandler 2023/05/29
    Call Print_Error_Message(Error_Info, "VBT_LIB_K_ADDRIO_AP", "UP1600_PPMU_Measure_R_SingleEnd") 'Add ErrHandler 2023/05/29
    If AbortTest Then Exit Function Else Resume Next 'Add ErrHandler 2023/05/29
End Function

Public Function UP1600_PPMU_Measure_R_Differential(Measure_P_Pin As String, Measure_N_Pin As String, ForceVoltStr As String, MeasureCurrRange As Double, Optional RAK_Flag As Enum_RAK = 0, _
    Optional ByRef RTN_Imped_Val As PinListData) As Long
On Error GoTo errHandler 'Add ErrHandler 2023/05/29

    Dim MeasureValue As New PinListData
    Dim Imped As New PinListData
    Dim pin  As Variant
    Dim site As Variant
    Dim TempStringArray() As String ''20170722 unwant string _P transfer to _N

    Dim i As Integer
    Dim j As Integer
    
    Dim Pin_Ary() As String
    Dim Pin_Cnt As Long
    
   
    Dim P_ForceVolt As Double
    Dim N_ForceVolt As Double


    
    
    
    Dim P_pin As String
    Dim N_pin As String
    Dim P_pin_P As String
    Dim N_pin_N As String
    Dim P_pin_name_seg() As String
    Dim N_pin_name_seg() As String
    

    
    Dim ForceVoltArray() As String
    
    
    
    ForceVoltArray = Split(ForceVoltStr, ",")
    
    If UBound(ForceVoltArray) >= 1 Then
        P_ForceVolt = ForceVoltArray(0)
        N_ForceVolt = ForceVoltArray(1)
    Else
        Stop
    End If
    
    
    TheHdw.Digital.Pins(Measure_P_Pin & "," & Measure_N_Pin).Disconnect
    TheHdw.Wait 10 * us
    
    '' Initial force I to 0 and force V by your specified
    With TheHdw.PPMU.Pins(Measure_P_Pin)
        .ForceI pc_Def_PPMU_InitialValue_FI, pc_Def_PPMU_InitialValue_FI_Range
        .Connect
        .Gate = tlOn
        .ForceV P_ForceVolt, MeasureCurrRange
    End With
    TheExec.Datalog.WriteComment " *** set current range *** pins:" & Measure_P_Pin & " = " & MeasureCurrRange & " A"
    
    With TheHdw.PPMU.Pins(Measure_N_Pin)
        .ForceI pc_Def_PPMU_InitialValue_FI, pc_Def_PPMU_InitialValue_FI_Range
        .Connect
        .Gate = tlOn
        .ForceV N_ForceVolt, MeasureCurrRange
    End With
    TheExec.Datalog.WriteComment " *** set current range *** pins:" & Measure_N_Pin & " = " & MeasureCurrRange & " A"
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
            TheExec.Datalog.WriteComment (" *** Site(" & site & "), Pin : " & pin & ", Measure Current = " & MeasureValue.Pins(pin).value(site))
        Next site
    Next pin
 
    '' Print force condition
    Call Print_Force_Condition("r", MeasureValue)
 
    '' Impedance measurement
    Imped = MeasureValue.Math.Invert.Multiply(P_ForceVolt - N_ForceVolt).Abs
    
    Dim RAK_Pin_N As String
    
    If 0 = 1 And TheExec.TesterMode = testModeOffline Then
        'Call SimulateOutputImped(Measure_P_Pin, Imped)
    Else
        If RAK_Flag = R_PathWithContact Then
            '' Compensate resistance after Kelvin for path resistance considerations
            For Each pin In Imped.Pins
                TempStringArray = Split(pin, "_") ''20170722 unwant string _P transfer to _N
                For i = 0 To UBound(TempStringArray)
                    If LCase(TempStringArray(i)) Like LCase("P*") And i = UBound(TempStringArray) Then
                        TempStringArray(i) = Replace(TempStringArray(i), "p", "n")
                        RAK_Pin_N = RAK_Pin_N & "_" & TempStringArray(i)
                    ElseIf i = 0 Then
                        RAK_Pin_N = TempStringArray(i)
                    Else
                        RAK_Pin_N = RAK_Pin_N & "_" & TempStringArray(i)
                    End If
                Next i

                '''RAK_Pin_N = Replace(UCase(Pin), "_P", "_N")
                For Each site In TheExec.sites
                    Imped.Pins.item(pin).value(site) = Imped.Pins.item(pin).value(site) - R_Path_PLD.Pins.item(pin).value(site) - R_Path_PLD.Pins.item(RAK_Pin_N).value(site)
                    'TheExec.Datalog.WriteComment " *** Site(" & Site & "), Pin : " & Pin & ", R-Path = " & R_Path_PLD.Pins.Item(Pin).Value(Site)
                    TheExec.Datalog.WriteComment " *** Site(" & site & "), R-Path : " & pin & " = " & R_Path_PLD.Pins.item(pin).value(site) & " ; " & RAK_Pin_N & " = " & R_Path_PLD.Pins.item(RAK_Pin_N).value(site)
                Next site
            Next pin
        End If
    End If
    
    TheHdw.PPMU.Pins(Measure_P_Pin & "," & Measure_N_Pin).Disconnect
    TheHdw.Digital.Pins(Measure_P_Pin & "," & Measure_N_Pin).Connect
    
    RTN_Imped_Val = Imped
    
Exit Function 'Add ErrHandler 2023/05/29
errHandler: 'Add ErrHandler 2023/05/29
    Call Print_Error_Message(Error_Info, "VBT_LIB_K_ADDRIO_AP", "UP1600_PPMU_Measure_R_Differential") 'Add ErrHandler 2023/05/29
    If AbortTest Then Exit Function Else Resume Next 'Add ErrHandler 2023/05/29
End Function




Private Sub ForceConditionToFocreValue(ForceCondition As String, ForceValueStr As String, ForceValueArray() As String)
On Error GoTo errHandler 'Add ErrHandler 2023/05/29
    Dim i As Integer
    ForceValueArray = Split(ForceCondition, ",")
    
    
    Call HIP_Evaluate_ForceVal(ForceValueArray)
    
    
    For i = 0 To UBound(ForceValueArray)
        If i = 0 Then
            ForceValueStr = ForceValueArray(i)
        Else
            ForceValueStr = ForceValueStr & "," & ForceValueArray(i)
        End If
    Next i


Exit Sub 'Add ErrHandler 2023/05/29
errHandler: 'Add ErrHandler 2023/05/29
    Call Print_Error_Message(Error_Info, "VBT_LIB_K_ADDRIO_AP", "ForceConditionToFocreValue") 'Add ErrHandler 2023/05/29
    If AbortTest Then Exit Sub Else Resume Next 'Add ErrHandler 2023/05/29
End Sub


Public Function MeasurePins_to_Differential_Pair(MeasurePins As PinList, MeasurePins_P As String, MeasurePins_N As String) As Integer
On Error GoTo errHandler 'Add ErrHandler 2023/05/29

    Dim i As Integer
    Dim j As Integer
    
    Dim Pin_Ary() As String
    Dim Pin_Cnt As Long
    
    Dim Measure_P_Pin As String
    Dim Measure_N_Pin As String
    Dim P_ForceVolt As Double
    Dim N_ForceVolt As Double

    
    Dim P_pin As String
    Dim N_pin As String
    Dim P_pin_P As String
    Dim N_pin_N As String
    Dim P_pin_name_seg() As String
    Dim N_pin_name_seg() As String

    TheExec.DataManager.DecomposePinList MeasurePins, Pin_Ary, Pin_Cnt
    

    For i = 0 To (Pin_Cnt - 1) Step 2
        
        P_pin = Pin_Ary(i)
        N_pin = Pin_Ary(i + 1)
        
        P_pin_name_seg = Split(P_pin, "_")
        N_pin_name_seg = Split(N_pin, "_")
        
        P_pin_P = P_pin_name_seg(UBound(P_pin_name_seg))
        N_pin_N = N_pin_name_seg(UBound(N_pin_name_seg))
        
        P_pin = left(P_pin, Len(P_pin) - Len(P_pin_P))
        N_pin = left(N_pin, Len(N_pin) - Len(N_pin_N))
        
        
        
        If P_pin = N_pin And UCase(left(P_pin_P, 1)) = "P" And UCase(left(N_pin_N, 1)) = "N" Then
                'check ok
                
                P_pin = Pin_Ary(i)
                N_pin = Pin_Ary(i + 1)
                
                Measure_P_Pin = Measure_P_Pin & P_pin & ","
                Measure_N_Pin = Measure_N_Pin & N_pin & ","
            
                
            
        Else
            MeasurePins_to_Differential_Pair = -1
            Stop
            ' please check Pin group should be
            ' pin_1_P,pin_1_N,pin_2_P,pin_2_N
        End If
    
    Next i
    
    Measure_P_Pin = left(Measure_P_Pin, Len(Measure_P_Pin) - 1)
    Measure_N_Pin = left(Measure_N_Pin, Len(Measure_N_Pin) - 1)
    
    
    MeasurePins_P = Measure_P_Pin
    MeasurePins_N = Measure_N_Pin
    
    MeasurePins_to_Differential_Pair = (Pin_Cnt + 1) / 2
    
Exit Function 'Add ErrHandler 2023/05/29
errHandler: 'Add ErrHandler 2023/05/29
    Call Print_Error_Message(Error_Info, "VBT_LIB_K_ADDRIO_AP", "MeasurePins_to_Differential_Pair") 'Add ErrHandler 2023/05/29
    If AbortTest Then Exit Function Else Resume Next 'Add ErrHandler 2023/05/29
End Function
'Public Function Meas_FreqVoltCurr_Universal_func_K(Optional patset As pattern, Optional TestSequence As String, Optional CPUA_Flag_In_Pat As Boolean, _
'Optional DisableComparePins As PinList, Optional DisableConnectPins As PinList, Optional DisableFRC As Boolean = False, Optional FRCPortName As String, _
'Optional MeasV_PinS As String, _
'Optional MeasF_PinS_SingleEnd As String, Optional MeasF_Interval As String, Optional MeasF_EventSourceWithTerminationMode As EventSourceWithTerminationMode, Optional MeasF_Flag_MeasureThreshold As Boolean = False, Optional MeasF_ThresholdPercentage As Double = 0.5, Optional MeasF_WaitTime As String, _
'Optional MeasI_pinS As String, Optional MeasI_Range As String, Optional MeasI_WaitTime As String, _
'Optional DigCap_Pin As PinList, Optional DigCap_DataWidth As Long, Optional DigCap_Sample_Size As Long, _
'Optional DigSrc_pin As PinList, Optional DigSrc_DataWidth As Long, Optional DigSrc_Sample_Size As Long, Optional DigSrc_Equation As String, Optional DigSrc_Assignment As String, Optional DigSrc_FlowForLoopIntegerName As String = "", _
'Optional SpecialCalcValSetting As CalculateMethodSetup = 0, _
'Optional InstSpecialSetting As InstrumentSpecialSetup = 0, _
'Optional CUS_Str_MainProgram As String = "", Optional CUS_Str_DigCapData As String = "", Optional CUS_Str_DigSrcData As String = "", _
'Optional Flag_SingleLimit As Boolean = False, Optional TestLimitPerPin_VFI As String = "FFF", _
'Optional MeasF_PinS_Differential As String, Optional ForceFunctional_Flag As Boolean = False, _
'Optional MeasF_WalkingStrobe_Flag As Boolean, Optional MeasF_WalkingStrobe_StartV As Double, Optional MeasF_WalkingStrobe_EndV As Double, Optional MeasF_WalkingStrobe_StepVoltage As Double, Optional MeasF_WalkingStrobe_BothVohVolDiffV As Double, Optional MeasF_WalkingStrobe_interval As Double, Optional MeasF_WalkingStrobe_miniFreq As Double, _
'Optional Meas_StoreName As String, Optional Calc_Eqn As String, _
'Optional Interpose_PrePat As String, Optional Interpose_PreMeas As String, Optional Interpose_PostTest As String, Optional CharSetName As String, _
'Optional ForceV_Val As String, Optional ForceI_Val As String, _
'Optional K_PreMeas_Setting As String, Optional K_PostMeas_Setting As String, Optional K_PatHalt_Setting As String, _
'Optional Validating_ As Boolean) As Long
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
'    Dim Ts As Variant, TestOption As Variant, site As Variant
'    Dim TestSeqNum As Integer
'    Dim MeasureV_pin As New PinList, MeasureF_Pin_SingleEnd As New PinList, MeasureI_pin As New PinList
'    Dim MeasureI_Pin_CurrentRange As String
'    Dim TestNum As Long
'    Dim InDSPWave As New DSPWave, OutDspWave As New DSPWave
'    Dim ShowDec As String, ShowOut As String
'    Dim patt As Variant
'    Dim pat As String
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
'''    Dim RTN_InterposeString As String
'
'
'    Call CZ_Style_TName_InstanceInfo_Reg(PatSetName:=patset.Value)   '2017-06-02: add by Kaino for CZ style Tname
'
'    Shmoo_Pattern = patset.Value  '' add by CL 20171101
'
'    '/* ----------------------------------------------------------------------------------------- */
'
'    Dim K_PreMeas_Setting_Sequence() As String
'    Dim K_PostMeas_Setting_Sequence() As String
'    'Dim K_PatHalt_Setting_Sequence() As String
'    'K_PatHalt_Setting_Sequence = Split(K_PatHalt_Setting, ";")
'
'    Dim K_Setting_List() As String
'    Dim K_Setting_index As Integer
'    Dim kk As Integer
'
'    Dim K_Setting_Args() As String
'
'    Dim K_Value_array(0) As String
'
'    Dim K_Pins As String
'    Dim K_Value(0) As String
'
'        Dim WaitTimeSeq() As String
'
'
'    '/* --- Eng Settings for ULPS not ULPS2 */
'    '''K_PreMeas_Setting = "VDD:VDD_ALLI:0.6&DCVS_IRange:VDD_ALLI:0.002&delay:10&DisconnectDigital:ADDRPins;x;x;VDD:VDD_ALLI:0.6&DCVS_IRange:VDD_ALLI:0.0002&delay:50&DisconnectDigital:ADDRPins;x;x;VDD:VDD_ALLI:0.45&DCVS_IRange:VDD_ALLI:0.0002&delay:50&DisconnectDigital:ADDRPins;VDD:VDD_ALLI:0.4&***KEEP VOLTAGE FROM THIS POINT TO END***;x"
'
'    'K_PreMeas_Setting = ""
'    'K_PreMeas_Setting = K_PreMeas_Setting & "VDD:VDD_ALLI:0.6 & delay:20 & DisconnectDigital:ADDRPins;"
'    'K_PreMeas_Setting = K_PreMeas_Setting & "x;"
'    'K_PreMeas_Setting = K_PreMeas_Setting & "x;"
'    '
'    'K_PreMeas_Setting = K_PreMeas_Setting & "VDD:VDD_ALLI:0.6 & delay:10 & DisconnectDigital:ADDRPins;"
'    'K_PreMeas_Setting = K_PreMeas_Setting & "x;"
'    'K_PreMeas_Setting = K_PreMeas_Setting & "x;"
'    '
'    'K_PreMeas_Setting = K_PreMeas_Setting & "VDD:VDD_ALLI:0.45 & delay:180 & DisconnectDigital:ADDRPins;"
'    'K_PreMeas_Setting = K_PreMeas_Setting & "VDD:VDD_ALLI:0.4 & ***KEEP VOLTAGE FROM THIS POINT TO END***;"
'    'K_PreMeas_Setting = K_PreMeas_Setting & "x"
'    '
'
'
''
''    K_PreMeas_Setting = ""
''    K_PreMeas_Setting = K_PreMeas_Setting & "DisconnectDigital:ADDRPins;"
''    K_PreMeas_Setting = K_PreMeas_Setting & "x;"
''    K_PreMeas_Setting = K_PreMeas_Setting & "x;"
''    K_PreMeas_Setting = K_PreMeas_Setting & "DisconnectDigital:ADDRPins;"
''    K_PreMeas_Setting = K_PreMeas_Setting & "x;"
''    K_PreMeas_Setting = K_PreMeas_Setting & "x;"
''    K_PreMeas_Setting = K_PreMeas_Setting & "VDD:VDD_ALLI:0.45 & delay:180 & DisconnectDigital:ADDRPins;"
''    K_PreMeas_Setting = K_PreMeas_Setting & "VDD:VDD_ALLI:0.4 & ***KEEP VOLTAGE FROM THIS POINT TO END***;"
''    K_PreMeas_Setting = K_PreMeas_Setting & "x"
'
'    'MeasI_WaitTime = "0.02+++0.04+++0.120++"
'    'MeasI_Range = "0.0002+++0.00002+++0.00002++"
'    '/* --- ------------ --- */
'
'    'MeasI_WaitTime = "1+++1+++1++"
'    'MeasI_WaitTime = "++++++++"
'    'MeasI_WaitTime = "0.02+++0.04+++0.12++"
'
'    WaitTimeSeq = Split(MeasI_WaitTime, "+")
'
'    K_PreMeas_Setting_Sequence = Split(K_PreMeas_Setting, ";")
'    K_PostMeas_Setting_Sequence = Split(K_PostMeas_Setting, ";")
'
'
'
'
'
'
'    '/* ----------------------------------------------------------------------------------------- */
'
'
'
'    Dim PreMeasForce_Restore_Before_Meas_array() As String
'    Dim PreMeasForce_Restore_After_Meas_array() As String
'    Dim RestoreTestSeqNum As Integer
'
''
''    If PreMeasForce_Restore_Before_Meas <> "" Then
''        PreMeasForce_Restore_Before_Meas_array = Split(PreMeasForce_Restore_Before_Meas, ",")
''    End If
''    If PreMeasForce_Restore_After_Meas <> "" Then
''        PreMeasForce_Restore_After_Meas_array = Split(PreMeasForce_Restore_After_Meas, ",")
''    End If
'
'
'
'
'
'    On Error GoTo errHandler
'    Dim CheckDSPWave As New DSPWave
'
'    Call tl_PinListDataSort(True)
'
'    ''20170322-Store MeasF mid value for VT
'    Dim SplitFreqVtValue() As String
'    Dim DictKey_StoreVT As String
'    Dim Dict_VT_Value As New SiteDouble
'
'
'    If (InStr(MeasI_Range, ":") <> 0) Then MeasI_Range = Select_MeasIRange(MeasI_Range, currentJobName)   ' support different Meter_Range in different Job, add by Roger 20170628
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
'    '' 20170703 - Evaluate value for ForceV_Val_Ary(), ForceI_Val_Ary()
'
'    Call HIP_Evaluate_ForceVal(ForceV_Val_Ary())
'
'    Call HIP_Evaluate_ForceVal(ForceI_Val_Ary())
'
'    '' 20150121 - Range Check
'    If Range_Check_Enable_Word = True Then
'        If TheExec.DataManager.MemberIndex = 0 Then
'            gl_UseLimitCheck_Counter = 0
'        End If
'    End If
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
'    Call TheHdw.Digital.Patgen.Continue(0, cpuA + cpuB + cpuC + cpuD) '20180606 add to prevent issue, reference from Cyprus.
'
'    Dim Loop_count As Long
'    Dim Loop_Init As Long
'    Dim Loop_Max As Long
'    Dim Loop_Step As Long
'    Dim Loop_BitNum As Long
'    Dim Loop_RegName As String
'    Dim SplitLoop_RegName() As String
'    Dim Split_Loop_DigSrc_Str() As String
'    Dim BINstr As String
'    Dim Loop_SplitByComma() As String
'    Dim Loop_SplitByEqual() As String
'
'    Loop_Init = 0
'    Loop_Max = 0
'    Loop_Step = 1
'
'    If CUS_Str_MainProgram <> "" And InStr(UCase(CUS_Str_MainProgram), UCase("Loop_DigSrc")) <> 0 Then
'
'        'Ex: CUS_Str_MainProgram ==> Loop_DigSrc;1;119;10;32;ddr0_mdll0_lsw:ddr0_mdll1_lsw:ddr1_mdll0_lsw:ddr1_mdll1_lsw
'        Split_Loop_DigSrc_Str = Split(CUS_Str_MainProgram, ";")
'        Loop_Init = Split_Loop_DigSrc_Str(1)
'        Loop_Max = Split_Loop_DigSrc_Str(2)
'        Loop_Step = Split_Loop_DigSrc_Str(3)
'        Loop_BitNum = Split_Loop_DigSrc_Str(4)
'        Loop_RegName = Split_Loop_DigSrc_Str(5)
'        SplitLoop_RegName = Split(Loop_RegName, ":")
'
'    End If
'
'    Dim loop_i As Long, Loop_j As Long
'    Dim Temp_Equal_Str As String
'    Dim Final_Comma_Str As String
'    Temp_Equal_Str = ""
'    Final_Comma_Str = ""
'
'    For Loop_count = Loop_Init To Loop_Max
'        If CUS_Str_MainProgram <> "" And InStr(UCase(CUS_Str_MainProgram), UCase("Loop_DigSrc")) <> 0 Then
'            BINstr = Dec2BinStr32Bit_Rev(Loop_BitNum, Loop_count)
'            Loop_SplitByComma = Split(DigSrc_Assignment, ";")
'
'            For loop_i = 0 To UBound(Loop_SplitByComma)
'                Loop_SplitByEqual = Split(Loop_SplitByComma(loop_i), "=")
'                For Loop_j = 0 To UBound(SplitLoop_RegName)
'                    If UCase(Loop_SplitByEqual(0)) = UCase(SplitLoop_RegName(Loop_j)) Then
'                        Loop_SplitByEqual(1) = BINstr
'                        Temp_Equal_Str = Loop_SplitByEqual(0) & "=" & Loop_SplitByEqual(1)
'                        Exit For
'                    Else
'                        Temp_Equal_Str = Loop_SplitByEqual(0) & "=" & Loop_SplitByEqual(1)
'                    End If
'                Next Loop_j
'                If loop_i = 0 Then
'                    Final_Comma_Str = Temp_Equal_Str
'                Else
'                    Final_Comma_Str = Final_Comma_Str & ";" & Temp_Equal_Str
'                End If
'            Next loop_i
'            DigSrc_Assignment = Final_Comma_Str
'        End If
'
'        '' 20160923 - Add Interpose_PrePat entry point
'        If Interpose_PrePat <> "" Then
'            Call SetForceCondition(Interpose_PrePat & ";STOREPREPAT")
'        End If
'
'        Call HardIP_InitialSetupForPatgen
'
'        ''20161205 - Force_Flow_Shmoo_Condition
'        If TheExec.sites.Item(site).SiteVariableValue("Flow_Shmoo_DevCharSetup") <> "" Then Force_Flow_Shmoo_Condition 'Do Flow Shmoo
'
'        If patset.Value <> "" Then
'            TheHdw.Patterns(patset).Load
'            Call PATT_GetPatListFromPatternSet(patset.Value, PattArray, PatCount)
'        Else
'            ReDim PattArray(0)
'            PattArray(0) = ""
'        End If
'
'        If (DisableConnectPins <> "") Then TheHdw.Digital.Pins(DisableConnectPins).Disconnect
'        If (DisableComparePins <> "") Then TheHdw.Digital.Pins(DisableComparePins).DisableCompare = True
'
'        ''20161107-Return sweep test name
'        Dim Rtn_SweepTestName As String
'        Rtn_SweepTestName = ""
'
'        For Each patt In PattArray
'            If patt <> "" Then
'                pat = CStr(patt)
'                TheHdw.Patterns(pat).Load
'
'                Call GeneralDigSrcSetting(pat, DigSrc_pin, DigSrc_Sample_Size, DigSrc_DataWidth, DigSrc_Equation, DigSrc_Assignment, _
'                                                       DigSrc_FlowForLoopIntegerName, CUS_Str_DigSrcData, InDSPWave, Rtn_SweepTestName)
'
'                If TPModeAsCharz_GLB = True Then
'                    If Rtn_SweepTestName <> "" Then
'        ''                Rtn_SweepTestName = "_" & Rtn_SweepTestName
'                        For i = 0 To UBound(FlowTestNme)
'                            FlowTestNme(i) = Replace(LCase(FlowTestNme(i)), "sweepcode", Rtn_SweepTestName)
'                        Next i
'                    Else
'                        Call SimulateFlowForSweep(FlowShmooString_GLB)
'                        If FlowShmooString_GLB <> "" Then
'                            For i = 0 To UBound(FlowTestNme)
'                                FlowTestNme(i) = Replace(LCase(FlowTestNme(i)), "sweepvoltage", FlowShmooString_GLB)
'                            Next i
'                        End If
'                    End If
'                End If
'
'                Set OutDspWave = Nothing
'                Call GeneralDigCapSetting(pat, DigCap_Pin, DigCap_Sample_Size, OutDspWave)
'
'                Call PrintDigCapSetting(DigCap_Pin, DigCap_Sample_Size, CUS_Str_DigCapData)
'
'                If CUS_Str_MainProgram <> "" Then
'                    If InStr(UCase(CUS_Str_MainProgram), "MTR_UVI80_SETUP") <> 0 Then
'                        Call MTR_UVI80_Setup
'                    End If
'                End If
'
'
'            '' 20170822 copy from Cyprus team for PCMRING
'
'            Dim SplitByCommaStr() As String
'            Dim ForcePin_X As String
'            Dim ForcePin_Y As String
'            Dim SweepIndexStr_X As String
'            Dim ForceVal_X As Double
'            If LCase(CUS_Str_MainProgram) Like "*x_sweep*" Then
'                SplitByCommaStr = Split(CUS_Str_DigSrcData, ",")
'                SweepIndexStr_X = SplitByCommaStr(0)
'                     ForcePin_X = SplitByCommaStr(1)
'
'                      ForceVal_X = CDbl(Val(TheExec.Flow.var(SweepIndexStr_X).Value)) / 1000
'                      TheExec.Datalog.WriteComment "ForcePin = " & ForcePin_X & "; ForceVal_X = " & ForceVal_X & "V"
'                      'TheExec.Datalog.WriteComment "ForcePin = " & SplitByCommaStr(2) & ";  ForceVal_X  = " & ForceVal_X & "V"
'                      TheHdw.DCVS.Pins(ForcePin_X).Voltage.Value = ForceVal_X
'                      'TheHdw.DCVS.Pins(SplitByCommaStr(2)).Voltage.Value = ForceVal_X
'                    FourceV = ForceVal_X
'             End If
'
'
'
'                '' 20160713 - If no cpuflags in the test item modify the code to run pattern by using .test
'                If (CPUA_Flag_In_Pat) Then
'                    Call TheHdw.Patterns(pat).start
'                Else
'                    Call TheHdw.Patterns(pat).test(pfAlways, 0)
'                End If
'            End If
'            TestSeqNum = 0
'
'            For Each Ts In TestSequenceArray
'
'                ''20150907 - Only need CPUA_Flag_In_Pat to do control
'                If (CPUA_Flag_In_Pat) Then
'                    Call TheHdw.Digital.Patgen.FlagWait(cpuA, 0) 'Meas during CPUA loop
'                Else
'                    Call TheHdw.Digital.Patgen.HaltWait 'Pattern without CPUA loop
'                End If
'
'
'                '/ *** --------------------------------------------------------------------------------------- ***
'                '/ ***    Before Meas
'                '/ *** --------------------------------------------------------------------------------------- ***
'
'                If TestSeqNum <= UBound(K_PreMeas_Setting_Sequence) Then
'                    TheExec.Datalog.WriteComment ""
'                    TheExec.Datalog.WriteComment " *** Setting: Before Measurement ***"
'                    Call K_Set_SpecialSettingSequence(K_PreMeas_Setting_Sequence(TestSeqNum))
'                End If
'                '/ *** --------------------------------------------------------------------------------------- *** /
'
'
'
'
'
'
'                ''20160923 - Add Interpose_PreMeas entry point by each sequence
'                If Interpose_PreMeas <> "" Then
'                    If UBound(Interpose_PreMeas_Ary) = 0 Then
'                        Call SetForceCondition(Interpose_PreMeas_Ary(0) & ";STOREPREMEAS")
'                    ElseIf Interpose_PreMeas_Ary(TestSeqNum) <> "" Then
'                        Call SetForceCondition(Interpose_PreMeas_Ary(TestSeqNum) & ";STOREPREMEAS")
'                    End If
'                End If
'
'                TestOptLen = Len(Ts)
'
'                For k = 1 To TestOptLen
'
'                    TestOption = Mid(Ts, k, 1)
'
'                    '' 20160106 - If "ForceFunctional_Flag" = True to let TestOption = "N" to make the test instance only run functional test
'                    If ForceFunctional_Flag = True Then
'                        TestOption = "N"
'                    End If
'
'                    '' 20170523 - Add force I value for UVI80
'                    If TestOption = "V" Then
'                        Call Decide_Measure_Pin(TestSeqNum, MeasPinAry_V, MeasureV_pin)
'                        Call Decide_ForceVal(TestSeqNum, ForceI_Val_Ary, MeasurePin_ForceI_Val)
'                    End If
'                    If TestOption = "F" And MeasF_PinS_SingleEnd <> "" Then Call Decide_Measure_Pin(TestSeqNum, MeasPinAry_F, MeasureF_Pin_SingleEnd)
'                    If TestOption = "F" And MeasF_PinS_Differential <> "" Then Call Decide_Measure_Pin(TestSeqNum, MeasPinAry_F_Differential, MeasureF_Pin_Differential)
'                    If TestOption = "I" Then
'                        Call Decide_Measure_Pin(TestSeqNum, MeasPinAry_I, MeasureI_pin)
'                        Call Decide_MeasureI_CurrentRange(TestSeqNum, MeasPinAry_IRange, MeasureI_Pin_CurrentRange)
'                        Call Decide_ForceVal(TestSeqNum, ForceV_Val_Ary, MeasurePin_ForceV_Val)
'                    End If
'
'                    For Each site In TheExec.sites.Active
'                        TestNum = TheExec.sites.Item(site).TestNumber
'                    Next site
'
'                    Select Case UCase(TestOption)
'
'                        Case "V"
'
'                            Call DiscriminateMeasureV_PinType(MeasureV_pin, MeasureV_Pin_PPMU, MeasureV_Pin_UVI80)
'    ''                        Call start_profile_DCVI(MeasureV_Pin, 0.01, 1000000, 1024, "capture_signal")
'
'                            Call HardIP_MeasureVolt_old(MeasureV_pin, TestLimitPerPin_VFI, TestSeqNum, k, pat, Flag_SingleLimit, HighLimitVal(0), LowLimitVal(0), FlowTestNme, InstSpecialSetting, CUS_Str_MainProgram, SpecialCalcValSetting, Rtn_MeasVolt, Rtn_SweepTestName, _
'                                                                 MeasurePin_ForceI_Val)
'    ''                        Call Plot_profile_DCVI(MeasureV_Pin, "capture_signal")
'
'                            ''20160906 - Check store measurement or not
'                            If Meas_StoreName <> "" Then
'                                If MeasStoreName_Ary(TestSeqNum) <> "" Then
'                                    Store_Rtn_Meas(StoreIndex) = Rtn_MeasVolt
'                                    Call StoreDataAllType(MeasStoreName_Ary(TestSeqNum), Store_Rtn_Meas(StoreIndex))
'                                    StoreIndex = StoreIndex + 1
'                                End If
'                            End If
'
'                        Case "F"
'
'                            Split_F_Str = Split(CUS_Str_MainProgram, ":")
'                            If CUS_Str_MainProgram <> "" And InStr(UCase(CUS_Str_MainProgram), "F_WAITTIME") <> 0 Then
'                                TheHdw.Wait CDbl(Split_F_Str(1))
'                            End If
'
'                            ''20170809 connect to ANALOGMUX_OUT_IO to MeasF
'                            If UCase(CUS_Str_MainProgram) = "MEASF_ANALOGMUX_OUT_IO" Then TheHdw.Utility.Pins("K7A").State = tlUtilBitOff
'
'                            If MeasF_Interval <> "" Then
'                                d_MeasF_Interval = CDbl(MeasF_Interval)
'                            Else
'                                d_MeasF_Interval = pc_Def_VFI_FreqInterval
'                            End If
'
'                            If MeasureF_Pin_SingleEnd <> "" Then
'
'                                If MeasF_Flag_MeasureThreshold = True Then
'                                    If MeasF_ThresholdPercentage = 0 Then MeasF_ThresholdPercentage = pc_Def_VFI_FreqThresholdPercentage
'                                    Call Freq_PPMU_Meas_VOH(MeasureF_Pin_SingleEnd, MeasF_ThresholdPercentage, MeasF_EnableVtMode, MeasF_EventSource)
'                                End If
'
'                                If MeasF_EnableVtMode = True Then
'                                    TheHdw.Digital.Pins(MeasureF_Pin_SingleEnd).Levels.DriverMode = tlDriverModeVt
'                                End If
'
'                                '' 20160228 HardIP Measure Freq with working strobe by JT
'                                If MeasF_WalkingStrobe_Flag = True Then
'                                    If CUS_Str_DigSrcData <> "" Then
'                                        SplitFreqVtValue = Split(CUS_Str_DigSrcData, ":")
'                                        If UCase(SplitFreqVtValue(0)) = "STORE_VT" Then
'                                            DictKey_StoreVT = SplitFreqVtValue(1)
'                                        End If
'                                    End If
'
'                                    Call Freq_WalkingStrobe_Meas_VOHVOL(MeasureF_Pin_SingleEnd, MeasF_WalkingStrobe_StartV, MeasF_WalkingStrobe_EndV, MeasF_WalkingStrobe_StepVoltage, MeasF_WalkingStrobe_BothVohVolDiffV, MeasF_WalkingStrobe_interval, MeasF_WalkingStrobe_miniFreq, DictKey_StoreVT)
'                                End If
'
'                                '' 20151113 - Modify get instrument type for actural waveform plot
'                                FreqPinsCheckType() = Split(MeasureF_Pin_SingleEnd, ",")
'                                ThisPinType = GetInstrument(CStr(FreqPinsCheckType(0)), 0)
'                                If ThisPinType = "HSD-U" Then
'                                    Call HardIP_FrequencyMeasure_old(MeasureF_Pin_SingleEnd, False, TestLimitPerPin_VFI, LowLimitVal(0), HighLimitVal(0), TestSeqNum, pat, Flag_SingleLimit, d_MeasF_Interval, FlowTestNme, MeasF_WaitTime, MeasF_EventSource, SpecialCalcValSetting, Rtn_MeasFreq, Rtn_SweepTestName, CUS_Str_MainProgram)
'                                Else
'                                    Call HardIP_FrequencyMeasure_Dctime_old(MeasureF_Pin_SingleEnd, TestLimitPerPin_VFI, LowLimitVal(0), HighLimitVal(0), TestSeqNum, pat, Flag_SingleLimit, d_MeasF_Interval, MeasF_WaitTime)
'                                End If
'
'                             '' 20151019 - Enable differential frequency counter
'                            ElseIf MeasureF_Pin_Differential <> "" Then
'
'                                If CUS_Str_DigSrcData <> "" Then
'                                    SplitFreqVtValue = Split(CUS_Str_DigSrcData, ":")
'                                    If UCase(SplitFreqVtValue(0)) = "SETUP_STORE_VT" Then
'                                        DictKey_StoreVT = SplitFreqVtValue(1)
'                                        Dict_VT_Value = GetStoreDataAllType(DictKey_StoreVT)
'
'                                        For Each site In TheExec.sites
'                                            TheHdw.Digital.Pins(MeasureF_Pin_Differential).DifferentialLevels.Value(chDiff_Vt) = Dict_VT_Value(site)
'                                            'TheExec.Datalog.WriteComment ("Site= " & Site & " Set " & MeasureF_Pin_Differential & " Diff_Vt = " & Dict_VT_Value(Site))
'                                        Next site
'                                    End If
'                                End If
'
'                                Call HardIP_FrequencyMeasure_old(MeasureF_Pin_Differential, True, TestLimitPerPin_VFI, LowLimitVal(0), HighLimitVal(0), TestSeqNum, pat, Flag_SingleLimit, d_MeasF_Interval, FlowTestNme, MeasF_WaitTime, MeasF_EventSource, , Rtn_MeasFreq, Rtn_SweepTestName, CUS_Str_MainProgram)
'                            End If
'
'                            ''20160906 - Check store measurement or not
'                            If Meas_StoreName <> "" Then
'                                If MeasStoreName_Ary(TestSeqNum) <> "" Then
'                                    Store_Rtn_Meas(StoreIndex) = Rtn_MeasFreq
'                                    Call StoreDataAllType(MeasStoreName_Ary(TestSeqNum), Store_Rtn_Meas(StoreIndex))
'                                    StoreIndex = StoreIndex + 1
'                                End If
'                            End If
'
'                            ''20170809 restore relay setting
'                            If UCase(CUS_Str_MainProgram) = "MEASF_ANALOGMUX_OUT_IO" Then TheHdw.Utility.Pins("K7A").State = tlUtilBitOn
'
'                        Case "I"
'
'                            If DisableFRC = True Then FreeRunClk_Disable (FRCPortName)
'
''                            '20170816 chyehq : disable FRC XO0 for LAPLL IDS
''                            If UCase(CUS_Str_MainProgram) = "XO0_DISABLE_FRC" And TestSeqNum = 4 Then
''                                TheExec.Datalog.WriteComment "=======Disable FRC XO0======="
''                                TheHdw.Digital.Pins("XO0").Disconnect
''                                With TheHdw.PPMU.Pins("XO0")
''                                    .Disconnect
''                                    .ForceV 0, 0.002
''                                    .Connect
''                                    .Gate = tlOn
''                                End With
''                                TheHdw.Utility.Pins("K1").state = tlUtilBitOn
''                            End If
'
'                            MeasI_WaitTime = WaitTimeSeq(TestSeqNum)
'                            ''20190603
'                            Call HardIP_MeasureCurrent_old(MeasureI_pin, LowLimitVal(0), HighLimitVal(0), MeasureI_Pin_CurrentRange, Flag_SingleLimit, TestSeqNum, pat, TestLimitPerPin_VFI, FlowTestNme, MeasI_WaitTime, SpecialCalcValSetting, Rtn_MeasCurr, Rtn_SweepTestName, MeasurePin_ForceV_Val)
'
'                             '20170816 chyehq : enable FRC XO0
''                            If UCase(CUS_Str_MainProgram) = "XO0_DISABLE_FRC" And TestSeqNum = 4 Then
''                                TheExec.Datalog.WriteComment "=======Enable FRC XO0======="
''                                TheHdw.PPMU.Pins("XO0").Gate = tlOff
''                                TheHdw.PPMU.Pins("XO0").Disconnect
''                                TheHdw.Digital.Pins("XO0").Connect
''                                TheHdw.Utility.Pins("K1").state = tlUtilBitOff
''                            End If
'
'                            ''20160906 - Check store measurement or not
'                            If Meas_StoreName <> "" Then
'                                If MeasStoreName_Ary(TestSeqNum) <> "" Then
'                                    Store_Rtn_Meas(StoreIndex) = Rtn_MeasCurr
'                                    Call StoreDataAllType(MeasStoreName_Ary(TestSeqNum), Store_Rtn_Meas(StoreIndex))
'                                    StoreIndex = StoreIndex + 1
'                                End If
'                            End If
'
'                        Case "N"
'
'                        Case Else
'
'                            TheExec.Datalog.WriteComment "Error Test Option, please select V,I or F"
'
'                    End Select
'
'                    If TheExec.sites.Active.Count = 0 Then Exit Function
'                Next k
'
'                ''20161206-Restore force condiction after measurement
'    ''            Call SetForceCondition("RESTORE")
'
'                '/ *** --------------------------------------------------------------------------------------- ***
'                '/ ***    After_Meas
'                '/ *** --------------------------------------------------------------------------------------- ***
'
'                If TestSeqNum <= UBound(K_PostMeas_Setting_Sequence) Then
'                    TheExec.Datalog.WriteComment " *** Setting: Afer Measurement ***"
'
'                    Call K_Set_SpecialSettingSequence(K_PostMeas_Setting_Sequence(TestSeqNum))
'
'                End If
'
'                '/ *** --------------------------------------------------------------------------------------- *** /
'
'                TestSeqNum = TestSeqNum + 1
'
'                If (CPUA_Flag_In_Pat) Then Call TheHdw.Digital.Patgen.Continue(0, cpuA)
'
'            Next Ts
'
'            If DebugPrintEnable = True Then TheExec.Datalog.WriteComment "  Pattern(" & PatCount & "): " & pat & ""
'
'            TheHdw.Digital.Patgen.HaltWait ' Haltwait at patten end
'
'
'            '/ *** --------------------------------------------------------------------------------------- ***
'            '/ ***    Pattern Halt
'            '/ *** --------------------------------------------------------------------------------------- ***
'
'
'            If K_PatHalt_Setting <> "" Then
'                TheExec.Datalog.WriteComment " *** Setting: Pattern Halt ***"
'                Call K_Set_SpecialSettingSequence(K_PatHalt_Setting)
'            End If
'            '/ *** --------------------------------------------------------------------------------------- *** /
'
'
'            PatCount = PatCount + 1
'
'            '' 20160923 - Add Interpose_PostTest entry point
'            Call SetForceCondition(Interpose_PostTest)
'
'            '' 20160211 - Process DigCapData by using DSP
'    ''        If b_ProcessDigCapByDSP = True Then
'                If DigCap_Sample_Size <> 0 Then
'                    Dim DigCapPinAry() As String, NumberPins As Long
'                    Call TheExec.DataManager.DecomposePinList(DigCap_Pin, DigCapPinAry(), NumberPins)
'
'                    If NumberPins > 1 Then
'                        Call CreateSimulateDataDSPWave_Parallel(OutDspWave, DigCap_Sample_Size)
'                        Call Checker_StoreDigCapAllToDictionary(CUS_Str_DigCapData, OutDspWave, NumberPins)
'                        Call DigCapDataProcessByDSP_Parallel(CUS_Str_DigCapData, OutDspWave, DigCap_Sample_Size, NumberPins)
'
'                    ElseIf NumberPins = 1 Then
'                        Call CreateSimulateDataDSPWave(OutDspWave, DigCap_Sample_Size, DigCap_DataWidth)
'                        Call Checker_StoreDigCapAllToDictionary(CUS_Str_DigCapData, OutDspWave, NumberPins)
'                        Call DigCapDataProcessByDSP(CUS_Str_DigCapData, OutDspWave, DigCap_Sample_Size, DigCap_DataWidth, CUS_Str_MainProgram)
'                    End If
'                End If
'    ''        End If
'        Next patt
'
'''        ''20170405-Record all functional test result from flow for loop opcode, use global string to store them
'        If CUS_Str_DigSrcData <> "" And UCase(CUS_Str_DigSrcData) = UCase("BinToGray") Then
'            If CPUA_Flag_In_Pat = False Then
'                Call DisplayForLoopFuncResult_EndOfTest(CUS_Str_DigSrcData, Rtn_SweepTestName, CPUA_Flag_In_Pat, DigSrc_FlowForLoopIntegerName)
'            End If
'        End If
'     If MeasureV_pin <> "" Then
'         Call EndSetupForMeasureVoltPins(MeasureV_Pin_PPMU, MeasureV_Pin_UVI80)
'     End If
'
'     If DisableConnectPins <> "" Then TheHdw.Digital.Pins(DisableConnectPins).Connect
'     If DisableComparePins <> "" Then TheHdw.Digital.Pins(DisableComparePins).DisableCompare = False
'
'     If DisableFRC = True Then
'         Call ReStart_FRC(FRCPortName)
'     End If
'
'
'     '' 20160907 - Process calculate equation by dictionary.
'     If Calc_Eqn <> "" Then
'         Call ProcessCalcEquation(Calc_Eqn) ''20190604Error
'     End If
'
'     '' 20160713 - Call write functional result if cpu flag in pattern
'     If (CPUA_Flag_In_Pat) Then
'         Call HardIP_WriteFuncResult
'     End If
'
'     DebugPrintFunc patset.Value  ' print all debug information
'
'     If TheExec.sites.Item(site).SiteVariableValue("Flow_Shmoo_DevCharSetup") <> "" Then  'Do Flow Shmoo
'         If Flow_Shmoo_Port_Name <> "" Then Restart_All_Freerun_Clk
'     End If
'
'     If Interpose_PrePat <> "" Then
'         Call SetForceCondition("RESTOREPREPAT")
'     End If
'
'     ''=============================== CharSetName ====================================
'
'     If TheExec.DevChar.Setups.IsRunning = False And CharSetName <> "" Then
'         Dim p As Variant, p_ary() As String, p_cnt As Long, ApplyPins As String, Setup_mode As String
'         'If TheExec.DevChar.Setups(CharSetName).TestMethod.Value = tlDevCharTestMethod_Reburst Then TheExec.Datalog.WriteComment "[PrintCharCondition:" & PrintCharSetup(Interpose_PrePat_GLB) & ",Test]"
'         Setup_mode = TheExec.DevChar.Setups(CharSetName).Shmoo.axes(tlDevCharShmooAxis_X).Parameter.Name
'         If (LCase(Setup_mode) <> "vid" And LCase(Setup_mode) <> "vicm") Then
'             ApplyPins = TheExec.DevChar.Setups(CharSetName).Shmoo.axes(tlDevCharShmooAxis_X).ApplyTo.Pins
'             TheExec.DataManager.DecomposePinList ApplyPins, p_ary, p_cnt
'             For Each p In p_ary
'                 TheExec.DevChar.Setups(CharSetName).Shmoo.axes(tlDevCharShmooAxis_X).ApplyTo.Pins = p
'                 run_shmoo CharSetName
'             Next p
'             TheExec.DevChar.Setups(CharSetName).Shmoo.axes(tlDevCharShmooAxis_X).ApplyTo.Pins = ApplyPins
'         Else
'             run_shmoo CharSetName
'         End If
'     End If
'
'    If CUS_Str_MainProgram <> "" And InStr(UCase(CUS_Str_MainProgram), UCase("Loop_DigSrc")) <> 0 And Loop_Step <> 1 Then
'        Loop_count = Loop_count + Loop_Step - 1
'    End If
'
'    Next Loop_count
'    ''================================================================================
'
'    Call CZ_Style_TName_InstanceInfo_Clear   '2017-06-02: add by Kaino for CZ style Tname
'    Exit Function
'
'errHandler:
'    TheExec.Datalog.WriteComment "error in Meas_FreqVoltCurr_Universal_func"
'    If AbortTest Then Exit Function Else Resume Next
'
'End Function






Public Sub K_Set_SpecialSettingSequence(Setting_Sequence As String)
On Error GoTo errHandler 'Add ErrHandler 2023/05/29
        
Dim K_Setting_index As Integer
    Dim K_Setting_List() As String
    Dim K_Setting_Args() As String
    
    Dim K_Pins As String
    
    Dim K_Value(0) As String
    Dim MsgHeader As String
               
    MsgHeader = "    ***"
                
    K_Setting_List = Split(Setting_Sequence, "&")
                    
    For K_Setting_index = 0 To UBound(K_Setting_List)
                    
        K_Setting_Args = Split(Trim(K_Setting_List(K_Setting_index)), ":")
        
        
    
        Select Case UCase(Trim(K_Setting_Args(0)))
        
            Case UCase("DisconnectDigital")
                K_Pins = K_Setting_Args(1)
                
                TheHdw.Pins(K_Pins).Digital.Disconnect
                TheExec.Datalog.WriteComment MsgHeader & " Digital.Disconnect : " & K_Pins
                            
            Case UCase("ConnectDigital")
                K_Pins = K_Setting_Args(1)
                
                TheHdw.Pins(K_Pins).Digital.Connect
                TheExec.Datalog.WriteComment MsgHeader & " Digital.Connect : " & K_Pins
    
             Case UCase("VDD")
                K_Pins = K_Setting_Args(1)
                K_Value(0) = K_Setting_Args(2)
                
                Call HIP_Evaluate_ForceVal(K_Value())
                TheHdw.DCVS.Pins(K_Pins).Voltage.Main = CDbl(K_Value(0))
                TheExec.Datalog.WriteComment MsgHeader & " Voltage.Main : " & K_Pins & ":" & Format(TheHdw.DCVS.Pins(K_Pins).Voltage.Main, "0.000")
                
            Case UCase("DCVS_IRange")
                K_Pins = K_Setting_Args(1)
                K_Value(0) = K_Setting_Args(2)
                
                With TheHdw.DCVS.Pins(K_Pins)
                    .Meter.mode = tlDCVSMeterCurrent
                    .SetCurrentRanges CDbl(K_Value(0)), CDbl(K_Value(0))
                    .CurrentLimit.Source.FoldLimit.level.value = CDbl(K_Value(0))
                End With
                
                TheExec.Datalog.WriteComment MsgHeader & " DCVS.SetCurrentRanges : " & K_Pins & ":" & Format(K_Value(0) * 1000, "0.000") & "mA"
            
            Case UCase("Delay")
                TheExec.Datalog.WriteComment MsgHeader & " Delay : " & K_Setting_Args(1) & "ms"
                TheHdw.Wait CDbl(K_Setting_Args(1)) * ms

            Case "X"
                ' do nothing
                'TheExec.Datalog.WriteComment MsgHeader
            Case Else
            
                TheExec.Datalog.WriteComment MsgHeader & " Message :" & K_Setting_List(K_Setting_index)
            
            
        End Select
    
    Next K_Setting_index

Exit Sub 'Add ErrHandler 2023/05/29
errHandler: 'Add ErrHandler 2023/05/29
    Call Print_Error_Message(Error_Info, "VBT_LIB_K_ADDRIO_AP", "K_Set_SpecialSettingSequence") 'Add ErrHandler 2023/05/29
    If AbortTest Then Exit Sub Else Resume Next 'Add ErrHandler 2023/05/29
End Sub


Public Function K_TrimImpedance_Turks_TTR(Optional patset As Pattern, Optional CPUA_Flag_In_Pat As Boolean, _
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
    Optional TTR_TrimStart As Long, _
    Optional TTR_HiVal As Double, _
    Optional TTR_LoVal As Double, _
    Optional TTR_Trend As Integer, _
    Optional SubBlockName As String, _
    Optional Validating_ As Boolean) As Long
On Error GoTo errHandler 'Add ErrHandler 2023/05/29


    Dim Pin_Ary() As String
    Dim Pin_Cnt As Long
    Dim TrimCode As Long
    Dim site As Variant
    Dim pin As Variant

    Dim TrimCodeMax As Long
    Dim TrimCodeMin As Long
    Dim i As Integer
    
    Dim TrimSeq As Long
    Dim Site_TrimCode As New SiteLong
    Dim TrimFinishForEachSite As New SiteBoolean
    Dim TrimState As Integer            ' -1 : 0 : 1
    Dim OneStep As Integer
    Dim MeasureValue_Minimum As New SiteVariant
    Dim MeasureValue_Maximum As New SiteVariant
    Dim MinimumValue_LessThan_Traget As New SiteBoolean
    Dim MaximumValue_GreaterThan_Traget As New SiteBoolean
    
    
    Dim TrimDSPwave As New DSPWave
    Dim InDSPWave As New DSPWave
    Dim MeasureValue As New PinListData
    Dim ColseToTargetValue As New PinListData
    Dim CloseToTargetTrimCode As New PinListData
    Dim CompareResult_ClosestToTarget As New PinListData
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





    'For TrimCode = TrimStart_Dec To TrimEnd_Dec Step TrimStep
    If TrimStep = 0 Then
        TrimStep = 1
    
    End If
    TrimSeq = 0


    If TheExec.enableWord("TTR_ADDRIO_TRIM") = True Then
        TrimCode = TTR_TrimStart
    Else
        TrimCode = TrimStart_Dec
    End If

    TrimState = 0
    Site_TrimCode = TrimCode


    TrimFinishForEachSite = False


    While TrimCode >= TrimStart_Dec And TrimCode <= TrimEnd_Dec And TrimState <> 2 And (TrimFinishForEachSite.Any(False) = True) 'TrimSeq <= Abs((TrimEnd_Dec - TrimStart_Dec) / TrimStep)  'Step TrimStep

        For Each site In TheExec.sites
            TrimDSPwave.Element(0) = Site_TrimCode
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

        'Report_TestLimit_by_CZ_Format MeasureValue, unit:=unitOhm, forceVal:=ForceValueArray(0), forceunit:=unitVolt, UserVar7:="Trim", TestSeqNum:=TrimCode
        'TheExec.DataLog.WriteComment " *** ^^^ Trim Code = " & TrimCode & " ^^^ *** "
        
        If TheExec.enableWord("TTR_ADDRIO_TRIM") = True And TTR_HiVal <> TTR_LoVal Then
            i = 0
            For Each site In TheExec.sites
                If i = 0 Then
                    TrimCodeMax = Site_TrimCode
                    TrimCodeMin = Site_TrimCode
                Else
                    If TrimCodeMax < Site_TrimCode Then
                        TrimCodeMax = Site_TrimCode
                    End If
                    If TrimCodeMin > Site_TrimCode Then
                        TrimCodeMin = Site_TrimCode
                    End If
                End If
                i = i + 1
            Next site
        
        
            If TrimCodeMax = TrimCodeMin Then
                TrimCode = TrimCodeMax
                Report_TestLimit_by_CZ_Format MeasureValue, unit:=unitOhm, ForceVal:=ForceValueArray(0), ForceUnit:=unitVolt, UserVar7:="Trim", TestSeqNum:=TrimCode
                TheExec.Datalog.WriteComment " *** ^^^ Trim Code = " & TrimCode & " ^^^ *** "
            Else
                For Each pin In MeasureValue.Pins
                    For Each site In TheExec.sites
                    
                        Report_TestLimit_by_CZ_Format MeasureValue.Pins(pin), unit:=unitOhm, ForceVal:=ForceValueArray(0), ForceUnit:=unitVolt, UserVar7:="Trim", TestSeqNum:=Site_TrimCode, TNGroup:=SubBlockName
                    
                    Next site
                Next pin
                TheExec.Datalog.WriteComment " *** ^^^ Trim Sequence = " & TrimSeq & " ^^^ *** "
           
            End If
            
        Else
            Report_TestLimit_by_CZ_Format MeasureValue, unit:=unitOhm, ForceVal:=ForceValueArray(0), ForceUnit:=unitVolt, UserVar7:="Trim", TestSeqNum:=TrimCode, TNGroup:=SubBlockName
            TheExec.Datalog.WriteComment " *** ^^^ Trim Code = " & TrimCode & " ^^^ *** "
        End If

        '/* ----------------------------------------- */
        '/* --- ColseToTargetValue / CloseToTargetTrimCode
        '/* ----------------------------------------- */

        'If ColseToTargetValue.Pins.Count = 0 Then
        'If TrimCode = TrimStart_Dec Then
        If TrimSeq = 0 Then
            '/* ----------------------------------------- */
            '/* --- Create Pins in pinlistdata
            '/* ----------------------------------------- */
            ColseToTargetValue = MeasureValue.Copy
            CloseToTargetTrimCode = MeasureValue.Copy   '0 ' MeasureValue.Math.Multiply(0).Add(TrimCode)
            CloseToTargetTrimCode = TrimCode

            '/* ----------------------------------------- */
            '/* --- Reset the value
            '/* ----------------------------------------- */

            'ColseToTargetValue = 0
            'CloseToTargetTrimCode = 0
        End If
        'Else



        '/* --- Closest To Target --- */
        CompareResult_ClosestToTarget = MeasureValue.Math.Subtract(MeasR_TrimTarget).Abs.compare(LessThan, ColseToTargetValue.Math.Subtract(MeasR_TrimTarget).Abs)

        '/* --- Below To Target --- */
        CompareResult_BelowTarget = MeasureValue.Math.compare(LessThanOrEqualTo, MeasR_TrimTarget)

        '/* --- Over To Target for ColseToTargetValue --- */ 2017-11-01
        CompareResult_OverTarget = ColseToTargetValue.Math.compare(GreaterThan, MeasR_TrimTarget)




        For Each pin In MeasureValue.Pins
            For Each site In TheExec.sites
            
                TrimCode = Site_TrimCode
            
                If CompareResult_ClosestToTarget.Pins(pin).value And CompareResult_BelowTarget.Pins(pin).value Then     '/* --- Closest To Target --- */ && '/* --- Below To Target --- */

                    CloseToTargetTrimCode.Pins(pin).value = TrimCode
                    ColseToTargetValue.Pins(pin).value = MeasureValue.Pins(pin).value
                ElseIf CompareResult_OverTarget.Pins(pin).value And (CompareResult_ClosestToTarget.Pins(pin).value Or CompareResult_BelowTarget.Pins(pin).value) Then

                    CloseToTargetTrimCode.Pins(pin).value = TrimCode
                    ColseToTargetValue.Pins(pin).value = MeasureValue.Pins(pin).value

                Else
                '
                '
                '     If TrimCode = TrimEnd_Dec Then
                '
                '        If ColseToTargetValue.Pins(Pin).Value = 0 Then
                '            CloseToTargetTrimCode.Pins(Pin).Value = TrimCode
                '            ColseToTargetValue.Pins(Pin).Value = MeasureValue.Pins(Pin).Value
                '        End If
                '
                '    End If
                End If
            Next site
        Next pin




        TheHdw.DSSC.Pins(DigSrc_pin).Pattern(patt_ary(0)).Source.Signals.Unload

        If TheExec.enableWord("TTR_ADDRIO_TRIM") = True Then
            If TTR_HiVal <> TTR_LoVal Then

                If InStr(1, TrimResultByPins, "grp", vbTextCompare) = 0 Then
                    For Each site In TheExec.sites
                        If MeasureValue.pin(TrimResultByPins).value > TTR_HiVal Then
                            Site_TrimCode = Site_TrimCode - TTR_Trend
                        ElseIf MeasureValue.pin(TrimResultByPins).value < TTR_LoVal Then
                            Site_TrimCode = Site_TrimCode + TTR_Trend
                        Else
                            Site_TrimCode = Site_TrimCode
                            TrimFinishForEachSite = True

                        End If
                        
                        
                        ' Check next code
                        If Site_TrimCode < TrimStart_Dec Then
                            Site_TrimCode = TrimStart_Dec
                            TrimFinishForEachSite = True
                        
                        ElseIf Site_TrimCode > TrimEnd_Dec Then
                        
                            Site_TrimCode = TrimEnd_Dec
                            TrimFinishForEachSite = True
                        Else
                        
                        End If
                        
                        
                        
                        
                        
                        
                    Next site
                Else
                    Stop
                End If

            Else

                '/* --- ZCPU / ZCPD / ZCODT --- */
                'MeasureValue_Minimum = MeasureValue.Analyze.Minimum
                'MeasureValue_Maximum = MeasureValue.Analyze.Maximum
                'MinimumValue_GreaterThan_Traget  = MeasureValue_Minimum.Compare(GreaterThan,MeasR_TrimTarget)

                If TrimState = 0 Then
                    MeasureValue_Minimum = MeasureValue.Analyze.Minimum
                    MinimumValue_LessThan_Traget = MeasureValue_Minimum.compare(LessThan, MeasR_TrimTarget)

                    If MinimumValue_LessThan_Traget.Any(True) Then
                        TrimState = -1      ' find a trim code let R(all pins) > target
                        TrimCode = TrimCode - 1
                    Else
                        TrimState = 1       ' find a trim code let R(all pins) < target
                        TrimCode = TrimCode + 1
                    End If


                ElseIf TrimState = -1 Then
                    MeasureValue_Minimum = MeasureValue.Analyze.Minimum
                    MinimumValue_LessThan_Traget = MeasureValue_Minimum.compare(LessThan, MeasR_TrimTarget)


                    If MinimumValue_LessThan_Traget.Any(True) Then
                        TrimState = -1
                        TrimCode = TrimCode - 1

                        If TrimCode < TrimStart_Dec Then
                            TrimState = 1
                            TrimCode = TTR_TrimStart + 1
                        End If

                    Else
                        TrimState = 1
                        TrimCode = TTR_TrimStart + 1
                    End If

                ElseIf TrimState = 1 Then
                    MeasureValue_Maximum = MeasureValue.Analyze.Maximum
                    MaximumValue_GreaterThan_Traget = MeasureValue_Maximum.compare(GreaterThan, MeasR_TrimTarget)
                    If MaximumValue_GreaterThan_Traget.Any(True) Then
                        TrimState = 1
                        TrimCode = TrimCode + 1
                    Else
                        TrimState = 2
                    End If


                Else
                    Stop
                End If

                Site_TrimCode = TrimCode

            End If
        Else 'Not TTR
            TrimCode = TrimCode + TrimStep
            Site_TrimCode = TrimCode
        End If


        TrimSeq = TrimSeq + 1
    Wend




    For Each site In TheExec.sites
        TrimCode = Site_TrimCode
        For Each pin In MeasureValue.Pins
            If TrimCode = TrimEnd_Dec Or TrimCode = TrimStart_Dec Then
                If ColseToTargetValue.Pins(pin).value = 0 Then
                    CloseToTargetTrimCode.Pins(pin).value = TrimCode
                    ColseToTargetValue.Pins(pin).value = MeasureValue.Pins(pin).value
                End If
            End If
        Next pin
    Next site





















    '/* ----------------------------------------- */
    '/* --- Show Select Impedance & Trim Code
    '/* ----------------------------------------- */

    'Report_TestLimit_by_CZ_Format ColseToTargetValue, Unit:=unitOhm, forceVal:=ForceValueArray(0), forceunit:=unitVolt, UserVar6:="Imp", UserVar5:="Closest", TNGroup:=SubBlockName
    Report_TestLimit_by_CZ_Format ColseToTargetValue, unit:=unitOhm, ForceVal:=ForceValueArray(0), ForceUnit:=unitVolt, UserVar6:="x", UserVar7:="Closest", TNGroup:=SubBlockName
                
    TheExec.Datalog.WriteComment " *** ^^^ Closest & Below to Target Impedance ( " & MeasR_TrimTarget & " ohm ) ^^^ *** Impedance"

    'Report_TestLimit_by_CZ_Format CloseToTargetTrimCode, MeasType:="c", UserVar6:="Code", UserVar5:="Closest", TNGroup:=SubBlockName
    Report_TestLimit_by_CZ_Format CloseToTargetTrimCode, MeasType:="c", UserVar6:="x", UserVar7:="Closest", TNGroup:=SubBlockName

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
'        If MeasR_Pins_SingleEnd <> "" Then
'            MeasureValue = TheHdw.PPMU.Pins(MeasR_Pins).Read(tlPPMUReadMeasurements, pc_Def_PPMU_ReadPoint)
'        Else
'            MeasureValue = TheHdw.PPMU.Pins(MeasR_Pins_P).Read(tlPPMUReadMeasurements, pc_Def_PPMU_ReadPoint)
'        End If
'
'
'        For Each Pin In MeasureValue.Pins
'            For Each Site In TheExec.sites
'                MeasureValue.Pins(Pin) = (upperbound - lowerbound + 1) * Rnd() + lowerbound
'            Next Site
'        Next Pin
        
        MeasureValue = ColseToTargetValue
    End If
    '/* ----------------------------------------- */
    '/* --- Show Trim Log
    '/* ----------------------------------------- */

    TheExec.Datalog.WriteComment " *** Final Trim Impedance ***"

    For Each pin In MeasureValue.Pins
        'Report_TestLimit_by_CZ_Format MeasureValue.Pins(Pin), Unit:=unitOhm, ForceVal:=ForceValueArray(0), forceunit:=unitVolt, ForceResults:=tlForceFlow, UserVar5:="FINAL", UserVar6:="IMP", TNGroup:=SubBlockName
        Report_TestLimit_by_CZ_Format MeasureValue.Pins(pin), unit:=unitOhm, ForceVal:=ForceValueArray(0), ForceUnit:=unitVolt, ForceResults:=tlForceFlow, UserVar5:="x", UserVar6:="x", TNGroup:=SubBlockName
    Next pin
    
    'Report_TestLimit_by_CZ_Format MeasureValue, Unit:=unitOhm, ForceVal:=ForceValueArray(0), forceunit:=unitVolt, ForceResults:=tlForceFlow, UserVar5:="x", UserVar6:="x", TNGroup:=SubBlockName
    
    TheExec.Datalog.WriteComment " *** Final Trim Code ***"
    
    
    'Report_TestLimit_by_CZ_Format TrimDSPwave.Element(0), ForceResults:=tlForceFlow, MeasType:="c", PinName:=TNamePin, UserVar5:="FINAL", UserVar6:="CODE", TNGroup:=SubBlockName
    Report_TestLimit_by_CZ_Format TrimDSPwave.Element(0), ForceResults:=tlForceNone_CZ, MeasType:="c", PinName:=TNamePin, UserVar5:="x", UserVar6:="x", TNGroup:=SubBlockName






    TheHdw.DSSC.Pins(DigSrc_pin).Pattern(patt_ary(0)).Source.Signals.Unload




    Call HardIP_WriteFuncResult(m_testName:=TheExec.DataManager.instancename)


    If eFuseName <> "" Then
         TheExec.Datalog.WriteComment " *** save efuse Trim Code ***"

        Call SaveEfuseDecimal_to_Dictionary(eFuseName, TrimResult)
        'Call WriteDecToEFuse(eFuseName, TrimResult)
    End If



    If storename <> "" Then
        Call StoreDataAllType(storename, InDSPWave)
    End If



    DebugPrintFunc patset.value

    Call CZ_Style_TName_InstanceInfo_Clear   'added by Kaino for CZ style Tname


Exit Function 'Add ErrHandler 2023/05/29
errHandler: 'Add ErrHandler 2023/05/29
    Call Print_Error_Message(Error_Info, "VBT_LIB_K_ADDRIO_AP", "K_TrimImpedance_Turks_TTR") 'Add ErrHandler 2023/05/29
    If AbortTest Then Exit Function Else Resume Next 'Add ErrHandler 2023/05/29
End Function



Public Function K_TrimImpedance_Turks_V1(Optional patset As Pattern, Optional CPUA_Flag_In_Pat As Boolean, _
    Optional MeasR_Pins_SingleEnd As PinList, Optional MeasR_Pins_Differential As PinList, _
    Optional MeasR_TrimTarget As Double, _
    Optional ForceVtoMeasR As String, _
    Optional DigSrc_pin As PinList, Optional DigSrc_Sample_Size As Long, _
    Optional TrimStart_Dec As Long, Optional TrimEnd_Dec As Long, Optional TrimStep As Long, _
    Optional TrimResultByPins As PinList, _
    Optional eFuseName As String, Optional eFuseType As String, _
    Optional storename As String, _
    Optional PullDown_Mode As Boolean, _
    Optional IRange_mA As Double, _
    Optional TTR_TrimStart As Long, _
    Optional TTR_HiVal As Double, _
    Optional TTR_LoVal As Double, _
    Optional TTR_Trend As Integer, _
    Optional SubBlockName As String, _
    Optional Validating_ As Boolean) As Long
On Error GoTo errHandler 'Add ErrHandler 2023/05/29


    Dim Pin_Ary() As String
    Dim Pin_Cnt As Long
    Dim TrimCode As Long
    Dim site As Variant
    Dim pin As Variant

    Dim TrimCodeMax As Long
    Dim TrimCodeMin As Long
    Dim i As Integer
    Dim j As Integer
    
    Dim TrimSeq As Long
    Dim Site_TrimCode As New SiteLong
    Dim TrimFinishForEachSite As New SiteBoolean
    Dim TrimState As Integer            ' -1 : 0 : 1
    Dim OneStep As Integer
    Dim MeasureValue_Minimum As New SiteVariant
    Dim MeasureValue_Maximum As New SiteVariant
    Dim MinimumValue_LessThan_Traget As New SiteBoolean
    Dim MaximumValue_GreaterThan_Traget As New SiteBoolean
    
    
    Dim TrimDSPwave As New DSPWave
    Dim InDSPWave As New DSPWave
    Dim MeasureValue As New PinListData
    Dim ColseToTargetValue As New PinListData
    Dim CloseToTargetTrimCode As New PinListData
    Dim CompareResult_ClosestToTarget As New PinListData
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

    Dim TrimStr As String

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

        If Pin_Cnt <= 0 Then
            Stop
        Else
            TheExec.Datalog.WriteComment " *** vvv *** Differential Pair P pis: " & MeasR_Pins_P
            TheExec.Datalog.WriteComment " *** vvv *** Differential Pair N pis: " & MeasR_Pins_N

        End If
    End If





    '/* ----------------------------------------- */
    Call ForceConditionToFocreValue(ForceVtoMeasR, ForceValueStr, ForceValueArray)





    TheHdw.Digital.ApplyLevelsTiming True, True, True, tlPowered
    TheHdw.Digital.Patgen.Halt
    TheHdw.Digital.Patgen.Continue 0, cpuA + cpuB + cpuC + cpuD
    TheHdw.Digital.Patgen.HaltMode = tlHaltOnOpcode
    TheHdw.Patterns(patset).Load



    'Call PATT_GetPatListFromPatternSet(PatSet.Value, PattArray, PatCount)


    TrimDSPwave.CreateConstant 0, 1, DspLong
    InDSPWave.CreateConstant 0, DigSrc_Sample_Size, DspLong





    'For TrimCode = TrimStart_Dec To TrimEnd_Dec Step TrimStep
    If TrimStep = 0 Then
        TrimStep = 1
    
    End If
    TrimSeq = 0


    If TheExec.enableWord("TTR_ADDRIO_TRIM") = True Then
        TrimCode = TTR_TrimStart
    Else
        TrimCode = TrimStart_Dec
    End If

    TrimState = 0
    Site_TrimCode = TrimCode


    TrimFinishForEachSite = False


    While TrimCode >= TrimStart_Dec And TrimCode <= TrimEnd_Dec And TrimState <> 2 And (TrimFinishForEachSite.Any(False) = True) 'TrimSeq <= Abs((TrimEnd_Dec - TrimStart_Dec) / TrimStep)  'Step TrimStep

        For Each site In TheExec.sites
            TrimDSPwave.Element(0) = Site_TrimCode
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
            TheExec.Datalog.WriteComment " *** off-line simulation value ***"
        End If
        '/* ----------------------------------------- */
        '/* --- Show Trim Log
        '/* ----------------------------------------- */

        'Report_TestLimit_by_CZ_Format MeasureValue, unit:=unitOhm, forceVal:=ForceValueArray(0), forceunit:=unitVolt, UserVar7:="Trim", TestSeqNum:=TrimCode
        'TheExec.DataLog.WriteComment " *** ^^^ Trim Code = " & TrimCode & " ^^^ *** "
        
        
            TrimStr = vbNullString
             For Each site In TheExec.sites
                For j = 0 To DigSrc_Sample_Size - 1
                    TrimStr = TrimStr & InDSPWave.Element(j)
                Next j
                Exit For
            Next site
        
        If TheExec.enableWord("TTR_ADDRIO_TRIM") = True And TTR_HiVal <> TTR_LoVal Then
            i = 0
            For Each site In TheExec.sites
                If i = 0 Then
                    TrimCodeMax = Site_TrimCode
                    TrimCodeMin = Site_TrimCode
                Else
                    If TrimCodeMax < Site_TrimCode Then
                        TrimCodeMax = Site_TrimCode
                    End If
                    If TrimCodeMin > Site_TrimCode Then
                        TrimCodeMin = Site_TrimCode
                    End If
                End If
                i = i + 1
            Next site
        
           
           
           For j = 0 To DigSrc_Sample_Size - 1
                TrimStr = TrimStr & InDSPWave.Element(j)
           Next j
        
        
            If TrimCodeMax = TrimCodeMin Then
                TrimCode = TrimCodeMax
                Report_TestLimit_by_CZ_Format MeasureValue, unit:=unitOhm, ForceVal:=ForceValueArray(0), ForceUnit:=unitVolt, UserVar7:="Trim", TestSeqNum:=TrimCode
                TheExec.Datalog.WriteComment " *** ^^^ Trim Code = " & TrimCode & "(LSB->MSB: " & TrimStr & ") ^^^ *** "
                
                
                
            Else
                For Each pin In MeasureValue.Pins
                    For Each site In TheExec.sites
                    
                        Report_TestLimit_by_CZ_Format MeasureValue.Pins(pin), unit:=unitOhm, ForceVal:=ForceValueArray(0), ForceUnit:=unitVolt, UserVar7:="Trim", TestSeqNum:=Site_TrimCode, TNGroup:=SubBlockName
                    
                    Next site
                Next pin
                TheExec.Datalog.WriteComment " *** ^^^ Trim Sequence = " & TrimSeq & " ^^^ *** "
           
            End If
            
        Else
            Report_TestLimit_by_CZ_Format MeasureValue, unit:=unitOhm, ForceVal:=ForceValueArray(0), ForceUnit:=unitVolt, UserVar7:="Trim", TestSeqNum:=TrimCode, TNGroup:=SubBlockName
            TheExec.Datalog.WriteComment " *** ^^^ Trim Code = " & TrimCode & "(LSB->MSB: " & TrimStr & ") ^^^ *** "
        End If

        '/* ----------------------------------------- */
        '/* --- ColseToTargetValue / CloseToTargetTrimCode
        '/* ----------------------------------------- */

        'If ColseToTargetValue.Pins.Count = 0 Then
        'If TrimCode = TrimStart_Dec Then
        If TrimSeq = 0 Then
            '/* ----------------------------------------- */
            '/* --- Create Pins in pinlistdata
            '/* ----------------------------------------- */
            ColseToTargetValue = MeasureValue.Copy
            CloseToTargetTrimCode = MeasureValue.Copy   '0 ' MeasureValue.Math.Multiply(0).Add(TrimCode)
            CloseToTargetTrimCode = TrimCode

            '/* ----------------------------------------- */
            '/* --- Reset the value
            '/* ----------------------------------------- */

            'ColseToTargetValue = 0
            'CloseToTargetTrimCode = 0
        End If
        'Else



        '/* --- Closest To Target --- */
        CompareResult_ClosestToTarget = MeasureValue.Math.Subtract(MeasR_TrimTarget).Abs.compare(LessThan, ColseToTargetValue.Math.Subtract(MeasR_TrimTarget).Abs)

        '/* --- Below To Target --- */
        CompareResult_BelowTarget = MeasureValue.Math.compare(LessThanOrEqualTo, MeasR_TrimTarget)

        '/* --- Over To Target for ColseToTargetValue --- */ 2017-11-01
        CompareResult_OverTarget = ColseToTargetValue.Math.compare(GreaterThan, MeasR_TrimTarget)




        For Each pin In MeasureValue.Pins
            For Each site In TheExec.sites
            
                TrimCode = Site_TrimCode
            
                If CompareResult_ClosestToTarget.Pins(pin).value And CompareResult_BelowTarget.Pins(pin).value Then     '/* --- Closest To Target --- */ && '/* --- Below To Target --- */

                    CloseToTargetTrimCode.Pins(pin).value = TrimCode
                    ColseToTargetValue.Pins(pin).value = MeasureValue.Pins(pin).value
                ElseIf CompareResult_OverTarget.Pins(pin).value And (CompareResult_ClosestToTarget.Pins(pin).value Or CompareResult_BelowTarget.Pins(pin).value) Then

                    CloseToTargetTrimCode.Pins(pin).value = TrimCode
                    ColseToTargetValue.Pins(pin).value = MeasureValue.Pins(pin).value

                Else
                '
                '
                '     If TrimCode = TrimEnd_Dec Then
                '
                '        If ColseToTargetValue.Pins(Pin).Value = 0 Then
                '            CloseToTargetTrimCode.Pins(Pin).Value = TrimCode
                '            ColseToTargetValue.Pins(Pin).Value = MeasureValue.Pins(Pin).Value
                '        End If
                '
                '    End If
                End If
            Next site
        Next pin




        TheHdw.DSSC.Pins(DigSrc_pin).Pattern(patt_ary(0)).Source.Signals.Unload

        If TheExec.enableWord("TTR_ADDRIO_TRIM") = True Then
            If TTR_HiVal <> TTR_LoVal Then

                If InStr(1, TrimResultByPins, "grp", vbTextCompare) = 0 Then
                    For Each site In TheExec.sites
                        If MeasureValue.pin(TrimResultByPins).value > TTR_HiVal Then
                            Site_TrimCode = Site_TrimCode - TTR_Trend
                        ElseIf MeasureValue.pin(TrimResultByPins).value < TTR_LoVal Then
                            Site_TrimCode = Site_TrimCode + TTR_Trend
                        Else
                            Site_TrimCode = Site_TrimCode
                            TrimFinishForEachSite = True

                        End If
                        
                        
                        ' Check next code
                        If Site_TrimCode < TrimStart_Dec Then
                            Site_TrimCode = TrimStart_Dec
                            TrimFinishForEachSite = True
                        
                        ElseIf Site_TrimCode > TrimEnd_Dec Then
                        
                            Site_TrimCode = TrimEnd_Dec
                            TrimFinishForEachSite = True
                        Else
                        
                        End If
                        
                        
                        
                        
                        
                        
                    Next site
                Else
                    Stop
                End If

            Else

                '/* --- ZCPU / ZCPD / ZCODT --- */
                'MeasureValue_Minimum = MeasureValue.Analyze.Minimum
                'MeasureValue_Maximum = MeasureValue.Analyze.Maximum
                'MinimumValue_GreaterThan_Traget  = MeasureValue_Minimum.Compare(GreaterThan,MeasR_TrimTarget)

                If TrimState = 0 Then
                    MeasureValue_Minimum = MeasureValue.Analyze.Minimum
                    MinimumValue_LessThan_Traget = MeasureValue_Minimum.compare(LessThan, MeasR_TrimTarget)

                    If MinimumValue_LessThan_Traget.Any(True) Then
                        TrimState = -1      ' find a trim code let R(all pins) > target
                        TrimCode = TrimCode - 1
                    Else
                        TrimState = 1       ' find a trim code let R(all pins) < target
                        TrimCode = TrimCode + 1
                    End If


                ElseIf TrimState = -1 Then
                    MeasureValue_Minimum = MeasureValue.Analyze.Minimum
                    MinimumValue_LessThan_Traget = MeasureValue_Minimum.compare(LessThan, MeasR_TrimTarget)


                    If MinimumValue_LessThan_Traget.Any(True) Then
                        TrimState = -1
                        TrimCode = TrimCode - 1

                        If TrimCode < TrimStart_Dec Then
                            TrimState = 1
                            TrimCode = TTR_TrimStart + 1
                        End If

                    Else
                        TrimState = 1
                        TrimCode = TTR_TrimStart + 1
                    End If

                ElseIf TrimState = 1 Then
                    MeasureValue_Maximum = MeasureValue.Analyze.Maximum
                    MaximumValue_GreaterThan_Traget = MeasureValue_Maximum.compare(GreaterThan, MeasR_TrimTarget)
                    If MaximumValue_GreaterThan_Traget.Any(True) Then
                        TrimState = 1
                        TrimCode = TrimCode + 1
                    Else
                        TrimState = 2
                    End If


                Else
                    Stop
                End If

                Site_TrimCode = TrimCode

            End If
        Else 'Not TTR
            TrimCode = TrimCode + TrimStep
            Site_TrimCode = TrimCode
        End If


        TrimSeq = TrimSeq + 1
    Wend




    For Each site In TheExec.sites
        TrimCode = Site_TrimCode
        For Each pin In MeasureValue.Pins
            If TrimCode = TrimEnd_Dec Or TrimCode = TrimStart_Dec Then
                If ColseToTargetValue.Pins(pin).value = 0 Then
                    CloseToTargetTrimCode.Pins(pin).value = TrimCode
                    ColseToTargetValue.Pins(pin).value = MeasureValue.Pins(pin).value
                End If
            End If
        Next pin
    Next site





















    '/* ----------------------------------------- */
    '/* --- Show Select Impedance & Trim Code
    '/* ----------------------------------------- */

    'Report_TestLimit_by_CZ_Format ColseToTargetValue, Unit:=unitOhm, forceVal:=ForceValueArray(0), forceunit:=unitVolt, UserVar6:="Imp", UserVar5:="Closest", TNGroup:=SubBlockName
    Report_TestLimit_by_CZ_Format ColseToTargetValue, unit:=unitOhm, ForceVal:=ForceValueArray(0), ForceUnit:=unitVolt, UserVar6:="x", UserVar7:="Closest", TNGroup:=SubBlockName
                
    TheExec.Datalog.WriteComment " *** ^^^ Closest & Below to Target Impedance ( " & MeasR_TrimTarget & " ohm ) ^^^ *** Impedance"

    'Report_TestLimit_by_CZ_Format CloseToTargetTrimCode, MeasType:="c", UserVar6:="Code", UserVar5:="Closest", TNGroup:=SubBlockName
    Report_TestLimit_by_CZ_Format CloseToTargetTrimCode, MeasType:="c", UserVar6:="x", UserVar7:="Closest", TNGroup:=SubBlockName

    TheExec.Datalog.WriteComment " *** ^^^ Closest & Below to Target Impedance ( " & MeasR_TrimTarget & " ohm ) ^^^ *** Trim Code"

    'TheExec.Datalog.WriteComment " *** ^^^ Closest & Below to Target Impedance ( " & MeasR_TrimTarget & " ohm )  Trim Code  ^^^ *** "



    If InStr(1, TrimResultByPins, "grp", vbTextCompare) > 0 Then

        CloseToTargetTrimCode = CloseToTargetTrimCode.Analyze.mean


        For Each site In TheExec.sites

            TrimDSPwave.Element(0) = Round(CloseToTargetTrimCode.Pins(0).value, 0)


            'TheExec.Datalog.WriteComment " *** ^^^ Trim Code : site (" & Site & "): " & TrimDSPwave.Element(0) & " , ( " & TrimResultByPins & " ) ^^^ *** "
            InDSPWave = TrimDSPwave.ConvertStreamTo(tldspSerial, DigSrc_Sample_Size, 0, Bit0IsMsb)

            TheExec.Datalog.WriteComment " *** ^^^ Trim Code : site (" & site & "): " & TrimDSPwave.Element(0) & "(LSB->MSB: " & K_Str_DumpDspWave(InDSPWave) & ")" & " , ( " & TrimResultByPins & " ) ^^^ *** "

            TrimResult = TrimDSPwave.Element(0)
        Next site


        TNamePin = "x"
        'Stop
    ElseIf TrimResultByPins <> "" Then

        For Each site In TheExec.sites

            TrimDSPwave.Element(0) = CloseToTargetTrimCode.Pins(TrimResultByPins).value

            'TheExec.Datalog.WriteComment " *** ^^^ Trim Code : site (" & Site & "): " & TrimDSPwave.Element(0) & " , ( " & TrimResultByPins & " ) ^^^ *** "

            InDSPWave = TrimDSPwave.ConvertStreamTo(tldspSerial, DigSrc_Sample_Size, 0, Bit0IsMsb)

            TheExec.Datalog.WriteComment " *** ^^^ Trim Code : site (" & site & "): " & TrimDSPwave.Element(0) & "(LSB->MSB: " & K_Str_DumpDspWave(InDSPWave) & ")" & " , ( " & TrimResultByPins & " ) ^^^ *** "

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
'        If MeasR_Pins_SingleEnd <> "" Then
'            MeasureValue = TheHdw.PPMU.Pins(MeasR_Pins).Read(tlPPMUReadMeasurements, pc_Def_PPMU_ReadPoint)
'        Else
'            MeasureValue = TheHdw.PPMU.Pins(MeasR_Pins_P).Read(tlPPMUReadMeasurements, pc_Def_PPMU_ReadPoint)
'        End If
'
'
'        For Each Pin In MeasureValue.Pins
'            For Each Site In TheExec.sites
'                MeasureValue.Pins(Pin) = (upperbound - lowerbound + 1) * Rnd() + lowerbound
'            Next Site
'        Next Pin
        
        MeasureValue = ColseToTargetValue
    End If
    '/* ----------------------------------------- */
    '/* --- Show Trim Log
    '/* ----------------------------------------- */

    TheExec.Datalog.WriteComment " *** Final Trim Impedance ***"

    For Each pin In MeasureValue.Pins
        'Report_TestLimit_by_CZ_Format MeasureValue.Pins(Pin), Unit:=unitOhm, ForceVal:=ForceValueArray(0), forceunit:=unitVolt, ForceResults:=tlForceFlow, UserVar5:="FINAL", UserVar6:="IMP", TNGroup:=SubBlockName
        Report_TestLimit_by_CZ_Format MeasureValue.Pins(pin), unit:=unitOhm, ForceVal:=ForceValueArray(0), ForceUnit:=unitVolt, ForceResults:=tlForceFlow, UserVar5:="x", UserVar6:="x", TNGroup:=SubBlockName
    Next pin
    
    'Report_TestLimit_by_CZ_Format MeasureValue, Unit:=unitOhm, ForceVal:=ForceValueArray(0), forceunit:=unitVolt, ForceResults:=tlForceFlow, UserVar5:="x", UserVar6:="x", TNGroup:=SubBlockName
    
    TheExec.Datalog.WriteComment " *** Final Trim Code ***"
    
    
    'Report_TestLimit_by_CZ_Format TrimDSPwave.Element(0), ForceResults:=tlForceFlow, MeasType:="c", PinName:=TNamePin, UserVar5:="FINAL", UserVar6:="CODE", TNGroup:=SubBlockName
    Report_TestLimit_by_CZ_Format TrimDSPwave.Element(0), ForceResults:=tlForceFlow, MeasType:="c", PinName:=TNamePin, UserVar5:="x", UserVar6:="x", TNGroup:=SubBlockName






    TheHdw.DSSC.Pins(DigSrc_pin).Pattern(patt_ary(0)).Source.Signals.Unload




    Call HardIP_WriteFuncResult(m_testName:=TheExec.DataManager.instancename)


    TheExec.Datalog.WriteComment " *** Triming from (" & TrimStart_Dec & ") to (" & TrimEnd_Dec & ") step (" & TrimStep & ")"
    If eFuseName <> "" Then
         TheExec.Datalog.WriteComment " *** save efuse Trim Code ***"

        Call SaveEfuseDecimal_to_Dictionary(eFuseName, TrimResult)
        'Call WriteDecToEFuse(eFuseName, TrimResult)
    End If


    

    If storename <> "" Then
        Call StoreDataAllType(storename, InDSPWave)
    End If



    DebugPrintFunc patset.value

    Call CZ_Style_TName_InstanceInfo_Clear   'added by Kaino for CZ style Tname


Exit Function 'Add ErrHandler 2023/05/29
errHandler: 'Add ErrHandler 2023/05/29
    Call Print_Error_Message(Error_Info, "VBT_LIB_K_ADDRIO_AP", "K_TrimImpedance_Turks_V1") 'Add ErrHandler 2023/05/29
    If AbortTest Then Exit Function Else Resume Next 'Add ErrHandler 2023/05/29
End Function

Public Function K_TrimImpedance_CodeCheck_V1(Optional patset As Pattern, Optional CPUA_Flag_In_Pat As Boolean, _
    Optional MeasR_Pins_SingleEnd As PinList, Optional MeasR_Pins_Differential As PinList, _
    Optional MeasR_TrimTarget As Double, _
    Optional ForceVtoMeasR As String, _
    Optional DigSrc_pin As PinList, Optional DigSrc_Sample_Size As Long, _
    Optional dummy1 As String, Optional dummy2 As Long, Optional dummy3 As Long, _
    Optional TrimResultByPins As PinList, _
    Optional dummy4 As String, Optional dummy5 As String, _
    Optional dummy6 As String, _
    Optional PullDown_Mode As Boolean, _
    Optional IRange_mA As Double, _
    Optional FuseDefault_check As String, _
    Optional StoreName_check As String, _
    Optional ReadCodeWithoutChecking As Boolean, _
    Optional dummy7 As String, _
    Optional SubBlockName As String, _
    Optional Validating_ As Boolean) As Long
On Error GoTo errHandler 'Add ErrHandler 2023/05/29
    '/* 2017-12-04 update by Kaino  for IRange_mA */
    '/* 2019-06-10 update by Kaino  for Turks */
 
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
    Call ForceConditionToFocreValue(ForceVtoMeasR, ForceValueStr, ForceValueArray)
    
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
            Set field = opbank.Fields(fieldStr)
            If UCase(field.name) = UCase(FuseDefault_check) Then
            'If UCase(CFGFuse.category(i).name) = UCase(FuseDefault_check) Then
            
                'Stop
    
                
                For Each site In TheExec.sites
                    'TrimDSPwave.Element(0) = Site_TrimCode
                    ''20180325 judge default code is hex or decimal
                    If LCase(field.Default) Like "*0x*" Then
                        TrimDSPwave.Element(0) = auto_HexStr2Value(field.Default)
                    Else
                        TrimDSPwave.Element(0) = val(field.Default)
                    End If
                    'If LCase(CFGFuse.category(i).DefaultValue) Like "*0x*" Then
                        'TrimDSPwave.Element(0) = auto_HexStr2Value(CFGFuse.category(i).DefaultValue)
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
            InDSPWave = GetStoreDataAllType(StoreName_check)
            
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
        
        'Report_TestLimit_by_CZ_Format TrimDSPwave.Element(0), ForceResults:=tlForceNone, MeasType:="c", PinName:=TNamePin, UserVar5:=UserVar5, UserVar6:="CODE"
        Report_TestLimit_by_CZ_Format TrimDSPwave.Element(0), ForceResults:=tlForceNone_CZ, MeasType:="c", PinName:=TNamePin, UserVar6:="CODE"
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
            Report_TestLimit_by_CZ_Format MeasureValue.Pins(pin), unit:=unitOhm, ForceVal:=ForceValueArray(0), ForceUnit:=unitVolt, ForceResults:=tlForceFlow, UserVar6:="IMP", TNGroup:=SubBlockName
        Next pin
        TheExec.Datalog.WriteComment " *** Final Trim Code ***"
        Report_TestLimit_by_CZ_Format TrimDSPwave.Element(0), ForceResults:=tlForceNone_CZ, MeasType:="c", PinName:=TNamePin, UserVar6:="CODE", TNGroup:=SubBlockName
        
        
        
        
        
        
        TheHdw.DSSC.Pins(DigSrc_pin).Pattern(patt_ary(0)).Source.Signals.Unload
        
        
        
        
        Call HardIP_WriteFuncResult(m_testName:=TheExec.DataManager.instancename)
    
    
    End If
    
    'Stop
    ''''' /* --- Mask by Kaino on 2019/06/10  for Turks --- */
    '''''   If eFuseName <> "" Then
    '''''        TheExec.Datalog.WriteComment " *** save efuse Trim Code *** " & UCase(eFuseName)
    '''''
    '''''       Call SaveEfuseDecimal_to_Dictionary(eFuseName, TrimResult)
    '''''       'Call WriteDecToEFuse(eFuseName, TrimResult)
    '''''   End If
    '''''
    '''''
    '''''   If StoreName <> "" Then
    '''''       TheExec.Datalog.WriteComment " *** save Trim Code *** " & StoreName
    '''''       Call StoreDataAllType(StoreName, InDSPwave)
    '''''   End If
    '''''
    
    
    
    DebugPrintFunc patset.value
    
    Call CZ_Style_TName_InstanceInfo_Clear   'added by Kaino for CZ style Tname
    

Exit Function 'Add ErrHandler 2023/05/29
errHandler: 'Add ErrHandler 2023/05/29
    Call Print_Error_Message(Error_Info, "VBT_LIB_K_ADDRIO_AP", "K_TrimImpedance_CodeCheck_V1") 'Add ErrHandler 2023/05/29
    If AbortTest Then Exit Function Else Resume Next 'Add ErrHandler 2023/05/29
End Function



''''Public Function K_Meas_FreqVoltCurr_Universal_func(Optional patset As Pattern, Optional TestSequence As String, Optional CPUA_Flag_In_Pat As Boolean, _
''''        Optional DisableComparePins As PinList, Optional DisableConnectPins As PinList, Optional DisableFRC As Boolean = False, Optional FRCPortName As String, _
''''        Optional MeasV_PinS As String, _
''''        Optional MeasF_PinS_SingleEnd As String, Optional MeasF_Interval As String, Optional MeasF_EventSourceWithTerminationMode As EventSourceWithTerminationMode, Optional MeasF_Flag_MeasureThreshold As Boolean = False, Optional MeasF_ThresholdPercentage As Double = 0.5, Optional MeasF_WaitTime As String, _
''''        Optional MeasI_pinS As String, Optional MeasI_Range As String, Optional MeasI_WaitTime As String, _
''''        Optional DigCap_Pin As PinList, Optional DigCap_DataWidth As Long, Optional DigCap_Sample_Size As Long, _
''''        Optional DigSrc_pin As PinList, Optional DigSrc_DataWidth As Long, Optional DigSrc_Sample_Size As Long, Optional DigSrc_Equation As String, Optional DigSrc_Assignment As String, Optional DigSrc_FlowForLoopIntegerName As String = "", _
''''        Optional SpecialCalcValSetting As CalculateMethodSetup = 0, _
''''        Optional InstSpecialSetting As InstrumentSpecialSetup = 0, _
''''        Optional CUS_Str_MainProgram As String = "", Optional CUS_Str_DigCapData As String = "", Optional CUS_Str_DigSrcData As String = "", _
''''        Optional Flag_SingleLimit As Boolean = False, Optional TestLimitPerPin_VFI As String = "FFF", _
''''        Optional MeasF_PinS_Differential As String, Optional ForceFunctional_Flag As Boolean = False, _
''''        Optional MeasF_WalkingStrobe_Flag As Boolean, Optional MeasF_WalkingStrobe_StartV As Double, Optional MeasF_WalkingStrobe_EndV As Double, Optional MeasF_WalkingStrobe_StepVoltage As Double, Optional MeasF_WalkingStrobe_BothVohVolDiffV As Double, Optional MeasF_WalkingStrobe_interval As Double, Optional MeasF_WalkingStrobe_miniFreq As Double, _
''''        Optional Meas_StoreName As String, Optional Calc_Eqn As String, _
''''        Optional Interpose_PrePat As String, Optional Interpose_PreMeas As String, Optional Interpose_PostTest As String, Optional CharSetName As String, _
''''        Optional ForceV_Val As String, Optional ForceI_Val As String, Optional UVI80_MeasV_WaitTime As String = "", _
''''        Optional RAK_Flag As Enum_RAK, Optional WaitTime_VIRZ As String, Optional MSB_First_Flag As Boolean = False, _
''''        Optional K_PreMeas_Setting As String, Optional K_PostMeas_Setting As String, Optional K_PatHalt_Setting As String, _
''''        Optional Validating_ As Boolean) As Long
''''    '/*** ------ Added by Kaino for special setting ------ ***/
''''    Dim K_PreMeas_Setting_Sequence() As String
''''    Dim K_PostMeas_Setting_Sequence() As String
''''    'Dim K_PatHalt_Setting_Sequence() As String
''''    'K_PatHalt_Setting_Sequence = Split(K_PatHalt_Setting, ";")
''''
''''    Dim K_Setting_List() As String
''''    Dim K_Setting_index As Integer
''''    Dim kk As Integer
''''
''''    Dim K_Setting_Args() As String
''''
''''    Dim K_Value_array(0) As String
''''
''''    Dim K_Pins As String
''''    Dim K_Value(0) As String
''''
''''    Dim WaitTimeSeq() As String
''''
''''    WaitTimeSeq = Split(MeasI_WaitTime, "+")
''''
''''    K_PreMeas_Setting_Sequence = Split(K_PreMeas_Setting, ";")
''''    K_PostMeas_Setting_Sequence = Split(K_PostMeas_Setting, ";")
''''
''''
''''    '/*** ^^^^^^ Added by Kaino for special setting ^^^^^^ ***/
''''
''''
''''    Dim PatCount As Long
''''    Dim PattArray() As String
''''
''''    If Validating_ Then
''''        Call PrLoadPattern(patset.value)
''''        Exit Function    ' Exit after validation
''''    End If
''''
''''    If MSB_First_Flag = True Then
''''        theexec.Datalog.WriteComment "Error: Pattern LSB sequence is reverse, please update pattern or reverse source data"
''''    End If
''''
''''    Call HardIP_InitialSetupForPatgen
''''
''''    Dim i As Long, j As Long, k As Long
''''    Dim TestOptLen As Integer
''''    Dim TestSequenceArray() As String, MeasPinAry_V() As String, MeasPinAry_F() As String, MeasPinAry_I() As String, MeasPinAry_IRange() As String
''''    Dim MeasPinAry_F_Differential() As String
''''    Dim MeasureF_Pin_Differential As New PinList
''''    Dim Ts As Variant, TestOption As Variant, site As Variant
''''    Dim TestSeqNum As Integer
''''    Dim MeasureV_pin As New PinList, MeasureF_Pin_SingleEnd As New PinList, MeasureI_pin As New PinList
''''    Dim MeasureI_Pin_CurrentRange As String
''''    Dim testnum As Long
''''    Dim InDSPWave As New DSPWave, OutDspWave As New DSPWave
''''    Dim ShowDec As String, ShowOut As String
''''    Dim patt As Variant
''''    Dim Pat As String
''''    Dim HighLimitVal() As Double, LowLimitVal() As Double
''''    Dim MeasureV_Pin_PPMU As String, MeasureV_Pin_UVI80 As String
''''    Dim d_MeasF_Interval As Double
''''    Dim FreqPinsCheckType() As String
''''    Dim ThisPinType As String
''''    Dim MeasF_EventSource As FreqCtrEventSrcSel
''''    Dim MeasF_EnableVtMode As Boolean
''''    Dim Split_F_Str() As String
''''    Dim Inst_Name_Str As String: Inst_Name_Str = theexec.DataManager.instancename  '20170728 Added for HardIP_WriteFuncResult Output
''''    Dim restore_Flag As Boolean
''''
''''
''''    ''20160906 - Return measurement to directionary if needed
''''    Dim Rtn_MeasVolt As New PinListData, Rtn_MeasCurr As New PinListData, Rtn_MeasFreq As New PinListData
''''    Dim MeasStoreName_Ary() As String
''''    Dim Interpose_PreMeas_Ary() As String
''''
''''''    Dim RTN_InterposeString As String
''''
''''    On Error GoTo errHandler
''''    Dim CheckDSPWave As New DSPWave
''''    Dim Sweep_Enable As Boolean: Sweep_Enable = False
''''    Dim Sweep_Loop_Calc_Eqn As String: Sweep_Loop_Calc_Eqn = ""
'''''    Dim Sweep_Calc_Eqn As Boolean: Sweep_Calc_Eqn = False
''''    Dim Sweep_Calc_Eqn_index As String: Sweep_Calc_Eqn_index = ""
''''    Dim Sweep_Dictionary As String: Sweep_Dictionary = ""
''''    Dim Sweep_Calc_Eqn As String: Sweep_Calc_Eqn = ""
''''
''''    Dim OutputTname() As String
''''
''''    Call tl_PinListDataSort(True)
''''    Dim instance_name As String
''''
''''    instance_name = theexec.DataManager.instancename
''''
''''    '================================================================ Roger
''''    If InStr(1, LCase(Interpose_PrePat), "sweep:") <> 0 Then
''''        Dim Sweep_Info() As Power_Sweep
''''        Dim Sweep_CUS_Str_DigCapData As String
''''        Call SortSweepInfo(Sweep_Info, Interpose_PrePat)
''''        Sweep_Enable = True
''''        Sweep_CUS_Str_DigCapData = CUS_Str_DigCapData
''''        Sweep_Calc_Eqn = Calc_Eqn
''''    End If
''''    '================================================================
''''    ''20170322-Store MeasF mid value for VT
''''    Dim SplitFreqVtValue() As String
''''    Dim DictKey_StoreVT As String
''''    Dim Dict_VT_Value As New SiteDouble
''''
''''
''''
''''    'If (UCase(MeasI_Range) Like "*CP*:*" Or UCase(MeasI_Range) Like "*FT*:*") Then MeasI_Range = Select_MeasIRange(MeasI_Range, CurrentJobName_U)   ' support different Meter_Range in different Job, add by Roger 20170628
''''
''''    '' 20160201 - Check input argumenets whether have "@" in the first character. Add it If no "@" in the beginning. Then remove it to process fomat.
''''    Call VFI_AnalyzedInputStringByAt(MeasV_PinS, MeasF_PinS_SingleEnd, MeasI_pinS, MeasI_Range, MeasF_PinS_Differential, ForceV_Val, ForceI_Val)
''''
''''    Dim ForceV_Val_Ary() As String
''''    Dim ForceI_Val_Ary() As String
''''    Dim MeasurePin_ForceV_Val As String
''''    Dim MeasurePin_ForceI_Val As String
''''    Dim MeasI_WaitTime_Ary() As String
''''    Dim MeasF_WaitTime_Ary() As String
''''    Dim UVI80_MeasV_WaitTime_Ary() As String
''''
''''
''''
''''    If TestSequence = "" Then                       '20170714
''''        ReDim TestSequenceArray(0) As String
''''        TestSequenceArray(0) = TestSequence
''''    Else
''''        TestSequenceArray = Split(TestSequence, ",")
''''    End If
''''    MeasStoreName_Ary = Split(Meas_StoreName, ",")
''''    Interpose_PreMeas_Ary = Split(Interpose_PreMeas, "|")
''''
''''    '----------------------------20180523
''''
''''    'Roger New,20180510 TName
''''    '--------------------------------------------------------------------
''''    'Call GetFlowTName
''''
''''    '----------------------------20180523
''''
''''    'Call VFI_ProcessInputString(TestSequence, MeasV_PinS, MeasI_pinS, MeasF_PinS_SingleEnd, MeasF_PinS_Differential, MeasI_Range, Meas_StoreName, Interpose_PreMeas, _
''''                                            ForceV_Val, ForceI_Val, _
''''                                            TestSequenceArray(), MeasPinAry_V(), MeasPinAry_I(), MeasPinAry_F(), _
''''                                            MeasPinAry_F_Differential(), MeasPinAry_IRange(), MeasStoreName_Ary(), Interpose_PreMeas_Ary(), ForceV_Val_Ary(), ForceI_Val_Ary())
''''
''''    'Call VFI_ProcessWaitTimeString(MeasI_WaitTime, MeasF_WaitTime, UVI80_MeasV_WaitTime, MeasI_WaitTime_Ary(), MeasF_WaitTime_Ary(), UVI80_MeasV_WaitTime_Ary(), TestSequenceArray())
''''
''''
'''''    Call HIP_Evaluate_ForceVal(ForceV_Val_Ary())
'''''
'''''    Call HIP_Evaluate_ForceVal(ForceI_Val_Ary())
''''
''''    ''20170807 - CZ test name index
'''''    gl_CZ_FlowTestNameIndex = 0
''''
'''''    Call Freq_ProcessEventSourceTerminationMode(MeasF_EventSourceWithTerminationMode, MeasF_EventSource, MeasF_EnableVtMode)
''''
''''    ''20141219 Get use-limit from flow table
'''''    Call GetFlowSingleUseLimit(HighLimitVal, LowLimitVal)
''''
''''    ''20161130-Get test name from flow table
''''    Dim FlowTestNme() As String
''''    ''========================================================================================
''''    Dim Store_Rtn_Meas() As New PinListData
''''    Dim SoreMaxNum As Long
''''    Dim StoreIndex As Long
''''    ''20170123-Get how many store name in MeasStoreName_Ary
''''    If Meas_StoreName <> "" Then
''''        SoreMaxNum = 0
''''        For i = 0 To UBound(MeasStoreName_Ary)
''''            If MeasStoreName_Ary(i) <> "" Then
''''                SoreMaxNum = SoreMaxNum + 1
''''            End If
''''        Next i
''''         ReDim Store_Rtn_Meas(SoreMaxNum - 1) As New PinListData
''''         StoreIndex = 0
''''     End If
''''    ''========================================================================================
''''    If theexec.DevChar.Setups.IsRunning = True And CharSetName <> "" And InStr(UCase(Interpose_PrePat), ":TERM:") <> 0 Then
''''        'HIO:can not applylevelsTiming for the first point  of run_shmoo
''''     Else
''''        TheHdw.Digital.ApplyLevelsTiming True, True, True, tlPowered
''''    End If
''''
''''    Dim Loop_Idx As Long
''''    Dim Loop_count As Long
''''    Dim Loop_Init As Long
''''    Dim Loop_Max As Long
''''    Dim Loop_Step As Long
''''    Dim Loop_BitNum As Long
''''    Dim Loop_RegName As String
''''    Dim SplitLoop_RegName() As String
''''    Dim Split_Loop_DigSrc_Str() As String
''''    Dim binstr As String
''''    Dim Loop_SplitByComma() As String
''''    Dim Loop_SplitByEqual() As String
''''
''''    Loop_Idx = 0
''''    Loop_Init = 0
''''    Loop_Max = 0
''''    Loop_Step = 1
''''
''''    If (Sweep_Enable = True) Then
''''        Loop_Max = Sweep_Info(0).Count - 1
''''
''''    End If
''''    Dim timer_ As Double
''''
''''    'timer_ = theexec.Timer()
'''''                                                                                                                                                                Dim VBT_LIB_HardIP_ProfileMark_1164 As Long: VBT_LIB_HardIP_ProfileMark_1164 = ProfileMarkEnter(2, instance_name & "_" & "ProsscessInputToGLB&Module=VBT_LIB_HardIP&ProcName=Meas_FreqVoltCurr_Universal_func&LineNumber=1158")   ' Profile Mark
''''
''''    Call ProcessInputToGLB(patset, TestSequence, CPUA_Flag_In_Pat, DisableComparePins, DisableConnectPins, DisableFRC, FRCPortName, MeasV_PinS, MeasF_PinS_SingleEnd, MeasF_Interval, MeasF_EventSourceWithTerminationMode, MeasF_Flag_MeasureThreshold, _
''''                            MeasF_ThresholdPercentage, MeasF_WaitTime, MeasI_pinS, MeasI_Range, MeasI_WaitTime, DigCap_Pin, DigCap_DataWidth, DigCap_Sample_Size, DigSrc_pin, DigSrc_DataWidth, DigSrc_Sample_Size, DigSrc_Equation, DigSrc_Assignment, _
''''                            DigSrc_FlowForLoopIntegerName, SpecialCalcValSetting, InstSpecialSetting, CUS_Str_MainProgram, CUS_Str_DigCapData, CUS_Str_DigSrcData, Flag_SingleLimit, TestLimitPerPin_VFI, MeasF_PinS_Differential, ForceFunctional_Flag, _
''''                            MeasF_WalkingStrobe_Flag, MeasF_WalkingStrobe_StartV, MeasF_WalkingStrobe_EndV, MeasF_WalkingStrobe_StepVoltage, MeasF_WalkingStrobe_BothVohVolDiffV, MeasF_WalkingStrobe_interval, MeasF_WalkingStrobe_miniFreq, Meas_StoreName, _
''''                            Calc_Eqn, Interpose_PrePat, Interpose_PreMeas, Interpose_PostTest, CharSetName, ForceV_Val, ForceI_Val, UVI80_MeasV_WaitTime, RAK_Flag, WaitTime_VIRZ)
'''''                                                                                                                                                                ProfileMarkLeave VBT_LIB_HardIP_ProfileMark_1164    ' Profile Mark
''''
''''    'theexec.Datalog.WriteComment "ProsscessInputToGLB Time : " & FormatNumber(theexec.Timer(timer_), 6) & ":" & theexec.DataManager.instanceName & ":" & TestSequence & ":" & CStr(DigSrc_Sample_Size) & ":" & DigSrc_Equation & ":" & DigSrc_Assignment
''''
''''    If InStr(UCase(CUS_Str_MainProgram), UCase("Loop_DigSrc")) <> 0 Then
''''
''''        'Ex: CUS_Str_MainProgram ==> Loop_DigSrc;1;119;10;32;ddr0_mdll0_lsw:ddr0_mdll1_lsw:ddr1_mdll0_lsw:ddr1_mdll1_lsw
''''        Split_Loop_DigSrc_Str = Split(CUS_Str_MainProgram, ";")
''''        Loop_Init = Split_Loop_DigSrc_Str(1)
''''        Loop_Max = Split_Loop_DigSrc_Str(2)
''''        Loop_Step = Split_Loop_DigSrc_Str(3)
''''        Loop_BitNum = Split_Loop_DigSrc_Str(4)
''''        Loop_RegName = Split_Loop_DigSrc_Str(5)
''''        SplitLoop_RegName = Split(Loop_RegName, ":")
''''
''''    End If
''''
''''    Dim loop_i As Long, Loop_j As Long
''''    Dim Temp_Equal_Str As String
''''    Dim Final_Comma_Str As String
''''    Temp_Equal_Str = ""
''''    Final_Comma_Str = ""
''''
''''    For Loop_count = Loop_Init To Loop_Max
''''
''''        If InStr(UCase(CUS_Str_MainProgram), UCase("Calc_Freq_SDLL_SWP")) <> 0 Then gl_Tname_Alg_Index = Loop_count
''''
''''        'TypeName (Loop_count)
''''
''''        If (Sweep_Enable = True) Then
''''            CUS_Str_DigCapData = Sweep_CUS_Str_DigCapData
''''
''''            Call SetForceSweepVoltAndTName(Sweep_Info, CUS_Str_DigCapData, Loop_count)
''''
''''            If InStr(UCase(theexec.DataManager.instancename), "MTRGR_T2P6") <> 0 Or InStr(UCase(theexec.DataManager.instancename), "MTRGR_T2P7") <> 0 Then
''''                Calc_Eqn = Replace(Calc_Eqn, Replace(Split(Split(Calc_Eqn, ":")(2), "(")(1), ")", ""), Split(CUS_Str_DigCapData, ":")(2))
''''                CUS_Str_DigCapData = Replace(CUS_Str_DigCapData, Split(CUS_Str_DigCapData, ":")(1), Split(CUS_Str_DigCapData, ":")(1) & CStr(Loop_count))
''''            Else
''''                Calc_Eqn = Sweep_Calc_Eqn & "," & CStr(Loop_count)
''''            End If
''''        End If
''''
''''
''''
''''        If CUS_Str_MainProgram <> "" And InStr(UCase(CUS_Str_MainProgram), UCase("Loop_DigSrc")) <> 0 Then
''''            binstr = Dec2BinStr32Bit_Rev(Loop_BitNum, Loop_count)
''''            Loop_SplitByComma = Split(DigSrc_Assignment, ";")
''''
''''            For loop_i = 0 To UBound(Loop_SplitByComma)
''''                Loop_SplitByEqual = Split(Loop_SplitByComma(loop_i), "=")
''''                For Loop_j = 0 To UBound(SplitLoop_RegName)
''''                    If UCase(Loop_SplitByEqual(0)) = UCase(SplitLoop_RegName(Loop_j)) Then
''''                        Loop_SplitByEqual(1) = binstr
''''                        Temp_Equal_Str = Loop_SplitByEqual(0) & "=" & Loop_SplitByEqual(1)
''''                        Exit For
''''                    Else
''''                        Temp_Equal_Str = Loop_SplitByEqual(0) & "=" & Loop_SplitByEqual(1)
''''                    End If
''''                Next Loop_j
''''                If loop_i = 0 Then
''''                    Final_Comma_Str = Temp_Equal_Str
''''                Else
''''                    Final_Comma_Str = Final_Comma_Str & ";" & Temp_Equal_Str
''''                End If
''''            Next loop_i
''''            DigSrc_Assignment = Final_Comma_Str
''''        End If
''''
''''
''''        '' 20190529 - Add for sweep force V
''''        If InStr(Interpose_PrePat, "x_sweep") <> 0 Then
''''            gl_Sweep_Glb_TName = CDbl(val(theexec.Flow.var("x_sweep").value)) / 1000
''''            Interpose_PrePat = Replace(Interpose_PrePat, "x_sweep", gl_Sweep_Glb_TName)
''''
''''            'USB_DP:V:x_sweep;Sweep_Name:
''''            If InStr(Interpose_PrePat, "Sweep_Name") <> 0 Then
''''                Interpose_PrePat = Replace(Interpose_PrePat, "Sweep_Name:", "")
''''            End If
''''        End If
''''
''''
''''
''''        '' 20190531 - Add for sweep Volt by shmoo
''''        If InStr(Interpose_PrePat, "Volt_sweep_GLB") <> 0 Then
''''        For Each site In theexec.sites
''''            gl_Sweep_Glb_TName = CDbl(theexec.Specs.Globals("Volt_sweep_GLB").CurrentValue)
''''            Exit For
''''        Next site
''''            Interpose_PrePat = Replace(Interpose_PrePat, "Volt_sweep_GLB", gl_Sweep_Glb_TName)
''''
''''
''''            If InStr(Interpose_PrePat, "Sweep_Name") <> 0 Then
''''                Interpose_PrePat = Replace(Interpose_PrePat, "Sweep_Name:", "")
''''            End If
''''        End If
''''
''''
''''
''''        '' 20160923 - Add Interpose_PrePat entry point
''''        If Interpose_PrePat <> "" Then
''''            Call SetForceCondition(Interpose_PrePat & ";STOREPREPAT")
''''        End If
''''
''''
''''        ''20161205 - Force_Flow_Shmoo_Condition
''''        If theexec.sites.item(site).SiteVariableValue("Flow_Shmoo_DevCharSetup") <> "" Then Force_Flow_Shmoo_Condition
''''        'Do Flow Shmoo
''''
''''        If patset.value <> "" Then
''''            Shmoo_Pattern = patset.value '' 20170808 add for shmoo pattern name print
''''            TheHdw.Patterns(patset).Load
''''            Call PATT_GetPatListFromPatternSet(patset.value, PattArray, PatCount)
''''        Else
''''            ReDim PattArray(0)
''''            PattArray(0) = ""
''''        End If
''''
''''        If (DisableConnectPins <> "") Then TheHdw.Digital.Pins(DisableConnectPins).Disconnect
''''        If (DisableComparePins <> "") Then TheHdw.Digital.Pins(DisableComparePins).DisableCompare = True
''''
''''        ''20161107-Return sweep test name
''''        Dim Rtn_SweepTestName As String
''''        Rtn_SweepTestName = ""
''''        gl_TName_Pat = patset.value
''''
''''        Dim current_pat_index As Integer
''''        current_pat_index = 0
''''
''''        For Each patt In PattArray
''''            If patt <> "" Then
''''                Pat = CStr(patt)
''''                TheHdw.Patterns(Pat).Load
'''''                                                                                                                                                                Dim VBT_LIB_HardIP_ProfileMark_1267 As Long: VBT_LIB_HardIP_ProfileMark_1267 = ProfileMarkEnter(2, instance_name & "_" & "GenDigSrc&Module=VBT_LIB_HardIP&ProcName=Meas_FreqVoltCurr_Universal_func&LineNumber=1264")    ' Profile Mark
''''                Set InDSPWave = Nothing
''''                Call GeneralDigSrcSetting(Pat, DigSrc_pin, DigSrc_Sample_Size, DigSrc_DataWidth, DigSrc_Equation, DigSrc_Assignment, _
''''                                                       DigSrc_FlowForLoopIntegerName, CUS_Str_DigSrcData, InDSPWave, Rtn_SweepTestName)
'''''                                                                                                                                                                ProfileMarkLeave VBT_LIB_HardIP_ProfileMark_1267    ' Profile Mark
''''
''''                Set OutDspWave = Nothing
''''                Call GeneralDigCapSetting(Pat, DigCap_Pin, DigCap_Sample_Size, OutDspWave)
''''
''''                Call PrintDigCapSetting(DigCap_Pin, DigCap_Sample_Size, CUS_Str_DigCapData)
''''
''''
''''                If InStr(UCase(CUS_Str_MainProgram), "MTR_UVI80_SETUP") <> 0 Then
''''                    Call MTR_UVI80_Setup
''''                End If
''''
''''
''''
''''                Dim SplitByCommaStr() As String
''''                Dim ForcePin_X As String
''''                Dim ForcePin_Y As String
''''                Dim SweepIndexStr_X As String
''''                Dim ForceVal_X As Double
''''                If LCase(CUS_Str_MainProgram) Like "*x_sweep*" Then
''''                    SplitByCommaStr = Split(CUS_Str_MainProgram, ",")
''''                    SweepIndexStr_X = SplitByCommaStr(0)
''''                         ForcePin_X = SplitByCommaStr(1)
''''
''''                          ForceVal_X = CDbl(val(theexec.Flow.var(SweepIndexStr_X).value)) / 1000
''''                          theexec.Datalog.WriteComment "ForcePin = " & ForcePin_X & "; ForceVal_X = " & ForceVal_X & "V"
''''                          'TheExec.Datalog.WriteComment "ForcePin = " & SplitByCommaStr(2) & ";  ForceVal_X  = " & ForceVal_X & "V"
''''                          'TheHdw.DCVS.Pins(ForcePin_X).Voltage.Value = ForceVal_X
''''                          'TheHdw.DCVS.Pins(SplitByCommaStr(2)).Voltage.Value = ForceVal_X
''''
''''                        TheHdw.Digital.Pins(ForcePin_X).Disconnect
''''
''''                            With TheHdw.PPMU.Pins(ForcePin_X)
''''                                .Gate = tlOff
''''                                .ForceI pc_Def_PPMU_InitialValue_FI, pc_Def_PPMU_Max_InitialValue_FI_Range
''''                                .ForceV CDbl(ForceVal_X), 0.02
''''                                .Connect
''''                                .Gate = tlOn
''''                            End With
''''
''''
''''                        FourceV = ForceVal_X
''''                End If
''''
''''
''''
''''                '' 20160713 - If no cpuflags in the test item modify the code to run pattern by using .test
''''                If (CPUA_Flag_In_Pat) Then
''''                    Call TheHdw.Patterns(Pat).start
''''                Else
''''                    Call TheHdw.Patterns(Pat).test(pfAlways, 0)
''''                End If
''''            End If
''''
''''            'TestSeqNum = 0
''''
''''            'Call ProcessTestNameInputString(OutputTname, UBound(TestSequenceArray))    Remove
''''
'''''            If PatCount > 1 Then
'''''                Dim ot_cnt As Long
'''''                For ot_cnt = 0 To UBound(OutputTname)
'''''                    OutputTname(ot_cnt) = OutputTname(ot_cnt) & Split(Split(Split(Pat, "\")(UBound(Split(Pat, "\"))), ":")(0), "_")(12)
'''''                Next ot_cnt
'''''            End If
''''
''''
''''            TestSeqNum = 0
''''
''''
''''            For Each Ts In TestSequenceArray
''''                Instance_Data.TestSeqNum = TestSeqNum
''''                ''20150907 - Only need CPUA_Flag_In_Pat to do control
''''                If (CPUA_Flag_In_Pat) Then
''''                    Call TheHdw.Digital.Patgen.FlagWait(cpuA, 0) 'Meas during CPUA loop
''''                Else
''''                    Call TheHdw.Digital.Patgen.HaltWait 'Pattern without CPUA loop
''''                End If
''''
''''                If InStr(MeasF_PinS_SingleEnd, "$") Then
''''                    Dim MeasF_Set() As String
''''                    MeasF_Set = Split(MeasF_PinS_SingleEnd, ",")
''''                End If
''''
''''
''''                If K_PreMeas_Setting = "" Then              '// Added by Kaino for special setting
''''                    ''20160923 - Add Interpose_PreMeas entry point by each sequence
''''                    If Interpose_PreMeas <> "" Then
''''                        If UBound(Interpose_PreMeas_Ary) = 0 Then
''''                            Call SetForceCondition(Interpose_PreMeas_Ary(0) & ";STOREPREMEAS")
''''                        ElseIf Interpose_PreMeas_Ary(TestSeqNum) <> "" Then
''''                            Call SetForceCondition(Interpose_PreMeas_Ary(TestSeqNum) & ";STOREPREMEAS")
''''                        End If
''''                    End If
''''                Else                                        '/*** ------ Added by Kaino for special setting ------ ***/
''''                    '/ *** --------------------------------------------------------------------------------------- ***
''''                    '/ ***    Before Meas
''''                    '/ *** --------------------------------------------------------------------------------------- ***
''''                    theexec.Datalog.WriteComment ""
''''                    theexec.Datalog.WriteComment "*** Test sequence : (" & TestSeqNum & ") special setting Before measurement ***"
''''
''''                    If TestSeqNum <= UBound(K_PreMeas_Setting_Sequence) Then
''''                        'TheExec.Datalog.WriteComment ""
''''                        'TheExec.Datalog.WriteComment " *** Setting: Before Measurement ***"
''''                        Call K_Set_SpecialSettingSequence(K_PreMeas_Setting_Sequence(TestSeqNum))
''''                    End If
''''                    'If TestSeqNum <= UBound(WaitTimeSeq) Then
''''                    '    MeasI_WaitTime = WaitTimeSeq(TestSeqNum)
''''                    'End If
''''                    '/ *** --------------------------------------------------------------------------------------- ***
''''                    '/ *** --------------------------------------------------------------------------------------- ***
''''                End If                                      '/*** ^^^^^^ Added by Kaino for special setting ^^^^^^ ***/
''''
''''                TestOptLen = Len(Ts)
''''
''''
''''
''''                For k = 1 To TestOptLen
''''                    Instance_Data.TestSeqSweepNum = k - 1
''''                    TestOption = Mid(Ts, k, 1)
''''
''''                    For Each site In theexec.sites.Active
''''                        testnum = theexec.sites.item(site).TestNumber
''''                    Next site
''''
''''                    '----------------0427 begin-------------------------------------
''''                    If InStr(MeasF_PinS_SingleEnd, "$") Then
''''                        MeasureF_Pin_SingleEnd = Replace(MeasF_Set(current_pat_index), "$", "")
''''                    End If
''''                    '----------------0427 end---------------------------------------
''''                    Select Case UCase(TestOption)
''''                        Case "V"
'''''                                                                                                                                                                Dim VBT_LIB_HardIP_ProfileMark_1347 As Long: VBT_LIB_HardIP_ProfileMark_1347 = ProfileMarkEnter(2, instance_name & "_" & "MeasV&Module=VBT_LIB_HardIP&ProcName=Meas_FreqVoltCurr_Universal_func&LineNumber=1345")    ' Profile Mark
''''
''''                            Call HardIP_MeasureVolt
'''''                                                                                                                                                                ProfileMarkLeave VBT_LIB_HardIP_ProfileMark_1347    ' Profile Mark
''''                        Case "F"
'''''                                                                                                                                                                Dim VBT_LIB_HardIP_ProfileMark_1352 As Long: VBT_LIB_HardIP_ProfileMark_1352 = ProfileMarkEnter(2, instance_name & "_" & "MeasF&Module=VBT_LIB_HardIP&ProcName=Meas_FreqVoltCurr_Universal_func&LineNumber=1350")    ' Profile Mark
''''
''''                            Call HardIP_MeasureFreq
'''''                                                                                                                                                                ProfileMarkLeave VBT_LIB_HardIP_ProfileMark_1352    ' Profile Mark
''''                        Case "I"
''''                            If DisableFRC = True Then FreeRunClk_Disable (FRCPortName)
'''''                                                                                                                                                                Dim VBT_LIB_HardIP_ProfileMark_1358 As Long: VBT_LIB_HardIP_ProfileMark_1358 = ProfileMarkEnter(2, instance_name & "_" & "MeasI&Module=VBT_LIB_HardIP&ProcName=Meas_FreqVoltCurr_Universal_func&LineNumber=1356")    ' Profile Mark
''''
''''                            Call HardIP_MeasureCurrent
'''''                                                                                                                                                                ProfileMarkLeave VBT_LIB_HardIP_ProfileMark_1358    ' Profile Mark
''''                        Case "R"
'''''                                                                                                                                                                Dim VBT_LIB_HardIP_ProfileMark_1363 As Long: VBT_LIB_HardIP_ProfileMark_1363 = ProfileMarkEnter(2, instance_name & "_" & "MeasR&Module=VBT_LIB_HardIP&ProcName=Meas_FreqVoltCurr_Universal_func&LineNumber=1361")    ' Profile Mark
''''
''''                            Call HardIP_SetupAndMeasureR
'''''                                                                                                                                                                ProfileMarkLeave VBT_LIB_HardIP_ProfileMark_1363    ' Profile Mark
''''                        Case "Z"
'''''                                                                                                                                                                Dim VBT_LIB_HardIP_ProfileMark_1368 As Long: VBT_LIB_HardIP_ProfileMark_1368 = ProfileMarkEnter(2, instance_name & "_" & "MeasZ&Module=VBT_LIB_HardIP&ProcName=Meas_FreqVoltCurr_Universal_func&LineNumber=1366")    ' Profile Mark
''''
''''                            Call HardIP_SetupAndMeasureZ
'''''                                                                                                                                                                ProfileMarkLeave VBT_LIB_HardIP_ProfileMark_1368    ' Profile Mark
''''                        Case "P"
''''                            HardIP_BySeqCurrentProfile
''''                        Case "N"
''''                            restore_Flag = True
''''                        Case Else
''''                            theexec.Datalog.WriteComment "Error Test Option, please select V,I or F"
''''                    End Select
''''                    If theexec.sites.Active.Count = 0 Then Exit Function
''''                Next k
''''
''''                ''20161206-Restore force condiction after measurement
''''    ''            Call SetForceCondition("RESTORE")
''''
''''                If K_PostMeas_Setting = "" Then             '/*** ------ Added by Kaino for special setting ------ ***/
''''                    If Interpose_PreMeas <> "" And Ts <> "N" Then
''''                        If UBound(Interpose_PreMeas_Ary) = 0 Then
''''                            Call SetForceCondition("RESTOREPREMEAS")
''''                        ElseIf Interpose_PreMeas_Ary(TestSeqNum) <> "" Then
''''                            Call SetForceCondition("RESTOREPREMEAS")
''''                        End If
''''                    End If
''''
''''                Else                                        '/*** ------ Added by Kaino for special setting ------ ***/
''''                    '/ *** --------------------------------------------------------------------------------------- ***
''''                    '/ ***    After_Meas
''''                    '/ *** --------------------------------------------------------------------------------------- ***
''''                    theexec.Datalog.WriteComment "*** Test sequence : (" & TestSeqNum & ") special setting After measurement ***"
''''
''''                    If TestSeqNum <= UBound(K_PostMeas_Setting_Sequence) Then
''''                        'TheExec.Datalog.WriteComment " *** Setting: Afer Measurement ***"
''''
''''                        Call K_Set_SpecialSettingSequence(K_PostMeas_Setting_Sequence(TestSeqNum))
''''
''''                    End If
''''                    '/ *** --------------------------------------------------------------------------------------- ***
''''                    '/ *** --------------------------------------------------------------------------------------- ***
''''
''''                End If                                      '/*** ^^^^^^ Added by Kaino for special setting ^^^^^^ ***/
''''
''''
''''                TestSeqNum = TestSeqNum + 1
''''
''''                If (CPUA_Flag_In_Pat) Then Call TheHdw.Digital.Patgen.Continue(0, cpuA)
''''                Instance_Data.TestSeqNum = TestSeqNum
''''
''''            Next Ts
''''
''''            If DebugPrintEnable = True Then theexec.Datalog.WriteComment "  Pattern(" & PatCount & "): " & Pat & ""
''''
''''            TheHdw.Digital.Patgen.HaltWait ' Haltwait at patten end
''''
''''            PatCount = PatCount + 1
''''
''''            '/*** ------ Added by Kaino for special setting ------ ***/
''''            If K_PatHalt_Setting <> "" Then
''''                theexec.Datalog.WriteComment "*** special setting After Pattern Halt  ***"
''''
''''                '/ *** --------------------------------------------------------------------------------------- ***
''''                '/ ***    Pattern Halt
''''                '/ *** --------------------------------------------------------------------------------------- ***
''''
''''
''''                If K_PatHalt_Setting <> "" Then
''''                    'TheExec.Datalog.WriteComment " *** Setting: Pattern Halt ***"
''''                    Call K_Set_SpecialSettingSequence(K_PatHalt_Setting)
''''                End If
''''                '/ *** --------------------------------------------------------------------------------------- *** /
''''            End If
''''            '/*** ^^^^^^ Added by Kaino for special setting ^^^^^^ ***/
''''
''''            '' 20160923 - Add Interpose_PostTest entry point
''''            Call SetForceCondition(Interpose_PostTest)
''''
'''''            If gl_FlowForLoop_DigSrc_SweepCode <> "" Then         '20180509
'''''                gl_FlowForLoop_DigSrc_SweepCode = ""
'''''            End If
''''
''''            '' 20160211 - Process DigCapData by using DSP
''''    ''        If b_ProcessDigCapByDSP = True Then
''''                If DigCap_Sample_Size <> 0 Then
''''                    Dim DigCapPinAry() As String, NumberPins As Long
''''                    Dim CUS_Str_DigCapData_temp As String
''''                    Call theexec.DataManager.DecomposePinList(DigCap_Pin, DigCapPinAry(), NumberPins)
''''
''''
''''                    If NumberPins > 1 Then
''''                        Call CreateSimulateDataDSPWave_Parallel(OutDspWave, DigCap_Sample_Size)
''''                        Call Checker_StoreDigCapAllToDictionary(CUS_Str_DigCapData, OutDspWave, NumberPins)
''''                        Call DigCapDataProcessByDSP_Parallel(CUS_Str_DigCapData, OutDspWave, DigCap_Sample_Size, NumberPins, , DigCap_Pin.value)
''''
''''                    ElseIf NumberPins = 1 Then
''''                        Call CreateSimulateDataDSPWave(OutDspWave, DigCap_Sample_Size, DigCap_DataWidth)
''''                        Call Checker_StoreDigCapAllToDictionary(CUS_Str_DigCapData, OutDspWave, NumberPins)
'''''                                                                                                                                                                Dim VBT_LIB_HardIP_ProfileMark_1429 As Long: VBT_LIB_HardIP_ProfileMark_1429 = ProfileMarkEnter(2, instance_name & "_" & "MeasC&Module=VBT_LIB_HardIP&ProcName=Meas_FreqVoltCurr_Universal_func&LineNumber=1427")    ' Profile Mark
''''
''''                        Call DigCapDataProcessByDSP(CUS_Str_DigCapData, OutDspWave, DigCap_Sample_Size, DigCap_DataWidth, CUS_Str_MainProgram, , DigCap_Pin.value)
'''''                                                                                                                                                                ProfileMarkLeave VBT_LIB_HardIP_ProfileMark_1429    ' Profile Mark
''''                    End If
''''                End If
''''
''''                '' 20160907 - Process calculate equation by dictionary.
''''                If Calc_Eqn <> "" And InStr(LCase(TestSequence), "p") = 0 Then
''''                    Call ProcessCalcEquation(Calc_Eqn)
''''                End If
''''
''''                '' 20160713 - Call write functional result if cpu flag in pattern
''''                If (CPUA_Flag_In_Pat) Then
''''                    Call HardIP_WriteFuncResult(, , Inst_Name_Str)
''''                End If
'''''Start - Restore this condition - Skua use
''''                If gl_FlowForLoop_DigSrc_SweepCode <> "" Then        '20180814
''''                    gl_FlowForLoop_DigSrc_SweepCode = ""
''''                End If
'''''End - Restore this condition - Skua use
''''    ''        End If
''''
''''            current_pat_index = current_pat_index + 1
''''
''''                If Interpose_PreMeas <> "" And restore_Flag = True Then
''''                    Call SetForceCondition("RESTOREPREMEAS")
''''
''''                End If
''''
''''
''''
''''               If LCase(CUS_Str_MainProgram) Like "*x_sweep*" Then
''''
''''                    With TheHdw.PPMU.Pins(ForcePin_X)
''''                            .ForceV pc_Def_PPMU_InitialValue_FV, pc_Def_PPMU_Max_InitialValue_FI_Range ''FVMI - Carter, 20190503
''''                            .Disconnect
''''                            .Gate = tlOff
''''                    End With
''''                    TheHdw.Digital.Pins(ForcePin_X).Connect ''Connect Digital pins after measurement - Carter, 20190503
''''
''''                End If
''''
''''
''''            gl_Sweep_Glb_TName = "" '' 20190529 - Add for sweep force V
''''
''''        Next patt
''''
''''''        ''20170405-Record all functional test result from flow for loop opcode, use global string to store them
''''        If CUS_Str_DigSrcData <> "" And UCase(CUS_Str_DigSrcData) = UCase("BinToGray") Then
''''            If CPUA_Flag_In_Pat = False Then
''''                Call DisplayForLoopFuncResult_EndOfTest(CUS_Str_DigSrcData, Rtn_SweepTestName, CPUA_Flag_In_Pat, DigSrc_FlowForLoopIntegerName)
''''            End If
''''        End If
''''     If MeasureV_pin <> "" Then
''''         Call EndSetupForMeasureVoltPins(MeasureV_Pin_PPMU, MeasureV_Pin_UVI80)
''''     End If
''''
''''     If DisableConnectPins <> "" Then TheHdw.Digital.Pins(DisableConnectPins).Connect
''''     If DisableComparePins <> "" Then TheHdw.Digital.Pins(DisableComparePins).DisableCompare = False
''''
''''     If DisableFRC = True Then
''''         Call ReStart_FRC(FRCPortName)
''''     End If
''''
''''     DebugPrintFunc patset.value  ' print all debug information
''''
''''     If theexec.sites.item(site).SiteVariableValue("Flow_Shmoo_DevCharSetup") <> "" Then
''''     'Do Flow Shmoo
''''         If Flow_Shmoo_Port_Name <> "" Then Restart_All_Freerun_Clk
''''     End If
''''
''''     If Interpose_PrePat <> "" Then
''''         Call SetForceCondition("RESTOREPREPAT")
''''     End If
''''
''''     ''=============================== CharSetName ====================================
''''     Dim p As Variant
''''     If theexec.DevChar.Setups.IsRunning = False And CharSetName <> "" Then
''''         Dim ApplyPins As String, Setup_mode As String, p_ary() As String, p_cnt As Long
''''         'If TheExec.DevChar.Setups(CharSetName).TestMethod.Value = tlDevCharTestMethod_Reburst Then TheExec.Datalog.WriteComment "[PrintCharCondition:" & PrintCharSetup(Interpose_PrePat_GLB) & ",Test]"
''''         Setup_mode = theexec.DevChar.Setups(CharSetName).Shmoo.axes(tlDevCharShmooAxis_X).Parameter.name
''''         If (LCase(Setup_mode) <> "vid" And LCase(Setup_mode) <> "vicm") Then
''''             ApplyPins = theexec.DevChar.Setups(CharSetName).Shmoo.axes(tlDevCharShmooAxis_X).ApplyTo.Pins
''''             theexec.DataManager.DecomposePinList ApplyPins, p_ary, p_cnt
''''             For Each p In p_ary
''''                 theexec.DevChar.Setups(CharSetName).Shmoo.axes(tlDevCharShmooAxis_X).ApplyTo.Pins = p
''''                 run_shmoo CharSetName
''''             Next p
''''             theexec.DevChar.Setups(CharSetName).Shmoo.axes(tlDevCharShmooAxis_X).ApplyTo.Pins = ApplyPins
''''         Else
''''             run_shmoo CharSetName
''''         End If
''''     End If
''''
''''    If CUS_Str_MainProgram <> "" And InStr(UCase(CUS_Str_MainProgram), UCase("Loop_DigSrc")) <> 0 And Loop_Step <> 1 Then
''''        Loop_count = Loop_count + Loop_Step - 1
''''    End If
''''
''''    Next Loop_count
''''    ''================================================================================
''''
''''
''''    ReDim TestConditionSeqData(0)
''''    Dim Instance_Data_temp() As Instance_Type
''''    ReDim Instance_Data_temp(0)
''''    Instance_Data = Instance_Data_temp(0)
''''
''''    Instance_Data.Meas_StoreName_Flag = False ''Carter, 20190521
''''    Exit Function
''''
''''errHandler:
''''    theexec.Datalog.WriteComment "error in Meas_FreqVoltCurr_Universal_func"
'''''    Resume Next
''''    If AbortTest Then Exit Function Else Resume Next
''''
''''End Function





Public Function K_Str_DumpDspWave(InDSPWave As DSPWave, Optional LSB As Integer = 0, Optional Delimiter As String = vbNullString)
On Error GoTo errHandler 'Add ErrHandler 2023/05/29

    Dim binstr As String
    Dim i As Integer
    
    binstr = vbNullString
    
    
    If LSB = 0 Then
        For i = 0 To InDSPWave.DataSize - 1
            binstr = binstr & InDSPWave.Element(i)
        Next i
    Else
         For i = InDSPWave.DataSize - 1 To 0
            binstr = binstr & InDSPWave.Element(i)
        Next i
    
    End If
    
    K_Str_DumpDspWave = binstr

Exit Function 'Add ErrHandler 2023/05/29
errHandler: 'Add ErrHandler 2023/05/29
    Call Print_Error_Message(Error_Info, "VBT_LIB_K_ADDRIO_AP", "K_Str_DumpDspWave") 'Add ErrHandler 2023/05/29
    If AbortTest Then Exit Function Else Resume Next 'Add ErrHandler 2023/05/29
End Function



Public Function CZ_Style_TName_InstanceInfo_Clear()
On Error GoTo errHandler 'Add ErrHandler 2023/05/29
        ' V0: 2017-06-02 by Kaino

        Dim i As Integer

        K_PatSetName = vbNullString
        K_InstanceName = vbNullString
        K_InstanceName_WO_Pset = vbNullString


        TNameSeg(0) = "HAC"
        TNameSeg(1) = "Meas?"
        TNameSeg(2) = "x"                                                   '[HV/NV/LV]
        TNameSeg(3) = "x"                                                   '[X1] : sub-block-name-1
        TNameSeg(4) = "x"                                                   '[Block]
        TNameSeg(5) = "{pinname}"
        TNameSeg(6) = "x"                                                   '[X2] : sub-block-name-2
        TNameSeg(7) = "x"                                                   '[X3] : X3 / DSSC Segment name
        TNameSeg(8) = "x"                                                   '[X4] :    / DSSC Register
        TNameSeg(9) = "x"                                                   '[X5] : subr-seq#



    For i = 0 To UBound(InstNameSegs)
        InstNameSegs(i) = vbNullString
    Next i

Exit Function 'Add ErrHandler 2023/05/29
errHandler: 'Add ErrHandler 2023/05/29
    Call Print_Error_Message(Error_Info, "VBT_LIB_K_ADDRIO_AP", "CZ_Style_TName_InstanceInfo_Clear") 'Add ErrHandler 2023/05/29
    If AbortTest Then Exit Function Else Resume Next 'Add ErrHandler 2023/05/29
End Function



