Attribute VB_Name = "LIB_HardIP_Calc"
Option Explicit
Public R_Path_PLD As New PinListData
Public CMRR_Values(10) As New SiteDouble
'ReDim CMRR_Value(10) As New SiteDouble
Public glb_ind As Long
'Public gl_Save_deltavalue As New DSPWave
'Public gl_Save_MeasNum As New DSPWave
'Public gl_Save_DCK As New DSPWave

Type Type_MonoWithBlock
    Block As Long
    DSP_Bin As New DSPWave
    DSP_Dec As New DSPWave
End Type

Public Const CalcStr = "Calc"
Public Const Calc = "Calc"
Public Const CalcV = "CalcV" '"V"
Public Const CalcI = "CalcI" '"I"
Public Const CalcF = "CalcF" '"F"
Public Const CalcR = "CalcR" '"R"
Public Const CalcC = "CalcC" '"C"
Public Const CalcT = "CalcT" '"T"

'Public meas_val_before() As New SiteDouble
Public meas_val_before As New PinListData
Public meas_val_delay_instance_name As String
Public meas_val_first(10) As New SiteDouble
Public meas_val_delay_instance As String

Public Function Calc_ADC_LDO_TRIM_TTR(argc As Integer, argv() As String) As Long
On Error GoTo errHandler 'Add ErrHandler 2023/05/29

'Alg::Calc_ADC_LDO_TRIM_TTR(SOC-ADC-ED-REG-OFFSET+ACCP0-ADC-ED-REG-OFFSET+ACCP1-ADC-ED-REG-OFFSET+ACCE0-ADC-ED-REG-OFFSET,ADC_SOC_LDO_105C,15,4)

Dim InWf_Split() As String: InWf_Split = Split(argv(0), "+")
Dim Volt_MeasDic() As New SiteDouble: ReDim Volt_MeasDic(UBound(Split(argv(0), "+")))
Dim i, z As Long
Dim Sweepsize As String
Sweepsize = argv(2)
Dim Ycurve_array() As Double: ReDim Ycurve_array(Sweepsize - 1)
'Dim Sweepsize As String
Dim ADCcalc As New DSPWave
Dim PositiveValue As New DSPWave
Dim Minima_value As String
Dim Target_Step As New DSPWave
Dim ADCLDO_fusename As String
Dim ADCLDO_Fusearray() As String
Dim Target_Step_Binarry As New DSPWave
Dim Temp_target() As New DSPWave

Dim Fuse_Digsrc_size As String
Dim Trim_code() As New SiteLong: ReDim Trim_code(UBound(Split(argv(0), "+")))
Dim TestNameInput As String


'Temp_target.CreateConstant 0, UBound(InWf_Split) + 1, DspLong
Target_Step.CreateConstant 0, UBound(InWf_Split) + 1, DspLong
'Sweepsize = argv(2)
ADCLDO_fusename = argv(1)
ADCLDO_Fusearray = Split(ADCLDO_fusename, "+")
Fuse_Digsrc_size = argv(3)

    For i = 0 To UBound(InWf_Split)
        Volt_MeasDic(i) = GetStoreDataAllType(InWf_Split(i))
    Next i
    
    
    For i = 1 To Sweepsize
        Ycurve_array(i - 1) = (4.9 * 10 ^ -3 * i) + 0.953
    Next

    ADCcalc.CreateConstant 0, Sweepsize, DspDouble

    For Each site In TheExec.sites.Active
    
        For z = 0 To UBound(InWf_Split)
            TheExec.Datalog.WriteComment "Start Find Trim Value For " & InWf_Split(z)
            For i = 0 To UBound(Ycurve_array)
                ADCcalc.Element(i) = Ycurve_array(i) * Volt_MeasDic(z) - 1
            Next i
'-------------For minima positive value -----------------------
''''            Count = 0
''''            For i = 0 To ADCcalc.SampleSize - 1
''''                If ADCcalc.Element(i) >= 0 Then
''''                    Count = Count + 1
''''                End If
''''            Next i
''''
''''            PositiveValue.CreateConstant 0, Count, DspDouble
''''
''''            Count = 0
''''            For i = 0 To UBound(Ycurve_array)
''''                If ADCcalc.Element(i) >= 0 Then
''''                    PositiveValue.Element(Count) = ADCcalc.Element(i)
''''                    Count = Count + 1
''''                End If
''''            Next i
'-----------------------------------------------------------------
           
            PositiveValue = ADCcalc.Abs
            Minima_value = PositiveValue.CalcMinimumValue
            
             For i = 0 To ADCcalc.SampleSize - 1
                If PositiveValue.Element(i) = Minima_value Then
                   TheExec.Datalog.WriteComment "site[" & site & "]" & " trim_code" & i & ": " & ADCcalc.Element(i) & "<-------Abs Value Closet to zero Value"
'                   If i = 0 Thensetx
                    Target_Step.Element(z) = i     'For trim value is 0 will error
                    Trim_code(z) = Target_Step.Element(z)
'                   Else
'                    Target_Step.Element(z) = i - 1
'                    Trim_code(z) = Target_Step.Element(z)
'                   End If
                Else
                    TheExec.Datalog.WriteComment "site[" & site & "]" & "trim_code" & i & ": " & ADCcalc.Element(i)
                End If
                
            Next i
            
        Next z
    
    Next site
'''    For Each site In TheExec.sites.Active
'''        Target_Step.ConvertDataTypeTo (DspLong)
'''        Target_Step_Binarry = Target_Step.ConvertStreamTo(tldspSerial, Fuse_Digsrc_size, 0, Bit0IsMsb)
'''        Call StoreDataAllType(ADCLDO_Fusearray, Target_Step_Binarry)
'''     Next site
'''

    ReDim Temp_target(UBound(ADCLDO_Fusearray)) As New DSPWave

    For Each site In TheExec.sites.Active
            Target_Step.ConvertDataTypeTo (DspLong)
            Target_Step_Binarry = Target_Step.ConvertStreamTo(tldspSerial, Fuse_Digsrc_size, 0, Bit0IsMsb)
    
        For i = 0 To UBound(ADCLDO_Fusearray)
''            Target_Step(site).ConvertDataTypeTo (DspLong)
''            Target_Step_Binarry = Target_Step(site).ConvertStreamTo(tldspSerial, Instance_Data.DigSrc_DataWidth, 0, Bit0IsMsb)
            'Target_Step_Binarry = Target_Step.Element(i).ConvertStreamTo(tldspSerial, Fuse_Digsrc_size, 0, Bit0IsMsb)
            Temp_target(i) = Target_Step_Binarry.Select(i * Fuse_Digsrc_size, 1, Fuse_Digsrc_size).COPY
            
            Call StoreDataAllType(ADCLDO_Fusearray(i), Temp_target(i))
        Next i

    Next site
    
'    For Each site In TheExec.sites.Active
        For z = 0 To UBound(InWf_Split)
           TestNameInput = Report_TName_From_Instance(CalcC, InWf_Split(z), , , z)
           TheExec.flow.TestLimit resultVal:=Trim_code(z), Tname:=TestNameInput, ForceResults:=tlForceFlow
        Next z
'     Next site
Exit Function 'Add ErrHandler 2023/05/29
errHandler: 'Add ErrHandler 2023/05/29
        Call Print_Error_Message(Error_Info, "LIB_HardIP_Calc", "Calc_ADC_LDO_TRIM_TTR") 'Add ErrHandler 2023/05/29
        If AbortTest Then Exit Function Else Resume Next 'Add ErrHandler 2023/05/29
 
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
        meas_val = GetStoreDataAllType(meas_name)
    
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
                
                TheExec.flow.TestLimit resultVal:=result, Tname:=TestNameInput, ForceResults:=tlForceFlow, scaletype:=scalePico
                
            End If
        
    Next i
    meas_val_delay_instance = TheExec.DataManager.instancename
    
End Function
Public Function Calc_delay_Sicily(argc As Integer, argv() As String) As Long
On Error GoTo errHandler 'Add ErrHandler 2023/05/29

    Dim i As Long
    Dim TestNameInput As String
    Dim Dict_Name As String
    Dim site As Variant
    Dim meas_val_now As New PinListData
    Dim meas_val As New PinListData

    Set meas_val = Nothing
    Dict_Name = argv(i)
    meas_val = GetStoreDataAllType(Dict_Name)

    
    If meas_val_delay_instance_name <> TheExec.DataManager.instancename Then 'first time to enter this function
        Set meas_val_before = Nothing
        meas_val_before = meas_val
    Else
        Set meas_val_now = Nothing
        meas_val_now = meas_val
        '=================prevent divide 0==============
        For Each site In TheExec.sites
            For i = 0 To meas_val_now.pins.Count - 1
                If meas_val_now.pins(i).value = 0 Then
                    meas_val_now.pins(i).value = 0.0000000001
                End If
            Next i
            For i = 0 To meas_val_before.pins.Count - 1
                If meas_val_before.pins(i).value = 0 Then
                    meas_val_before.pins(i).value = 0.0000000001
                End If
            Next i
        Next site
        '=================prevent divide 0==============
        meas_val_now = meas_val_now.Math.Invert.Subtract(meas_val_before.Math.Invert).Multiply(0.5)
        Dim PLD_For_TestLimit As New PinListData
        For i = 0 To meas_val_now.pins.Count - 1
            If UCase(meas_val_now.pins(i)) Like "*DQS_P*" Then
                PLD_For_TestLimit.AddPin (meas_val_now.pins(i))
                For Each site In TheExec.sites
                    PLD_For_TestLimit.pins(meas_val_now.pins(i)).value = meas_val_now.pins(i).value
                Next site
            End If
        Next i
        TestNameInput = Report_TName_From_Instance(CalcC, vbNullString)
        TheExec.flow.TestLimit resultVal:=PLD_For_TestLimit, Tname:=TestNameInput, ForceResults:=tlForceFlow, scaletype:=scalePico
        meas_val_before = meas_val
    End If

    meas_val_delay_instance_name = TheExec.DataManager.instancename
Exit Function 'Add ErrHandler 2023/05/29
errHandler: 'Add ErrHandler 2023/05/29
    Call Print_Error_Message(Error_Info, "LIB_HardIP_Calc", "Calc_delay_Sicily") 'Add ErrHandler 2023/05/29
    If AbortTest Then Exit Function Else Resume Next 'Add ErrHandler 2023/05/29

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
        meas_val = GetStoreDataAllType(meas_name)
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
        DSPWave_Dict = GetStoreDataAllType(argv(i))
        TestNameInput = TestNameInput & argv(i)
        Call rundsp.Transfer2GrayCode(DSPWave_Dict, DSPWave_GrayCode, DSPWave_GrayCodeDec)

        TestNameInput = Report_TName_From_Instance(CalcC, "X", "GrayCode", CInt(i))
        TheExec.flow.TestLimit resultVal:=DSPWave_GrayCodeDec.Element(0), Tname:=TestNameInput, ForceResults:=tlForceFlow
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
        'CMRR_Value = GetStoreDataAllType(argv(0))
        Voltage_Value = TheExec.Specs.DC.item(argv(1)).CurrentValue(site)
        CMRR_Value = GetStoreDataAllType(argv(0) + "_para")
        
        CMRR_Value = CMRR_Value * 1.25 / (2 ^ 17)
        
        OutputTname_format = Split(TestNameInput, "_")
        OutputTname_format(6) = "CMRR"
        OutputTname_format(7) = CStr(GetStoreDataAllType(argv(0) + "_para"))
        OutputTname_format(8) = Replace(CStr(TheExec.Specs.DC.item(argv(1)).CurrentValue(site)), ".", "p")
        TestNameInput = Merge_TName(OutputTname_format)
        CMRR_Value = CMRR_Value / Voltage_Value
        
    Next
    
    TheExec.flow.TestLimit resultVal:=CMRR_Value, Tname:=TestNameInput, ForceResults:=tlForceFlow

''    For i = 0 To argc - 1
''        For Each Site In TheExec.sites
''            DSPWave_Dict = GetStoreDataAllType(argv(i))
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
        'CMRR_Value = GetStoreDataAllType(argv(0))
        Voltage_Value = TheExec.Specs.DC.item(argv(1)).CurrentValue(site)
        PSRR_Value = GetStoreDataAllType(argv(0) + "_para")
        
        PSRR_Value = PSRR_Value * 1.25 / (2 ^ 17)
        
        OutputTname_format = Split(TestNameInput, "_")
        OutputTname_format(6) = "PSRR"
        OutputTname_format(7) = CStr(GetStoreDataAllType(argv(0) + "_para"))
        OutputTname_format(8) = Replace(CStr(TheExec.Specs.DC.item(argv(1)).CurrentValue(site)), ".", "p")
        TestNameInput = Merge_TName(OutputTname_format)
        OutputTname_format(6) = "VDDIO12_MTR_GR"
        TestNameInput1 = Merge_TName(OutputTname_format)
        'PSRR_Value = PSRR_Value / Voltage_Value
        PSRR_Value = PSRR_Value.power(-1).Multiply(0.2).Log10.Multiply(20)
    Next
    
    TheExec.flow.TestLimit resultVal:=Voltage_Value, Tname:=TestNameInput1, ForceResults:=tlForceFlow
    TheExec.flow.TestLimit resultVal:=PSRR_Value, Tname:=TestNameInput, ForceResults:=tlForceFlow

''    For i = 0 To argc - 1
''        For Each Site In TheExec.sites
''            DSPWave_Dict = GetStoreDataAllType(argv(i))
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
    LSW_dspwave = GetStoreDataAllType(InputKey)
    InputKey = LCase(argv(2))
    MSW_dspwave = GetStoreDataAllType(InputKey)
    
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
    LSW_dspwave = GetStoreDataAllType(InputKey)
    InputKey = LCase(argv(4))
    MSW_dspwave = GetStoreDataAllType(InputKey)
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
    
    'TheExec.Flow.TestLimit resultVal:=dqs0rx_sweep, Tname:="Number_of_First_Continuous_One_DQS0RX", ForceResults:=tlForceNone
    TheExec.flow.TestLimit resultVal:=dqs0rx_sweep, Tname:="Number_of_Max_Continuous_One_DQS0RX", ForceResults:=tlForceFlow 'transfer_to_forceflow
    
    'dqs0rx_sweep * step_size > mdll0 / 2
    
    ReportVal = dqs0rx_sweep.Multiply(Step_Size)
    
    TheExec.Datalog.WriteComment " *** DQS0RX_Sweep x Step_Size ( " & Step_Size & " ) ***"
    
    For Each site In TheExec.sites
        LoVal = mdll0
        
        If ReportVal = 0 Then ReportVal = -1        ' update by Kaino on 2017/09/20
        
        'Report_TestLimit_by_CZ_Format resultVal:=ReportVal, lowVal:=Str(LoVal), MeasType:="C", UserVar5:="EYEDQS0", scaletype:=scaleNoScaling
        TestNameInput = Report_TName_From_Instance(CalcC, "X", "EYEDQS0", 0, , , , , tlForceFlow) 'transfer_to_forceflow
        TheExec.flow.TestLimit resultVal:=ReportVal, lowVal:=str(LoVal), Tname:=TestNameInput, ForceResults:=tlForceFlow 'transfer_to_forceflow
    Next site
    
    
    
    '/*** --------------------------------------------- ***/
    '/*** ------------------- MDLL1 ------------------- ***/
    '/*** --------------------------------------------- ***/
    
    InputKey = LCase(argv(5))
    LSW_dspwave = GetStoreDataAllType(InputKey)
    InputKey = LCase(argv(6))
    MSW_dspwave = GetStoreDataAllType(InputKey)
    
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
    LSW_dspwave = GetStoreDataAllType(InputKey)
    InputKey = LCase(argv(8))
    MSW_dspwave = GetStoreDataAllType(InputKey)
    
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
    
    'TheExec.Flow.TestLimit resultVal:=dqs1rx_sweep, Tname:="Number_of_First_Continuous_One_DQS1RX", ForceResults:=tlForceNone
    TheExec.flow.TestLimit resultVal:=dqs1rx_sweep, Tname:="Number_of_Max_Continuous_One_DQS1RX", ForceResults:=tlForceFlow 'transfer_to_forceflow
    
    
    
    'dqs1rx_sweep * step_size > mdll1 / 2
    
    ReportVal = dqs1rx_sweep.Multiply(Step_Size)
    
    TheExec.Datalog.WriteComment " *** DQS1RX_Sweep x Step_Size ( " & Step_Size & " ) ***"
    
    For Each site In TheExec.sites
        LoVal = mdll1
        
        If ReportVal = 0 Then ReportVal = -1        ' update by Kaino on 2017/09/20
        
        'Report_TestLimit_by_CZ_Format resultVal:=ReportVal, lowVal:=Str(LoVal), MeasType:="C", UserVar5:="EYEDQS1", scaletype:=scaleNoScaling
        TestNameInput = Report_TName_From_Instance(CalcC, "X", "EYEDQS1", 0, , , , , tlForceFlow) 'transfer_to_forceflow
        TheExec.flow.TestLimit resultVal:=ReportVal, lowVal:=str(LoVal), Tname:=TestNameInput, ForceResults:=tlForceFlow 'transfer_to_forceflow
   
    Next site
    
End Function

Public Function Calc_R_Path_Cal(argc As Integer, argv() As String) As Long
On Error GoTo errHandler 'Add ErrHandler 2023/05/29
    
    '' NOTE :
    '' argv(0) : I1 Dictionary KeyName
    '' argv(1) : I2 Dictionary KeyName
    '' argv(2) : I3 Dictionary KeyName
    '' argv(3) : Force Condition Equation Ex: VDDQL/2 => Evaluate(=_VDDQL_VAR_H/2)

    Dim site As Variant
    Dim Meas_I1_PLD As New PinListData: Meas_I1_PLD = GetStoreDataAllType(argv(0))
    Dim Meas_I2_PLD As New PinListData: Meas_I2_PLD = GetStoreDataAllType(argv(1))
    Dim Meas_I3_PLD As New PinListData: Meas_I3_PLD = GetStoreDataAllType(argv(2))
    Dim Force_Cond_Eq As String: Force_Cond_Eq = argv(3)
    Dim Cust_Str As String
    
    If UBound(argv) = 4 Then
        Cust_Str = argv(4)
    End If

    Dim R_Contact_PLD As New PinListData
    Dim Rak_val() As Double
    Dim Total_RAK_Val As Double
    Dim Split_Name() As String
    Dim Force_Cond As Double
    Dim AddPin_Flag As Integer: AddPin_Flag = 1                                         'Flag for Pinlist global variable add pin
    Dim PinName As Variant
    Dim PinName_Glb_PLD As Variant
    Dim ForceCond_str As String
    Dim Ary_str(0) As String
    Dim DDR_R1 As New PinListData   ' DDR Test R1 Variable
    Dim DDR_R2 As New PinListData   ' DDR Test R2 Variable
    Dim OutputTname_format() As String
    Dim TestNameInput As String
    
    
    'Force_Cond_Eq = VDDQL_DDR0/2 ; Split_Name(0) is PowerName
    Split_Name = Split(Force_Cond_Eq, "/")
    ForceCond_str = "_" + Split_Name(0) + "_VAR" + "/" + Split_Name(1)
    Ary_str(0) = ForceCond_str
    Call HIP_Evaluate_ForceVal(Ary_str)
    Force_Cond = CDbl(Ary_str(0))
    
    For Each PinName In Meas_I1_PLD.pins
        
        'If PinName is not exist then add new one to Global PinListData
        For Each PinName_Glb_PLD In R_Path_PLD.pins
            If LCase(PinName_Glb_PLD) = LCase(PinName) Then
                'AddPin_Flag = 0 : no need to add pin since pin has already available
                AddPin_Flag = 0
                Exit For
            End If
        Next PinName_Glb_PLD
        
        If AddPin_Flag = 1 Then
            R_Path_PLD.AddPin (CStr(PinName))
        End If
        
        'Add PinName to Local PinListData
        R_Contact_PLD.AddPin (CStr(PinName))
        
        'Customize String for DDR Test only
        If Cust_Str = UCase("DDR_TEST") Then
            DDR_R1.AddPin (CStr(PinName))
            DDR_R2.AddPin (CStr(PinName))
        End If
        
        For Each site In TheExec.sites.Active
            
            Dim i1 As Double: i1 = Meas_I1_PLD.pins(PinName).value(site)
            Dim i2 As Double: i2 = Meas_I2_PLD.pins(PinName).value(site)
            Dim i3 As Double: i3 = Meas_I3_PLD.pins(PinName).value(site)
            Dim I3_I1 As Double: I3_I1 = Meas_I3_PLD.pins(PinName).value(site) - Meas_I1_PLD.pins(PinName).value(site)
            Dim I3_I2 As Double: I3_I2 = Meas_I3_PLD.pins(PinName).value(site) - Meas_I2_PLD.pins(PinName).value(site)
            
            'Initialize the Value on PinListData to prevent any cross usage between samples
            R_Path_PLD.pins(PinName).value(site) = 0
            R_Contact_PLD.pins(PinName).value(site) = 0
            
            'RAK_Val = TheHdw.PPMU.ReadRakValuesByPinnames(PinName, site)
       
''            If InStr(UCase(TheExec.CurrentChanMap), "FT") <> 0 Then
''                Total_RAK_Val = RAK_Val(0) + FT_Card_RAK.Pins(PinName).Value(Site)
''            Else
''                Total_RAK_Val = RAK_Val(0) + CP_Card_RAK.Pins(PinName).Value(Site)
''            End If
            
            
            Total_RAK_Val = CurrentJob_Card_RAK.pins(PinName).value(site)     ' Edited by Dylan 2019/12/04 (For debug running)
           
            If i1 <> 0 And i2 <> 0 And i3 <> 0 Then
                If I3_I1 > 0 And I3_I2 > 0 And (i1 * i2) > 0 Then
                    R_Path_PLD.pins(PinName).value(site) = (Force_Cond / i3) * (1 - ((I3_I1 * I3_I2) / (i1 * i2)) ^ 0.5)
                Else
                    R_Path_PLD.pins(PinName).value(site) = 999 ' report R= 999 when divide by 0
                    TheExec.Datalog.WriteComment (" Error : PinName " & CStr(PinName) & " , Site" & CStr(site) & " I3 should greater than I1,I2 And (I1*I2) should greater than 0!  ")
                End If
            Else
                R_Path_PLD.pins(PinName).value(site) = 999
                TheExec.Datalog.WriteComment (" Error : PinName " & CStr(PinName) & " , Site" & CStr(site) & " Division by Zero !   ")
            End If
            
            R_Contact_PLD.pins(PinName).value(site) = R_Path_PLD.pins(PinName).value(site) - Total_RAK_Val
                        
            'Customize String for DDR Test only
            If Cust_Str = UCase("DDR_TEST") Then
                If i1 <> 0 And i2 <> 0 Then
                    DDR_R1.pins(PinName).value(site) = (1 * Force_Cond / i1) - R_Path_PLD.pins(PinName).value(site)
                    DDR_R2.pins(PinName).value(site) = (1 * Force_Cond / i2) - R_Path_PLD.pins(PinName).value(site)
                Else
                    DDR_R1.pins(PinName).value(site) = 999
                    DDR_R2.pins(PinName).value(site) = 999
                End If
            End If
                        
        Next site
        
    Next PinName
    
    Dim temp
    
    temp = TheExec.flow.TestLimitIndex
    If EnableDigitalTestLimitTTR = True Then
        'TTR,20200423, Oscar
        TestNameInput = Report_TName_From_Instance(CalcR, "X", , 0)
        TheExec.flow.TestLimit resultVal:=R_Contact_PLD, unit:=unitCustom, customUnit:="ohm", Tname:=TestNameInput, ForceResults:=tlForceFlow
    Else
        For Each PinName In R_Contact_PLD.pins
                TheExec.flow.TestLimitIndex = temp
                TestNameInput = Report_TName_From_Instance(CalcR, CStr(PinName), , 0)
                TheExec.flow.TestLimit resultVal:=R_Contact_PLD.pins(PinName), unit:=unitCustom, customUnit:="ohm", Tname:=TestNameInput, ForceResults:=tlForceFlow
                'TheExec.Flow.TestLimit resultVal:=R_Contact_PLD.Pins(PinName), lowval:=0, hival:=5, Unit:=unitCustom, customUnit:="ohm", TName:=TestNameInput, ForceResults:=tlForceNone
        Next PinName
    End If
    If Cust_Str = UCase("DDR_TEST") Then
    
        If EnableDigitalTestLimitTTR = True Then
                        'TTR,20200423, Oscar
            TestNameInput = Report_TName_From_Instance(CalcR, "X", , 0)
            TheExec.flow.TestLimit resultVal:=DDR_R1.pins(PinName), unit:=unitCustom, customUnit:="ohm", Tname:=TestNameInput, ForceResults:=tlForceFlow
            TheExec.flow.TestLimit resultVal:=DDR_R2.pins(PinName), unit:=unitCustom, customUnit:="ohm", Tname:=TestNameInput, ForceResults:=tlForceFlow
        Else
            temp = TheExec.flow.TestLimitIndex
            For Each PinName In DDR_R1.pins
                TheExec.flow.TestLimitIndex = temp
                TestNameInput = Report_TName_From_Instance(CalcR, CStr(PinName))
                TheExec.flow.TestLimit resultVal:=DDR_R1.pins(PinName), unit:=unitCustom, customUnit:="ohm", Tname:=TestNameInput, ForceResults:=tlForceFlow
            Next PinName
            
            temp = TheExec.flow.TestLimitIndex
            For Each PinName In DDR_R2.pins
                TheExec.flow.TestLimitIndex = temp
                TestNameInput = Report_TName_From_Instance(CalcR, CStr(PinName))
                TheExec.flow.TestLimit resultVal:=DDR_R2.pins(PinName), unit:=unitCustom, customUnit:="ohm", Tname:=TestNameInput, ForceResults:=tlForceFlow
            Next PinName
        End If
    End If
   
    Exit Function

Exit Function 'Add ErrHandler 2023/05/29
errHandler:
    Call Print_Error_Message(Error_Info, "LIB_HardIP_Calc", "Calc_R_Path_Cal") 'Add ErrHandler 2023/05/29
    If AbortTest Then Exit Function Else Resume Next
    
End Function


Public Function Calc_ConcatenateDSP(argc As Integer, argv() As String) As Long
On Error GoTo errHandler 'Add ErrHandler 2023/05/29
    Dim site As Variant
    Dim i, j As Long
    Dim DSPWave_First As New DSPWave
    Dim DSPWave_Second As New DSPWave
    Dim DSPWave_Combine() As New DSPWave
    Dim TestNameInput As String
    Dim SplitByAt() As String
    Dim First_StartElement As Long
    Dim First_EndElement As Long
    Dim Second_StartElement As Long
    Dim Second_EndElement As Long
    
    Dim DictKey_DSPWave_Combine As String
    
    Dim DataString_First As String
    Dim DataString_Second As String
    Dim DataString_Combine As String
   ' Dim TestNameInput As String
    Dim OutputTname_format() As String
    
    
    ReDim DSPWave_Combine(argc - 1) As New DSPWave
    Dim DSPWave_Combine_Dec As New DSPWave
    
    For i = 0 To argc - 1
        TestNameInput = "ConcatenateDSP_"
        SplitByAt = Split(argv(i), "@")
        DSPWave_First = GetStoreDataAllType(SplitByAt(0))
        First_StartElement = SplitByAt(1)
        First_EndElement = SplitByAt(2)
        DSPWave_Second = GetStoreDataAllType(SplitByAt(3))
        Second_StartElement = SplitByAt(4)
        Second_EndElement = SplitByAt(5)

'''        Call rundsp.ConcatenateDSP(DSPWave_First, First_StartElement, First_EndElement, DSPWave_Second, Second_StartElement, Second_EndElement, DSPWave_Combine(i))
                Call ConcatenateDSP_TTR(DSPWave_First, First_StartElement, First_EndElement, DSPWave_Second, Second_StartElement, Second_EndElement, DSPWave_Combine(i))
        ''20170718 - Store Concatenate DSP to Dict.
        If UBound(SplitByAt) = 6 Then
            DictKey_DSPWave_Combine = SplitByAt(6)
            Call StoreDataAllType(DictKey_DSPWave_Combine, DSPWave_Combine(i))
        End If
        
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
            
            If gl_Disable_HIP_debug_log = False Then TheExec.Datalog.WriteComment ("Site " & site & " Dictionary " & SplitByAt(0) & " Output Bits = " & DataString_First & " Extract Bits [" & First_StartElement & "-" & First_EndElement & "]" & _
                                                           " ,Dictionary " & SplitByAt(3) & " Output Bits = " & DataString_Second & " Extract Bits [" & Second_StartElement & "-" & Second_EndElement & "]" & _
                                                           " ,Dictionary " & DictKey_DSPWave_Combine & " Output Bits = " & DataString_Combine)
        Next site
        'Call rundsp.BinToDec(DSPWave_Combine(i), DSPWave_Combine_Dec)
        Dim DSPWave_Combine_BIN As New DSPWave
        For Each site In TheExec.sites
            ''===================== BinToDec =====================
            DSPWave_Combine_BIN(site) = DSPWave_Combine(i)(site).ConvertDataTypeTo(DspLong).COPY
            DSPWave_Combine_Dec(site) = DSPWave_Combine_BIN(site).ConvertStreamTo(tldspParallel, DSPWave_Combine_BIN(site).SampleSize, 0, Bit0IsMsb)
            ''===================== BinToDec (End) =====================
        Next site
        TestNameInput = Report_TName_From_Instance(CalcC, "X", "ConcatenateDSP", 0)
        
        TheExec.flow.TestLimit resultVal:=DSPWave_Combine_Dec.Element(0), PinName:="X", Tname:=TestNameInput, ForceResults:=tlForceFlow
    Next i
Exit Function 'Add ErrHandler 2023/05/29
errHandler: 'Add ErrHandler 2023/05/29
    Call Print_Error_Message(Error_Info, "LIB_HardIP_Calc", "Calc_ConcatenateDSP") 'Add ErrHandler 2023/05/29
    If AbortTest Then Exit Function Else Resume Next 'Add ErrHandler 2023/05/29
    
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
    
    Val_SerialDSP_1 = GetStoreDataAllType(argv(0))
    Val_SerialDSP_2 = GetStoreDataAllType(argv(1))
    
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
    TheExec.flow.TestLimit resultVal:=temp, Tname:=TestNameInput, ForceResults:=tlForceFlow

End Function

Public Function Calc_BitwiseDSP(argc As Integer, argv() As String) As Long
On Error GoTo errHandler 'Add ErrHandler 2023/05/29

    Dim i As Long, j As Long
    Dim site As Variant
    Dim SplitByAt() As String
    Dim HEX_Str As String
    Dim bin_str As String
    Dim bitwidth As Long
    Dim DSP_Fixed_Bin As New DSPWave
    Dim DictKey As String
    Dim OperationKeyWord As String
    Dim DSP_DictKey As New DSPWave
    Dim DSP_ProcessOutput_BIN As New DSPWave
    Dim DSP_ProcessOutput_DEC As New DSPWave
    
    Dim Dict_Str As String
    Dim Fixed_Str As String
    Dim ProcessOutput_Str As String
    Dim TestName As String
    Dim TestNameInput As String
    Dim OutputTname_format() As String
    
''    Dim SV_BitWidth As New SiteLong
''    SV_BitWidth = 32
    For i = 0 To argc - 1
        SplitByAt = Split(argv(i), "@")
        HEX_Str = SplitByAt(2)
        bitwidth = Len(SplitByAt(2)) * 4
        bin_str = HEX_to_BIN(HEX_Str)
        
        Set DSP_Fixed_Bin = Nothing
        DSP_Fixed_Bin.CreateConstant 0, bitwidth, DspLong
        
        For Each site In TheExec.sites
            For j = 0 To DSP_Fixed_Bin.SampleSize - 1
                'DSP_Fixed_Bin(Site).Element(j) = Mid(bin_str, i + 1, 1)
                DSP_Fixed_Bin(site).Element(j) = mid(bin_str, DSP_Fixed_Bin.SampleSize - j, 1)          'ZB correct  for Cyprus AMP dqpi, capi binary bits LSM-->MSB re-order  - 20170905
            Next j
        Next site
        
        DictKey = SplitByAt(0)
        
        OperationKeyWord = UCase(SplitByAt(1))
        
        DSP_DictKey = GetStoreDataAllType(DictKey)
        
        Select Case OperationKeyWord
            Case "OR"
                Call rundsp.DSP_BitWiseOr(DSP_DictKey, DSP_Fixed_Bin, bitwidth, DSP_ProcessOutput_BIN)
            Case "AND"
                Call rundsp.DSP_BitWiseAnd(DSP_DictKey, DSP_Fixed_Bin, bitwidth, DSP_ProcessOutput_BIN)
            Case "XOR"
                Call rundsp.DSP_BitWiseXOR(DSP_DictKey, DSP_Fixed_Bin, bitwidth, DSP_ProcessOutput_BIN)
            Case Else
        End Select
        
        For Each site In TheExec.sites
            Dict_Str = vbNullString
            Fixed_Str = vbNullString
            ProcessOutput_Str = vbNullString
            For j = 0 To DSP_DictKey.SampleSize - 1
                Dict_Str = Dict_Str & DSP_DictKey(site).Element(j)
            Next j
            For j = 0 To DSP_Fixed_Bin.SampleSize - 1
                Fixed_Str = Fixed_Str & DSP_Fixed_Bin(site).Element(j)
            Next j
            For j = 0 To DSP_ProcessOutput_BIN.SampleSize - 1
                ProcessOutput_Str = ProcessOutput_Str & DSP_ProcessOutput_BIN(site).Element(j)
            Next j
        
           If gl_Disable_HIP_debug_log = False Then TheExec.Datalog.WriteComment ("Site " & site & " Dictionary Output Bits = " & Dict_Str & "[" & SplitByAt(0) & "]" & vbCrLf & _
                                                                                     "Hex Val                       = " & Fixed_Str & "[" & SplitByAt(1) & " " & SplitByAt(2) & "]" & vbCrLf & _
                                                                                     "Process Result                = " & ProcessOutput_Str)
        Next site

        Set DSP_ProcessOutput_DEC = Nothing
        DSP_ProcessOutput_DEC.CreateConstant 0, 1, DspDouble
        
        TestName = OperationKeyWord & "_" & i
        
        Call rundsp.BinToDec(DSP_ProcessOutput_BIN, DSP_ProcessOutput_DEC)
        
        TestNameInput = Report_TName_From_Instance(CalcC, "X", , CInt(i))
                
        TheExec.flow.TestLimit resultVal:=DSP_ProcessOutput_DEC.Element(0), Tname:=TestNameInput, ForceResults:=tlForceFlow
        
    Next i
Exit Function 'Add ErrHandler 2023/05/29
errHandler: 'Add ErrHandler 2023/05/29
    Call Print_Error_Message(Error_Info, "LIB_HardIP_Calc", "Calc_BitwiseDSP") 'Add ErrHandler 2023/05/29
    If AbortTest Then Exit Function Else Resume Next 'Add ErrHandler 2023/05/29

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

DSPWave_Dict_DIG_25C = GetStoreDataAllType(argv(0))
DSPWave_Dict_DIG_85C = GetStoreDataAllType(argv(1))


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

Call StoreDataAllType(Fuse_REFERENCE_CTRL_DIG_Name, ADC_Trim_Code_DIG_final)
TheExec.Datalog.WriteComment ("DigCap data store in dictionary " & "<<" & Fuse_REFERENCE_CTRL_DIG_Name & ">>")

eFuse_CTRL_DIG = GetStoreDataAllType(Fuse_REFERENCE_CTRL_DIG_Name)

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

Public Function Calc_2S_Complement_DSP(argc As Integer, argv() As String) As Long
On Error GoTo errHandler 'Add ErrHandler 2023/05/29
    Dim i As Long, j As Long
    Dim site As Variant
    Dim SplitByAt() As String
    Dim DictKey As String
    Dim DictKey_2S_DEC As String
    
    Dim DSP_DictKey_BIN As New DSPWave
    Dim DSP_DictKey_DEC As New DSPWave
    
    Dim DSPWave_2S_Complement() As New DSPWave
    ReDim DSPWave_2S_Complement(argc - 1) As New DSPWave
    
    Dim TestName As String
    Dim TestNameInput As String
    Dim OutputTname_format() As String
    
    Dim SL_BitWidth As New SiteLong
''    Call rundsp.WordWidthExample
    
    For i = 0 To argc - 1
        If InStr(argv(i), "@") <> 0 Then
            SplitByAt = Split(argv(i), "@")
            DictKey = SplitByAt(0)
            DictKey_2S_DEC = SplitByAt(1)
'''''            TestName = SplitByAt(2)
        Else
            DictKey = argv(i)
        End If
        
        DSP_DictKey_BIN = GetStoreDataAllType(DictKey)
        Set DSP_DictKey_DEC = Nothing
        DSP_DictKey_DEC.CreateConstant 0, 1, DspDouble
        'Call rundsp.BinToDec(DSP_DictKey_BIN, DSP_DictKey_DEC)
        
        For Each site In TheExec.sites
        
        
                    ''===================== BinToDec =====================
            DSP_DictKey_BIN(site) = DSP_DictKey_BIN(site).ConvertDataTypeTo(DspLong)
            DSP_DictKey_DEC(site) = DSP_DictKey_BIN(site).ConvertStreamTo(tldspParallel, DSP_DictKey_BIN(site).SampleSize, 0, Bit0IsMsb)
            ''===================== BinToDec (End) =====================
        
        
            SL_BitWidth(site) = DSP_DictKey_BIN(site).SampleSize
''            DSP_DictKey_DEC(0).Element(0) = 255
''            DSP_DictKey_DEC(1).Element(0) = 254
        Next site
        
        Set DSPWave_2S_Complement(i) = Nothing
        DSPWave_2S_Complement(i).CreateConstant 0, 1, DspDouble
        
        'Call rundsp.DSP_Convert_2S_Complement(DSP_DictKey_DEC, SL_BitWidth, DSPWave_2S_Complement(i))
        
        
        
                 ''===================== Convert_2S_Complement =====================
        For Each site In TheExec.sites
            DSPWave_2S_Complement(i)(site) = DSP_DictKey_DEC(site).ConvertDataTypeTo(DspLong)
            DSPWave_2S_Complement(i)(site).WordWidth = SL_BitWidth(site)
            DSPWave_2S_Complement(i)(site) = DSPWave_2S_Complement(i)(site).ConvertDataTypeTo(DspLong)
'            Debug.Print DSPWave_2S_Complement(i)(Site).Element(0)
        Next site
        ''===================== Convert_2S_Complement  (End) =====================
        
        
        
        
        If InStr(argv(i), "@") <> 0 Then
            Call StoreDataAllType(DictKey_2S_DEC, DSPWave_2S_Complement(i))
            
            TestNameInput = Report_TName_From_Instance(CalcC, "X", "DEC" & CStr(i), CInt(i))
            
            TheExec.flow.TestLimit resultVal:=DSP_DictKey_DEC.Element(0), Tname:=TestNameInput, ForceResults:=tlForceFlow
            
            Call Update_BC_PassFail_Flag
            
            TestNameInput = Report_TName_From_Instance(CalcC, "X", , CInt(i))
            
            TheExec.flow.TestLimit resultVal:=DSPWave_2S_Complement(i).Element(0), Tname:=TestNameInput, ForceResults:=tlForceFlow
            
            Call Update_BC_PassFail_Flag
        Else
            
            TestNameInput = Report_TName_From_Instance(CalcC, "X", , CInt(i))
            
            TheExec.flow.TestLimit resultVal:=DSPWave_2S_Complement(i).Element(0), Tname:=TestNameInput, ForceResults:=tlForceFlow
            
            Call Update_BC_PassFail_Flag
        End If
    Next i
Exit Function 'Add ErrHandler 2023/05/29
errHandler: 'Add ErrHandler 2023/05/29
    Call Print_Error_Message(Error_Info, "LIB_HardIP_Calc", "Calc_2S_Complement_DSP") 'Add ErrHandler 2023/05/29
    If AbortTest Then Exit Function Else Resume Next 'Add ErrHandler 2023/05/29
End Function


Public Function MSBNEGATE(argc As Integer, argv() As String) As Long
On Error GoTo errHandler 'Add ErrHandler 2023/05/29
    
    Dim temp_dsp As New DSPWave
    Dim i As Long
    Dim site As Variant
    
    For i = 0 To argc - 1
        temp_dsp = GetStoreDataAllType(argv(i))
        For Each site In TheExec.sites
            If temp_dsp.Element(temp_dsp.SampleSize - 1) = 0 Then
                temp_dsp.Element(temp_dsp.SampleSize - 1) = 1
            ElseIf temp_dsp.Element(temp_dsp.SampleSize - 1) = 1 Then
                temp_dsp.Element(temp_dsp.SampleSize - 1) = 0
            Else
            'Do nothing
            End If
            'Call StoreDataAllType(argv(i), temp_dsp)
        Next site
        Call StoreDataAllType(argv(i), temp_dsp)
    Next

Exit Function 'Add ErrHandler 2023/05/29
errHandler: 'Add ErrHandler 2023/05/29
    Call Print_Error_Message(Error_Info, "LIB_HardIP_Calc", "MSBNEGATE") 'Add ErrHandler 2023/05/29
    If AbortTest Then Exit Function Else Resume Next 'Add ErrHandler 2023/05/29
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
        'DataOut_25C_Sensor1 = GetStoreDataAllType(argv(0))
        Call HardIP_Bin2Dec(DataOut_25C_Sensor1, GetStoreDataAllType(argv(0))) ' for Turks
    End If
    
    Call HardIP_Bin2Dec(DataOut_85C_Sensor1, GetStoreDataAllType(argv(1)))

    Call TMPS_Coeff_Calculation(Coeff_A0_Sensor1, Coeff_A1_Sensor1, Coeff_A2_Sensor1, Coeff_A3_Sensor1, Coeff_A4_Sensor1, DataOut_85C_Sensor1, DataOut_25C_Sensor1)

    Call HardIP_Dec2Bin(Coeff_A0_Sensor1_Dict, Coeff_A0_Sensor1, 15)
    Call HardIP_Dec2Bin(Coeff_A1_Sensor1_Dict, Coeff_A1_Sensor1, 14)
    Call HardIP_Dec2Bin(Coeff_A2_Sensor1_Dict, Coeff_A2_Sensor1, 12)
    Call HardIP_Dec2Bin(Coeff_A3_Sensor1_Dict, Coeff_A3_Sensor1, 10)
    Call HardIP_Dec2Bin(Coeff_A4_Sensor1_Dict, Coeff_A4_Sensor1, 11)

    Call StoreDataAllType(argv(2), Coeff_A0_Sensor1_Dict)
    Call StoreDataAllType(argv(3), Coeff_A1_Sensor1_Dict)
    Call StoreDataAllType(argv(4), Coeff_A2_Sensor1_Dict)
    Call StoreDataAllType(argv(5), Coeff_A3_Sensor1_Dict)
    Call StoreDataAllType(argv(6), Coeff_A4_Sensor1_Dict)
        
    Exit Function
errHandler:
        Call Print_Error_Message(Error_Info, "LIB_HardIP_Calc", "Calc_TMPS_Coeff") 'Add ErrHandler 2023/05/29
        If AbortTest Then Exit Function Else Resume Next
End Function

Public Function Calc_TMPS_Coeff_1point(argc As Integer, argv() As String) As Long
On Error GoTo errHandler 'Add ErrHandler 2023/05/29
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
    
    
   '' DataOut_25C_Sensor1 = GetStoreDataAllType(argv(0))

    Call HardIP_Bin2Dec(DataOut_25C_Sensor1, GetStoreDataAllType(argv(0)))

    Call TMPS_Coeff_Calculation_1point(Coeff_A0_Sensor1, Coeff_A1_Sensor1, Coeff_A2_Sensor1, Coeff_A3_Sensor1, Coeff_A4_Sensor1, DataOut_25C_Sensor1)

    Call HardIP_Dec2Bin(Coeff_A0_Sensor1_Dict, Coeff_A0_Sensor1, 15)
    Call HardIP_Dec2Bin(Coeff_A1_Sensor1_Dict, Coeff_A1_Sensor1, 14)
    Call HardIP_Dec2Bin(Coeff_A2_Sensor1_Dict, Coeff_A2_Sensor1, 12)
    Call HardIP_Dec2Bin(Coeff_A3_Sensor1_Dict, Coeff_A3_Sensor1, 10)
    Call HardIP_Dec2Bin(Coeff_A4_Sensor1_Dict, Coeff_A4_Sensor1, 11)

    Call StoreDataAllType(argv(1), Coeff_A0_Sensor1_Dict)
    Call StoreDataAllType(argv(2), Coeff_A1_Sensor1_Dict)
    Call StoreDataAllType(argv(3), Coeff_A2_Sensor1_Dict)
    Call StoreDataAllType(argv(4), Coeff_A3_Sensor1_Dict)
    Call StoreDataAllType(argv(5), Coeff_A4_Sensor1_Dict)
        
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
        DSPWave_Bin(i) = GetStoreDataAllType(argv(i))
        Call rundsp.BinToDec(DSPWave_Bin(i), DSPWave_Dec(i))
        Call rundsp.DSP_Add(DSPWave_AverageDec, DSPWave_Dec(i))
    Next i
    Call rundsp.DSP_DivideConstant(DSPWave_AverageDec, argc - 1)
''    Call rundsp.DSP_ConvertDataTypeToLong(DSPWave_AverageDec)
    For Each site In TheExec.sites
        ''20170210-Rounding
        DSPWave_AverageDec(site).Element(0) = Int(DSPWave_AverageDec(site).Element(0) + 0.5)
    Next site

    Call StoreDataAllType(argv(argc - 1), DSPWave_AverageDec)
    TheExec.flow.TestLimit resultVal:=DSPWave_AverageDec.Element(0), Tname:="ADDRIO_AverageTrimCode", ForceResults:=tlForceFlow 'transfer_to_forceflow
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
        DSPWaveBin(i) = GetStoreDataAllType(argv(i))
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
    
    TheExec.flow.TestLimit resultVal:=MDLL_DecreaseResultPass, lowVal:=1, hiVal:=1, Tname:=TestNameInput, ForceResults:=tlForceFlow 'transfer_to_forceflow
    For Each site In TheExec.sites
        If MDLL_DecreaseResultPass.bitwiseand(1) Then
        Else
            MDLL_Index(site) = -99
        End If
    Next site
    
    TestNameInput = Report_TName_From_Instance(CalcC, "X", "MDLLUnique", 1)
    TheExec.flow.TestLimit resultVal:=MDLL_Index, lowVal:=1, hiVal:=2, Tname:=TestNameInput, ForceResults:=tlForceFlow 'transfer_to_forceflow
End Function
Public Function Calc_DDR_MDCC_Freq(argc As Integer, argv() As String) As Long
On Error GoTo errHandler 'Add ErrHandler 2023/05/29

    Dim i As Long, j As Long
    Dim SplitByAt() As String
    Dim F_Case As String
    Dim DictKey As String
    Dim Dict_DSP_DEC() As New DSPWave
    ReDim Dict_DSP_DEC(argc - 1) As New DSPWave
    Dim Dict_DSP_BINARY() As New DSPWave
    ReDim Dict_DSP_BINARY(argc - 1) As New DSPWave
    Dim site As Variant
    Dim Calc_DSP_DEC() As New DSPWave
    ReDim Calc_DSP_DEC(argc - 1) As New DSPWave
    Dim TestName As String
    Dim TestNameInput As String
    Dim OutputTname_format() As String
    
    For i = 0 To argc - 1
        Calc_DSP_DEC(i).CreateConstant 0, 1, DspDouble
        SplitByAt = Split(argv(i), "@")
        F_Case = UCase(SplitByAt(0))
        DictKey = SplitByAt(1)
        Dict_DSP_BINARY(i) = GetStoreDataAllType(DictKey)
        TestName = SplitByAt(2)
        
'        Call rundsp.BinToDec(Dict_DSP_BINARY(i), Dict_DSP_DEC(i))

        For Each site In TheExec.sites
            ''===================== BinToDec =====================
            Dict_DSP_BINARY(i)(site) = Dict_DSP_BINARY(i)(site).ConvertDataTypeTo(DspLong)
            Dict_DSP_DEC(i)(site) = Dict_DSP_BINARY(i)(site).ConvertStreamTo(tldspParallel, Dict_DSP_BINARY(i)(site).SampleSize, 0, Bit0IsMsb)
            ''===================== BinToDec (End) =====================
        Next site


        
        For Each site In TheExec.sites
            Select Case F_Case
                Case "F0"
                    Calc_DSP_DEC(i)(site).Element(0) = ((Dict_DSP_DEC(i)(site).Element(0)) / (114 * 2)) * 2.133 * 1000000000#
                Case "F1"
                    Calc_DSP_DEC(i)(site).Element(0) = (Dict_DSP_DEC(i)(site).Element(0)) / (100 * 2) * 1.466 * 1000000000#
                Case "F2"
                    Calc_DSP_DEC(i)(site).Element(0) = (Dict_DSP_DEC(i)(site).Element(0)) / (90 * 2) * 0.712 * 1000000000#
                Case "M9_F1"
                    Calc_DSP_DEC(i)(site).Element(0) = (Dict_DSP_DEC(i)(site).Element(0)) / (100 * 2) * 1.2 * 1000000000#
                Case Else
            
            End Select
        Next site
''        TheExec.Flow.TestLimit resultVal:=Calc_DSP_DEC(i).Element(0), Tname:="Calc_DDR_MDCC_Freq", unit:=unitHz, ForceResults:=tlForceFlow
        TestNameInput = Report_TName_From_Instance(CalcF, "X", , CInt(i))
        TheExec.flow.TestLimit resultVal:=Calc_DSP_DEC(i).Element(0), Tname:=TestNameInput, unit:=unitHz, ForceResults:=tlForceFlow
    Next i
Exit Function 'Add ErrHandler 2023/05/29
errHandler: 'Add ErrHandler 2023/05/29
    Call Print_Error_Message(Error_Info, "LIB_HardIP_Calc", "Calc_DDR_MDCC_Freq") 'Add ErrHandler 2023/05/29
    If AbortTest Then Exit Function Else Resume Next 'Add ErrHandler 2023/05/29
End Function

Public Function Calc_MDLL_Monotonicity_DevideBlock_SEG(argc As Integer, argv() As String) As Long
On Error GoTo errHandler 'Add ErrHandler 2023/05/29
   Dim i As Long
   Dim site As Variant
   Dim DSP_Captured() As New DSPWave
   ReDim DSP_Captured((argc - 2))
    
   Dim DSP_Arry_Bin() As New DSPWave
   ReDim DSP_Arry_Bin((argc - 2) * 4 - 1)
    
   Dim DSP_Arry_Dec() As New DSPWave
   ReDim DSP_Arry_Dec((argc - 2) * 4 - 1)
    
   Dim Uni_DLL_Indicator As New SiteLong
   Dim Max_Dec_Val As New SiteLong
    
   Dim TestNameInput As String
   Dim OutputTname_format() As String
    
   Dim TestName As String
    
   For i = 0 To argc - 2
      DSP_Captured(i) = GetStoreDataAllType(argv(i + 1))
   Next i
    
''   '--- For Offline Simulation ------
''   If TheExec.TesterMode = testModeOffline Then
''      For Each Site In TheExec.sites.Active
''         'code1
''         DSP_Captured(0).Element(0) = 1
''         DSP_Captured(0).Element(1) = 1
''         DSP_Captured(0).Element(2) = 1
''         DSP_Captured(0).Element(3) = 1
''         DSP_Captured(0).Element(4) = 1
''         DSP_Captured(0).Element(5) = 1
''         DSP_Captured(0).Element(6) = 0
''         DSP_Captured(0).Element(7) = 0
''
''         'code6
''         DSP_Captured(0).Element(8) = 1
''         DSP_Captured(0).Element(9) = 1
''         DSP_Captured(0).Element(10) = 1
''         DSP_Captured(0).Element(11) = 1
''         DSP_Captured(0).Element(12) = 1
''         DSP_Captured(0).Element(13) = 1
''         DSP_Captured(0).Element(14) = 0
''         DSP_Captured(0).Element(15) = 0
''
''         'code0
''         DSP_Captured(0).Element(16) = 1
''         DSP_Captured(0).Element(17) = 1
''         DSP_Captured(0).Element(18) = 1
''         DSP_Captured(0).Element(19) = 1
''         DSP_Captured(0).Element(20) = 1
''         DSP_Captured(0).Element(21) = 1
''         DSP_Captured(0).Element(22) = 0
''         DSP_Captured(0).Element(23) = 0
''
''         'code4
''         DSP_Captured(0).Element(24) = 1
''         DSP_Captured(0).Element(25) = 1
''         DSP_Captured(0).Element(26) = 1
''         DSP_Captured(0).Element(27) = 1
''         DSP_Captured(0).Element(28) = 1
''         DSP_Captured(0).Element(29) = 1
''         DSP_Captured(0).Element(30) = 0
''         DSP_Captured(0).Element(31) = 0
''
''         'code5
''         DSP_Captured(1).Element(0) = 1
''         DSP_Captured(1).Element(1) = 1
''         DSP_Captured(1).Element(2) = 1
''         DSP_Captured(1).Element(3) = 1
''         DSP_Captured(1).Element(4) = 1
''         DSP_Captured(1).Element(5) = 1
''         DSP_Captured(1).Element(6) = 0
''         DSP_Captured(1).Element(7) = 0
''
''         'code2
''         DSP_Captured(1).Element(8) = 1
''         DSP_Captured(1).Element(9) = 1
''         DSP_Captured(1).Element(10) = 1
''         DSP_Captured(1).Element(11) = 1
''         DSP_Captured(1).Element(12) = 1
''         DSP_Captured(1).Element(13) = 1
''         DSP_Captured(1).Element(14) = 0
''         DSP_Captured(1).Element(15) = 0
''
''         'code7
''         DSP_Captured(1).Element(16) = 1
''         DSP_Captured(1).Element(17) = 1
''         DSP_Captured(1).Element(18) = 1
''         DSP_Captured(1).Element(19) = 1
''         DSP_Captured(1).Element(20) = 1
''         DSP_Captured(1).Element(21) = 1
''         DSP_Captured(1).Element(22) = 0
''         DSP_Captured(1).Element(23) = 0
''
''         'code3
''         DSP_Captured(1).Element(24) = 1
''         DSP_Captured(1).Element(25) = 1
''         DSP_Captured(1).Element(26) = 1
''         DSP_Captured(1).Element(27) = 1
''         DSP_Captured(1).Element(28) = 1
''         DSP_Captured(1).Element(29) = 1
''         DSP_Captured(1).Element(30) = 0
''         DSP_Captured(1).Element(31) = 0
''
''      Next Site
    
      '       Dim j, k As Long
      '       For j = 0 To 1
      '           For k = 0 To 3
      '               DSP_Captured(j).Element(k * 8) = 1
      '
      '           Next k
      '       Next j
       
'   End If
    
    
    
   For Each site In TheExec.sites.Active
        'cfgh_cadll_sts_mdll_code_grp1_w210,{3'b000, oct2[26:18], 1'b0, oct1[17:9],   1'b0, oct0[8:0]};
        'cfgh_cadll_sts_mdll_code_grp2_w543,{3'b000, oct5[53:45], 1'b0, oct4[44:36], 1'b0, oct3[35:27]};
        'cfgh_cadll_sts_mdll_code_grp3_w76,{13'b0,     oct7[71:63], 1'b0, oct6[62:54]};
        '
        'for the 32 bits of w210 DigCap, it maps to : {3'b0, oct2 values, 1'b0, oct1 values, 1'b0, oct0 values}.
        'for the 32 bits of w543 DigCap, it maps to : {3'b0, oct5 values, 1'b0, oct4 values, 1'b0, oct3 values}.
        'for the 32 bits of w76 DigCap, it maps to : {13'b0, oct7 values, 1'b0, oct6 values}.
        
        
'      DSP_Arry_Bin(4) = DSP_Captured(0).Select(0, 1, DSP_Captured(0).SampleSize / 4).Copy
'      DSP_Arry_Bin(0) = DSP_Captured(0).Select(DSP_Captured(0).SampleSize / 4, 1, DSP_Captured(0).SampleSize / 4).Copy
'      DSP_Arry_Bin(6) = DSP_Captured(0).Select(DSP_Captured(0).SampleSize / 2, 1, DSP_Captured(0).SampleSize / 4).Copy
'      DSP_Arry_Bin(1) = DSP_Captured(0).Select((DSP_Captured(0).SampleSize / 4) * 3, 1, DSP_Captured(0).SampleSize / 4).Copy
'      DSP_Arry_Bin(3) = DSP_Captured(1).Select(0, 1, DSP_Captured(1).SampleSize / 4).Copy
'      DSP_Arry_Bin(7) = DSP_Captured(1).Select(DSP_Captured(1).SampleSize / 4, 1, DSP_Captured(1).SampleSize / 4).Copy
'      DSP_Arry_Bin(2) = DSP_Captured(1).Select(DSP_Captured(1).SampleSize / 2, 1, DSP_Captured(1).SampleSize / 4).Copy
'      DSP_Arry_Bin(5) = DSP_Captured(1).Select((DSP_Captured(1).SampleSize / 4) * 3, 1, DSP_Captured(1).SampleSize / 4).Copy

      
      DSP_Arry_Bin(4) = DSP_Captured(0).Select(0, 1, 9).COPY
      DSP_Arry_Bin(0) = DSP_Captured(0).Select(10, 1, 9).COPY
      DSP_Arry_Bin(6) = DSP_Captured(0).Select(20, 1, 9).COPY
      
      DSP_Arry_Bin(1) = DSP_Captured(1).Select(0, 1, 9).COPY
      DSP_Arry_Bin(3) = DSP_Captured(1).Select(10, 1, 9).COPY
      DSP_Arry_Bin(7) = DSP_Captured(1).Select(20, 1, 9).COPY
      
      DSP_Arry_Bin(2) = DSP_Captured(2).Select(0, 1, 9).COPY
      DSP_Arry_Bin(5) = DSP_Captured(2).Select(10, 1, 9).COPY
      
      
   
      For i = 0 To UBound(DSP_Arry_Bin)
         DSP_Arry_Bin(i) = DSP_Arry_Bin(i).ConvertDataTypeTo(DspLong)
         DSP_Arry_Dec(i) = DSP_Arry_Bin(i).ConvertStreamTo(tldspParallel, DSP_Arry_Bin(i).SampleSize, 0, Bit0IsMsb)
         'nope, only for debugging purpose
         'Report_TestLimit_by_CZ_Format resultVal:=DSP_Arry_Dec(i).Element(0), ForceResults:=tlForceNone, UserVar6:="DSP_Arry_Dec" & i, UserVar5:=argv(0), MeasType:="C"
      Next i
      
      Uni_DLL_Indicator(site) = 1
      For i = 0 To UBound(DSP_Arry_Bin) - 1
         If Uni_DLL_Indicator(site) = 1 Then
            If DSP_Arry_Dec(i).Element(0) = DSP_Arry_Dec(i + 1).Element(0) Then
            ElseIf DSP_Arry_Dec(i).Element(0) = DSP_Arry_Dec(i + 1).Element(0) + 1 Then Uni_DLL_Indicator(site) = 2
            Else
               Uni_DLL_Indicator(site) = -2
               Exit For
            End If
         ElseIf Uni_DLL_Indicator(site) = 2 Then
            If DSP_Arry_Dec(i).Element(0) = DSP_Arry_Dec(i + 1).Element(0) Then
            Else
               Uni_DLL_Indicator(site) = -1
               Exit For
            End If
         Else
         End If
      Next i
      Max_Dec_Val(site) = DSP_Arry_Dec(0).Element(0)
   Next site
    
   Call GetFlowTName
    
   If gl_UseStandardTestName_Flag = True Then                     'Roger add
      Call Report_ALG_TName_From_Instance(OutputTname_format, "C", CStr(argv(0)) & "Max_Diff", gl_Tname_Meas_FromFlow(TheExec.flow.TestLimitIndex))
      TestNameInput = Merge_TName(OutputTname_format)
            
   Else
      TestNameInput = TestName & "Max_Diff"
   End If
    
   TheExec.flow.TestLimit resultVal:=Max_Dec_Val, Tname:=TestNameInput, ForceResults:=tlForceFlow
    
   If gl_UseStandardTestName_Flag = True Then                     'Roger add
      Call Report_ALG_TName_From_Instance(OutputTname_format, "C", CStr(argv(0)) & "Decrease", gl_Tname_Meas_FromFlow(TheExec.flow.TestLimitIndex))
      TestNameInput = Merge_TName(OutputTname_format)
            
   Else
      TestNameInput = TestName & "Decrease"
   End If
    
    
   TheExec.flow.TestLimit resultVal:=Uni_DLL_Indicator, Tname:=TestNameInput, ForceResults:=tlForceFlow
    
    
    
    
   '''    Report_TestLimit_by_CZ_Format resultVal:=Max_Dec_Val, ForceResults:=tlForceFlow, MeasType:="C"
   '''
   '''    Report_TestLimit_by_CZ_Format resultVal:=Uni_DLL_Indicator, lowVal:=1, hiVal:=2, ForceResults:=tlForceFlow, MeasType:="C"
Exit Function 'Add ErrHandler 2023/05/29
errHandler: 'Add ErrHandler 2023/05/29
        Call Print_Error_Message(Error_Info, "LIB_HardIP_Calc", "Calc_MDLL_Monotonicity_DevideBlock_SEG") 'Add ErrHandler 2023/05/29
        If AbortTest Then Exit Function Else Resume Next 'Add ErrHandler 2023/05/29
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
        DSP_Captured(i) = GetStoreDataAllType(argv(i))
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
        gl_Tname_Alg_Index = CStr(TheExec.flow.TestLimitIndex)
        TestNameInput = Report_TName_From_Instance("N", "x", left(argv(0), InStr(1, argv(0), "_")) & "Decrease", CInt(gl_Tname_Alg_Index), , "qq")
        TheExec.flow.TestLimit resultVal:=DecreaseRank, lowVal:=1, hiVal:=1, Tname:=TestNameInput, ForceResults:=tlForceFlow
        gl_Tname_Alg_Index = CStr(TheExec.flow.TestLimitIndex)
        TestNameInput = Report_TName_From_Instance("N", "x", left(argv(0), InStr(1, argv(0), "_")) & "Unique", CInt(gl_Tname_Alg_Index))
        TheExec.flow.TestLimit resultVal:=Uni_DLL_Indicator, lowVal:=1, hiVal:=2, Tname:=TestNameInput, ForceResults:=tlForceFlow
        gl_Tname_Alg_Index = CStr(TheExec.flow.TestLimitIndex)
        TestNameInput = Report_TName_From_Instance("N", "x", left(argv(0), InStr(1, argv(0), "_")) & "MaxDiff", CInt(gl_Tname_Alg_Index))
        TheExec.flow.TestLimit resultVal:=MaxDiffRank, lowVal:=0, hiVal:=1, Tname:=TestNameInput, ForceResults:=tlForceFlow
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
    
    
    DSP_FXCode_Bin = GetStoreDataAllType(Dict_FXCode)
    Call rundsp.BinToDec(DSP_FXCode_Bin, DSP_FXCode_Dec)
     
'     ''Simulation
'    DSP_FXCode_Dec(0).Element(0) = 12
'    DSP_FXCode_Dec(1).Element(0) = 15
    
    ''Truncate FXCode to 5 bit
    Call rundsp.DSPWaveDecToBinary(DSP_FXCode_Dec, 5, DSP_FXCode_5Bit_Bin)
    Call StoreDataAllType(Dict_FXCode_5Bit, DSP_FXCode_5Bit_Bin)
    
 
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
    
    ''Call StoreDataAllType(Dict_Margin_1Bit, DSP_Margin_1Bit_Bin)
    
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


    ''Call StoreDataAllType(Dict_Margin_5Bit, DSP_Margin_5Bit_Bin)
    Call StoreDataAllType(Dict_Post_Bin, DSP_Post_Bin)
    Call StoreDataAllType(Dict_Pre_Bin, DSP_Pre_Bin)
    Call StoreDataAllType(Dict_Post_2R, DSP_Post_2R_Bin)
    Call StoreDataAllType(Dict_Pre_2R, DSP_Pre_2R_Bin)
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
    DSPWave_Dict = GetStoreDataAllType(Dict_Name)
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
        Dict_DSP_Bin(i) = GetStoreDataAllType(UCase(argv(i)))
        'Call rundsp.DSP_GrayCode2Bin(b_IsUnSigned, Dict_DSP_Bin(i), GrayCode_DSP_Bin(i), GrayCode_DSP_Dec(i))
        Call GrayCode2Bin_TTR(b_IsUnSigned, Dict_DSP_Bin(i), GrayCode_DSP_Bin(i), GrayCode_DSP_Dec(i))
        TestNameInput = Report_TName_From_Instance(CalcC, "X", vbNullString, CInt(i))
        
        TheExec.flow.TestLimit resultVal:=GrayCode_DSP_Dec(i).Element(0), Tname:=TestNameInput, ForceResults:=tlForceFlow
        
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
        CalcDutyVal(i) = GetStoreDataAllType(argv(i))
        If TheExec.TesterMode = testModeOffline Then
            For j = 0 To CalcDutyVal(i).pins.Count - 1
                CalcDutyVal(i).pins(j) = 1000000 - 1000 * j - i * 2000
            Next j
        End If
        For j = 1 To CalcDutyVal(i).pins.Count - 1
            If InStr(UCase(CalcDutyVal(i).pins(j)), "_P") <> 0 Then
                PinName = CalcDutyVal(i).pins(j)
                TestNameInput = Replace(LCase(PinName), "ddr", "ch")
                TestNameInput = Replace(LCase(TestNameInput), "dqs_p", "core")
                TestNameInput = TestNameInput & "_" & Freq_TestName_Input
                
                For Each site In TheExec.sites
                    If CalcDutyVal(i).pins(j).value(site) = 0 Then
                        b_DivideZeroError(site) = True
                        If gl_Disable_HIP_debug_log = False Then TheExec.Datalog.WriteComment ("Site " & site & " Freq Meas 0 Hz , No CalcDutyDelay ")
                        CalcDutyVal(i).pins(j).value = 1
                    End If
                Next site
            
                CalcDutyVal(i).pins(j).value = CalcDutyVal(i).pins(j).Multiply(2).Invert
                
                For Each site In TheExec.sites
                    If b_DivideZeroError(site) = True Then
                        CalcDutyVal(i).pins(j).value = -999
    ''                TheExec.Flow.TestLimit resultVal:=-999, ScaleType:=scalePico, Tname:="Delay" & CStr(i - 1) & "_" & TestNameInput, ForceResults:=tlForceNone
                    End If
                Next site
                TestNameInput = Report_TName_From_Instance(CalcF, CalcDutyVal(i).pins(j), vbNullString, 0)
                TheExec.flow.TestLimit resultVal:=CalcDutyVal(i).pins(j), scaletype:=scalePico, Tname:=TestNameInput, ForceResults:=tlForceNone_CZ
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
        For j = 1 To Octant_Val(OctantIndex).pins.Count - 1
             If InStr(UCase(Octant_Val(OctantIndex).pins(j)), "_P") <> 0 Then
                PinName = Octant_Val(OctantIndex).pins(j)
                TestNameInput = Replace(LCase(PinName), "ddr", "ch")
                TestNameInput = Replace(LCase(TestNameInput), "dqs_p", "core")
                TestNameInput = TestNameInput & "_" & Freq_TestName_Input
                
                For Each site In TheExec.sites
                    If b_DivideZeroError(site) = True Then
                        Octant_Val(OctantIndex).pins(j).value = -999
''                        TheExec.Flow.TestLimit resultVal:=-999, ScaleType:=scalePico, Tname:="Oct" & CStr(OctantIndex) & "_" & TestNameInput, ForceResults:=tlForceNone
                    End If
                Next site
                TestNameInput = Report_TName_From_Instance(CalcF, Octant_Val(OctantIndex).pins(j), , 0)
                TheExec.flow.TestLimit resultVal:=Octant_Val(OctantIndex).pins(j), scaletype:=scalePico, Tname:=TestNameInput, ForceResults:=tlForceNone_CZ
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
        For j = 1 To OctPhaseError(i).pins.Count - 1
            If InStr(UCase(OctPhaseError(i).pins(j)), "_P") <> 0 Then
                PinName = OctPhaseError(i).pins(j)
                TestNameInput = Replace(LCase(PinName), "ddr", "ch")
                TestNameInput = Replace(LCase(TestNameInput), "dqs_p", "core")
                TestNameInput = TestNameInput & "_" & Freq_TestName_Input
                
                For Each site In TheExec.sites
                    If OctPhaseError(i).pins(j).value > OctPhaseError_Max.pins(j).value Then
                        OctPhaseError_Max.pins(j).value = OctPhaseError(i).pins(j).value
                    End If
                    If OctPhaseError(i).pins(j).value < OctPhaseError_Min.pins(j).value Then
                        OctPhaseError_Min.pins(j).value = OctPhaseError(i).pins(j).value
                    End If
                    If b_DivideZeroError(site) = True Then
''                        TheExec.Flow.TestLimit resultVal:=-999, ScaleType:=scalePico, Tname:="PE" & CStr(i) & "_" & TestNameInput, ForceResults:=tlForceNone
                        OctPhaseError(i).pins(j).value = -999
                    End If
''                    Else
''                        TheExec.Flow.TestLimit resultVal:=OctPhaseError(i).Pins(j).Value, ScaleType:=scalePico, Tname:="PE" & CStr(i) & "_" & TestNameInput, ForceResults:=tlForceNone
''                    End If
''                    If i = OctantMaxNum Then
''                        If b_DivideZeroError(Site) = True Then
''                            TheExec.Flow.TestLimit resultVal:=-999, ScaleType:=scalePico, Tname:="PE_MAX" & "_" & TestNameInput, ForceResults:=tlForceNone
''                            TheExec.Flow.TestLimit resultVal:=-999, ScaleType:=scalePico, Tname:="PE_MIN" & "_" & TestNameInput, ForceResults:=tlForceNone
''                        Else
''                            TheExec.Flow.TestLimit resultVal:=OctPhaseError_Max.Pins(j).Value, ScaleType:=scalePico, Tname:="PE_MAX" & "_" & TestNameInput, ForceResults:=tlForceNone
''                            TheExec.Flow.TestLimit resultVal:=OctPhaseError_Min.Pins(j).Value, ScaleType:=scalePico, Tname:="PE_MIN" & "_" & TestNameInput, ForceResults:=tlForceNone
''                        End If
''                    End If
                Next site
                TestNameInput = Report_TName_From_Instance(CalcF, OctPhaseError(i).pins(j), , 0)
                TheExec.flow.TestLimit resultVal:=OctPhaseError(i).pins(j), scaletype:=scalePico, Tname:=TestNameInput, ForceResults:=tlForceNone_CZ
''                If i = OctantMaxNum Then
''                    For Each Site In TheExec.sites
''                        If b_DivideZeroError(Site) = True Then
''                            TheExec.Flow.TestLimit resultVal:=-999, ScaleType:=scalePico, Tname:="PE_MAX" & "_" & TestNameInput, ForceResults:=tlForceNone
''                            TheExec.Flow.TestLimit resultVal:=-999, ScaleType:=scalePico, Tname:="PE_MIN" & "_" & TestNameInput, ForceResults:=tlForceNone
''                        End If
''                    Next Site
''                    TheExec.Flow.TestLimit resultVal:=OctPhaseError_Max.Pins(j), ScaleType:=scalePico, Tname:="PE_MAX" & "_" & TestNameInput, ForceResults:=tlForceNone
''                    TheExec.Flow.TestLimit resultVal:=OctPhaseError_Min.Pins(j), ScaleType:=scalePico, Tname:="PE_MIN" & "_" & TestNameInput, ForceResults:=tlForceNone
''
''                End If
            End If
        Next j
    Next i

    For j = 1 To OctPhaseError_Max.pins.Count - 1
        If InStr(UCase(OctPhaseError_Max.pins(j)), "_P") <> 0 Then
            PinName = OctPhaseError_Max.pins(j)
            TestNameInput = Replace(LCase(PinName), "ddr", "ch")
            TestNameInput = Replace(LCase(TestNameInput), "dqs_p", "core")
            TestNameInput = TestNameInput & "_" & Freq_TestName_Input
            
            For Each site In TheExec.sites
                If b_DivideZeroError(site) = True Then
                    OctPhaseError_Max.pins(j).value = -999
                    OctPhaseError_Min.pins(j).value = -999
                End If
            Next site
            
            TestNameInput = Report_TName_From_Instance(CalcF, OctPhaseError_Max.pins(j), vbNullString, 0)
            TheExec.flow.TestLimit resultVal:=OctPhaseError_Max.pins(j), scaletype:=scalePico, Tname:="PE_MAX" & "_" & TestNameInput & "_" & Voltage_Name(UBound(Voltage_Name)), ForceResults:=tlForceNone_CZ
            
            TestNameInput = Report_TName_From_Instance(CalcF, OctPhaseError_Min.pins(j), vbNullString, 0)
            TheExec.flow.TestLimit resultVal:=OctPhaseError_Min.pins(j), scaletype:=scalePico, Tname:=TestNameInput, ForceResults:=tlForceNone_CZ
        End If
    Next j

    
    For i = StartNumOfDuty To MaxNumOfDuty
        If i = 1 Then
        Else

            DeltaDelayVal(i) = CalcDutyVal(i).Math.Subtract(CalcDutyVal(i - 1))
            For j = 1 To DeltaDelayVal(i).pins.Count - 1
                If InStr(UCase(DeltaDelayVal(i).pins(j)), "_P") <> 0 Then
                    PinName = DeltaDelayVal(i).pins(j)
                    TestNameInput = Replace(LCase(PinName), "ddr", "ch")
                    TestNameInput = Replace(LCase(TestNameInput), "dqs_p", "core")
                    TestNameInput = TestNameInput & "_" & Freq_TestName_Input
                    
                    For Each site In TheExec.sites
                        If b_DivideZeroError(site) = True Then
''                            TheExec.Flow.TestLimit resultVal:=-999, ScaleType:=scalePico, Tname:="Delta_Delay_" & CStr(i - 2) & "_" & TestNameInput, ForceResults:=tlForceNone
                            DeltaDelayVal(i).pins(j).value = -999
                        End If
                    Next site
                    TestNameInput = Report_TName_From_Instance(CalcF, DeltaDelayVal(i).pins(j), vbNullString, 0)
                    TheExec.flow.TestLimit resultVal:=DeltaDelayVal(i).pins(j), scaletype:=scalePico, Tname:=TestNameInput, ForceResults:=tlForceNone_CZ
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
        
        For j = 1 To DeltaDelayVal(i).pins.Count - 1
            
            If InStr(UCase(DeltaDelayVal(i).pins(j)), "_P") <> 0 Then
                PinName = DeltaDelayVal(i).pins(j)
                TestNameInput = Replace(LCase(PinName), "ddr", "ch")
                TestNameInput = Replace(LCase(TestNameInput), "dqs_p", "core")
                TestNameInput = TestNameInput & "_" & Freq_TestName_Input
                
                For Each site In TheExec.sites
    
                   If DNL_Val(i).pins(j).value > DNL_Val_Max.pins(j).value Then
                       DNL_Val_Max.pins(j).value = DNL_Val(i).pins(j).value
                   End If
                   If DNL_Val(i).pins(j).value < DNL_Val_Min.pins(j).value Then
                       DNL_Val_Min.pins(j).value = DNL_Val(i).pins(j).value
                   End If
                   
                    If b_DivideZeroError(site) = True Then
                        DNL_Val(i).pins(j).value = -999
                    End If
                    
                       Select Case UCase(Freq_Dll_Str)
                           Case "DDR_F0"
                               If b_DivideZeroError(site) = True Then
                                    TestNameInput = Report_TName_From_Instance(CalcF, DNL_Val(i).pins(j), vbNullString, 0)
                                    TheExec.flow.TestLimit resultVal:=-999, lowVal:=-1, hiVal:=1, scaletype:=scaleNoScaling, unit:=unitCustom, customUnit:="LSB", Tname:=TestNameInput, ForceResults:=tlForceNone_CZ
                               Else
                                    TestNameInput = Report_TName_From_Instance(CalcF, DNL_Val(i).pins(j), vbNullString, 0)
                                    TheExec.flow.TestLimit resultVal:=DNL_Val(i).pins(j).value, lowVal:=-1, hiVal:=1, scaletype:=scaleNoScaling, unit:=unitCustom, customUnit:="LSB", Tname:=TestNameInput, ForceResults:=tlForceNone_CZ
                               End If
                               If DNL_Val(i).pins(j).value > 1 Or DNL_Val(i).pins(j).value < -1 Then
                                   b_Linearity_Fail = True
                               End If
                           Case "DDR_F1", "DDR_F1M9"
                               If b_DivideZeroError(site) = True Then
                                   TestNameInput = Report_TName_From_Instance(CalcF, DNL_Val(i).pins(j), vbNullString, 0)
                                   TheExec.flow.TestLimit resultVal:=-999, lowVal:=-1, hiVal:=1, scaletype:=scaleNoScaling, unit:=unitCustom, customUnit:="LSB", Tname:=TestNameInput, ForceResults:=tlForceNone_CZ
                               Else
                                   TestNameInput = Report_TName_From_Instance(CalcF, DNL_Val(i).pins(j), vbNullString, 0)
                                   TheExec.flow.TestLimit resultVal:=DNL_Val(i).pins(j).value, lowVal:=-1, hiVal:=1, scaletype:=scaleNoScaling, unit:=unitCustom, customUnit:="LSB", Tname:=TestNameInput, ForceResults:=tlForceNone_CZ
                               End If
                               If DNL_Val(i).pins(j).value > 1 Or DNL_Val(i).pins(j).value < -1 Then
                                   b_Linearity_Fail = True
                               End If
                           Case "DDR_F2", "DDR_F2M9"
                               If b_DivideZeroError(site) = True Then
                                   TestNameInput = Report_TName_From_Instance(CalcF, DNL_Val(i).pins(j), vbNullString, 0)
                                   TheExec.flow.TestLimit resultVal:=-999, lowVal:=-1, hiVal:=1.5, scaletype:=scaleNoScaling, unit:=unitCustom, customUnit:="LSB", Tname:=TestNameInput, ForceResults:=tlForceNone_CZ
                               Else
                                   TestNameInput = Report_TName_From_Instance(CalcF, DNL_Val(i).pins(j), vbNullString, 0)
                                   TheExec.flow.TestLimit resultVal:=DNL_Val(i).pins(j).value, lowVal:=-1, hiVal:=1.5, scaletype:=scaleNoScaling, unit:=unitCustom, customUnit:="LSB", Tname:=TestNameInput, ForceResults:=tlForceNone_CZ
                               End If
                               If DNL_Val(i).pins(j).value > 1.5 Or DNL_Val(i).pins(j).value < -1 Then
                                   b_Linearity_Fail = True
                               End If
                       End Select
                Next site
                
                ''20170818-Sum of DNL to be INL
                 If i = 0 + AryShiftNum Then
                    INL(i).pins(j) = INL(i).pins(j).Add(DNL_Val(i).pins(j))
                Else
                    INL(i).pins(j) = INL(i).pins(j).Add(DNL_Val(i).pins(j)).Add(INL(i - 1).pins(j))
                End If

                ''20170830 - Bypass
'                TheExec.Flow.TestLimit resultVal:=DNL_Val(i).Pins(j), ScaleType:=scaleNoScaling, unit:=unitCustom, customUnit:="LSB", Tname:="DNL" & CStr(i - 2) & "_" & TestNameInput, ForceResults:=tlForceNone
            End If
            
        Next j
    Next i
    
''    If i = No_Of_Valid_Delta_Delay + AryShiftNum Then
    For j = 1 To DNL_Val_Max.pins.Count - 1
            
        If InStr(UCase(DNL_Val_Max.pins(j)), "_P") <> 0 Then
            PinName = DNL_Val_Max.pins(j)
            TestNameInput = Replace(LCase(PinName), "ddr", "ch")
            TestNameInput = Replace(LCase(TestNameInput), "dqs_p", "core")
            TestNameInput = TestNameInput & "_" & Freq_TestName_Input
            
            For Each site In TheExec.sites
                If b_DivideZeroError(site) = True Then
                    DNL_Val_Max.pins(j).value = -999
                    DNL_Val_Min.pins(j).value = -999
                End If
            Next site
            

            TestNameInput = Report_TName_From_Instance(CalcF, DNL_Val_Max.pins(j), vbNullString, 0)

            TheExec.flow.TestLimit resultVal:=DNL_Val_Max.pins(j), Tname:=TestNameInput, scaletype:=scaleNoScaling, unit:=unitCustom, customUnit:="LSB", ForceResults:=tlForceNone_CZ
            
            TestNameInput = Report_TName_From_Instance(CalcF, DNL_Val_Min.pins(j), vbNullString, 0)
            TheExec.flow.TestLimit resultVal:=DNL_Val_Min.pins(j), Tname:=TestNameInput, scaletype:=scaleNoScaling, unit:=unitCustom, customUnit:="LSB", ForceResults:=tlForceNone_CZ
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
       
        For j = 1 To DeltaDelayVal(i).pins.Count - 1
            If InStr(UCase(DeltaDelayVal(i).pins(j)), "_P") <> 0 Then
                PinName = DeltaDelayVal(i).pins(j)
                TestNameInput = Replace(LCase(PinName), "ddr", "ch")
                TestNameInput = Replace(LCase(TestNameInput), "dqs_p", "core")
                TestNameInput = TestNameInput & "_" & Freq_TestName_Input
                
                For Each site In TheExec.sites
                   If INL(i).pins(j).value > INL_Val_Max.pins(j).value Then
                       INL_Val_Max.pins(j).value = INL(i).pins(j).value
                   End If
                   If INL(i).pins(j).value < INL_Val_Min.pins(j).value Then
                       INL_Val_Min.pins(j).value = INL(i).pins(j).value
                   End If
                                     
                    TestNameInput = Report_TName_From_Instance(CalcF, INL_Val_Min.pins(j), vbNullString, 0)
                    TheExec.flow.TestLimit resultVal:=INL(i).pins(j).value, scaletype:=scaleNoScaling, unit:=unitCustom, customUnit:="LSB", Tname:=TestNameInput, ForceResults:=tlForceNone_CZ
               Next site
             End If
             
        Next j
    Next i
    
    For j = 1 To INL_Val_Max.pins.Count - 1
            
        If InStr(UCase(INL_Val_Max.pins(j)), "_P") <> 0 Then
            PinName = INL_Val_Max.pins(j)
            TestNameInput = Replace(LCase(PinName), "ddr", "ch")
            TestNameInput = Replace(LCase(TestNameInput), "dqs_p", "core")
            TestNameInput = TestNameInput & "_" & Freq_TestName_Input
            
            For Each site In TheExec.sites
                If b_DivideZeroError(site) = True Then
                    DNL_Val_Max.pins(j).value = -999
                    DNL_Val_Min.pins(j).value = -999
                End If
            Next site
            
            TestNameInput = Report_TName_From_Instance(CalcF, DNL_Val_Max.pins(j), vbNullString, 0)
                    
            TheExec.flow.TestLimit resultVal:=INL_Val_Max.pins(j), Tname:=TestNameInput, scaletype:=scaleNoScaling, unit:=unitCustom, customUnit:="LSB", ForceResults:=tlForceNone_CZ
            
            TestNameInput = Report_TName_From_Instance(CalcF, DNL_Val_Min.pins(j), vbNullString, 0)
                    
            TheExec.flow.TestLimit resultVal:=INL_Val_Min.pins(j), Tname:=TestNameInput, scaletype:=scaleNoScaling, unit:=unitCustom, customUnit:="LSB", ForceResults:=tlForceNone_CZ
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
''                TheExec.Flow.TestLimit resultVal:=INL.Pins(j), Tname:="INL" & "_" & TestNameInput, ScaleType:=scaleNoScaling, unit:=unitCustom, customUnit:="LSB", ForceResults:=tlForceNone
''                Counter = Counter + 1
''            End If
''        Next j
''    Next i
''    TheExec.Flow.TestLimit resultVal:=INL_Val_Max, Tname:="INL_MAX" & "_" & Freq_TestName_Input, ScaleType:=scaleNoScaling, unit:=unitCustom, customUnit:="LSB", ForceResults:=tlForceNone
''    TheExec.Flow.TestLimit resultVal:=INL_Val_Min, Tname:="INL_MIN" & "_" & Freq_TestName_Input, ScaleType:=scaleNoScaling, unit:=unitCustom, customUnit:="LSB", ForceResults:=tlForceNone

'''''    For Each Site In TheExec.Sites
'''''         If b_DivideZeroError(Site) = True Then
'''''            TheExec.Flow.TestLimit resultVal:=-999, lowVal:=False, hiVal:=False, Tname:="Linearity_Pass" & "_" & TestNameInput, ForceResults:=tlForceNone
'''''         Else
'''''            TheExec.Flow.TestLimit resultVal:=b_Linearity_Fail, lowVal:=False, hiVal:=False, Tname:="Linearity_Pass" & "_" & TestNameInput, ForceResults:=tlForceNone
'''''        End If
'''''    Next Site
    
End Function

Public Function CalcDelayDelta_Sicily(argc As Integer, argv() As String) As Long
On Error GoTo errHandler 'Add ErrHandler 2023/05/29

Dim DDR0_MC_DQS_DIFFx_F() As New PinListData
Dim SiteDouble_Frequency() As New SiteDouble
Dim SiteDouble_Delay() As New SiteDouble
Dim SiteDouble_Delta() As New SiteDouble

Dim NumberOfFreq As Long
Dim i, j As Long
Dim site As Variant
Dim TestNameInput As String

NumberOfFreq = CLng(argv(2)) - CLng(argv(1)) + 1

ReDim DDR0_MC_DQS_DIFFx_F(NumberOfFreq - 1) As New PinListData
ReDim SiteDouble_Frequency(NumberOfFreq - 1) As New SiteDouble
ReDim SiteDouble_Delay(NumberOfFreq - 1) As New SiteDouble
ReDim SiteDouble_Delta(NumberOfFreq - 2) As New SiteDouble


For i = 0 To NumberOfFreq - 1
    DDR0_MC_DQS_DIFFx_F(i) = GetStoreDataAllType(argv(0) & i)
    For j = 0 To DDR0_MC_DQS_DIFFx_F(i).pins.Count - 1
        If InStr(UCase(DDR0_MC_DQS_DIFFx_F(i).pins(j)), "DQS_P") <> 0 Then
            For Each site In TheExec.sites
                If DDR0_MC_DQS_DIFFx_F(i).pins(j).value = 0 Then
                    SiteDouble_Frequency(i) = 0.000000001
                    TheExec.Datalog.WriteComment "Site" & site & " : DDR F" & i & " frequency is 0"
                Else
                    SiteDouble_Frequency(i) = DDR0_MC_DQS_DIFFx_F(i).pins(j).value
                End If
            Next site
        End If
    Next j
Next i


For i = 0 To NumberOfFreq - 1
    SiteDouble_Delay(i) = SiteDouble_Frequency(i).Multiply(2).Invert
    TestNameInput = Report_TName_From_Instance(CalcF, vbNullString, , 0)
    TheExec.flow.TestLimit resultVal:=SiteDouble_Delay(i), Tname:=TestNameInput, ForceResults:=tlForceFlow, scaletype:=scalePico
Next i

For i = 0 To NumberOfFreq - 2
    SiteDouble_Delta(i) = SiteDouble_Delay(i + 1).Subtract(SiteDouble_Delay(i))
    TestNameInput = Report_TName_From_Instance(CalcF, vbNullString, , 0)
    TheExec.flow.TestLimit resultVal:=SiteDouble_Delta(i), Tname:=TestNameInput, ForceResults:=tlForceFlow, scaletype:=scalePico
Next i
Exit Function 'Add ErrHandler 2023/05/29errHandler: 'Add ErrHandler 2023/05/29
errHandler: 'Add ErrHandler 2023/05/29
    Call Print_Error_Message(Error_Info, "LIB_HardIP_Calc", "CalcDelayDelta_Sicily") 'Add ErrHandler 2023/05/29
    If AbortTest Then Exit Function Else Resume Next 'Add ErrHandler 2023/05/29

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
        CalcDutyVal(i) = GetStoreDataAllType(argv(i))
        If TheExec.TesterMode = testModeOffline Then
            For j = 0 To CalcDutyVal(i).pins.Count - 1
                CalcDutyVal(i).pins(j) = 1000000 - 1000 * j - i * 2000
            Next j
        End If
        'For j = 1 To CalcDutyVal(i).Pins.Count - 1 Step 2
        For j = 0 To CalcDutyVal(i).pins.Count - 1 Step 1           'Modify 20170908
            If j Mod 4 = 2 Or j Mod 4 = 3 Then                                  'Modify 20170908
                
                PinName = CalcDutyVal(i).pins(j)
                For Each site In TheExec.sites
    
                    If CalcDutyVal(i).pins(j).value(site) = 0 Then
                        b_DivideZeroError(site) = True
                       If gl_Disable_HIP_debug_log = False Then TheExec.Datalog.WriteComment ("Site " & site & " Freq Meas 0 Hz , No CalcDutyDelay ")
                        CalcDutyVal(i).pins(j).value = 1
                    End If
                    
                    CalcDutyVal(i).pins(j).value = CalcDutyVal(i).pins(j).Multiply(2).Invert
                    
                    If b_DivideZeroError(site) = True Then
                        TestNameInput = Report_TName_From_Instance(CalcF, CalcDutyVal(i).pins(j), vbNullString, 0, i)
                        TheExec.flow.TestLimit resultVal:=-999, scaletype:=scalePico, PinName:=PinName, Tname:=TestNameInput, ForceResults:=tlForceFlow 'transfer_to_forceflow
                    Else
                        TestNameInput = Report_TName_From_Instance(CalcF, CalcDutyVal(i).pins(j), vbNullString, 0, i)
                        TheExec.flow.TestLimit resultVal:=CalcDutyVal(i).pins(j).value, scaletype:=scalePico, PinName:=PinName, Tname:=TestNameInput, ForceResults:=tlForceFlow 'transfer_to_forceflow
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
            For j = 0 To DeltaDelayVal(i).pins.Count - 1 Step 1     'Modify 20170908
                If j Mod 4 = 2 Or j Mod 4 = 3 Then                                  'Modify 20170908
                
                    PinName = DeltaDelayVal(i).pins(j)
                    For Each site In TheExec.sites
                        If b_DivideZeroError(site) = True Then
                            TestNameInput = Report_TName_From_Instance(CalcF, DeltaDelayVal(i).pins(j), vbNullString, 0, i)
                            TheExec.flow.TestLimit resultVal:=-999, scaletype:=scalePico, PinName:=PinName, Tname:=TestNameInput, ForceResults:=tlForceFlow 'transfer_to_forceflow
                        Else
                            TestNameInput = Report_TName_From_Instance(CalcF, DeltaDelayVal(i).pins(j), vbNullString, 0, i)
                            TheExec.flow.TestLimit resultVal:=DeltaDelayVal(i).pins(j).value, scaletype:=scalePico, PinName:=PinName, Tname:=TestNameInput, ForceResults:=tlForceFlow 'transfer_to_forceflow
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
        CalcDutyVal(i) = GetStoreDataAllType(argv(i))
        If TheExec.TesterMode = testModeOffline Then
            For j = 0 To CalcDutyVal(i).pins.Count - 1
                CalcDutyVal(i).pins(j) = 1000000 - 1000 * j - i * 2000
            Next j
        End If
        For j = 1 To CalcDutyVal(i).pins.Count - 1
            If InStr(UCase(CalcDutyVal(i).pins(j)), "_P") <> 0 Then
                PinName = CalcDutyVal(i).pins(j)
                TestNameInput = Replace(LCase(PinName), "ddr", "ch")
                TestNameInput = Replace(LCase(TestNameInput), "dqs_p", "core")
                
                For Each site In TheExec.sites
                    If CalcDutyVal(i).pins(j).value(site) = 0 Then
                        b_DivideZeroError(site) = True
                        If gl_Disable_HIP_debug_log = False Then TheExec.Datalog.WriteComment ("Site " & site & " Freq Meas 0 Hz , No CalcDutyDelay ")
                        CalcDutyVal(i).pins(j).value = 1
                    End If
                Next site
                    
                CalcDutyVal(i).pins(j).value = CalcDutyVal(i).pins(j).Multiply(2).Invert
                    
                For Each site In TheExec.sites
                    If b_DivideZeroError(site) = True Then
''                        TheExec.Flow.TestLimit resultVal:=-999, ScaleType:=scalePico, Tname:="Delay" & "_" & TestNameInput, ForceResults:=tlForceNone
                        CalcDutyVal(i).pins(j).value = -999
                    End If
                Next site
                TestNameInput = Report_TName_From_Instance(CalcF, CalcDutyVal(i).pins(j), vbNullString, 0)
                TheExec.flow.TestLimit resultVal:=CalcDutyVal(i).pins(j), scaletype:=scalePico, Tname:=TestNameInput, ForceResults:=tlForceFlow 'transfer_to_forceflow
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
    
    CalcDutyVal_1 = GetStoreDataAllType(Dict_CalcDutyVal_1)
    CalcDutyVal_2 = GetStoreDataAllType(Dict_CalcDutyVal_2)
    Dim PinName As String
    
    For j = 1 To CalcDutyVal_1.pins.Count - 1
        If InStr(UCase(CalcDutyVal_1.pins(j)), "_P") <> 0 Then
            PinName = CalcDutyVal_1.pins(j)
            TestNameInput = Replace(LCase(PinName), "ddr", "ch")
            TestNameInput = Replace(LCase(TestNameInput), "dqs_p", "core")
            
            For Each site In TheExec.sites
            
                If CalcDutyVal_1.pins(j).value(site) = 0 Then
                    b_DivideZeroError(site) = True
                    If gl_Disable_HIP_debug_log = False Then If gl_Disable_HIP_debug_log = False Then TheExec.Datalog.WriteComment ("Site " & site & " Freq Meas 0 Hz , No CalcDutyDelay ")
                    CalcDutyVal_1.pins(j).value = 1
                End If
                If CalcDutyVal_2.pins(j).value(site) = 0 Then
                    b_DivideZeroError(site) = True
                    TheExec.Datalog.WriteComment ("Site " & site & " Freq Meas 0 Hz , No CalcDutyDelay ")
                    CalcDutyVal_2.pins(j).value = 1
                End If
            Next site
            
            CalcDutyVal_1.pins(j).value = CalcDutyVal_1.pins(j).Multiply(2).Invert
            CalcDutyVal_2.pins(j).value = CalcDutyVal_2.pins(j).Multiply(2).Invert
        End If
    Next j
    
    CalcDuty_Diff = CalcDutyVal_1.Math.Subtract(CalcDutyVal_2)
    
    For j = 1 To CalcDuty_Diff.pins.Count - 1
        If InStr(UCase(CalcDuty_Diff.pins(j)), "_P") <> 0 Then
            PinName = CalcDuty_Diff.pins(j)
            TestNameInput = Replace(LCase(PinName), "ddr", "ch")
            TestNameInput = Replace(LCase(TestNameInput), "dqs_p", "core")
            For Each site In TheExec.sites
                If b_DivideZeroError(site) = True Then
''                    TheExec.Flow.TestLimit resultVal:=-999, ScaleType:=scalePico, Tname:="Jitter" & "_" & TestNameInput, ForceResults:=tlForceNone
                    CalcDuty_Diff.pins(j).value = -999
                End If
            Next site

            TestNameInput = Report_TName_From_Instance(CalcF, CalcDuty_Diff.pins(j), vbNullString, 0)
            TheExec.flow.TestLimit resultVal:=CalcDuty_Diff.pins(j), scaletype:=scalePico, Tname:=TestNameInput, ForceResults:=tlForceNone 'transfer_to_forceflow
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
        
        DSP_DictKey_2S_BIN = GetStoreDataAllType(DictKey_2S_BIN)
        
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
        
        Call StoreDataAllType(DictKey_SIGN_DEC, DSP_DictKey_SIGN_DEC(i))
        
''        TheExec.Flow.TestLimit resultVal:=DSP_DictKey_DEC.Element(0), Tname:="DEC_" & i, ForceResults:=tlForceFlow
        
        TestNameInput = Report_TName_From_Instance(CalcC, "X", , CInt(i))

        TheExec.flow.TestLimit resultVal:=DSP_DictKey_SIGN_DEC(i).Element(0), Tname:=TestNameInput, ForceResults:=tlForceFlow
        
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

Call HardIP_Bin2Dec(DataOut_TemperatureCode, GetStoreDataAllType(argv(0)))

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
    TheExec.flow.TestLimit resultVal:=DataOut_Temperature, lowVal:=15, hiVal:=35, ForceResults:=tlForceFlow 'transfer_to_forceflow
    Update_BC_PassFail_Flag
    
    If TheExec.CurrentJob = "CP1" Then
    Else: TheHdw.Wait 0.15
    End If
Else
    TestNameInput = Report_TName_From_Instance(CalcT, "X", , 0)
    TheExec.flow.TestLimit resultVal:=DataOut_Temperature, ForceResults:=tlForceFlow, Tname:=TestNameInput
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
    DataOut_ADC_Code_Average_Dict = GetStoreDataAllType("ADC_OFFSET_AVERAGE_X")
End If

Call HardIP_Bin2Dec(DataOut_ADC_Code_0, GetStoreDataAllType(argv(0)))
Call HardIP_Bin2Dec(DataOut_ADC_Code_1, GetStoreDataAllType(argv(1)))

For Each site In TheExec.sites
    DataOut_ADC_Voltage_0(site) = TheHdw.DCVS.pins("VDD12_PCIE").Voltage.value * DataOut_ADC_Code_0(site).Element(0) / 255
    DataOut_ADC_Voltage_1(site) = TheHdw.DCVS.pins("VDD12_PCIE").Voltage.value * DataOut_ADC_Code_1(site).Element(0) / 255
    DataOut_ADC_Voltage_Average(site) = (DataOut_ADC_Voltage_0(site) + DataOut_ADC_Voltage_1(site)) / 2
    If argv(0) Like "*adc_offset_adc*" Then
        DataOut_ADC_Code_0_OffSet(site) = DataOut_ADC_Code_0(site).Element(0) - 128
        DataOut_ADC_Code_1_OffSet(site) = DataOut_ADC_Code_1(site).Element(0) - 128
        DataOut_ADC_Code_OffSet_Average(site) = (DataOut_ADC_Code_0_OffSet(site) + DataOut_ADC_Code_1_OffSet(site)) / 2
        DataOut_ADC_Code_Average_Dict(site).Element(0) = DataOut_ADC_Code_OffSet_Average(site)
    Else
    DataOut_ADC_Code_Average(site) = (DataOut_ADC_Code_0(site).Element(0) + DataOut_ADC_Code_1(site).Element(0)) / 2
    DataOut_ADC_Code_Final(site) = DataOut_ADC_Code_Average(site) - DataOut_ADC_Code_Average_Dict(site).Element(0)
    DataOut_ADC_Voltage_Out(site) = 0.25 * TheHdw.DCVS.pins("VDD12_PCIE").Voltage.value + DataOut_ADC_Code_Final(site) * TheHdw.DCVS.pins("VDD12_PCIE").Voltage.value * 0.5 / 256
    End If
Next site

If argv(0) Like "*adc_offset_adc*" Then
    Call StoreDataAllType("ADC_OFFSET_AVERAGE_X", DataOut_ADC_Code_Average_Dict)
End If

Str_Split = Split(argv(0), "_")

TheExec.flow.TestLimit resultVal:=DataOut_ADC_Voltage_0, Tname:="Voltage_" & argv(0), ForceResults:=tlForceFlow
TheExec.flow.TestLimit resultVal:=DataOut_ADC_Voltage_1, Tname:="Voltage_" & argv(1), ForceResults:=tlForceFlow
TheExec.flow.TestLimit resultVal:=DataOut_ADC_Voltage_Average, Tname:="Average_Voltage_" & Str_Split(1) & "_" & Str_Split(2) & "_adc", ForceResults:=tlForceFlow

If argv(0) Like "*adc_offset_adc*" Then
    TheExec.flow.TestLimit resultVal:=DataOut_ADC_Code_0_OffSet, Tname:="OffSet_" & argv(0), ForceResults:=tlForceFlow
    TheExec.flow.TestLimit resultVal:=DataOut_ADC_Code_1_OffSet, Tname:="OffSet_" & argv(1), ForceResults:=tlForceFlow
    TheExec.flow.TestLimit resultVal:=DataOut_ADC_Code_OffSet_Average, Tname:="Average_OffSet_" & Str_Split(1) & "_" & Str_Split(2) & "_adc", ForceResults:=tlForceFlow
Else
    TheExec.flow.TestLimit resultVal:=DataOut_ADC_Code_Average, Tname:="Average_" & Str_Split(1) & "_" & Str_Split(2) & "_adc", ForceResults:=tlForceFlow
    TheExec.flow.TestLimit resultVal:=DataOut_ADC_Code_Final, Tname:="Final_" & Str_Split(1) & "_" & Str_Split(2) & "_adc", ForceResults:=tlForceFlow
    TheExec.flow.TestLimit resultVal:=DataOut_ADC_Voltage_Out, Tname:="Voltage_Out_" & Str_Split(1) & "_" & Str_Split(2) & "_adc", ForceResults:=tlForceFlow
End If

End Function

Public Function Calc_LPDPRX_Bin2Hex(argc As Integer, argv() As String) As Long
Dim i As Integer
Dim Data_Temp As String
Dim DSPWave_Dict As New DSPWave: DSPWave_Dict = GetStoreDataAllType(argv(0))
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

Public Function TX_Level(argc As Integer, argv() As String) As Long
On Error GoTo errHandler 'Add ErrHandler 2023/05/29
    
    ''20170711
    '' Calculate V_DP_K =  ((V_1 I_2-V_2 I_1 ) R_term)/(V_1-V_2+(R_term-R_(path) )(I_2-I_1 ) )
    '' where R_term = 45ohm; Rpath = trace + RAK
    ''TheExec.Datalog.WriteComment "Some Error in " & TheExec.DataManager.InstanceName
    
    Dim site As Variant
    Dim i As Long, j As Long
    Dim DictKey_V1 As String, DictKey_V2 As String
    Dim pld_V1 As New PinListData, pld_V2 As New PinListData
    Dim i1 As Double, i2 As Double
    Dim PinName As String
    Dim R_Term As Double
    Dim DictKey_Diff As String, DictKey_Diff_Calc As String
    Dim V_Diff As New SiteDouble
    
    DictKey_V1 = argv(0)
    DictKey_V2 = argv(1)
    i1 = CDbl(argv(2))
    i2 = CDbl(argv(3))
    PinName = argv(4)
    R_Term = CDbl(argv(5))
    
    If argc = 7 Then
        If argv(6) <> "" Then
            DictKey_Diff = argv(6)
        End If
    End If
    
    If argc = 8 Then
        If argv(7) <> "" Then
            'DictKey_Diff_Calc = argv(7) 'For Sicily
            DictKey_Diff_Calc = argv(6) ' For Turks
        End If
    End If
    
    pld_V1 = GetStoreDataAllType(DictKey_V1)
    pld_V2 = GetStoreDataAllType(DictKey_V2)
    
    Dim R_Path As New SiteDouble
    'Dim R_Channel_RAK() As Double
    For Each site In TheExec.sites
        'R_Channel_RAK = TheHdw.PPMU.ReadRakValuesByPinnames(PinName, site)
        R_Path(site) = CurrentJob_Card_RAK.pins(PinName).value(site)
    Next site
    
    Dim V_DP_K As New SiteDouble
    
    For Each site In TheExec.sites
        V_DP_K(site) = ((pld_V1.pins(PinName).value(site) * i2 - pld_V2.pins(PinName).value(site) * i1) * R_Term) / (pld_V1.pins(PinName).value(site) - pld_V2.pins(PinName).value(site) + (R_Term - R_Path(site)) * (i2 - i1))
    Next site
    Dim OutputTname_format() As String
    Dim TestNameInput As String
    
''    TheExec.Flow.TestLimit resultVal:=V_DP_K, PinName:=PinName, Tname:="Volt_meas_TX_Level", ForceResults:=tlForceFlow
    TestNameInput = Report_TName_From_Instance(CalcV, PinName, vbNullString, CInt(j))
    TheExec.flow.TestLimit resultVal:=V_DP_K, PinName:=PinName, ForceResults:=tlForceFlow, Tname:=TestNameInput
    If argc = 7 Then
        If argv(6) <> "" Then
            Call StoreDataAllType(DictKey_Diff, V_DP_K)
        End If
    End If
    
    If argc = 8 Then
        If argv(7) <> "" Then
            V_Diff = GetStoreDataAllType(DictKey_Diff_Calc)
            'V_Diff = V_Diff.Subtract(V_DP_K)
            V_Diff = V_Diff.Subtract(V_DP_K).Abs 'For Tonga
            TestNameInput = Report_TName_From_Instance(CalcV, PinName, vbNullString, CInt(j))
            TheExec.flow.TestLimit resultVal:=V_Diff, PinName:=PinName, ForceResults:=tlForceFlow, Tname:=TestNameInput
        End If
    End If
    
''    Dim TX_V1a As New PinListData
''    Dim TX_V1b As New PinListData
''    Dim TX_V1c As New PinListData
''    Dim TX_V1d As New PinListData
''
''    Dim TX_V2a As New PinListData
''    Dim TX_V2b As New PinListData
''    Dim TX_V2c As New PinListData
''    Dim TX_V2d As New PinListData
''
''    Dim TX0P As New SiteDouble
''    Dim TX1P As New SiteDouble
''
''    Dim TX0M As New SiteDouble
''    Dim TX1M As New SiteDouble
''
''    Dim R_path As Long
''    Dim R_term As Long
''    Dim Rdiff As New PinListData
''    Dim r_trace_TX0P() As Double
''    Dim r_trace_TX1P() As Double
''    Dim r_trace_TX0M() As Double
''    Dim r_trace_TX1M() As Double
''    Dim I1 As Double
''    Dim I2 As Double
''
''    R_path = 0
''    R_term = 50
''    I1 = 0.007
''    I2 = 0.009
''    Rdiff.AddPin ("TX0_P")
''    Rdiff.AddPin ("TX1_P")
''    Rdiff.AddPin ("TX0_M")
''    Rdiff.AddPin ("TX1_M")
''    For Each Site In TheExec.sites.Active
''        r_trace_TX0P = TheHdw.PPMU.ReadRakValuesByPinnames("TX0_P", Site)
''        r_trace_TX1P = TheHdw.PPMU.ReadRakValuesByPinnames("TX1_P", Site)
''        r_trace_TX0M = TheHdw.PPMU.ReadRakValuesByPinnames("TX0_M", Site)
''        r_trace_TX1M = TheHdw.PPMU.ReadRakValuesByPinnames("TX1_M", Site)
''
''        Rdiff.Pins("TX0_P") = R_term - CP_Card_RAK.Pins("TX0_P").Value - r_trace_TX0P(0)
''        Rdiff.Pins("TX1_P") = R_term - CP_Card_RAK.Pins("TX1_P").Value - r_trace_TX1P(0)
''        Rdiff.Pins("TX0_M") = R_term - CP_Card_RAK.Pins("TX0_M").Value - r_trace_TX0M(0)
''        Rdiff.Pins("TX1_M") = R_term - CP_Card_RAK.Pins("TX1_M").Value - r_trace_TX1M(0)
''    Next Site
''
''    TX_V1a = GetStoreDataAllType(argv(0))  'TX0_P,0.07,I1,TX0P
''    TX_V1b = GetStoreDataAllType(argv(1))  'TX1_P,0.07,I1,TX1P
''    TX_V1c = GetStoreDataAllType(argv(2))  'TX0_M,0.07,I1,TX0M
''    TX_V1d = GetStoreDataAllType(argv(3))  'TX1_M,0.07,I1,TX1M
''
''    TX_V2a = GetStoreDataAllType(argv(4))  'TX0_P,0.09,I2,TX0P
''    TX_V2b = GetStoreDataAllType(argv(5))  'TX1_P,0.09,I2,TX1P
''    TX_V2c = GetStoreDataAllType(argv(6))  'TX0_P,0.09,I2,TX0M
''    TX_V2d = GetStoreDataAllType(argv(7))  'TX1_P,0.09,I2,TX1M
''
''
'''(V1I2-V2I1)Rterm /  V1-V2+(Rterm-Rpath)(I2-I1)
''
''    For Each Site In TheExec.sites.Active
''        TX0P = ((TX_V1a.Pins("TX0_P").Value(Site) * I2 - TX_V2a.Pins("TX0_P").Value(Site) * I1) * R_term) / ((TX_V1a.Pins("TX0_P").Value(Site) - TX_V2a.Pins("TX0_P").Value(Site)) + Rdiff.Pins("TX0_P").Value(Site) * (I2 - I1))
''        TX1P = ((TX_V1b.Pins("TX1_P").Value(Site) * I2 - TX_V2b.Pins("TX1_P").Value(Site) * I1) * R_term) / ((TX_V1b.Pins("TX1_P").Value(Site) - TX_V2b.Pins("TX1_P").Value(Site)) + Rdiff.Pins("TX1_P").Value(Site) * (I2 - I1))
''        TX0M = ((TX_V1c.Pins("TX0_M").Value(Site) * I2 - TX_V2c.Pins("TX0_M").Value(Site) * I1) * R_term) / ((TX_V1c.Pins("TX0_M").Value(Site) - TX_V2c.Pins("TX0_M").Value(Site)) + Rdiff.Pins("TX0_M").Value(Site) * (I2 - I1))
''        TX1M = ((TX_V1d.Pins("TX1_M").Value(Site) * I2 - TX_V2d.Pins("TX1_M").Value(Site) * I1) * R_term) / ((TX_V1d.Pins("TX1_M").Value(Site) - TX_V2d.Pins("TX1_M").Value(Site)) + Rdiff.Pins("TX1_M").Value(Site) * (I2 - I1))
''
''    Next Site
''
''    TheExec.Flow.TestLimit resultVal:=TX0P, Tname:="TX0P_Level_H", ForceResults:=tlForceFlow
''    TheExec.Flow.TestLimit resultVal:=TX1P, Tname:="TX1P_Level_H", ForceResults:=tlForceFlow
''    TheExec.Flow.TestLimit resultVal:=TX0M, Tname:="TX0M_Level_H", ForceResults:=tlForceFlow
''    TheExec.Flow.TestLimit resultVal:=TX1M, Tname:="TX1M_Level_H", ForceResults:=tlForceFlow
        
Exit Function 'Add ErrHandler 2023/05/29errHandler: 'Add ErrHandler 2023/05/29
errHandler: 'Add ErrHandler 2023/05/29
    Call Print_Error_Message(Error_Info, "LIB_HardIP_Calc", "TX_Level") 'Add ErrHandler 2023/05/29
    If AbortTest Then Exit Function Else Resume Next 'Add ErrHandler 2023/05/29
End Function
Public Function TX_Level_Pingroup(argc As Integer, argv() As String) As Long
    ' [20230725][T-BraC][Neil] Function TTR - Support multi pins calculate
    
    ''20170711
    '' Calculate V_DP_K =  ((V_1 I_2-V_2 I_1 ) R_term)/(V_1-V_2+(R_term-R_(path) )(I_2-I_1 ) )
    '' where R_term = 45ohm; Rpath = trace + RAK
    ''TheExec.Datalog.WriteComment "Some Error in " & TheExec.DataManager.InstanceName
    On Error GoTo errHandler
    
    Dim site As Variant
    Dim i As Long, j As Long, k As Long
    Dim DictKey_V1 As String, DictKey_V2 As String
    Dim pld_V1 As New PinListData, pld_V2 As New PinListData
    Dim i1 As Double, i2 As Double
    Dim PinName As String
    Dim R_Term As Double
    Dim DictKey_Diff As String, DictKey_Diff_Calc As String
    Dim V_Diff As New PinListData
    Dim DictKey_1() As String
    Dim DictKey_2() As String
    
    Dim R_Path As New SiteDouble
    Dim V_DP_K As New SiteDouble

    
    Dim PLD_Calc_A1 As New PinListData
    Dim PLD_Calc_A2 As New PinListData
    Dim PLD_Calc_A As New PinListData
    
    Dim PLD_Calc_B1 As New PinListData
    Dim PLD_Calc_B As New PinListData
    
    Dim PLD_V_DP_K As New PinListData
    Dim PLD_V_DP_K_P As New PinListData
    Dim PLD_V_DP_K_N As New PinListData
    Dim PLD_R_Path As New PinListData
    Dim SD_I2_I1 As New SiteDouble
    
    DictKey_1 = Split(argv(0), "+")
    DictKey_2 = Split(argv(1), "+")
    
    i1 = CDbl(argv(2))
    i2 = CDbl(argv(3))
    R_Term = CDbl(argv(4))


    For k = 0 To UBound(DictKey_1)
        Set PLD_R_Path = Nothing
        Set PLD_Calc_A1 = Nothing
        Set PLD_Calc_A2 = Nothing
        Set PLD_Calc_A = Nothing
        Set PLD_Calc_B1 = Nothing
        Set PLD_Calc_B = Nothing
        Set PLD_V_DP_K = Nothing
        
        DictKey_V1 = DictKey_1(k)
        DictKey_V2 = DictKey_2(k)

        
        pld_V1 = GetStoredMeasurement(DictKey_V1)
        pld_V2 = GetStoredMeasurement(DictKey_V2)
        
        
        '' Calculate V_DP_K =  ((V_1 I_2-V_2 I_1 ) R_term)/(V_1-V_2+(R_term-R_(path) )(I_2-I_1 ) )
        PLD_Calc_A1 = pld_V1.Math.Multiply(i2)
        PLD_Calc_A2 = pld_V2.Math.Multiply(i1)
        SD_I2_I1 = i2 - i1
        PLD_Calc_A = PLD_Calc_A1.Math.Subtract(PLD_Calc_A2).Multiply(R_Term)
        
        For i = 0 To PLD_Calc_A.pins.Count - 1
            PinName = UCase(PLD_Calc_A.pins.item(i))
            PLD_R_Path.AddPin (PinName)
            PLD_R_Path.pins(PinName) = CurrentJob_Card_RAK.pins(PinName).Subtract(R_Term).Negate
        Next i
        
    
        PLD_Calc_B1 = PLD_R_Path.Math.Multiply(SD_I2_I1)
        PLD_Calc_B = pld_V1.Math.Subtract(pld_V2).Add(PLD_Calc_B1)
        PLD_V_DP_K = PLD_Calc_A.Math.divide(PLD_Calc_B)
       
        Dim OutputTname_format() As String
        Dim TestNameInput As String
        Dim TESTLINT_MAX As Long: TESTLINT_MAX = TheExec.flow.TestLimitIndex
        For i = 0 To PLD_Calc_A.pins.Count - 1
            PinName = UCase(PLD_Calc_A.pins.item(i))
            TheExec.flow.TestLimitIndex = TESTLINT_MAX
            TestNameInput = Report_TName_From_Instance(CalcV, PinName, vbNullString, CInt(j))
            TheExec.flow.TestLimit resultVal:=PLD_V_DP_K.pins(i), PinName:=PinName, ForceResults:=tlForceFlow, Tname:=TestNameInput
        Next i
        If k = 0 Then
            PLD_V_DP_K_P = PLD_V_DP_K
        Else
            PLD_V_DP_K_N = PLD_V_DP_K
        End If
    Next k
        
        TESTLINT_MAX = TheExec.flow.TestLimitIndex

        V_Diff = PLD_V_DP_K_N.Math.Subtract(PLD_V_DP_K_P).Abs
        For i = 0 To PLD_Calc_A.pins.Count - 1
            PinName = UCase(PLD_Calc_A.pins.item(i))
            TheExec.flow.TestLimitIndex = TESTLINT_MAX
            TestNameInput = Report_TName_From_Instance(CalcV, PinName, vbNullString, CInt(j))
            TheExec.flow.TestLimit resultVal:=V_Diff.pins(i), PinName:=PinName, ForceResults:=tlForceFlow, Tname:=TestNameInput
        Next i



Exit Function
errHandler:
    Call Print_Error_Message(Error_Info, "LIB_HardIP_Calc", "TX_Level_Pingroup")
    If AbortTest Then Exit Function Else Resume Next
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
            
    TX_Va_EQ = GetStoreDataAllType(argv(0))
    TX_Vb_EQ = GetStoreDataAllType(argv(1))
     TX_PM_EQ = TX_Va_EQ
'    TX_Va_EQ.AddPin ("Hello")
'    TX_Vb_EQ.AddPin ("Hi")
    If TheExec.TesterMode = testModeOffline Then
    
    For i = 0 To TX_PM_EQ.pins.Count - 1
        For Each site In TheExec.sites.Active

            TX_Va_EQ.pins(i).value(site) = 10
            TX_Vb_EQ.pins(i).value(site) = 10
            
        Next site
'
    Next i
    
    End If
    
   
    
        
    For i = 0 To TX_PM_EQ.pins.Count - 1
        For Each site In TheExec.sites.Active

            TX_PM_EQ.pins(i).value(site) = 20 * log(TX_Va_EQ.pins(i).value(site) / TX_Vb_EQ.pins(i).value(site))

        Next site
'
    Next i

    For j = 0 To TX_PM_EQ.pins.Count - 1
        For Each site In TheExec.sites.Active
                TestNameInput = Report_TName_From_Instance(CalcV, TX_PM_EQ.pins(j), vbNullString, CInt(j))
                TheExec.flow.TestLimit resultVal:=TX_PM_EQ.pins(j).value, Tname:=TestNameInput, ForceResults:=tlForceFlow 'transfer_to_forceflow
        Next site
    Next j




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
    
    TX_Va_EQ = GetStoreDataAllType(argv(0))
    TX_Vb_EQ = GetStoreDataAllType(argv(1))
    TX_PM_EQ = TX_Va_EQ
    
    If (argc = 3) Then
        TX_Vc_EQ = GetStoreDataAllType(argv(2))
        TX_PM_PRE = TX_Vc_EQ
    End If
    
    
    
    If TheExec.TesterMode = testModeOffline Then
    For i = 0 To TX_PM_EQ.pins.Count - 1
        For Each site In TheExec.sites.Active

            TX_Va_EQ.pins(i).value(site) = 10
            TX_Vb_EQ.pins(i).value(site) = 10
            If (argc = 3) Then
                TX_Vc_EQ.pins(i).value(site) = 10
            End If
        Next site
'
    Next i
    End If
    
    
    
        
    For i = 0 To TX_PM_EQ.pins.Count - 1
        For Each site In TheExec.sites.Active

            If ((TX_Vb_EQ.pins(i).value(site) / TX_Va_EQ.pins(i).value(site)) > 0) Then
                TX_PM_EQ.pins(i).value(site) = 20 * Log10((TX_Vb_EQ.pins(i).value(site) / TX_Va_EQ.pins(i).value(site)))
            Else
                TX_PM_EQ.pins(i).value(site) = 0
            End If
            If (argc = 3) Then
                If (TX_Vc_EQ.pins(i).value(site) / TX_Vb_EQ.pins(i).value(site) > 0) Then
                    TX_PM_PRE.pins(i).value(site) = 20 * Log10(TX_Vc_EQ.pins(i).value(site) / TX_Vb_EQ.pins(i).value(site))
                Else
                    TX_PM_PRE.pins(i).value(site) = 0
                End If
            End If
        Next site
'
    Next i
    

    For j = 0 To TX_PM_EQ.pins.Count - 1
        TestNameInput = Report_TName_From_Instance(CalcV, TX_PM_EQ.pins(j), vbNullString, CInt(j))
        TheExec.flow.TestLimit resultVal:=TX_PM_EQ.pins(j), ForceResults:=tlForceFlow, Tname:=TestNameInput
    Next j

    If (argc = 3) Then
        For j = 0 To TX_PM_PRE.pins.Count - 1
            TestNameInput = Report_TName_From_Instance(CalcV, TX_PM_PRE.pins(j), vbNullString, CInt(j))
            TheExec.flow.TestLimit resultVal:=TX_PM_PRE.pins(j), ForceResults:=tlForceFlow, Tname:=TestNameInput
        Next j
    End If


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
     
            DSP_DictKey_2S_BIN = GetStoreDataAllType(DictKey_2S_BIN)
        
        Else
        
        
            DictKey_2S_BIN = argv(0)
            DictKey_SIGN_DEC = DictKey_2S_BIN
            TestName = DictKey_2S_BIN
            DSP_DictKey_2S_BIN = GetStoreDataAllType(DictKey_2S_BIN)
        
        End If

        
        For Each site In TheExec.sites
            SL_BitWidth(site) = DSP_DictKey_2S_BIN(site).SampleSize

        Next site
        
        Set DSP_DictKey_SIGN_DEC(i) = Nothing
        DSP_DictKey_SIGN_DEC(i).CreateConstant 0, 1, DspLong
        
        Call rundsp.DSP_2S_Complement_To_SignDec(DSP_DictKey_2S_BIN, SL_BitWidth, DSP_DictKey_SIGN_DEC(i))
        
        
         Call StoreDataAllType(DictKey_SIGN_DEC, DSP_DictKey_SIGN_DEC(i))
        

        If Not ByPassTestLimit Then
            
            TestNameInput = Report_TName_From_Instance(CalcC, "X", , CInt(i))
            TheExec.flow.TestLimit resultVal:=DSP_DictKey_SIGN_DEC(i).Element(0), Tname:=TestNameInput, ForceResults:=tlForceFlow
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
                Call StoreDataAllType(DictKey_2S_BIN, DSP_CMRR_CALC(i))
                If Not ByPassTestLimit Then
                    TestNameInput = Report_TName_From_Instance(CalcC, "X", "CMRR", CInt(i))
                    TheExec.flow.TestLimit resultVal:=DSP_CMRR_CALC(i).Element(0), Tname:=TestNameInput, ForceResults:=tlForceFlow 'transfer_to_forceflow
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
    
                Call StoreDataAllType(DictKey_2S_BIN, DSP_PSRR_CALC(i))
                If Not ByPassTestLimit Then
                    TestNameInput = Report_TName_From_Instance(CalcC, "X", "PSRR", CInt(i))
                 ' Call StoreDataAllType(SplitByAt(2), DSP_PSRR_CALC(i))
                    If Not ByPassTestLimit Then
                            TheExec.flow.TestLimit resultVal:=DSP_PSRR_CALC(i).Element(0), Tname:=TestNameInput, ForceResults:=tlForceFlow 'transfer_to_forceflow
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
     
            DSP_DictKey_2S_BIN = GetStoreDataAllType(DictKey_2S_BIN)
        
        Else
        
        
            DictKey_2S_BIN = argv(0)
    
            DictKey_SIGN_DEC = DictKey_2S_BIN
            
            TestName = DictKey_2S_BIN
     
            DSP_DictKey_2S_BIN = GetStoreDataAllType(DictKey_2S_BIN)
        
        End If

        
        For Each site In TheExec.sites
            SL_BitWidth(site) = DSP_DictKey_2S_BIN(site).SampleSize

        Next site
        
        Set DSP_DictKey_SIGN_DEC(i) = Nothing
        DSP_DictKey_SIGN_DEC(i).CreateConstant 0, 1, DspLong
        
        Call rundsp.DSP_2S_Complement_To_SignDec(DSP_DictKey_2S_BIN, SL_BitWidth, DSP_DictKey_SIGN_DEC(i))
        
        
         Call StoreDataAllType(DictKey_SIGN_DEC, DSP_DictKey_SIGN_DEC(i))
        

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
        DSP_Input = GetStoreDataAllType(InputKey)
        
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
            'TheExec.Flow.TestLimit resultVal:=DSPWaveDec(i)(Site).Element(0), lowVal:=0, hiVal:=119, Tname:=TestNameInput, ForceResults:=tlForceNone
            TestNameInput = Report_TName_From_Instance(CalcC, "X", "LockCodeRange", CInt(i))
            TheExec.flow.TestLimit resultVal:=DSPWaveDec(i)(site).Element(0), Tname:=TestNameInput, ForceResults:=tlForceFlow
            TheExec.flow.TestLimitIndex = TheExec.flow.TestLimitIndex - 1
        Next i
    Next site
    TheExec.flow.TestLimitIndex = TheExec.flow.TestLimitIndex + 1
    
    
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

    TheExec.flow.TestLimit resultVal:=MDLL_DecreaseResultPass, lowVal:=1, hiVal:=1, Tname:=TestNameInput, ForceResults:=tlForceFlow 'transfer_to_forceflow
''    For Each Site In TheExec.sites
''        If MDLL_DecreaseResultPass.BitwiseAnd(1) Then
''        Else
''            MDLL_Index(Site) = -99
''        End If
''    Next Site
    TestNameInput = Report_TName_From_Instance(CalcC, "X", "Unique", 0)
    
    TheExec.flow.TestLimit resultVal:=DiffVal_Num, lowVal:=1, hiVal:=2, Tname:=TestName & "Unique", ForceResults:=tlForceFlow 'transfer_to_forceflow
    
    TestNameInput = Report_TName_From_Instance(CalcC, "X", "MaxDiff", 0)

    TheExec.flow.TestLimit resultVal:=DiffVal_MaxSubMin, lowVal:=0, hiVal:=1, Tname:=TestName & "Max_Diff", ForceResults:=tlForceFlow 'transfer_to_forceflow
End Function
Public Function Calc_Metrology_GainError(argc As Integer, argv() As String) As Long
    Dim Dict_ReturnKey As String
    Dim Dict_InputKey As String
    Dim InputVal As New PinListData
    Dim CalcVal As New PinListData
    
    Dict_ReturnKey = argv(0)
    Dict_InputKey = argv(1)
    InputVal = GetStoreDataAllType(Dict_InputKey)
    
    CalcVal.AddPin (InputVal.pins(0))
    CalcVal = InputVal.pins(0).Subtract(0.4).divide(0.7975).Subtract(1)
    Call StoreDataAllType(Dict_ReturnKey, CalcVal)
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
        InputDSPWave_BIN = GetStoreDataAllType(argv(i))
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
        
        TheExec.flow.TestLimit MIPI_threshold_Code_value1(p), 0, 2 ^ 10 - 1, PinName:="code1_" & p + 1, ForceResults:=tlForceFlow
    Next p

    TestNameInput = Report_TName_From_Instance(CalcC, "MIPI_Tolerance1_1", , CInt(x))

    TheExec.flow.TestLimit MIPI_threshold_lower(0), scaletype:=scaleNone, PinName:="MIPI_Tolerance1_1", ForceResults:=tlForceFlow
    'TheExec.Flow.TestLimit MIPI_threshold_lower(0), ScaleType:=None, PinName:="MIPI_Tolerance1_1", ForceResults:=tlForceFlow ''OscarLi_Compile,20190629
    TestNameInput = Report_TName_From_Instance(CalcC, "MIPI_Tolerance1_2", , CInt(x))
      
    TheExec.flow.TestLimit MIPI_threshold_high(0), scaletype:=scaleNone, PinName:="MIPI_Tolerance1_2", ForceResults:=tlForceFlow
    'TheExec.Flow.TestLimit MIPI_threshold_high(0), ScaleType:=None, PinName:="MIPI_Tolerance1_1", ForceResults:=tlForceFlow ''OscarLi_Compile,20190629
    TestNameInput = Report_TName_From_Instance(CalcC, "MIPI_threshold_found", , CInt(x))
    
    TheExec.flow.TestLimit MIPI_threshold_found, True, True, PinName:="MIPI_threshold_found", ForceResults:=tlForceFlow

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

    DSP_gainErrorOffset1 = GetStoreDataAllType(CapturedCode1)
    DSP_gainErrorOffset2 = GetStoreDataAllType(CapturedCode2)
    DSP_gainErrorOffset3 = GetStoreDataAllType(CapturedCode3)
    DSP_gainErrorOffset4 = GetStoreDataAllType(CapturedCode4)

    DSP_tfe_vol_0_in_2S_binary = GetStoreDataAllType(Dict_tfe_vol_0)
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

    TheExec.flow.TestLimit resultVal:=DSP_tfe_vol_0_in_decimal.Element(0), Tname:=TestNameInput, ForceResults:=tlForceFlow 'transfer_to_forceflow

    Call rundsp.BinToDec(DSP_gainErrorOffset1, DSP_gainErrorOffset1_decimal)
    Call rundsp.BinToDec(DSP_gainErrorOffset2, DSP_gainErrorOffset2_decimal)
    Call rundsp.BinToDec(DSP_gainErrorOffset3, DSP_gainErrorOffset3_decimal)
    Call rundsp.BinToDec(DSP_gainErrorOffset4, DSP_gainErrorOffset4_decimal)

    DSP_tfe_vol_1_in_decimal.CreateConstant 0, 1, DspLong

    TestNameInput = Report_TName_From_Instance(CalcC, "X", "CapCode1_Dec", CInt(x))

    TheExec.flow.TestLimit resultVal:=DSP_gainErrorOffset1_decimal.Element(0), Tname:=TestNameInput, ForceResults:=tlForceFlow 'transfer_to_forceflow
    
    TestNameInput = Report_TName_From_Instance(CalcC, "X", "CapCode2_Dec", CInt(x))
    
    TheExec.flow.TestLimit resultVal:=DSP_gainErrorOffset2_decimal.Element(0), Tname:=TestNameInput, ForceResults:=tlForceFlow 'transfer_to_forceflow
    
    TestNameInput = Report_TName_From_Instance(CalcC, "X", "CapCode3_Dec", CInt(x))
    
    TheExec.flow.TestLimit resultVal:=DSP_gainErrorOffset3_decimal.Element(0), Tname:=TestNameInput, ForceResults:=tlForceFlow 'transfer_to_forceflow
    
    TestNameInput = Report_TName_From_Instance(CalcC, "X", "CapCode4_Dec", CInt(x))
    
    TheExec.flow.TestLimit resultVal:=DSP_gainErrorOffset4_decimal.Element(0), Tname:=TestNameInput, ForceResults:=tlForceFlow 'transfer_to_forceflow

    
    TestNameInput = Report_TName_From_Instance(CalcC, "X", "Gain_vol_1_LimitExceeded_Dec", CInt(x))
    
    For Each site In TheExec.sites


        DSP_tfe_vol_1_in_decimal(site).Element(0) = DSP_gainErrorOffset1_decimal(site).Element(0) + DSP_gainErrorOffset2_decimal(site).Element(0) + DSP_gainErrorOffset3_decimal(site).Element(0) + DSP_gainErrorOffset4_decimal(site).Element(0) - 4 * DSP_tfe_vol_0_in_decimal(site).Element(0)
        If (DSP_tfe_vol_1_in_decimal(site).Element(0) > 262143) Then
            TheExec.Datalog.WriteComment ("Site:" + CStr(site) + "  Gain_vol_1_LimitExceeded_Dec = " + CStr(DSP_tfe_vol_1_in_decimal(site).Element(0)) + ", Force tfe_vol_1_dec = 174762")
            DSP_tfe_vol_1_in_decimal(site).Element(0) = 174762
        End If

    Next site
    
    TestNameInput = Report_TName_From_Instance(CalcC, "X", "tfe_vol_1_dec", CInt(x))
    
    TheExec.flow.TestLimit resultVal:=DSP_tfe_vol_1_in_decimal.Element(0), Tname:=TestNameInput, ForceResults:=tlForceFlow 'transfer_to_forceflow
    

    Call rundsp.DSPWf_Dec2Binary(DSP_tfe_vol_1_in_decimal, 18, DSP_tfe_vol_1_in_binary)

    Call StoreDataAllType(Dict_tfe_vol_1, DSP_tfe_vol_1_in_binary)
    
    Dim tfe_vol_1_bin_str As String
    Dim i As Long
   
    For Each site In TheExec.sites

            tfe_vol_1_bin_str = vbNullString
         For i = DSP_tfe_vol_1_in_binary(site).SampleSize - 1 To 0 Step -1
         
                tfe_vol_1_bin_str = tfe_vol_1_bin_str + CStr(DSP_tfe_vol_1_in_binary(site).Element(i))
            
         Next i
      
        If gl_Disable_HIP_debug_log = False Then TheExec.Datalog.WriteComment ("Site:" + CStr(site) + "  tfe_vol_1_binary_fuse_Code  " + tfe_vol_1_bin_str)
    Next site
     
   ' TheExec.Flow.TestLimit resultVal:=DSP_tfe_vol_1_in_binary.Element(0), Tname:="tfe_vol_1", ForceResults:=tlForceNone


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

Call StoreDataAllType(Dict_actual_temp, DSP_conv_temp_rounded_binary)


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

Dict_encoded_temp_in_2S_binary = GetStoreDataAllType(Dict_encoded_temp)


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



Call StoreDataAllType(Dict_decoded_temp, Dict_decoded_temp_in_Decimal)


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





    DSP_tfe_vol_x1_binary = GetStoreDataAllType(Dict_name_tfe_vol_x1)
    DSP_fuse_read_tfe_vol_0_2S_binary = GetStoreDataAllType(fuse_read_tfe_vol_0)
    DSP_fuse_read_tfe_vol_1_binary = GetStoreDataAllType(fuse_read_tfe_vol_1)
    DSP_fuse_read_tfe_x0_binary = GetStoreDataAllType(fuse_read_tfe_x0)



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

    DSP_fuse_read_tfe_y0_in_double = GetStoreDataAllType(decoded_Dic_tfe_y0_in_double)



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
'    DSP_tfe_y1_in_2S_binary = GetStoreDataAllType(encoded_tfe_y1_in_2S_binary)
'
'    'Check End


    'Start the algo


    'Define Constants

    Dim C0 As Double
    Dim C1 As Double
    Dim C2 As Double
    Dim C3 As Double

    'Values for Constants

    C0 = CDbl("-21.5822184999726")
    C1 = CDbl("428.0092266096283") 'truncated one digit
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
                 
                TheExec.flow.TestLimit resultVal:=DSP_fuse_read_tfe_x0_in_decimal(site).Element(0), Tname:=TestNameInput, ForceResults:=tlForceFlow 'transfer_to_forceflow
                
                TestNameInput = Report_TName_From_Instance(CalcC, "X", "tfey0", CInt(x))
                  
                TheExec.flow.TestLimit resultVal:=DSP_fuse_read_tfe_y0_in_double(site).Element(0), Tname:=TestNameInput, ForceResults:=tlForceFlow 'transfer_to_forceflow
                
                TestNameInput = Report_TName_From_Instance(CalcC, "X", "tfex1", CInt(x))
                    
                TheExec.flow.TestLimit resultVal:=DSP_tfe_vol_x1_in_decimal(site).Element(0), Tname:=TestNameInput, ForceResults:=tlForceFlow 'transfer_to_forceflow
                
                TestNameInput = Report_TName_From_Instance(CalcC, "X", "tfey1", CInt(x))
                    
                TheExec.flow.TestLimit resultVal:=DSP_tfe_y1_in_double(site).Element(0), Tname:=TestNameInput, ForceResults:=tlForceFlow 'transfer_to_forceflow
                
                TestNameInput = Report_ALG_TName_From_Instance(OutputTname_format, "C", "X", "tfevol0", CInt(x))
                
                TheExec.flow.TestLimit resultVal:=DSP_fuse_read_tfe_vol_0_in_decimal(site).Element(0), Tname:=TestNameInput, ForceResults:=tlForceFlow 'transfer_to_forceflow
                
                TestNameInput = Report_TName_From_Instance(CalcC, "X", "tfevol1", CInt(x))
                
                TheExec.flow.TestLimit resultVal:=DSP_fuse_read_tfe_vol_1_in_decimal(site).Element(0), Tname:=TestNameInput, ForceResults:=tlForceFlow 'transfer_to_forceflow
                
               
                
                If (DSP_fuse_read_tfe_x0_in_decimal(site).Element(0) = DSP_tfe_vol_x1_in_decimal(site).Element(0)) Or (DSP_fuse_read_tfe_vol_1_in_decimal(site).Element(0) = 0) Then
                         
                         tfe_temp0_double(site) = 178956970

                        tfe_temp1_double(site) = 178956970
                        
                            Dsp_tfe_temp0_in_decimal(site).Element(0) = FormatNumber(tfe_temp0_double(site))

                        Dsp_tfe_temp1_in_decimal(site).Element(0) = FormatNumber(tfe_temp1_double(site))
                    TestNameInput = Report_TName_From_Instance(CalcC, "X", "Error_code_temp_0", CInt(x))
                    
                    TheExec.flow.TestLimit resultVal:=Dsp_tfe_temp0_in_decimal(site).Element(0), Tname:=TestNameInput, ForceResults:=tlForceFlow 'transfer_to_forceflow
                    TestNameInput = Report_TName_From_Instance(CalcC, "X", "Error_code_temp_1", CInt(x))
                        
                    TheExec.flow.TestLimit resultVal:=Dsp_tfe_temp1_in_decimal(site).Element(0), Tname:=TestNameInput, ForceResults:=tlForceFlow 'transfer_to_forceflow
                Else
                

                    x0(site) = ((DSP_fuse_read_tfe_x0_in_decimal(site).Element(0) - DSP_fuse_read_tfe_vol_0_in_decimal(site).Element(0)) / CDbl(DSP_fuse_read_tfe_vol_1_in_decimal(site).Element(0))) * 16
            
        x1(site) = ((DSP_tfe_vol_x1_in_decimal(site).Element(0) - DSP_fuse_read_tfe_vol_0_in_decimal(site).Element(0)) / CDbl(DSP_fuse_read_tfe_vol_1_in_decimal(site).Element(0))) * 16


        Y0(site) = 273.15 + DSP_fuse_read_tfe_y0_in_double(site).Element(0) - C2 * x0(site) * x0(site) - C3 * x0(site) * x0(site) * x0(site)


        Y1(site) = 273.15 + DSP_tfe_y1_in_double(site).Element(0) - C2 * x1(site) * x1(site) - C3 * x1(site) * x1(site) * x1(site)


        c1_cal(site) = (Y1(site) - Y0(site)) / (x1(site) - x0(site))

        C0_CAL(site) = (x1(site) * Y0(site) - x0(site) * Y1(site)) / (x1(site) - x0(site))

        tfe_temp0_double(site) = (C0_CAL(site) - C0) * (2 ^ 13)

        tfe_temp1_double(site) = (c1_cal(site) - C1) * (2 ^ 13)

    'tfe_temp0_long(Site) = FormatNumber(tfe_temp0_double(Site))

    'tfe_temp1_long(Site) = FormatNumber(tfe_temp1_double(Site))
                
                
                    If (tfe_temp0_double(site) > 134217727) Or (tfe_temp0_double(site) < -134217728) Then
                    

                        TestNameInput = Report_TName_From_Instance(CalcC, "X", "UpperLimit_Reached_temp_0", CInt(x))
                            
                        TheExec.flow.TestLimit resultVal:=tfe_temp0_double(site), Tname:=TestNameInput, ForceResults:=tlForceFlow 'transfer_to_forceflow
 
                                               
                        tfe_temp0_double(site) = 178956970
                                           
                        
                    
                    End If
                    If (tfe_temp1_double(site) > 134217727) Or (tfe_temp1_double(site) < -134217728) Then
                        TestNameInput = Report_TName_From_Instance(CalcC, "X", "UpperLimit_Reached_temp_1", CInt(x))
                            
                        TheExec.flow.TestLimit resultVal:=tfe_temp1_double(site), Tname:=TestNameInput, ForceResults:=tlForceFlow 'transfer_to_forceflow
                    
                      tfe_temp1_double(site) = 178956970
                    End If
                
                
                    Dsp_tfe_temp0_in_decimal(site).Element(0) = FormatNumber(tfe_temp0_double(site))

                    Dsp_tfe_temp1_in_decimal(site).Element(0) = FormatNumber(tfe_temp1_double(site))
        
        
                TestNameInput = Report_TName_From_Instance(CalcC, "X", "X0", CInt(x))
                    
                TheExec.flow.TestLimit resultVal:=x0(site), Tname:=TestNameInput, ForceResults:=tlForceFlow 'transfer_to_forceflow
                
                TestNameInput = Report_TName_From_Instance(CalcC, "X", "X1", CInt(x))
                   
                TheExec.flow.TestLimit resultVal:=x1(site), Tname:=TestNameInput, ForceResults:=tlForceFlow 'transfer_to_forceflow

                TestNameInput = Report_TName_From_Instance(CalcC, "X", "Y0", CInt(x))
                    
                TheExec.flow.TestLimit resultVal:=Y0(site), Tname:=TestNameInput, ForceResults:=tlForceFlow 'transfer_to_forceflow
                
                TestNameInput = Report_TName_From_Instance(CalcC, "X", "Y1", CInt(x))
                
                TheExec.flow.TestLimit resultVal:=Y1(site), Tname:=TestNameInput, ForceResults:=tlForceFlow 'transfer_to_forceflow
                
                TestNameInput = Report_TName_From_Instance(CalcC, "X", "c1_calc", CInt(x))
                    
                TheExec.flow.TestLimit resultVal:=c1_cal(site), Tname:=TestNameInput, ForceResults:=tlForceFlow 'transfer_to_forceflow
                
                TestNameInput = Report_TName_From_Instance(CalcC, "X", "c0_calc", CInt(x))
                    
                TheExec.flow.TestLimit resultVal:=C0_CAL(site), Tname:=TestNameInput, ForceResults:=tlForceFlow 'transfer_to_forceflow
                
                TestNameInput = Report_TName_From_Instance(CalcC, "X", "temp_0", CInt(x))
                
                TheExec.flow.TestLimit resultVal:=Dsp_tfe_temp0_in_decimal(site).Element(0), Tname:=TestNameInput, ForceResults:=tlForceFlow 'transfer_to_forceflow
                
                TestNameInput = Report_TName_From_Instance(CalcC, "X", "temp_1", CInt(x))
                    
                TheExec.flow.TestLimit resultVal:=Dsp_tfe_temp1_in_decimal(site).Element(0), Tname:=TestNameInput, ForceResults:=tlForceFlow 'transfer_to_forceflow
                
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

    Call StoreDataAllType(fuse_write_tfe_temp_0, Dsp_tfe_temp0_in_binary)
    Call StoreDataAllType(fuse_write_tfe_temp_1, Dsp_tfe_temp1_in_binary)



End Function



Public Function Calc_Fmax_Divide_Fmin(argc As Integer, argv() As String) As Long
On Error GoTo errHandler 'Add ErrHandler 2023/05/29
    Dim site As Variant
    Dim DSP_Freq As New PinListData
    Dim Dict_Freq_Value() As New PinListData
    Dim i As Integer
    Dim Max_Temp As New PinListData
    Dim Min_Temp As New PinListData
    Dim Divide_Temp As New PinListData
    Dim ArrayNum As Integer
    Dim j As Integer
    Dim increase_flag As New SiteBoolean
    Dim DeltaF As New SiteDouble
    Dim TestNameInput As String
    Dim OutputTname_format() As String

    ArrayNum = argc - 1
    ReDim Dict_Freq_Value(ArrayNum) As New PinListData
    site = 0
    For i = 0 To argc - 1
        Dict_Freq_Value(i) = GetStoreDataAllType(argv(i))
''            ''===========Verification===========
''            For Each Site In TheExec.sites
''                    Dict_Freq_Value(i).Pins(0).Value(Site) = 1
''            Next Site
''            If i = 5 Then
''                Dict_Freq_Value(i).Pins(0).Value(0) = 30
''                Dict_Freq_Value(i).Pins(0).Value(1) = 40
''                Dict_Freq_Value(i).Pins(0).Value(2) = 50
''                Dict_Freq_Value(i).Pins(0).Value(3) = 60
''                Dict_Freq_Value(i).Pins(0).Value(4) = 70
''                Dict_Freq_Value(i).Pins(0).Value(5) = 80
''            End If
''            ''==================================
        If i = 0 Then
            Max_Temp.AddPin (Dict_Freq_Value(i).pins(0).name)
            Min_Temp.AddPin (Dict_Freq_Value(i).pins(0).name)
            Divide_Temp.AddPin (Dict_Freq_Value(i).pins(0).name)
        End If
        For Each site In TheExec.sites
            increase_flag(site) = True
            If i = 0 Then
                Max_Temp.pins(0).value(site) = Dict_Freq_Value(i).pins(0).value(site)
                Min_Temp.pins(0).value(site) = Dict_Freq_Value(i).pins(0).value(site)
            Else
                '''''''''''''''''''''print datalog'''''''''''''''''''''''''''''''''''''
                'TheExec.Datalog.WriteComment "Site " & site & ":" & argv(i) & "-" & argv(i - 1) & "=" & Dict_Freq_Value(i).Pins(0).Value(site) - Dict_Freq_Value(i - 1).Pins(0).Value(site)
                DeltaF = Dict_Freq_Value(i).pins(0).value(site) - Dict_Freq_Value(i - 1).pins(0).value(site)
                '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                If Dict_Freq_Value(i).pins(0).value(site) > Max_Temp.pins(0).value(site) Then
                    Max_Temp.pins(0).value(site) = Dict_Freq_Value(i).pins(0).value(site)
                Else
                    increase_flag = False
                End If
                If Dict_Freq_Value(i).pins(0).value(site) < Min_Temp.pins(0).value(site) Then
                    Min_Temp.pins(0).value(site) = Dict_Freq_Value(i).pins(0).value(site)
                End If
            End If
        Next site
        If i <> 0 Then
            If EnableDigitalTestLimitTTR Then
                TestNameInput = Report_TName_From_Instance("Calc", "X", "Delta-" & i, CInt(i), , , , , tlForceNone)
            Else
                TestNameInput = Report_TName_From_Instance("Calc", Dict_Freq_Value(i).pins(0).name, "Delta-" & i, CInt(i), , , , , tlForceNone)
            End If
            TheExec.flow.TestLimit resultVal:=DeltaF, Tname:=TestNameInput, ForceResults:=tlForceFlow, PinName:=Dict_Freq_Value(i).pins(0).name 'transfer_to_forceflow
        End If
    Next i
    
    For Each site In TheExec.sites
        If Min_Temp.pins(0).value(site) = 0 Or increase_flag(site) = False Then
        
            Divide_Temp.pins(0).value(site) = 999
            If Min_Temp.pins(0).value(site) = 0 Then
                TheExec.Datalog.WriteComment ("Error! Site " & site & " Min Freq Meas(Denominator)=0 Hz ")
            Else
                TheExec.Datalog.WriteComment ("Error! Site " & site & " Not FRO0<FRO1<FRO2....<FRO23 ")
            End If
            
        Else
'            Divide_Temp.Pins(0).Value(site) = Max_Temp.Pins(0).Value(site) / Min_Temp.Pins(0).Value(site)
            Divide_Temp.pins(0).value(site) = Dict_Freq_Value(argc - 1).pins(0).value(site) / Dict_Freq_Value(0).pins(0).value(site)
        End If
    Next site
    
    For i = 0 To Max_Temp.pins.Count - 1

        TestNameInput = Report_TName_From_Instance(CalcF, Max_Temp.pins(i), "Fmax", CInt(i))

        TheExec.flow.TestLimit resultVal:=Max_Temp, Tname:=TestNameInput, ForceResults:=tlForceFlow

        TestNameInput = Report_TName_From_Instance(CalcF, Max_Temp.pins(i), "Fmin", CInt(i))

        TheExec.flow.TestLimit resultVal:=Min_Temp, Tname:=TestNameInput, ForceResults:=tlForceFlow
        
        'TestNameInput = Report_TName_From_Instance("Calc", Max_Temp.Pins(i), "FmaxDivideFmin", CInt(i))
        If EnableDigitalTestLimitTTR Then
            TestNameInput = Report_TName_From_Instance("Calc", "X", "FmaxDivideFmin", CInt(i))
        Else
            TestNameInput = Report_TName_From_Instance("Calc", Max_Temp.pins(i), "FmaxDivideFmin", CInt(i))     '20200702 update by CT
        End If
        TheExec.flow.TestLimit resultVal:=Divide_Temp, Tname:=TestNameInput, ForceResults:=tlForceFlow
        
    Next i
Exit Function 'Add ErrHandler 2023/05/29errHandler: 'Add ErrHandler 2023/05/29
errHandler: 'Add ErrHandler 2023/05/29
    Call Print_Error_Message(Error_Info, "LIB_HardIP_Calc", "Calc_Fmax_Divide_Fmin") 'Add ErrHandler 2023/05/29
    If AbortTest Then Exit Function Else Resume Next 'Add ErrHandler 2023/05/29
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
    DSP_freq_Dut = GetStoreDataAllType(freq_Dut)
    DSP_freq_ref = GetStoreDataAllType(freq_ref)

   
    For Each site In TheExec.sites

    If DSP_freq_Dut(site).Element(0) <> 0 Then
        DSP_fdiff_percent(site).Element(0) = ((DSP_freq_Dut(site).Element(0) - DSP_freq_ref(site).Element(0)) / DSP_freq_Dut(site).Element(0)) * 100
     
    

    Else
        DSP_fdiff_percent(site).Element(0) = 99999
         If gl_Disable_HIP_debug_log = False Then TheExec.Datalog.WriteComment ("Site:" + CStr(site) + "  freq_of_Dut " + freq_Dut + " is 0")

    End If

                TheExec.flow.TestLimit resultVal:=DSP_fdiff_percent(site).Element(0), Tname:=TestName, ForceResults:=tlForceFlow 'transfer_to_forceflow
    Next site


    Call StoreDataAllType(fdiff_percent, DSP_fdiff_percent)



End Function

Public Function Calc_MIPI_VCMTX(argc As Integer, argv() As String) As Long
On Error GoTo errHandler 'Add ErrHandler 2023/05/29

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
    
    For i = 0 To argc - 2
        DSPWave_Binary(i) = GetStoreDataAllType(argv(i))
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
    Dim VDD18_MIPI_value As Double
    VDD18_MIPI_value = TheHdw.DCVS.pins("VDD12_MIPI").Voltage.Main.value
    
    If VDD18_MIPI_value = 0 Then
            VDD18_MIPI_value = 999
            TheExec.Datalog.WriteComment ("Error! Apply VDD18_MIPI=0 V  ")
    End If
    
    For Each site In TheExec.sites
        VCMTX(site).Element(0) = DSPWave_Combine_Dec(site).Element(0) / 1024 * VDD18_MIPI_value
    Next site
'    Call rundsp.DSPWaveDecToBinary(DSPWave_Combine_Dec, 10, DSPWave_Combine_verify)
    
    TestNameInput = Report_TName_From_Instance(CalcV, "X", , 0)

    TheExec.flow.TestLimit resultVal:=VCMTX.Element(0), Tname:=TestNameInput, ForceResults:=tlForceFlow
Exit Function 'Add ErrHandler 2023/05/29errHandler: 'Add ErrHandler 2023/05/29
errHandler: 'Add ErrHandler 2023/05/29
    Call Print_Error_Message(Error_Info, "LIB_HardIP_Calc", "Calc_MIPI_VCMTX") 'Add ErrHandler 2023/05/29
    If AbortTest Then Exit Function Else Resume Next 'Add ErrHandler 2023/05/29
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
        DSPWave_Binary(i) = GetStoreDataAllType(argv(i))
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
    VDD18_MIPID_value = TheHdw.DCVS.pins("VDD18_MIPID").Voltage.Main.value '20190523 CWCIOU
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

    TheExec.flow.TestLimit resultVal:=VCMTX.Element(0), Tname:=TestNameInput, ForceResults:=tlForceFlow
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
        DSPWave_Binary(i) = GetStoreDataAllType(argv(i))
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
    
    TheExec.flow.TestLimit resultVal:=DSPWave_Combine_Dec.Element(0), Tname:=TestNameInput, ForceResults:=tlForceFlow 'transfer_to_forceflow
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
    Input_V1 = GetStoreDataAllType(Dict_V1)
      Input_V2 = GetStoreDataAllType(Dict_V2)
      
      
    DummyPinListData.AddPin (Input_V1.pins(0))
      DummyPinListData = Input_V1.pins(0).Subtract(Input_V2.pins(0)).Abs
      
      
      
      
      For Each site In TheExec.sites
        result(site).Element(0) = DummyPinListData.pins(0).value
      Next site
      


'    CalcVal.AddPin (InputVal.Pins(0))
'    CalcVal = InputVal.Pins(0).Subtract(0.4).Divide(0.7975).Subtract(1)
    
         If Not ByPassTestLimit Then
            TestNameInput = Report_TName_From_Instance(CalcV, "X", , CInt(x))
            TheExec.flow.TestLimit resultVal:=result.Element(0), Tname:=TestNameInput, ForceResults:=tlForceFlow
        End If
'    Call StoreDataAllType(Dict_ReturnKey, CalcVal)
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
        DSPWave_Binary(i) = GetStoreDataAllType(argv(i))
        Call rundsp.BinToDec(DSPWave_Binary(i), DSPWave_Dec(i))
    Next i
    
    TestName = argv(argc - 1)
    bitwidth = argv(argc - 2)
    Dict = argv(argc - 3)
    
    For Each site In TheExec.sites
            DSPWave_Avg_Dec.Element(0) = Int(((DSPWave_Dec(0).Element(0) + DSPWave_Dec(1).Element(0)) / 2) + 0.5) ''Example 1). 78.4=>78  2). 78.5=79
    Next site
    Call rundsp.DSPWaveDecToBinary(DSPWave_Avg_Dec, bitwidth, DSPWave_Avg_Bin)
    Call StoreDataAllType(Dict, DSPWave_Avg_Bin)
    Dim TestNameInput As String
    Dim OutputTname_format() As String

    TestNameInput = Report_TName_From_Instance(CalcC, "X", , CInt(i))
    
    TheExec.flow.TestLimitIndex = TheExec.flow.TestLimitIndex + 1
    TheExec.flow.TestLimit resultVal:=DSPWave_Avg_Dec.Element(0), Tname:=TestNameInput, ForceResults:=tlForceFlow 'transfer_to_forceflow
    
End Function
Public Function Calc_CalR_FVMI_IO(argc As Integer, argv() As String) As Long
On Error GoTo errHandler 'Add ErrHandler 2023/05/29

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
    StoredCurrent = GetStoreDataAllType(argv(0))
    ForceVoltVal = argv(2)
    
    For Each pin In StoredCurrent.pins
        For Each site In TheExec.sites
            If StoredCurrent.pins(pin).value(site) = 0 Then
                StoredCurrent.pins(pin).value(site) = 0.000000000001
            End If
        Next site
    Next pin

    
    CalR = StoredCurrent.Math.Invert.Multiply(ForceVoltVal).Abs
          
    '===============RAK read
    Dim GetRakVal As New PinListData
    GetRakVal = CurrentJob_Card_RAK
       
            For Each site In TheExec.sites
                GetRakVal = CurrentJob_Card_RAK.pins(PowerPinName).value(site)
                If gl_Disable_HIP_debug_log = False Then TheExec.Datalog.WriteComment PowerPinName & " = " & CalR.pins.item(PowerPinName).value(site) & ", RAK val = " & GetRakVal.pins(PowerPinName).value
                CalR.pins.item(PowerPinName).value(site) = CalR.pins.item(PowerPinName).value(site) - GetRakVal.pins(PowerPinName).value
            Next site
    
        For p = 0 To CalR.pins.Count - 1
            If LCase(CalR.pins.item(p).name) Like LCase((PowerPinName)) Then
                    TestNameInput = Report_TName_From_Instance("R", CalR.pins(p), , CInt(p))
                    Hilimitval_temp = 96
                    Lowlimitval_temp = 64
                    TheExec.flow.TestLimit CalR.pins(p), Lowlimitval_temp, Hilimitval_temp, , , , unitCustom, , TestNameInput, , , , , " ohm", , ForceResults:=tlForceFlow
            End If
        Next p
    
Exit Function 'Add ErrHandler 2023/05/29errHandler: 'Add ErrHandler 2023/05/29
errHandler:
    Call Print_Error_Message(Error_Info, "LIB_HardIP_Calc", "Calc_CalR_FVMI_IO") 'Add ErrHandler 2023/05/29
    If AbortTest Then Exit Function Else Resume Next 'Add ErrHandler 2023/05/29
End Function
Public Function Calc_CalR_FVMI(argc As Integer, argv() As String) As Long
On Error GoTo errHandler 'Add ErrHandler 2023/05/29
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
    StoredCurrent = GetStoreDataAllType(argv(0))
    ForceVoltVal = TheHdw.DCVS.pins(PowerPinName).Voltage.value
    
    For Each pin In StoredCurrent.pins
        For Each site In TheExec.sites
            If StoredCurrent.pins(pin).value(site) = 0 Then
                StoredCurrent.pins(pin).value(site) = 0.000000000001
            End If
        Next site
    Next pin
    
    CalR = StoredCurrent.Math.Invert.Multiply(ForceVoltVal).Abs
'
'    Dim RakV() As Double
'    Dim GetRakVal As Double
'    Dim RAK_Pin As String
'
'    Dim PinGetRakVal As New PinListData
'    Set PinGetRakVal = Nothing
'
'    For Each Pin In CalR.Pins
'        PinGetRakVal.AddPin CStr(Pin)
'        For Each site In TheExec.sites
'            RAK_Pin = CStr(Pin)
'            RakV = TheHdw.PPMU.ReadRakValuesByPinnames(RAK_Pin, site)
'
'            GetRakVal = RakV(0) + CurrentJob_Card_RAK.Pins(Pin).Value(site)
'            PinGetRakVal.Pins(Pin).Value = GetRakVal
'                If argc <= 2 And gl_Disable_HIP_debug_log = False Then
'                    TheExec.DataLog.WriteComment Pin & " = " & CalR.Pins.Item(Pin).Value(site) & ", RAK val = " & GetRakVal
'                ElseIf argv(2) <> "TTR" And gl_Disable_HIP_debug_log = False Then
'                    TheExec.DataLog.WriteComment Pin & " = " & CalR.Pins.Item(Pin).Value(site) & ", RAK val = " & GetRakVal
'                End If
'            CalR.Pins.Item(Pin).Value(site) = CalR.Pins.Item(Pin).Value(site) - GetRakVal
'        Next site
'    Next Pin
'
    If EnableDigitalTestLimitTTR Then
                'TTR,20200423, Oscar
        TestNameInput = Report_TName_From_Instance(CalcR, "X")
        TheExec.flow.TestLimit CalR, , , , , , unitCustom, , TestNameInput, , , , , " ohm", , ForceResults:=tlForceFlow
    ElseIf argc <= 2 Then
    Dim Temp_index As Long
    Temp_index = TheExec.flow.TestLimitIndex
        For i = 0 To CalR.pins.Count - 1
            TheExec.flow.TestLimitIndex = Temp_index
            TestNameInput = Report_TName_From_Instance(CalcR, CalR.pins(i), , CInt(i))
'            If i = 0 Then
                TheExec.flow.TestLimit CalR.pins(i), , , , , , unitCustom, , TestNameInput, , , , , " ohm", , ForceResults:=tlForceFlow
'            Else
'                TheExec.Flow.TestLimit CalR.Pins(i), GetLowLimitFromFlow, GetHiLimitFromFlow, , , , unitCustom, , TestNameInput, , , , , " ohm", , ForceResults:=tlForceNone
'            End If
        Next i
    ElseIf argv(2) = "TTR" Then
            TheExec.flow.TestLimitIndex = TheExec.flow.TestLimitIndex + 2   'use second test limit spec for R,first test limit for current(don't need)

        Lowlimitval_temp = GetLowLimitFromFlow
        Hilimitval_temp = GetHiLimitFromFlow
         If TheExec.enableWord("HIP_TTR_FailResultOnly") = True Then
        For Each site In TheExec.sites.Active
            For p = 0 To CalR.pins.Count - 1
                If CalR.pins(p).value > Hilimitval_temp Or CalR.pins(p).value < Lowlimitval_temp Then
                    TestNameInput = Report_TName_From_Instance(CalcR, CalR.pins(p), , CInt(p))
                    
                    'TheExec.Flow.TestLimit StoredCurrent.Pins(p), , , , , , unitAmp, , , ForceResults:=tlForceNone
                    If gl_Disable_HIP_debug_log = False Then TheExec.Datalog.WriteComment CalR.pins(p) & " = " & CalR.pins(p).value(site)
                    TheExec.flow.TestLimit CalR.pins(p), Lowlimitval_temp, Hilimitval_temp, , , , unitCustom, , TestNameInput, , , , , " ohm", , ForceResults:=tlForceFlow 'transfer_to_forceflow

                    
                End If
            Next p
        Next site
        Else
            For p = 0 To CalR.pins.Count - 1

                    If gl_Disable_HIP_debug_log = False Then TheExec.Datalog.WriteComment CalR.pins(p) & " = " & CalR.pins(p).value(site)
                    TestNameInput = Report_TName_From_Instance(CalcR, CalR.pins(p), , CInt(p))
                    TheExec.flow.TestLimit CalR.pins(p), Lowlimitval_temp, Hilimitval_temp, , , , unitCustom, , TestNameInput, , , , , " ohm", , ForceResults:=tlForceFlow 'transfer_to_forceflow
            Next p
    End If
    Else
    End If
Exit Function 'Add ErrHandler 2023/05/29errHandler: 'Add ErrHandler 2023/05/29
errHandler: 'Add ErrHandler 2023/05/29
    Call Print_Error_Message(Error_Info, "LIB_HardIP_Calc", "Calc_CalR_FVMI") 'Add ErrHandler 2023/05/29
    If AbortTest Then Exit Function Else Resume Next 'Add ErrHandler 2023/05/29

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
        

    StoredCurrent_I1 = GetStoreDataAllType(argv(3))
    StoredCurrent_I2 = GetStoreDataAllType(argv(2))
  
    ForceVoltVal = argv(0) - argv(1)
    

    StoredCurrent = StoredCurrent_I2.Math.Subtract(StoredCurrent_I1)
    
        For Each pin In StoredCurrent.pins ' To prevent i=0
            For Each site In TheExec.sites
                If StoredCurrent.pins(pin).value(site) = 0 Then
                    StoredCurrent.pins(pin).value(site) = 0.000000000001
                End If
            Next site
        Next pin
    
        CalR = StoredCurrent.Math.Invert.Multiply(ForceVoltVal).Abs

        If gl_Disable_HIP_debug_log = False Then TheExec.Datalog.WriteComment CalR.pins(p) & " = " & CalR.pins(p).value(site)
        TestNameInput = Report_TName_From_Instance(CalcR, "X", vbNullString, 0)
        TheExec.flow.TestLimit CalR, , , , , , unitCustom, , TestNameInput, , , , , " ohm", , ForceResults:=tlForceFlow


End Function


Public Function Calc_MDLL_Monotonicity_DevideBlock_TTR(argc As Integer, argv() As String) As Long
On Error GoTo errHandler 'Add ErrHandler 2023/05/29
'TTR,20200423, Oscar
    Dim DSP_Input As New DSPWave
    Dim i As Long, j As Long
    Dim site As Variant
    Dim TestNameInput As String
    Dim InputKey As String
    Dim ConcatenateDSP_BeforeSort As New DSPWave
    Dim ConcatenateDSP_AfterSort As New DSPWave
        
'    For i = 0 To argc - 2
'        InputKey = LCase(argv(i + 1))
'        DSP_Input = GetStoreDataAllType(InputKey)
'
'        If i = 0 Then
'            ConcatenateDSP_BeforeSort = DSP_Input
'        Else
'            Call rundsp.FullConcatenateDSP(ConcatenateDSP_BeforeSort, DSP_Input)
'        End If
'    Next i
    ConcatenateDSP_BeforeSort.CreateConstant 0, argc * 2 \ 3, DspLong
    Dim SL_MDLL As New SiteLong
    Dim MDLLindex As Long: MDLLindex = 0
    For i = 0 To argc - 1
        If i Mod 3 <> 0 Then
            InputKey = LCase(argv(i))
            Set DSP_Input = Nothing
            DSP_Input = GetStoreDataAllType(InputKey)
            
            
'            If i = 1 Then
'                ConcatenateDSP_BeforeSort = DSP_Input
'            Else
'                Call rundsp.FullConcatenateDSP(ConcatenateDSP_BeforeSort, DSP_Input)
'            End If
            SL_MDLL = GetStoreDataAllType(InputKey & "_para")
            For Each site In TheExec.sites
                ConcatenateDSP_BeforeSort.ElementLite(MDLLindex) = SL_MDLL
            Next site
            MDLLindex = MDLLindex + 1
        End If
     Next i
    
    Dim sl_MDLL_DecreaseDirection As New DSPWave
    Dim sl_Num_DiffVal As New DSPWave
    Dim sl_Diff_MaxMin As New DSPWave
    
    sl_Diff_MaxMin.CreateConstant 0, argc \ 3
    sl_MDLL_DecreaseDirection.CreateConstant 1, argc \ 3
    sl_Num_DiffVal.CreateConstant 1, argc \ 3
    
    
    Call rundsp.DSP_CalcMDLLMonotonicityDevideBlock(ConcatenateDSP_BeforeSort, ConcatenateDSP_AfterSort, sl_MDLL_DecreaseDirection, sl_Num_DiffVal, sl_Diff_MaxMin)
    

    
    For j = 0 To argc \ 3 - 1
       
        Dim TempResult As New SiteLong
        For i = 0 To 7
            TempResult = ConcatenateDSP_AfterSort.Element(i + 8 * j)
            TestNameInput = Report_TName_From_Instance(CalcC, "X", , CInt(j), ForceResult:=tlForceFlow)
            TheExec.flow.TestLimit resultVal:=TempResult, Tname:=TestNameInput, ForceResults:=tlForceFlow
            TheExec.flow.TestLimitIndex = TheExec.flow.TestLimitIndex - 1
        Next i
        TheExec.flow.TestLimitIndex = TheExec.flow.TestLimitIndex + 1
        
        TempResult = sl_MDLL_DecreaseDirection.Element(j)
        TestNameInput = Report_TName_From_Instance(CalcC, "X", , CInt(j), ForceResult:=tlForceFlow)
        TheExec.flow.TestLimit resultVal:=TempResult, lowVal:=1, hiVal:=1, Tname:=TestNameInput, ForceResults:=tlForceFlow
        
        
        TempResult = sl_Num_DiffVal.Element(j)
        TestNameInput = Report_TName_From_Instance(CalcC, "X", , CInt(j), ForceResult:=tlForceFlow)
        TheExec.flow.TestLimit resultVal:=TempResult, lowVal:=1, hiVal:=2, Tname:=TestNameInput, ForceResults:=tlForceFlow
    
        TempResult = sl_Diff_MaxMin.Element(j)
        TestNameInput = Report_TName_From_Instance(CalcC, "X", , CInt(j), ForceResult:=tlForceFlow)
        TheExec.flow.TestLimit resultVal:=TempResult, lowVal:=0, hiVal:=1, Tname:=TestNameInput, ForceResults:=tlForceFlow
    Next j
Exit Function 'Add ErrHandler 2023/05/29errHandler: 'Add ErrHandler 2023/05/29
errHandler: 'Add ErrHandler 2023/05/29
    Call Print_Error_Message(Error_Info, "LIB_HardIP_Calc", "Calc_MDLL_Monotonicity_DevideBlock_TTR") 'Add ErrHandler 2023/05/29
    If AbortTest Then Exit Function Else Resume Next 'Add ErrHandler 2023/05/29

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
    DSPWave_Matrix_rot = GetStoreDataAllType(SensorTempName_rot)
    DSPWave_Matrix_rov = GetStoreDataAllType(SensorTempName_rov)
    
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
                TheExec.flow.TestLimit resultVal:=a1.Element(Row), Tname:=TestName, ForceResults:=tlForceFlow
            Next Row
            Set DSPWave_Coeff_1 = a1
            For Row = 0 To 2
                TestName = "a2_row_" + SensorTempName_rov + "_" + CStr(Row + 1) + ":"
                TheExec.flow.TestLimit resultVal:=a2.Element(Row), Tname:=TestName, ForceResults:=tlForceFlow
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
            TheExec.flow.TestLimit resultVal:=a3.Element(Row), Tname:=TestName, ForceResults:=tlForceFlow
        Next Row
        Set DSPWave_Coeff_1 = a3
        For Row = 0 To 2
            TestName = "a4_row_" + SensorTempName_rov + "_" + CStr(Row + 1) + ":"
            TheExec.flow.TestLimit resultVal:=a4.Element(Row), Tname:=TestName, ForceResults:=tlForceFlow
        Next Row
        Set DSPWave_Coeff_2 = a4
    End If
End Function

Public Function MTR_Cal_DecimalToBinary(ByVal inputDecimal As Double, ByVal bitsize As Long, ByVal placesAfterDecimal As Long, ByRef outBinaryStr As String) As Long
    
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
        Do While currCount < bitsize
        
            currDecimal = fractional * 2
            currIntegral = Int(currDecimal)
            decimalFract = decimalFract + CStr(currIntegral) * (2 ^ (bitsize - currCount))
            binaryStr = binaryStr + CStr(currIntegral)
            fractional = currDecimal - currIntegral
        
            
            currCount = currCount + 1
            
        Loop
        outBinaryStr = binaryStr
        

    Else
        currCount = 0
        binaryStr = vbNullString
        Do While currCount < bitsize
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
    pld_V1 = GetStoreDataAllType(DictKey_V1)
    pld_V2 = GetStoreDataAllType(DictKey_V2)
    
    pld_upd_V1.AddPin (Pin_Name_1)
    pld_upd_V2.AddPin (Pin_Name_2)
    
    For Each site In TheExec.sites
        'Rak_Pin_Name_1 = TheHdw.PPMU.ReadRakValuesByPinnames(Pin_Name_1, site)
        'Rak_Pin_Name_2 = TheHdw.PPMU.ReadRakValuesByPinnames(Pin_Name_2, site)
        GetRakVal = (CurrentJob_Card_RAK.pins(Pin_Name_1).value(site) + CurrentJob_Card_RAK.pins(Pin_Name_2).value(site)) / 2
        pld_upd_V1.pins(Pin_Name_1).value(site) = pld_V1.pins(Pin_Name_1).Multiply(45).divide(45 + 45 + GetRakVal).value(site)
        pld_upd_V2.pins(Pin_Name_2).value(site) = pld_V2.pins(Pin_Name_2).Multiply(45).divide(45 + 45 + GetRakVal).value(site)
    Next site
    
    TestNameInput = Report_TName_From_Instance(CalcV, "X", vbNullString, 0)
    TheExec.flow.TestLimit resultVal:=pld_upd_V1, ForceResults:=tlForceFlow, Tname:=TestNameInput
    
    TestNameInput = Report_TName_From_Instance(CalcV, "X", vbNullString, 0)
    TheExec.flow.TestLimit resultVal:=pld_upd_V2, ForceResults:=tlForceFlow, Tname:=TestNameInput
    
    
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
        DSPWave_First = GetStoreDataAllType(SplitByAdd(0))
        First_StartElement = 0
        First_EndElement = 7
        DSPWave_Second = GetStoreDataAllType(SplitByAdd(1))
        Second_StartElement = 0
        Second_EndElement = 1
        

        Call ConcatenateDSP_TTR(DSPWave_First, First_StartElement, First_EndElement, DSPWave_Second, Second_StartElement, Second_EndElement, DSPWave_Combine(i))
        
        ''20170718 - Store Concatenate DSP to Dict.
'        If UBound(SplitByAt) = 6 Then
'            DictKey_DSPWave_Combine = SplitByAt(6)
'            Call StoreDataAllType(DictKey_DSPWave_Combine, DSPWave_Combine(i))
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
               
            TheExec.flow.TestLimit resultVal:=DSPWave_Combine_Dec.Element(0), Tname:=TestNameInput, ForceResults:=tlForceFlow 'transfer_to_forceflow
      
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
        TheExec.flow.TestLimit MIPI_threshold_Code_value_P(p), 0, 2 ^ 10 - 1, PinName:="code_P_" & p, ForceResults:=tlForceFlow, Tname:=TestNameInput 'transfer_to_forceflow
    Next p
    End If
    TestNameInput = Report_TName_From_Instance(CalcC, "DATA0_Term_Tol1", vbNullString, 0)
    TheExec.flow.TestLimit MIPI_threshold_lower_p(0), scaletype:=scaleNone, PinName:="DATA0_Term_Tol1", ForceResults:=tlForceFlow, Tname:=TestNameInput
        
    TestNameInput = Report_TName_From_Instance(CalcC, "DATA0_Term_Tol2", vbNullString, 0)
    TheExec.flow.TestLimit MIPI_threshold_high_p(0), scaletype:=scaleNone, PinName:="DATA0_Term_Tol2", ForceResults:=tlForceFlow, Tname:=TestNameInput
    
    TestNameInput = Report_TName_From_Instance(CalcC, "DATA0_Found_Thresh", vbNullString, 0)
    TheExec.flow.TestLimit MIPI_threshold_found_p_value, 1, 1, PinName:="DATA0_Found_Thresh", ForceResults:=tlForceFlow, Tname:=TestNameInput
    
    If gl_Disable_HIP_debug_log = False Then
    
    For n = 0 To 7
        TestNameInput = Report_TName_From_Instance(CalcC, "code_N_" & n, vbNullString, CLng(n))
        TheExec.flow.TestLimit MIPI_threshold_Code_value_N(n), 0, 2 ^ 10 - 1, PinName:="code_N_" & n, ForceResults:=tlForceFlow, Tname:=TestNameInput 'transfer_to_forceflow
        
    Next n
    End If

    TestNameInput = Report_TName_From_Instance(CalcC, "DATA1_Term_Tol1", vbNullString, 0)
    TheExec.flow.TestLimit MIPI_threshold_lower_n(0), scaletype:=scaleNone, PinName:="DATA1_Term_Tol1", ForceResults:=tlForceFlow, Tname:=TestNameInput
    
    TestNameInput = Report_TName_From_Instance(CalcC, "DATA1_Term_Tol2", vbNullString, 0)
    TheExec.flow.TestLimit MIPI_threshold_high_n(0), scaletype:=scaleNone, PinName:="DATA1_Term_Tol2", ForceResults:=tlForceFlow, Tname:=TestNameInput
        
    TestNameInput = Report_TName_From_Instance(CalcC, "DATA1_Found_Thresh", vbNullString, 0)
    TheExec.flow.TestLimit MIPI_threshold_found_n_value, 1, 1, PinName:="DATA1_Found_Thresh", ForceResults:=tlForceFlow, Tname:=TestNameInput
    
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

    Call HardIP_Bin2Dec(ADC_Trim_Code, GetStoreDataAllType(ADC_Trim_Code_Str))
    For Each site In TheExec.sites
        Error_Code(site).Element(0) = ADC_Trim_Code(site).Element(0) - REFERENCE_CTRL
    Next site
    
    TestNameInput = Report_TName_From_Instance(CalcC, ADC_Error_Code_Str, vbNullString)
    TheExec.flow.TestLimit resultVal:=Error_Code.Element(0), lowVal:=-127, hiVal:=127, ForceResults:=tlForceFlow, Tname:=TestNameInput 'transfer_to_forceflow
        
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
    Call StoreDataAllType(ADC_Error_Code_Str, ERROR_CODE_Dict)
    
    
    If argc >= 4 Then
        ADC_Error_Code_Str_25 = argv(3)
        ADC_Error_Code_Str_85 = argv(1)
        ADC_Final_RefCtrl_Str = argv(4)
        
        Error_Code_25C = GetStoreDataAllType(ADC_Error_Code_Str_25)
        Error_Code_85C = GetStoreDataAllType(ADC_Error_Code_Str_85)
    
        For Each site In TheExec.sites
            SL_BitWidth(site) = Error_Code_25C(site).SampleSize
        Next site
        
        Call rundsp.DSP_2S_Complement_To_SignDec(Error_Code_25C, SL_BitWidth, Error_Code_25C_Dec)
        Call rundsp.DSP_2S_Complement_To_SignDec(Error_Code_85C, SL_BitWidth, Error_Code_85C_Dec)
    
        For Each site In TheExec.sites
            ADC_Final_RefCtrl(site).Element(0) = REFERENCE_CTRL + (Error_Code_25C_Dec(site).Element(0) + Error_Code_85C_Dec(site).Element(0)) / 2
        Next site
        TestNameInput = Report_TName_From_Instance(CalcC, "FinalReferenceControlCode", vbNullString)
        TheExec.flow.TestLimit resultVal:=ADC_Final_RefCtrl.Element(0), ForceResults:=tlForceFlow, Tname:=TestNameInput 'transfer_to_forceflow
        
        Call HardIP_Dec2Bin(ADC_Final_RefCtrl_Dict, ADC_Final_RefCtrl, 8)
        Call StoreDataAllType(ADC_Final_RefCtrl_Str, ADC_Final_RefCtrl_Dict)
    
    
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
DSP_Input = GetStoreDataAllType(InputKey)
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
    
        TestName = "f_diff_" + Replace(freq_Dut_array(k), index_name, TheExec.flow.var(index_name).value)
        DSP_freq_Dut = GetStoreDataAllType(dut_array(k))
        DSP_freq_ref = GetStoreDataAllType(ref_array(k))
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
        
        TestNameInput = Report_TName_From_Instance(CalcF, "X", Replace(freq_Dut_array(k), index_name, vbNullString), CInt(TheExec.flow.var(index_name).value))
        
        TheExec.flow.TestLimit resultVal:=DSP_freq_Dut.Element(0), Tname:=Replace(freq_Dut_array(k), index_name, TheExec.flow.var(index_name).value), ForceResults:=tlForceFlow, scaletype:=scaleMega 'transfer_to_forceflow
        
        TestNameInput = Report_TName_From_Instance(CalcF, "X", Replace(freq_ref_array(k), index_name, vbNullString), CInt(TheExec.flow.var(index_name).value))
        
        TheExec.flow.TestLimit resultVal:=DSP_freq_ref.Element(0), Tname:=Replace(freq_ref_array(k), index_name, TheExec.flow.var(index_name).value), ForceResults:=tlForceFlow, scaletype:=scaleMega 'transfer_to_forceflow
        
        TestNameInput = Report_TName_From_Instance(CalcF, "X", "Percent", CInt(TheExec.flow.var(index_name).value))
        
        TheExec.flow.TestLimit resultVal:=DSP_fdiff_percent.Element(0), Tname:=TestName, ForceResults:=tlForceFlow 'transfer_to_forceflow
            
        
    
        
        
        Call StoreDataAllType(Replace(freq_Dut_array(k), index_name, TheExec.flow.var(index_name).value), DSP_freq_Dut)
        Call StoreDataAllType(Replace(freq_ref_array(k), index_name, TheExec.flow.var(index_name).value), DSP_freq_ref)
    
        Call StoreDataAllType(Replace(fdiff_percent_array(k), index_name, TheExec.flow.var(index_name).value), DSP_fdiff_percent)
        
        
        
        
        If (TheExec.flow.var(index_name).value + 1 = Index_count) Then
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
                freq_Dut_wav = GetStoreDataAllType(Replace(freq_Dut_array(k), index_name, CStr(i)))
                freq_ref_wav = GetStoreDataAllType(Replace(freq_ref_array(k), index_name, CStr(i)))
                fdiff_percent_wav = GetStoreDataAllType(Replace(fdiff_percent_array(k), index_name, CStr(i)))
                
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

            TheExec.flow.TestLimit resultVal:=freq_Dut_mean, Tname:="freq_Dut_mean", ForceResults:=tlForceFlow
            TheExec.flow.TestLimit resultVal:=freq_ref_mean, Tname:="freq_ref_mean", ForceResults:=tlForceFlow
            TheExec.flow.TestLimit resultVal:=RSD_DUT, Tname:="RSD_DUT", ForceResults:=tlForceFlow
            TheExec.flow.TestLimit resultVal:=RSD_REF, Tname:="RSD_REF", ForceResults:=tlForceFlow
            TheExec.flow.TestLimit resultVal:=R_Ref, Tname:="R_Ref", ForceResults:=tlForceFlow
            TheExec.flow.TestLimit resultVal:=fdiff_percent_mean, Tname:="Avg_R0t0_E3", ForceResults:=tlForceFlow

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
    If (TheExec.flow.var(index_name).value + 1 = Index_count) Then
    
        If gl_Disable_HIP_debug_log = False Then
            For Each site In TheExec.sites.Active
                TheExec.Datalog.WriteComment "Final Efuse Write Value , Site:" + CStr(site) + " Value: " + CStr(Final_Fuse_Code_DEC.Element(0))
            Next site
        End If
        
        Call StoreDataAllType(Efuse_Dict_Name, Final_Fuse_Code_DEC)
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
    
    If (TheExec.flow.var(Sweep_Info(0).Loop_Index_Name).value = Loop_count - 1 And Loop_Index = Sweep_Count - 1) Then
        
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
                
                DSP_Capture = GetStoreDataAllType(dict_key & "_" & CStr(i))
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
                        
            TheExec.flow.TestLimit resultVal:=Sweep_Mean, Tname:=TestNameInput, ForceResults:=tlForceFlow
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
'            DSP_Capture = GetStoreDataAllType(CStr(key))
'
'        Next key
'
'
'    End If
    
    

End Function



Public Function Report_ALG_TName_From_Instance(ByRef TNameSeg() As String, MeasType As String, PinName As String, Tname As String, Optional TestSeqNum As Integer, Optional k As Long)
On Error GoTo errHandler 'Add ErrHandler 2023/05/29

        'Modify from M9 module
        Dim instancename As String
        Dim InstanceName_WO_Pset As String
        Dim InstNameSegs() As String
        Dim InTNameSegs() As String
        ReDim TNameSeg(9) As String

        instancename = UCase(TheExec.DataManager.instancename)
        InTNameSegs = Split(gl_Tname_Alg, ",")
        
        ''20190107 - Global name for saving Customize Subblock name
        If gl_Current_Instance_Tname <> instancename Then
            gl_Current_Instance_Tname = instancename
            gl_Current_Instance_Tname_subblock = Application.Worksheets(TheExec.flow.Raw.SheetInRun).range("AM" & CStr(TheExec.flow.Raw.GetCurrentLineNumber + 5)).value
        End If
'        PatSetName = Trim(PatSetName)
'        InstanceName_WO_Pset = Replace(InstanceName, UCase(PatSetName), "")
'        InstanceName_WO_Pset = Replace(InstanceName_WO_Pset, "__", "_")

        ' At Head   :   "DCTEST"
        If InstanceName_WO_Pset Like "DCTEST_*" Then
            instancename = Replace(InstanceName_WO_Pset, "DCTEST_", vbNullString)
        End If
        ' All places:   "VIR"
        instancename = Replace(instancename, "_VIR_", "_")

        InstNameSegs = Split(instancename, "_")

        'Instance name:    [Block]_[X1]_{patset}_[X2]_[X3]_[HV/NV/LV]
        'Test name:        HAC_____[Meas type]_[HV/NV/LV]_[X1]_[Block]_[Pin name]_[X2]_[X3]_[X4]_[X5]_

        TNameSeg(0) = "HAC"
        TNameSeg(1) = "Meas"
        TNameSeg(2) = InstNameSegs(UBound(InstNameSegs))                '[HV/NV/LV]
        TNameSeg(3) = "x"                                               '[X1] : sub-block-name-1
        TNameSeg(4) = InstNameSegs(0)                                   '[Block]
        TNameSeg(5) = "{pinname}"                                       '[Pin-name]
        TNameSeg(6) = "x"                                               '[X2] : sub-block-name-2
        TNameSeg(7) = "x"                                               '[X3] : X3 / DSSC Segment name
        TNameSeg(8) = "x"                                               '[X4] :    / DSSC Register
        TNameSeg(9) = "x"                                               '[X5] : subr-seq#

        '[H/N/L]
        TNameSeg(2) = Replace(UCase(TNameSeg(2)), "V", vbNullString)

        '[X1]
        If UBound(InstNameSegs) >= 2 Then
            If gl_Current_Instance_Tname_subblock <> "" Then            ''20190107 - Global name for saving Customize Subblock name
                TNameSeg(3) = gl_Current_Instance_Tname_subblock
            Else
                TNameSeg(3) = InstNameSegs(1)
            End If
        End If
        
        If TheExec.DataManager.instancename Like "IDS_*IDS*" Then
            TNameSeg(9) = CStr(TestSeqNum)
        Else
            TNameSeg(9) = CStr(gl_Tname_Alg_Index)
        End If
        TNameSeg(5) = Replace(PinName, "_", vbNullString)
        
        If gl_Tname_Alg <> "" Then
            If UBound(InTNameSegs) < gl_Tname_Alg_Index Then
                TNameSeg(6) = "X"
                Else
                TNameSeg(6) = InTNameSegs(gl_Tname_Alg_Index)
            End If
        End If
        
        If MeasType = "I" Then
            TNameSeg(1) = TNameSeg(1) & MeasType
        Else
            TNameSeg(1) = "Calc"
        End If
          
        
        If gl_Sweep_Name <> "" Then
            If sweep_power_val_per_loop_count <> "" Then
                TNameSeg(8) = Replace(sweep_power_val_per_loop_count, ".", "p")
            Else
                TNameSeg(9) = TheExec.flow.var(gl_Sweep_Name).value
            End If
            
        Else
            If gl_Tname_Alg <> "" Then TNameSeg(9) = gl_Tname_Alg_Index
        End If
        
        If LCase(TNameSeg(4)) = "pp" Or LCase(TNameSeg(4)) = "dd" Or LCase(TNameSeg(4)) = "dp" Or LCase(TNameSeg(4)) = "cz" Or LCase(TNameSeg(4)) = "ht" Then TNameSeg(4) = "X"
        If LCase(TNameSeg(3)) = "pp" Or LCase(TNameSeg(3)) = "dd" Or LCase(TNameSeg(3)) = "dp" Or LCase(TNameSeg(3)) = "cz" Or LCase(TNameSeg(3)) = "ht" Then TNameSeg(3) = "X"
        
        If InStr(LCase(TheExec.DataManager.instancename), "lapll") <> 0 Or InStr(LCase(TheExec.DataManager.instancename), "usb2") <> 0 Or InStr(LCase(TheExec.DataManager.instancename), "mipi") <> 0 Then
            TNameSeg(3) = UCase(TNameSeg(3))
            
            If InStr(LCase(TNameSeg(3)), "v") <> 0 Then
                TNameSeg(7) = Split(TNameSeg(3), "V")(0)
                TNameSeg(3) = "V" & Split(TNameSeg(3), "V")(1)
            ElseIf InStr(LCase(TNameSeg(3)), "t") <> 0 Then
                TNameSeg(7) = Split(TNameSeg(3), "T")(0)
                TNameSeg(3) = "T" & Split(TNameSeg(3), "T")(1)
            Else
            End If
        End If
        
        If InStr(LCase(TheExec.DataManager.instancename), "lpdprx") <> 0 Then
            TNameSeg(3) = UCase(TNameSeg(3))
            
            If LCase(TNameSeg(3)) Like "rx2*" And InStr(TNameSeg(3), "L") <> 0 Then
                TNameSeg(7) = "L" & Split(TNameSeg(3), "L")(1)
                
                If UCase(TNameSeg(6)) Like "LN*" Then
                    TNameSeg(6) = Replace(UCase(TNameSeg(6)), "LN" & Split(TNameSeg(3), "L")(1), vbNullString)
                End If
                TNameSeg(3) = Split(TNameSeg(3), "L")(0)
            End If
        End If
        
        If InStr(LCase(TheExec.DataManager.instancename), "pcie") <> 0 Then

                If UCase(TNameSeg(6)) Like "LN*" Then
                    TNameSeg(6) = UCase(Replace(TNameSeg(6), "_", vbNullString))
                    TNameSeg(7) = UCase(mid(TNameSeg(6), 1, 3))
                    TNameSeg(6) = UCase(mid(TNameSeg(6), 4, Len(TNameSeg(6)) - 3))
                End If
                
            
        End If
        
        If InStr(LCase(TheExec.DataManager.instancename), "amp") <> 0 Then
            TNameSeg(3) = UCase(TNameSeg(3))
            
            If LCase(TNameSeg(6)) Like "ddr*" Then
                TNameSeg(7) = UCase(mid(TNameSeg(6), 1, 4))
                TNameSeg(6) = UCase(mid(TNameSeg(6), 5, Len(TNameSeg(6)) - 4))
            End If
        End If
        
        '-------------------------------Pin Split--------------------------------------------------------
        If InStr(LCase(TheExec.DataManager.instancename), "amp") <> 0 Then
            If LCase(TNameSeg(5)) Like "ddr*" Then
                TNameSeg(5) = Replace(TNameSeg(5), "_", vbNullString)
                TNameSeg(7) = UCase(mid(TNameSeg(5), 1, 4))
                TNameSeg(5) = UCase(mid(TNameSeg(5), 5, Len(TNameSeg(5)) - 4))
            End If
        End If
        '-------------------------------Pin Split--------------------------------------------------------
        

        '[X2]_[X3]_[X4]
'        If UBound(InstNameSegs) >= 5 Then
'                TNameSeg(6) = InstNameSegs(UBound(InstNameSegs) - 3)        '[X2]
'                TNameSeg(7) = InstNameSegs(UBound(InstNameSegs) - 2)        '[X3]
'                TNameSeg(8) = InstNameSegs(UBound(InstNameSegs) - 1)        '[X4]
'        ElseIf UBound(InstNameSegs) >= 4 Then
'                TNameSeg(6) = InstNameSegs(UBound(InstNameSegs) - 2)        '[X2]
'                TNameSeg(7) = InstNameSegs(UBound(InstNameSegs) - 1)        '[X3]
'        ElseIf UBound(InstNameSegs) >= 3 Then
'                TNameSeg(6) = InstNameSegs(UBound(InstNameSegs) - 1)        '[X2]
'        End If

'        Call SetupDatalogFormat(TestNameW:=90, PatternW:=100)
    gl_Tname_Alg_Index = gl_Tname_Alg_Index + 1
    
    Call SetupDatalogFormat(80, 100)
    
Exit Function 'Add ErrHandler 2023/05/29
errHandler: 'Add ErrHandler 2023/05/29
    Call Print_Error_Message(Error_Info, "LIB_HardIP_Calc", "Report_ALG_TName_From_Instance") 'Add ErrHandler 2023/05/29
    If AbortTest Then Exit Function Else Resume Next 'Add ErrHandler 2023/05/29
End Function

Public Function Calc_ADC_Convert_Average(argc As Integer, argv() As String) As Long
On Error GoTo errHandler 'Add ErrHandler 2023/05/29

Dim Addvalue As Double
Dim Minuspin As String
Dim RefferenceCode_string As String
    Dim Transfer_Code_string As String
    Dim i As Long
    Dim j As Long
    Dim RefferenceCode As New DSPWave
    Dim Transfer_Code As New DSPWave
    Dim ADC_code As New DSPWave
    Dim ADC_code_DEC As New SiteDouble
    Dim ADC_code_average As New DSPWave
    Dim ADC_code_average_DEC As New SiteDouble
    Dim OutputTname_format() As String
    Dim TestNameInput As String
    Dim averagecount As Long
    Dim temp_RefferanceCode As New DSPWave
    Dim temp_RefferanceCode_DEC As New SiteDouble
    Dim MinusValue As Double
    Dim Tname_String As String
    
    Dim RefferenceCode_DEC As New SiteDouble
    Dim Transfer_Code_DEC As New SiteDouble
    
    RefferenceCode.CreateConstant 0, 1, DspDouble
    Transfer_Code.CreateConstant 0, 1, DspDouble
    temp_RefferanceCode.CreateConstant 0, 1, DspDouble
    ADC_code_average.CreateConstant 0, 1, DspDouble
    
    ''''''''''''''''CalcArg:2,127,VDD12_PCIE,adc_offset_adc0_0,adc_offset_ adc1_1,ctlevos_in_p_adc0_2,ctlevos_in_p_adc1_3,ctlevos_in_n_adc0_4,ctlevos_in_n_adc1_5,vss_adc0_6,vss_adc1_7,ctlevos_in_cm_adc0_8'''''
    
    averagecount = argv(0)
    Addvalue = argv(1)
    Minuspin = argv(2)
    MinusValue = ProcessEvaluateDCSpec(Minuspin)
    
    '////////////////////////////////////// for DDR/SOC PLL referece code
    If LCase(argv(3)) = "dummyvref" Then
        Dim Dummytemp As New SiteDouble
            Dummytemp = 127
         Call StoreDataAllType("DummyVref" & "_para", Dummytemp)
    End If
    '/////////////////////////////////////////
    For i = 3 To 3 + averagecount - 1
        RefferenceCode_string = argv(i)
        RefferenceCode_DEC = GetStoreDataAllType(RefferenceCode_string & "_para")
        Set temp_RefferanceCode_DEC = temp_RefferanceCode_DEC.Add(RefferenceCode_DEC)
    Next i
    Set temp_RefferanceCode_DEC = temp_RefferanceCode_DEC.divide(averagecount)
    If averagecount >= 2 Then
        If EnableDigitalTestLimitTTR = True Then
            TestNameInput = Report_TName_From_Instance(CalcC, "X", left(RefferenceCode_string, Len(RefferenceCode_string) - 3), CInt(i - 3))
        Else
            TestNameInput = Report_TName_From_Instance(CalcC, left(RefferenceCode_string, Len(RefferenceCode_string) - 3), left(RefferenceCode_string, Len(RefferenceCode_string) - 3), CInt(i - 3))
        End If
        TheExec.flow.TestLimit resultVal:=temp_RefferanceCode_DEC, ForceResults:=tlForceFlow, Tname:=TestNameInput
    End If
    
    For j = 3 + averagecount To UBound(argv) Step averagecount
        Set ADC_code_average_DEC = ADC_code_average_DEC.Multiply(0)
        For i = 0 To averagecount - 1
            Transfer_Code_string = argv(j + i)
            Transfer_Code_DEC = GetStoreDataAllType(Transfer_Code_string & "_para")
            Set ADC_code_DEC = Transfer_Code_DEC.Add(Addvalue).Add(temp_RefferanceCode_DEC.Multiply(-1)).divide(256).Multiply(0.5).Multiply(MinusValue).Add(0.25 * (MinusValue))
            Set ADC_code_average_DEC = ADC_code_average_DEC.Add(ADC_code_DEC)
        Next i
        Set ADC_code_average_DEC = ADC_code_average_DEC.divide(averagecount)
        Tname_String = argv(j)
        If averagecount >= 2 Then
            If EnableDigitalTestLimitTTR = True Then
                TestNameInput = Report_TName_From_Instance(CalcC, "X", left(Tname_String, Len(Tname_String) - 3))
            Else
                TestNameInput = Report_TName_From_Instance(CalcC, left(Tname_String, Len(Tname_String) - 3), left(Tname_String, Len(Tname_String) - 3))
            End If
            TheExec.flow.TestLimit resultVal:=ADC_code_average_DEC, ForceResults:=tlForceFlow, Tname:=TestNameInput
        Else
            If EnableDigitalTestLimitTTR = True Then
                TestNameInput = Report_TName_From_Instance(CalcC, "X", "_ADC")
            Else
                TestNameInput = Report_TName_From_Instance(CalcC, Tname_String, "_ADC")
            End If
            TheExec.flow.TestLimit resultVal:=ADC_code_average_DEC, ForceResults:=tlForceFlow, Tname:=TestNameInput
        End If
    Next j

Exit Function 'Add ErrHandler 2023/05/29
errHandler: 'Add ErrHandler 2023/05/29
    Call Print_Error_Message(Error_Info, "LIB_HardIP_Calc", "Calc_ADC_Convert_Average") 'Add ErrHandler 2023/05/29
    If AbortTest Then Exit Function Else Resume Next 'Add ErrHandler 2023/05/29
End Function


Public Function compensate_Volt(argc As Integer, argv() As String) As Long
On Error GoTo errHandler 'Add ErrHandler 2023/05/29

' This function is for accurate voh, vol calculations
'TTR,20200423, Oscar
Dim RAK_Pin As String
Dim Meas_pins() As String
Dim inVoh As New PinListData
Dim MeasureValue As New PinListData
Dim Compensate_V As New PinListData
Dim voh_temp As New SiteDouble
Dim forceV_temp(8) As Double
Dim RakV() As Double
Dim GetRakVal As Double
Dim pin  As Variant
Dim i As Integer
Dim j As Integer
Dim k As Integer
Dim p As Integer
Dim OutputTname_format() As String
Dim TestNameInput As String
Dim site As Variant

inVoh = GetStoreDataAllType(argv(1))
ReDim Meas_pins(inVoh.pins.Count - 1)
    
' 0.x get pins
For i = 0 To inVoh.pins.Count - 1
    Meas_pins(i) = inVoh.pins.item(i).name
    ' 0.1 add pins
    MeasureValue.AddPin (Meas_pins(i))
    Compensate_V.AddPin (Meas_pins(i))
    ' 0.2 disconect digital pins first
    TheHdw.Digital.pins(Meas_pins(i)).Disconnect
Next i

' 1.x Force V
For Each pin In Meas_pins
        ' 1.0 Get the measured voltage firstly
        Set voh_temp = inVoh.pins(pin)
        ' 1.1 Check if measured voltages are out of PPMU spec, if so, use 0v as default setting
        For Each site In TheExec.sites
             If voh_temp < -1 Or voh_temp > 6 Then
                    TheExec.Datalog.WriteComment "the force value " & voh_temp & "is out of PPMU range -1V ~ 6V, bypass force PPMU and set measurement result to 9999"
                    voh_temp(site) = 0
            End If
        Next site
        ' 1.2 Setup PPMU then measure current
         With TheHdw.PPMU.pins(pin)
            '' 20150615 - Force 0 mA before expected force value to solve over clamp issue.
             .ForceI pc_Def_PPMU_InitialValue_FI, pc_Def_PPMU_InitialValue_FI_Range
            .Connect
            .Gate = tlOn
            If TheExec.TesterMode = testModeOffline Then voh_temp = 0.1
            For Each site In TheExec.sites '20191231
              .ForceV voh_temp, 0.05
            Next site
            '' 20160108 - Only keep 1 force value but current range can be different for force pin
        End With
        TheHdw.Wait (100 * us)
        MeasureValue.pins(pin) = TheHdw.PPMU.pins(pin).Read(tlPPMUReadMeasurements, pc_Def_PPMU_ReadPoint)
        ' 1.3 Set abnormal voltage to 9999
        For Each site In TheExec.sites
            If inVoh.pins(pin).value(site) < -1 Or inVoh.pins(pin).value(site) > 6 Then MeasureValue.pins(pin).value(site) = 9999
        Next site
Next pin
        
' 2.x calc VOH
Dim GetRakValSite As New SiteDouble
For Each pin In Meas_pins
    ' 2.1 Get RAK value
    RAK_Pin = CStr(pin)
    GetRakValSite = CurrentJob_Card_RAK.pins(pin)
    ' 2.2 voh = I*R + V
    Compensate_V.pins(pin) = GetRakValSite.Multiply(MeasureValue.pins(pin)).Abs.Add(inVoh.pins(pin))
Next pin
      
Dim TempLimitIndex As Long
' 3.x print out datalog
For Each pin In Meas_pins
      If TheExec.TesterMode = testModeOffline Then voh_temp = 0.1
      TestNameInput = Report_TName_From_Instance("I", inVoh.pins(pin), "OutputCurrent", CInt(i))
      'For Each site In TheExec.sites '20191231
          voh_temp = inVoh.pins(pin)
          TempLimitIndex = TheExec.flow.TestLimitIndex
          For Each site In TheExec.sites '20191231
              TheExec.flow.TestLimitIndex = TempLimitIndex
              TheExec.flow.TestLimit MeasureValue.pins(pin), formatStr:="%.4f", Tname:=TestNameInput, ForceVal:=voh_temp, ForceUnit:=unitVolt, ForceResults:=tlForceFlow
          Next site '20191231
Next pin
'      For Each pin In Meas_pins
TestNameInput = Report_TName_From_Instance(CalcV, inVoh.pins(pin), vbNullString, CInt(i))
TheExec.flow.TestLimit Compensate_V, , scaletype:=scaleNone, unit:=unitVolt, formatStr:="%.4f", Tname:=TestNameInput, ForceResults:=tlForceFlow

 
' 4.x restore
Dim TestSeq As Long
For TestSeq = 0 To (inVoh.pins.Count - 1)
      With TheHdw.PPMU.pins(Meas_pins(TestSeq))
          .ForceI 0, 0.05
          .ForceV 0, 0.05
          .Gate = tlOff
          .Disconnect
      End With
      TheHdw.Digital.pins(Meas_pins(TestSeq)).Connect
Next TestSeq
Exit Function 'Add ErrHandler 2023/05/29errHandler: 'Add ErrHandler 2023/05/29
errHandler: 'Add ErrHandler 2023/05/29
    Call Print_Error_Message(Error_Info, "LIB_HardIP_Calc", "compensate_Volt") 'Add ErrHandler 2023/05/29
    If AbortTest Then Exit Function Else Resume Next 'Add ErrHandler 2023/05/29

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
    DSP_Input_2 = GetStoreDataAllType(InputKey & "_para")
    'DSP_Input = GetStoreDataAllType(InputKey & "_para")
    ADC_Output = ADC_Output.Add(DSP_Input_2)
    ADC_Output = ADC_Output.Multiply(MinusValue).divide(devideV)
    TestNameInput = Report_TName_From_Instance(CalcC, InputKey, "_ADC", CInt(i - 2))
    TheExec.flow.TestLimit resultVal:=ADC_Output, ForceResults:=tlForceFlow, Tname:=TestNameInput
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
        Voltage_Value = TheHdw.DCVS.pins(Input_Pins(z)).Voltage.Main
        TheExec.Datalog.WriteComment Input_Pins(z) & " Vmain=" & Voltage_Value
        Voltage_Value = TheHdw.DCVS.pins(Input_Pins(z)).Voltage.Alt
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
        temp_dsp = GetStoreDataAllType(argv(i))
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
        TheExec.flow.TestLimit resultVal:=cap_dec_data, ForceResults:=tlForceFlow, unit:=unitCustom, customUnit:=vbNullString   ', Tname:="FailBitCount", Unit:=unitNone, ScaleType:=scaleNone
    Next i


'/////////////// Judgement passing flag///////////////////
    
    Dim temp_HexStr As String
        For Each CurSite In TheExec.sites
            temp_HexStr = vbNullString
            HexStr = vbNullString
            For i = 0 To argc - 1
             temp_dsp = GetStoreDataAllType(argv(i))
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
    TheExec.flow.TestLimit resultVal:=dc_read, ForceResults:=tlForceFlow, unit:=unitCustom, customUnit:=vbNullString
    TheExec.flow.TestLimit resultVal:=second_flag, ForceResults:=tlForceFlow, unit:=unitCustom, customUnit:=vbNullString

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
                DSP_Captured(j) = GetStoreDataAllType(tmp_name)
                
                DSP_INVTemp(i).CreateConstant 0, DSP_Captured(j).SampleSize, DspLong
                
'                For k = DSP_Captured(j).SampleSize To 1 Step -1
                For k = 0 To DSP_Captured(j).SampleSize - 1
                                  
                    DSP_INVTemp(i).Element(k) = DSP_Captured(j).Element(UBound(DSP_Captured(j).data) - k)
                    DSP_Record(i * (UBound(SplitByAt) + 1) + j) = CStr(DSP_Record(i * (UBound(SplitByAt) + 1) + j)) & CStr(DSP_INVTemp(i).Element(k))
                   
                Next k
                 
 
                Set DSP_INV(i) = DSP_INV(i).Concatenate(DSP_INVTemp(i)) 'Merge all need flipped bit into one DSP
                
      
            Else
                
                DSP_Captured(j) = GetStoreDataAllType(SplitByAt(j))
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
            TestNameInput = Report_TName_From_Instance("X", "x", TestNameInput, CInt(TheExec.flow.TestLimitIndex), 0)
        
            For Each site In TheExec.sites
                TheExec.flow.TestLimit resultVal:=Eye_str_long.Element(i), formatStr:="%i", Tname:=TestNameInput, ForceResults:=tlForceFlow, tNum:=TnumRecord, scaletype:=scaleNoScaling
        '            TheExec.Flow.TestLimit lowVal:=mdll_low(i)(Site), resultVal:=Eye_str_long(Site).Element(i) * 4, FormatStr:="%i", TName:=TestNameInput, ForceResults:=tlForceFlow, TNum:=TnumRecord
                TheExec.flow.TestLimitIndex = TheExec.flow.TestLimitIndex - 1
            Next site
            TheExec.flow.TestLimitIndex = TheExec.flow.TestLimitIndex + 1
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
                    DSP_Captured(j) = GetStoreDataAllType(tmp_name)
                    'DSP_Captured(j).CreateRandom 0, 1, 10, 1, DspLong '<- should be replaced by previous Line
                    'Set DSP_INV(i) = DSP_INV(i).Concatenate(DSP_Captured(j)) 'Merge all need flipped bit into one DSP
                    DSP_INVTemp(i).CreateConstant 0, DSP_Captured(j).SampleSize, DspLong
                    For k = 0 To DSP_Captured(j).SampleSize - 1
                        DSP_INVTemp(i).Element(k) = DSP_Captured(j).Element(UBound(DSP_Captured(j).data) - k)
                        DSP_Record(i * (UBound(SplitByAt) + 1) + j) = CStr(DSP_Record(i * (UBound(SplitByAt) + 1) + j)) & CStr(DSP_INVTemp(i).Element(k))
                    Next k
                    Set DSP_INV(i) = DSP_INV(i).Concatenate(DSP_INVTemp(i)) 'Merge all need flipped bit into one D
                Else
                    DSP_Captured(j) = GetStoreDataAllType(SplitByAt(j))
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
            DSP_Mdll_Capture(i) = GetStoreDataAllType(Mdll_value(z))
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
            
            TestNameInput = Report_TName_From_Instance("X", "x", TestNameInput, CInt(TheExec.flow.TestLimitIndex), 0)
            
        'Next Site
           For Each site In TheExec.sites
''''''''''            TheExec.Flow.TestLimit LowVal:=mdll_low(i), HiVal:=mdll_high(i), resultVal:=Eye_str_long.Element(i) * 8, FormatStr:="%i", TName:=TestNameInput, ForceResults:=tlForceFlow, TNum:=TnumRecord, ScaleType:=scaleNoScaling
                 TheExec.flow.TestLimit lowVal:=mdll_low(i), hiVal:=mdll_high(i), resultVal:=Eye_str_long.Element(i) * 8, formatStr:="%i", Tname:=TestNameInput, ForceResults:=tlForceFlow, tNum:=TnumRecord, scaletype:=scaleNoScaling
            TheExec.flow.TestLimitIndex = TheExec.flow.TestLimitIndex - 1
            Next site
            
            TheExec.flow.TestLimitIndex = TheExec.flow.TestLimitIndex + 1
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
                    DSP_Captured(j) = GetStoreDataAllType(tmp_name)
                    'DSP_Captured(j).CreateRandom 0, 1, 10, 1, DspLong '<- should be replaced by previous Line
                    'Set DSP_INV(i) = DSP_INV(i).Concatenate(DSP_Captured(j)) 'Merge all need flipped bit into one DSP
                    DSP_INVTemp(i).CreateConstant 0, DSP_Captured(j).SampleSize, DspLong
                    For k = 0 To DSP_Captured(j).SampleSize - 1
                        DSP_INVTemp(i).Element(k) = DSP_Captured(j).Element(UBound(DSP_Captured(j).data) - k)
                        DSP_Record(i * (UBound(SplitByAt) + 1) + j) = CStr(DSP_Record(i * (UBound(SplitByAt) + 1) + j)) & CStr(DSP_INVTemp(i).Element(k))
                    Next k
                    Set DSP_INV(i) = DSP_INV(i).Concatenate(DSP_INVTemp(i)) 'Merge all need flipped bit into one D
                Else
                    DSP_Captured(j) = GetStoreDataAllType(SplitByAt(j))
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
            DSP_Mdll_Capture(i) = GetStoreDataAllType(Mdll_value(z))
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
            
            TestNameInput = Report_TName_From_Instance("X", "x", TestNameInput, CInt(TheExec.flow.TestLimitIndex), 0)
       
            For Each site In TheExec.sites
                'TheExec.Flow.TestLimit LowVal:=mdll_low(i)(Site), HiVal:=mdll_high(i)(Site), resultVal:=Eye_str_long(Site).Element(i) * 4, FormatStr:="%i", TName:=TestNameInput, ForceResults:=tlForceFlow, TNum:=TnumRecord, ScaleType:=scaleNoScaling
                TheExec.flow.TestLimit lowVal:=mdll_low(i), hiVal:=mdll_high(i), resultVal:=Eye_str_long.Element(i), formatStr:="%i", Tname:=TestNameInput, ForceResults:=tlForceFlow, tNum:=TnumRecord, scaletype:=scaleNoScaling
        '           TheExec.Flow.TestLimit lowVal:=mdll_low(i)(Site), resultVal:=Eye_str_long(Site).Element(i) * 4, FormatStr:="%i", TName:=TestNameInput, ForceResults:=tlForceFlow, TNum:=TnumRecord
                TheExec.flow.TestLimitIndex = TheExec.flow.TestLimitIndex - 1
            Next site
                TheExec.flow.TestLimitIndex = TheExec.flow.TestLimitIndex + 1
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
Public Function Calc_GPIO_DriverStrength_BAK(argc As Integer, argv() As String) As Long

    'GPIO_Pins=DS_Pins
    'Read Stored_PinListData=GPIO_IOL/IOH
    'Write Dict_name=gpio_iol_8
    
    Dim i As Long
    Dim site As Variant
    Dim TestNameInput As String
    Dim OutputTname_format() As String
    Dim DS_Pins As New PinList: DS_Pins = argv(0)
    Dim Get_StoredData As New PinListData: Get_StoredData = GetStoreDataAllType(argv(1))
    Dim DS_Fuse_Name As String: DS_Fuse_Name = argv(2)
    Dim Fuse_Bit As Long: Fuse_Bit = argv(3)
    Dim DS_Data As New PinListData
    Dim Pin_Ary() As String
    Dim pinCnt As Long
    Dim DS_Data_DSPwave As New DSPWave
    Dim DS_Max As New SiteDouble, DS_Min As New SiteDouble, DS_Avg As New SiteDouble, DS_Result As New SiteDouble
    Dim DS_Max_Diff As New SiteDouble, DS_Min_Diff As New SiteDouble
    Dim Fuse_Bin As New DSPWave: Fuse_Bin.CreateConstant 0, Fuse_Bit
    Dim Fuse_Dec As New DSPWave: Fuse_Dec.CreateConstant 0, 1
    
    
    Call TheExec.DataManager.DecomposePinList(DS_Pins, Pin_Ary, pinCnt)
    DS_Data_DSPwave.CreateConstant 0, pinCnt
    
    For i = 0 To pinCnt - 1
        DS_Data.AddPin (Pin_Ary(i))
        DS_Data.pins(Pin_Ary(i)) = Get_StoredData.pins(Pin_Ary(i))
        For Each site In TheExec.sites
            DS_Data_DSPwave.Element(i) = DS_Data.pins(Pin_Ary(i)).value
        Next site
    Next i
    
    Dim DSP_Result As New DSPWave: DSP_Result.CreateConstant 0, UBound(Pin_Ary)
    
    For Each site In TheExec.sites
        DS_Data_DSPwave = DS_Data_DSPwave.Multiply(1000)
        DS_Avg = Format(DS_Data_DSPwave.CalcMean, "0.0")
        DSP_Result = DS_Data_DSPwave.Subtract(DS_Avg).divide(1000)
        Fuse_Dec.Element(0) = DS_Avg.Abs.Multiply(10)
    Next site
    
    Call HardIP_Dec2Bin(Fuse_Bin, Fuse_Dec, Fuse_Bit)
    Call StoreDataAllType(DS_Fuse_Name, Fuse_Bin)
    
    Dim pin As Variant
    Dim j As Long
    Dim TestName() As String
        TestName() = Split(CStr(argv(1)), "_")
    Dim tempTestLimitIndex As Long
        tempTestLimitIndex = TheExec.flow.TestLimitIndex
    If EnableDigitalTestLimitTTR = True Then
                'TTR,20200423, Oscar
        If TestName(1) = "ioh" Then
              TestNameInput = Report_TName_From_Instance(CalcI, "X", "CurrError" & Replace(DS_Fuse_Name, "_", vbNullString), CInt(i))
              TestNameInput = Replace(TestNameInput, "IOL", "CurrError" & UCase(TestName(1)))
        Else
              TestNameInput = Report_TName_From_Instance(CalcI, "X", "CurrError" & Replace(DS_Fuse_Name, "_", vbNullString), CInt(i))
              TestNameInput = Replace(TestNameInput, "IOL", "CurrError" & UCase(TestName(1)))
         End If
         If Fuse_Bit = 9 Then  'if is DS14
            TheExec.flow.TestLimit DS_Data.Math.Subtract(DS_Avg.divide(1000)), hiVal:=0.005, lowVal:=-0.005, Tname:=TestNameInput, unit:=unitAmp, scaletype:=scaleMicro
        ElseIf Fuse_Bit = 8 Then 'if is DS8
            TheExec.flow.TestLimit DS_Data.Math.Subtract(DS_Avg.divide(1000)), hiVal:=0.002, lowVal:=-0.002, Tname:=TestNameInput, unit:=unitAmp, scaletype:=scaleMicro
         Else
            TheExec.flow.TestLimit DS_Data.Math.Subtract(DS_Avg.divide(1000)), hiVal:=0.001, lowVal:=-0.001, Tname:=TestNameInput, unit:=unitAmp, scaletype:=scaleMicro
        End If
    Else
        For j = 0 To UBound(Pin_Ary)
            'For Each site In TheExec.sites
            TheExec.flow.TestLimitIndex = tempTestLimitIndex
                 If TestName(1) = "ioh" Then
                    TestNameInput = Report_TName_From_Instance(CalcI, CStr(Pin_Ary(j)), "CurrError" & Replace(DS_Fuse_Name, "_", vbNullString), CInt(i))
                    TestNameInput = Replace(TestNameInput, "IOL", "CurrError" & UCase(TestName(1)))
                 Else
                    TestNameInput = Report_TName_From_Instance(CalcI, CStr(Pin_Ary(j)), "CurrError" & Replace(DS_Fuse_Name, "_", vbNullString), CInt(i))
                    TestNameInput = Replace(TestNameInput, "IOL", "CurrError" & UCase(TestName(1)))
                 End If
                 If Fuse_Bit = 9 Then  'if is DS14
                    TheExec.flow.TestLimit DSP_Result.Element(j), hiVal:=0.005, lowVal:=-0.005, Tname:=TestNameInput, PinName:=Pin_Ary(j), unit:=unitAmp, scaletype:=scaleMicro
                ElseIf Fuse_Bit = 8 Then 'if is DS8
                    TheExec.flow.TestLimit DSP_Result.Element(j), hiVal:=0.002, lowVal:=-0.002, Tname:=TestNameInput, PinName:=Pin_Ary(j), unit:=unitAmp, scaletype:=scaleMicro
                 Else
                    TheExec.flow.TestLimit DSP_Result.Element(j), hiVal:=0.001, lowVal:=-0.001, Tname:=TestNameInput, PinName:=Pin_Ary(j), unit:=unitAmp, scaletype:=scaleMicro
                End If
           ' Next site
        Next j
    End If

    If TestName(1) = "ioh" Then
        TestNameInput = Report_TName_From_Instance(CalcI, vbNullString, "CurrAvg", CInt(i))
        TestNameInput = Replace(TestNameInput, "IOL", UCase(TestName(1)))
    Else
        TestNameInput = Report_TName_From_Instance(CalcI, vbNullString, "CurrAvg", CInt(i))
    End If
    
    
    If TestName(1) = "ioh" Then
    
        If Fuse_Bit = 7 Then  'if is DS4
            TheExec.flow.TestLimit DS_Avg.divide(1000), Tname:=TestNameInput, PinName:="CurrAvg", unit:=unitAmp, lowVal:=-0.0127, hiVal:=0
        ElseIf Fuse_Bit = 8 Then   'if is DS8
            TheExec.flow.TestLimit DS_Avg.divide(1000), Tname:=TestNameInput, PinName:="CurrAvg", unit:=unitAmp, lowVal:=-0.0255, hiVal:=0
        Else  'if is DS14
            TheExec.flow.TestLimit DS_Avg.divide(1000), Tname:=TestNameInput, PinName:="CurrAvg", unit:=unitAmp, lowVal:=-0.0511, hiVal:=0
        End If
    Else
        If Fuse_Bit = 7 Then  'if is DS4
            TheExec.flow.TestLimit DS_Avg.divide(1000), Tname:=TestNameInput, PinName:="CurrAvg", unit:=unitAmp, hiVal:=0.0127, lowVal:=0
        ElseIf Fuse_Bit = 8 Then   'if is DS8
            TheExec.flow.TestLimit DS_Avg.divide(1000), Tname:=TestNameInput, PinName:="CurrAvg", unit:=unitAmp, hiVal:=0.0255, lowVal:=0
        Else  'if is DS14
            TheExec.flow.TestLimit DS_Avg.divide(1000), Tname:=TestNameInput, PinName:="CurrAvg", unit:=unitAmp, hiVal:=0.0511, lowVal:=0
        End If
    End If
    
End Function

Public Function Calc_GPIO_DriverStrength(argc As Integer, argv() As String) As Long

    'GPIO_Pins=DS_Pins
    'Read Stored_PinListData=GPIO_IOL/IOH
    'Write Dict_name=gpio_iol_8

    Dim i As Long
    Dim site As Variant
    Dim TestNameInput As String
    Dim OutputTname_format() As String
    Dim DS_Pins As New PinList: DS_Pins = argv(0)
    Dim Get_StoredData As New PinListData: Get_StoredData = GetStoreDataAllType(argv(1))
    Dim DS_Fuse_Name As String: DS_Fuse_Name = argv(2)
    Dim Fuse_Bit As Long: Fuse_Bit = argv(3)
    Dim DS_Data As New PinListData
    Dim Pin_Ary() As String
    Dim pinCnt As Long
    Dim DS_Data_DSPwave As New DSPWave
    Dim DS_Max As New SiteDouble, DS_Min As New SiteDouble, DS_Avg As New SiteDouble, DS_Result As New SiteDouble
    Dim DS_Max_Diff As New SiteDouble, DS_Min_Diff As New SiteDouble
    Dim Fuse_Bin As New DSPWave: Fuse_Bin.CreateConstant 0, Fuse_Bit
    Dim Fuse_Dec As New DSPWave: Fuse_Dec.CreateConstant 0, 1
    
    
    Call TheExec.DataManager.DecomposePinList(DS_Pins, Pin_Ary, pinCnt)
    DS_Data_DSPwave.CreateConstant 0, pinCnt
    
    For i = 0 To pinCnt - 1
        DS_Data.AddPin (Pin_Ary(i))
        DS_Data.pins(Pin_Ary(i)) = Get_StoredData.pins(Pin_Ary(i))
        For Each site In TheExec.sites
            DS_Data_DSPwave.Element(i) = DS_Data.pins(Pin_Ary(i)).value
        Next site
    Next i
    
    Dim DSP_Result As New DSPWave: DSP_Result.CreateConstant 0, UBound(Pin_Ary)
    
    For Each site In TheExec.sites
        DS_Data_DSPwave = DS_Data_DSPwave.Multiply(1000)
        DS_Avg = Format(DS_Data_DSPwave.CalcMean, "0.0")
        DSP_Result = DS_Data_DSPwave.Subtract(DS_Avg).divide(1000)
        Fuse_Dec.Element(0) = DS_Avg.Abs.Multiply(10)
    Next site
 
    Call HardIP_Dec2Bin(Fuse_Bin, Fuse_Dec, Fuse_Bit)
    Call StoreDataAllType(DS_Fuse_Name, Fuse_Bin)
    
    Dim pin As Variant
    Dim j As Long
    Dim TestName() As String
        TestName() = Split(CStr(argv(1)), "_")
    If EnableDigitalTestLimitTTR = True Then
        If TestName(1) = "ioh" Then
              TestNameInput = Report_TName_From_Instance(CalcI, "X", "CurrError" & UCase(Replace(DS_Fuse_Name, "_", "")), CInt(i), , , , , tlForceFlow)
              TestNameInput = Replace(TestNameInput, "IOL", "CurrError" & UCase(TestName(1)))
        Else
              TestNameInput = Report_TName_From_Instance(CalcI, "X", "CurrError" & UCase(Replace(DS_Fuse_Name, "_", "")), CInt(i), , , , , tlForceFlow)
'              TestNameInput = Replace(TestNameInput, "IOL", "CurrError" & UCase(TestName(1)))
         End If
         If Fuse_Bit = 9 Then  'if is DS14
            TheExec.flow.TestLimit DS_Data.Math.Subtract(DS_Avg.divide(1000)), Tname:=TestNameInput, unit:=unitAmp, scaletype:=scaleMicro, ForceResults:=tlForceFlow
        ElseIf Fuse_Bit = 8 Then 'if is DS8
            TheExec.flow.TestLimit DS_Data.Math.Subtract(DS_Avg.divide(1000)), Tname:=TestNameInput, unit:=unitAmp, scaletype:=scaleMicro, ForceResults:=tlForceFlow
         Else
            TheExec.flow.TestLimit DS_Data.Math.Subtract(DS_Avg.divide(1000)), Tname:=TestNameInput, unit:=unitAmp, scaletype:=scaleMicro, ForceResults:=tlForceFlow
        End If
    Else
        For j = 0 To UBound(Pin_Ary)
          '  For Each site In theexec.sites
                 If TestName(1) = "ioh" Then
                    TestNameInput = Report_TName_From_Instance(CalcI, CStr(Pin_Ary(j)), "CurrError" & UCase(Replace(DS_Fuse_Name, "_", "")), CInt(i), , , , , tlForceFlow)
                    TestNameInput = Replace(TestNameInput, "IOL", "CurrError" & UCase(TestName(1)))
                 Else
                    TestNameInput = Report_TName_From_Instance(CalcI, CStr(Pin_Ary(j)), "CurrError" & UCase(Replace(DS_Fuse_Name, "_", "")), CInt(i), , , , , tlForceFlow)
                    TestNameInput = Replace(TestNameInput, "IOL", "CurrError" & UCase(TestName(1)))
                 End If
                 If Fuse_Bit = 9 Then  'if is DS14
                    TheExec.flow.TestLimit DSP_Result.Element(j), Tname:=TestNameInput, PinName:=Pin_Ary(j), unit:=unitAmp, scaletype:=scaleMicro, ForceResults:=tlForceFlow
                ElseIf Fuse_Bit = 8 Then 'if is DS8
                    TheExec.flow.TestLimit DSP_Result.Element(j), Tname:=TestNameInput, PinName:=Pin_Ary(j), unit:=unitAmp, scaletype:=scaleMicro, ForceResults:=tlForceFlow
                 Else
                    TheExec.flow.TestLimit DSP_Result.Element(j), Tname:=TestNameInput, PinName:=Pin_Ary(j), unit:=unitAmp, scaletype:=scaleMicro, ForceResults:=tlForceFlow
                End If
          '  Next site
            
        Next j
    End If
    ''' Update for Donan. Fix Tname TTR issue by adding TestLimitIndex ''remove by Ian because force flow 20240301
    ''TheExec.Flow.TestLimitIndex = TheExec.Flow.TestLimitIndex + 1
    
    If TestName(1) = "ioh" Then
        TestNameInput = Report_TName_From_Instance(CalcI, "", "CurrAvgIOH", CInt(i), , , , , tlForceFlow)
        TestNameInput = Replace(TestNameInput, "IOL", UCase(TestName(1)))
    Else
        TestNameInput = Report_TName_From_Instance(CalcI, "", "CurrAvgIOL", CInt(i), , , , , tlForceFlow)
    End If
    
    If TestName(1) = "ioh" Then
    
        If Fuse_Bit = 7 Then  'if is DS4
            TheExec.flow.TestLimit DS_Avg.divide(1000), Tname:=TestNameInput, PinName:="CurrAvgIOH", unit:=unitAmp, ForceResults:=tlForceFlow
        ElseIf Fuse_Bit = 8 Then   'if is DS8
            TheExec.flow.TestLimit DS_Avg.divide(1000), Tname:=TestNameInput, PinName:="CurrAvgIOH", unit:=unitAmp, ForceResults:=tlForceFlow
        Else  'if is DS14
            TheExec.flow.TestLimit DS_Avg.divide(1000), Tname:=TestNameInput, PinName:="CurrAvgIOH", unit:=unitAmp, ForceResults:=tlForceFlow
        End If
    Else
        If Fuse_Bit = 7 Then  'if is DS4
            TheExec.flow.TestLimit DS_Avg.divide(1000), Tname:=TestNameInput, PinName:="CurrAvgIOL", unit:=unitAmp, ForceResults:=tlForceFlow
        ElseIf Fuse_Bit = 8 Then   'if is DS8
            TheExec.flow.TestLimit DS_Avg.divide(1000), Tname:=TestNameInput, PinName:="CurrAvgIOL", unit:=unitAmp, ForceResults:=tlForceFlow
        Else  'if is DS14
            TheExec.flow.TestLimit DS_Avg.divide(1000), Tname:=TestNameInput, PinName:="CurrAvgIOL", unit:=unitAmp, ForceResults:=tlForceFlow
        End If
      End If
    ''' Update for Donan. Fix Tname TTR issue by adding TestLimitIndex ''remove by Ian because force flow 20240301
    ''TheExec.Flow.TestLimitIndex = TheExec.Flow.TestLimitIndex + 1

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
    
    Input_V0 = GetStoreDataAllType(Dict_V0)
    Input_V1 = GetStoreDataAllType(Dict_V1)
    Input_V2 = GetStoreDataAllType(Dict_V2)
    
    Voff_PinListData.AddPin (Input_V1.pins(0))
    Voff_PinListData = Input_V1.pins(0).Subtract(Input_V0.pins(0))
    Voff_PinListData = Voff_PinListData.Math.divide(0.001)
    
    Voff_PinListData_Round.AddPin (Input_V1.pins(0))
    Voff_PinListData_Round = Voff_PinListData.pins(0).divide(0.5)
    For Each site In TheExec.sites
        Voff_PinListData_Round.pins(0).value(site) = CDbl(FormatNumber(Voff_PinListData_Round.pins(0).value(site), 0))
    Next site
    
    Ratio_off_PinListData.AddPin (Input_V1.pins(0))
    Ratio_off_PinListData = Input_V1.pins(0).divide(Input_V2.pins(0)).divide(2).Subtract(1)
    For Each site In TheExec.sites
            Ratio_off_Per_DSP(site).Element(0) = Ratio_off_PinListData.pins(0).value(site)
    Next site
    
    Call StoreDataAllType(Dict_Ratio_off_Per, Ratio_off_Per_DSP)
                                       
    TheExec.flow.TestLimit resultVal:=Voff_PinListData.pins(0), ForceResults:=tlForceFlow
    TheExec.flow.TestLimit resultVal:=Voff_PinListData_Round.pins(0), ForceResults:=tlForceFlow
    For Each site In TheExec.sites
        If Voff_PinListData_Round.pins(0).value(site) < 0 Then
            UnSinged_Voff_Round(site).Element(0) = Voff_PinListData_Round.pins(0).value(site) + (2 ^ Fuse_BitCount)
        Else
            UnSinged_Voff_Round(site).Element(0) = Voff_PinListData_Round.pins(0).value(site)
        End If
    Next site
    
    Call StoreDataAllType(Fuse_Voff_Round, UnSinged_Voff_Round)
    
    TheExec.flow.TestLimit resultVal:=Ratio_off_PinListData.pins(0), ForceResults:=tlForceFlow
    
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
        DSPWave_Ratio_off_per(i) = GetStoreDataAllType(argv(i))
        Call rundsp.DSP_Add(DSPWave_Average, DSPWave_Ratio_off_per(i))
    Next i
    Call rundsp.DSP_DivideConstant(DSPWave_Average, Sweep_Count)
    
    If TheExec.TesterMode = testModeOffline Then            'for offline run
        For Each site In TheExec.sites
            DSPWave_Average(site).Element(0) = -0.00060171
        Next site
    End If
    
    TheExec.flow.TestLimit resultVal:=DSPWave_Average.Element(0), Tname:="Ratio_off_per_avg", ForceResults:=tlForceFlow 'transfer_to_forceflow
    
    For Each site In TheExec.sites
        DSPWave_Round_Average(site).Element(0) = FormatNumber(DSPWave_Average(site).Element(0) * 1600, 0)
    Next site
    
    TheExec.flow.TestLimit resultVal:=DSPWave_Round_Average.Element(0), Tname:="Round_Ratio_off_per_avg", ForceResults:=tlForceFlow 'transfer_to_forceflow
    
    For Each site In TheExec.sites
        If DSPWave_Round_Average(site).Element(0) < 0 Then
            DSPWave_Unsinged_Round_Average(site).Element(0) = DSPWave_Round_Average(site).Element(0) + (2 ^ Fuse_BitCount)
        Else
            DSPWave_Unsinged_Round_Average(site).Element(0) = DSPWave_Round_Average(site).Element(0)
        End If
    Next site
    
    Call StoreDataAllType(Dict_Unsinged_Round_Avg, DSPWave_Unsinged_Round_Average)
    
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
        
        DSP_DictKey_2S_BIN = GetStoreDataAllType(DictKey_2S_BIN)
        
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
        
        Call StoreDataAllType(DictKey_SIGN_DEC, DSP_DictKey_SIGN_DEC(i))

        TestNameInput = Report_TName_From_Instance(CalcC, "X", , CInt(i))
        
        Call rundsp.DSP_DivideConstant(DSP_DictKey_SIGN_DEC(i), DivConst)
        
        TheExec.flow.TestLimit resultVal:=DSP_DictKey_SIGN_DEC(i).Element(0), Tname:=TestNameInput, ForceResults:=tlForceFlow
        
    Next i


End Function

Public Function Calc_BinStr2HexStr(ByVal binstr As String, ByVal HexBit As Long) As String
On Error GoTo errHandler 'Add ErrHandler 2023/05/29

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

    Calc_BinStr2HexStr = HexStr
Exit Function 'Add ErrHandler 2023/05/29errHandler: 'Add ErrHandler 2023/05/29
errHandler: 'Add ErrHandler 2023/05/29
    Call Print_Error_Message(Error_Info, "LIB_HardIP_Calc", "Calc_BinStr2HexStr") 'Add ErrHandler 2023/05/29
    If AbortTest Then Exit Function Else Resume Next 'Add ErrHandler 2023/05/29

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
    
    PinList_V1 = GetStoreDataAllType(Dict_V1)
    PinList_V2 = GetStoreDataAllType(Dict_V2)
    
    PinList_Vdiff.AddPin (PinList_V1.pins(0))
    PinList_Vdiff.pins(0) = PinList_V1.Math.Subtract(PinList_V2)
    
    Call StoreDataAllType(Dict_Vdiff, PinList_Vdiff)
    
    TheExec.Datalog.WriteComment ("Voltage Difference Calculation")
    
End Function

Public Function Calc_DigCap_Avg_Store_old(argc As Integer, argv() As String) As Long
    Dim i As Long
    'Dim site As Variant
    '''Syntax: Alg::Calc_DigCap_Avg_Store(aneivdm_1,...,aneivdm_30, ("2SCOMPLEMENT"), aneivdm_trim_low1)
        '''                     aneivdm_trim_low1 : Dict Store Name used for Calc_Eqn "Calc_DigCap_Offset_Store"
        Dim DSPWave_Bin() As New DSPWave
    Dim DSPWave_Dec() As New DSPWave
    Dim DSPWave_Bin_MSB1st() As New DSPWave
    ReDim DSPWave_Bin(argc - 3) As New DSPWave
    ReDim DSPWave_Dec(argc - 3) As New DSPWave
    ReDim DSPWave_Bin_MSB1st(argc - 3) As New DSPWave
    'DSPWave_Dec(0).CreateConstant 0, 1, DspDouble
    Dim DSPWave_AverageDec As New DSPWave
    DSPWave_AverageDec.CreateConstant 0, 1, DspDouble
    
    Dim DSPWave_SumDec As New DSPWave
    DSPWave_SumDec.CreateConstant 0, 1, DspLong
    
    Dim DSPWave_AverageBin As New DSPWave
    'DSPWave_AverageBin.CreateConstant 0, 18
    Dim TestNameInput As String
    
    Dim SL_BitWidth As New SiteLong
    
    Dim Is_2sComplement As Boolean: Is_2sComplement = False
    TheHdw.dsp.ExecutionMode = tlDSPModeHostDebug
    
    If UCase(argv(argc - 2)) = "2SCOMPLEMENT" Then Is_2sComplement = True
    
    For i = 0 To argc - 3
        DSPWave_Bin(i) = GetStoreDataAllType(argv(i))
        For Each site In TheExec.sites
            SL_BitWidth(site) = DSPWave_Bin(i)(site).SampleSize
        Next site
        
        If UCase(Instance_Data.CUS_Str_MainProgram) = UCase("DigCap_LSBtoMSB") Then
            rundsp.Split_Dspwave_Reverse DSPWave_Bin(i), DSPWave_Bin_MSB1st(i)
            DSPWave_Bin(i) = DSPWave_Bin_MSB1st(i)
        End If
        
        DSPWave_Dec(i).CreateConstant 0, 1, DspLong
        
        If Is_2sComplement = True Then
            Call rundsp.DSP_2S_Complement_To_SignDec(DSPWave_Bin(i), SL_BitWidth, DSPWave_Dec(i))
        Else
            Call rundsp.BinToDec(DSPWave_Bin(i), DSPWave_Dec(i))
        End If
        
        For Each site In TheExec.sites
            DSPWave_Dec(i)(site) = DSPWave_Dec(i)(site).ConvertDataTypeTo(DspLong)
        Next site
        Call rundsp.DSP_Add(DSPWave_SumDec, DSPWave_Dec(i))
    Next i
    For Each site In TheExec.sites
        DSPWave_AverageDec = DSPWave_SumDec.ConvertDataTypeTo(DspDouble)
    Next site
    Call rundsp.DSP_DivideConstant(DSPWave_AverageDec, argc - 2)

    For Each site In TheExec.sites
        TheExec.Datalog.WriteComment "Site : " & site & ", Average result:" & DSPWave_AverageDec(site).Element(0)
        DSPWave_AverageDec(site).Element(0) = FormatNumber(DSPWave_AverageDec(site).Element(0), 0)
    Next site
    
    Call rundsp.DSPWf_Dec2Binary(DSPWave_AverageDec, SL_BitWidth, DSPWave_AverageBin)
    'Call StoreDataAllType(argv(argc - 1), DSPWave_AverageDec)

    TestNameInput = Report_TName_From_Instance(CalcC, "X", , CInt(i), , , , , tlForceFlow)

    TheExec.flow.TestLimit resultVal:=DSPWave_AverageDec.Element(0), Tname:=TestNameInput, ForceResults:=tlForceFlow  'tlForceFlow
    Call StoreDataAllType(argv(argc - 1), DSPWave_AverageBin)
    
    TheHdw.dsp.ExecutionMode = tlDSPModeAutomatic
End Function


Public Function Calc_DigCap_Avg_Store(argc As Integer, argv() As String) As Long
On Error GoTo errHandler 'Add ErrHandler 2023/05/29
    'Update for IVDM function TTR (DSP process from local to DSPPC )-- 20220105
    '''Syntax: Alg::Calc_DigCap_Avg_Store(aneivdm_1,...,aneivdm_30, ("2SCOMPLEMENT"), aneivdm_trim_low1,True[STDEV_Flag])
    '''                     aneivdm_trim_low1 : Dict Store Name used for Calc_Eqn "Calc_DigCap_Offset_Store_Central"
    '''                     STDEV_Flag : True:Enable and printing /False:Disable
    '20230116 Update new function IVDM calculate STDEV

    Dim i As Long
    Dim serial_data As Long
    Dim Input_Arry() As String
    Dim sLen As Integer
    Dim Pasing_data_Arry() As String
    Dim Input_Str As String
    Input_Str = vbNullString
    Dim STDEV_Flag As Boolean           'New for IVDM STDEV -- 20230116
    Dim DSPWave_STDEV As New DSPWave    'New for IVDM STDEV -- 20230116
    
    Dim DSPWave_Bin() As New DSPWave
    Dim DSPWave_AverageBin() As New DSPWave
    Dim DSPWave_AverageDec As New DSPWave
    Dim DSPWave_AverageDec2Bin As New DSPWave
    Dim TestNameInput As String
    Dim DSPWave_ALL_DATA As New DSPWave
    Dim DSPWave_All_DATA_1D As New DSPWave
    Dim DSPWave_2SComplement_Flag As New DSPWave
    Dim DSPWave_LSBtoMSB_Flag As New DSPWave
    Dim DSPWave_Dictionry_amount As New DSPWave
    Dim DSPWave_Split_Bit_perDic As New DSPWave
    Dim Reg_total_bits As Long: Reg_total_bits = 0
    Dim vsite As Variant
    Dim j As Long: j = 0
    Dim Sample_Size As Long: Sample_Size = 0
    Dim Split_bit As Long: Split_bit = 0
    
    DSPWave_STDEV.CreateConstant 0, 1, DspDouble
    DSPWave_ALL_DATA.CreateConstant 0, 0, DspLong
    
    
    For i = 0 To argc - 1
        If i <> argc - 1 Then
            Input_Str = Input_Str & argv(i) & ","
        Else
            Input_Str = Input_Str & argv(i)
        End If
    Next i
    Input_Arry = Split(Input_Str, "@")
    
    DSPWave_2SComplement_Flag.CreateConstant 0, UBound(Input_Arry) + 1, DspLong
    DSPWave_LSBtoMSB_Flag.CreateConstant 0, UBound(Input_Arry) + 1, DspLong
    DSPWave_Dictionry_amount.CreateConstant 0, UBound(Input_Arry) + 1, DspLong
    DSPWave_Split_Bit_perDic.CreateConstant 0, UBound(Input_Arry) + 1, DspLong
    
    For serial_data = 0 To UBound(Input_Arry)
        Pasing_data_Arry = Split(Input_Arry(serial_data), ",")
        sLen = UBound(Pasing_data_Arry)
        ReDim DSPWave_Bin(sLen - 3) As New DSPWave
        
        If sLen - 3 = 0 Then
            TheExec.Datalog.WriteComment ("Error! Divide 0.")
            Exit Function
        End If

        For i = 0 To sLen - 3
            DSPWave_Bin(i) = GetStoreDataAllType(Pasing_data_Arry(i))
            For Each vsite In TheExec.sites
                DSPWave_ALL_DATA = DSPWave_ALL_DATA.Concatenate(DSPWave_Bin(i))
            Next vsite
            If i = 0 Then
                For Each vsite In TheExec.sites
                    Sample_Size = DSPWave_Bin(i).SampleSize
                    Exit For
                Next vsite
            End If
        Next i
        DSPWave_Split_Bit_perDic.Element(serial_data) = Sample_Size
    Next serial_data

    DSPWave_All_DATA_1D.CreateConstant 0, Reg_total_bits, DspLong
    
    For serial_data = 0 To UBound(Input_Arry)
        Pasing_data_Arry = Split(Input_Arry(serial_data), ",")
        sLen = UBound(Pasing_data_Arry)
        ReDim DSPWave_Bin(sLen - 3) As New DSPWave
        
        If UCase(Pasing_data_Arry(sLen - 2)) = "2SCOMPLEMENT" Then
            DSPWave_2SComplement_Flag.Element(serial_data) = 1
        Else
            DSPWave_2SComplement_Flag.Element(serial_data) = 0
        End If
        
        DSPWave_Dictionry_amount.Element(serial_data) = sLen - 2
        
        If UCase(Instance_Data.CUS_Str_MainProgram) = UCase("DigCap_LSBtoMSB") Then
            DSPWave_LSBtoMSB_Flag.Element(serial_data) = 1
        Else
            DSPWave_LSBtoMSB_Flag.Element(serial_data) = 0
        End If
    Next serial_data
    j = 0
    
    TheHdw.dsp.ExecutionMode = tlDSPModeForceAutomatic
    DSPWave_AverageDec.CreateConstant 0, 1, DspDouble
    DSPWave_AverageDec2Bin.CreateConstant 0, 8, DspLong
   
    Call rundsp.DSPWF_AVG(DSPWave_ALL_DATA, DSPWave_Dictionry_amount, DSPWave_Split_Bit_perDic, DSPWave_2SComplement_Flag, DSPWave_LSBtoMSB_Flag, DSPWave_AverageDec, DSPWave_AverageDec2Bin, DSPWave_STDEV)
    
    Dim tmp_dsp As New DSPWave
    ReDim DSPWave_AverageBin(UBound(Input_Arry)) As New DSPWave

    tmp_dsp.CreateConstant 0, 1, DspDouble
    For serial_data = 0 To UBound(Input_Arry)
        For Each vsite In TheExec.sites
            TheExec.Datalog.WriteComment "Site : " & vsite & ", Average result:" & DSPWave_AverageDec(vsite).Element(serial_data)
            DSPWave_AverageDec(vsite).Element(serial_data) = FormatNumber(DSPWave_AverageDec(vsite).Element(serial_data), 0)
        Next vsite

        TestNameInput = Report_TName_From_Instance(CalcC, "X", , CInt(i), , , , , tlForceFlow)
        TheExec.flow.TestLimit resultVal:=DSPWave_AverageDec.Element(serial_data), Tname:=TestNameInput, ForceResults:=tlForceFlow

        Pasing_data_Arry = Split(Input_Arry(serial_data), ",")
        sLen = UBound(Pasing_data_Arry)
        
        If Pasing_data_Arry(sLen) = True Then
            '--- Print STDEV ---- 20230116
            TestNameInput = Report_TName_From_Instance(CalcC, "X", , CInt(i), , , , , tlForceFlow)
            TheExec.flow.TestLimit resultVal:=DSPWave_STDEV.Element(serial_data), Tname:=TestNameInput, ForceResults:=tlForceFlow
        End If

        Call StoreDataAllType(Pasing_data_Arry(sLen - 1), DSPWave_AverageDec2Bin)
    Next serial_data
Exit Function 'Add ErrHandler 2023/05/29errHandler: 'Add ErrHandler 2023/05/29
errHandler: 'Add ErrHandler 2023/05/29
    Call Print_Error_Message(Error_Info, "LIB_HardIP_Calc", "Calc_DigCap_Avg_Store") 'Add ErrHandler 2023/05/29
    If AbortTest Then Exit Function Else Resume Next 'Add ErrHandler 2023/05/29

End Function

Public Function Calc_MetrologyGR_t5p5(argc As Integer, argv() As String) As Long
    Dim Vsrp As New SiteDouble
    Dim Vsrn As New SiteDouble
    Dim Vdiff As New SiteDouble
    Dim TestNameInput As String
    Dim site As Variant
    Dim OutputTname_format() As String
    
    'GetStoreDataAllType
    For Each site In TheExec.sites
        Vsrp = GetStoreDataAllType(argv(0))
        Vsrn = GetStoreDataAllType(argv(1))
        Vdiff = Vsrn.Subtract(Vsrp).divide(0.000005)
    Next site
    
    TestNameInput = Report_TName_From_Instance(CalcV, "X", "Rpsr")
    OutputTname_format = Split(TestNameInput, "_")
    OutputTname_format(6) = "Rpsr"
    TestNameInput = Merge_TName(OutputTname_format)
    TheExec.flow.TestLimit resultVal:=Vdiff, Tname:=TestNameInput, ForceResults:=tlForceFlow
    
End Function

Public Function Calc_MetrologyGR_t1p0(argc As Integer, argv() As String) As Long
    Dim Vsrp As New SiteDouble
    Dim Vsrn As New SiteDouble
    Dim Vdiff As New SiteDouble
    Dim TestNameInput As String
    Dim site As Variant
    Dim OutputTname_format() As String
    
    'GetStoreDataAllType
    For Each site In TheExec.sites
        Vsrp = GetStoreDataAllType(argv(0))
        Vsrn = GetStoreDataAllType(argv(1))
        Vdiff = Vsrp.Subtract(Vsrn)
    Next site
    
    TestNameInput = Report_TName_From_Instance(CalcV, "X", "Rpsr")
    OutputTname_format = Split(TestNameInput, "_")
    OutputTname_format(6) = "Vdiff"
    OutputTname_format(7) = CStr(TheExec.flow.var("SrcCodeIndx").value)
    TestNameInput = Merge_TName(OutputTname_format)
    TheExec.flow.TestLimit resultVal:=Vdiff, Tname:=TestNameInput, ForceResults:=tlForceFlow
    
End Function



Public Function Calc_PCIE_RXTERM(argc As Integer, argv() As String) As Long
    Dim DSP_RCAL_TX_DIV4_CODE As New DSPWave: DSP_RCAL_TX_DIV4_CODE = GetStoreDataAllType(argv(0))
    Dim DSP_RXTERM_CODE_Binary As New DSPWave
    Dim DSP_RXTERM_CODE_Dec As New DSPWave
    Dim TestNameInput As String
    For Each site In TheExec.sites.Active
        DSP_RXTERM_CODE_Binary = DSP_RCAL_TX_DIV4_CODE.Select(0, 1, DSP_RCAL_TX_DIV4_CODE.SampleSize - 1).COPY
        DSP_RXTERM_CODE_Dec = DSP_RXTERM_CODE_Binary.ConvertStreamTo(tldspParallel, DSP_RXTERM_CODE_Binary.SampleSize, 0, Bit0IsMsb)
    Next site
    TestNameInput = Report_TName_From_Instance(CalcC, vbNullString)
    TheExec.flow.TestLimit resultVal:=DSP_RXTERM_CODE_Dec.Element(0), Tname:=TestNameInput, ForceResults:=tlForceFlow
    Call StoreDataAllType(argv(1), DSP_RXTERM_CODE_Binary)
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
    
    OutDspWave = GetStoreDataAllType(get_store_name)
    
    
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
            
            'Call StoreDataAllType(argv(2), source_dsp)
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
     Call StoreDataAllType(argv(2), source_dsp)

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
            DSPWave_UnitSegment(i) = GetStoreDataAllType(argv(i + 1))
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
            CalcResult(i).pins("Mean").value(site) = mean
            CalcResult(i).pins("STDEV").value(site) = STDEV
            CalcResult(i).pins("Variance").value(site) = STDEV * STDEV
        Next i
        
    Next site
   
    
    TheExec.Datalog.WriteComment vbNullString
    
    For i = 0 To GroupCount - 1
        'TestNameInput = Report_TName_From_Instance("C", "X", , CInt(i))
        
        For Each site In TheExec.sites.Active
            For j = 0 To SegmentCount - 1
                TheExec.flow.TestLimit resultVal:=DSPWave_MergedSegment(i).data(j), _
                Tname:=Report_TName_From_Instance("C", "X", argv(i + 1), Instance_Data.TestSeqNum + CInt(i * (SegmentCount + 2) + j)), ForceResults:=tlForceFlow 'transfer_to_forceflow
            Next j
                TheExec.flow.TestLimit resultVal:=CalcResult(i).pins("Mean").value(site), _
                Tname:=Report_TName_From_Instance("C", "X", argv(i + 1) & "Mean", Instance_Data.TestSeqNum + CInt(i * (SegmentCount + 2) + SegmentCount)), ForceResults:=tlForceFlow 'transfer_to_forceflow
                TheExec.flow.TestLimit resultVal:=CalcResult(i).pins("Variance").value(site), _
                Tname:=Report_TName_From_Instance("C", "X", argv(i + 1) & "Variance", Instance_Data.TestSeqNum + CInt(i * (SegmentCount + 2) + SegmentCount + 1)), ForceResults:=tlForceFlow 'transfer_to_forceflow
        Next site
        
    Next i
    
    'ReDim DSPWave_Avg_Bin(argc - 3) As New DSPWave
    
'    Dim TestName As String
'    Dim Site As Variant
'    Dim Dict As String
'    Dim BitWidth As Long
'
'    For i = 0 To 1
'        DSPWave_Binary(i) = GetStoreDataAllType(argv(i))
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
'    Call StoreDataAllType(Dict, DSPWave_Avg_Bin)
'    Dim TestNameInput As String
'    Dim OutputTname_format() As String
'
'    TestNameInput = Report_TName_From_Instance("C", "X", , CInt(i))
'
'    TheExec.Flow.TestLimit resultVal:=DSPWave_Avg_Dec.Element(0), TName:=TestNameInput, ForceResults:=tlForceNone
    
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
                    DSPWaveFromDict = GetStoreDataAllType(LCase(SplitCap(j) & SplitBias(i)))
                    ''DSPWaveFromDict.Element(i + 4) = 1 '' Debug
                    DSPWaveDecType = DSPWaveFromDict.ConvertStreamTo(tldspParallel, 16, 0, Bit0IsMsb)
                    If DSPWaveDecType.Element(0) >= CInt(SplitTrimFreq(1)) Then
                        F_TrimComplete(j) = True
                        VroVoltage(j) = GetStoreDataAllType(CStr(LCase(SplitCap(j) & SplitBias(i) & SplitVro(0))))
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
           
            UseLimitTname = CStr(Instance_Data.Tname(TheExec.flow.TestLimitIndex))          ' Dylan Edit by 20190616
            TestNameInput = Report_TName_From_Instance("Calc", "X", UseLimitTname, 0, 0)
            If FinalBiasIndex = 7 And FinalVroIndex = 0 Then
                TheExec.flow.TestLimit resultVal:=1, hiVal:=0, lowVal:=0, Tname:=TestNameInput, ForceResults:=tlForceFlowFail
            Else
                TheExec.flow.TestLimit resultVal:=0, hiVal:=0, lowVal:=0, Tname:=TestNameInput, ForceResults:=tlForceFlowPass
            End If
            TheExec.flow.TestLimitIndex = TheExec.flow.TestLimitIndex - 1       ' Edited for avoid useLimit index messy
            
            Call Dec2Bin(CLng(SplitDCO(1)), DCO_arry())
            For k = 0 To UBound(DCO_arry)
                FinalDSPWave.Element(k + UBound(Bias_arry) + UBound(Cap_arry) + 1 + 1) = DCO_arry((UBound(DCO_arry)) - k) 'Bit start from 5
            Next k
        End If
        TheExec.Datalog.WriteComment ("Final Bias Index : " & FinalBiasIndex)
        TheExec.Datalog.WriteComment ("Final Cap Index : " & FinalVroIndex)
        TheExec.Datalog.WriteComment ("DCO : " & CLng(SplitDCO(1)))
    Next site
    Call StoreDataAllType(LCase(argv(5)), FinalDSPWave)
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
''''''''''                    DSPWaveTemp = GetStoreDataAllType(SplitCH(i + 1) & SplitDQ(j + 1) & "x" & SplitCountIN(k + 1) & SplitCountIN(0))
''''''''''                    DSPWaveTemp1 = GetStoreDataAllType(SplitCH(i + 1) & SplitDQ(j + 1) & "x" & SplitCount100(k + 1) & SplitCount100(0))
                    DSPWaveDec = GetStoreDataAllType("2SDEC_" & SplitCH(i + 1) & SplitDQ(j + 1) & "x" & SplitCountIN(k + 1) & SplitCountIN(0))
                    DSPWaveDec1 = GetStoreDataAllType("2SDEC_" & SplitCH(i + 1) & SplitDQ(j + 1) & "x" & SplitCount100(k + 1) & SplitCount100(0))
                    
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
            
            TestNameInput = Report_TName_From_Instance("Calc", SplitCH(i + 1) & SplitDQ(j + 1), "DC_Input_CLK", CInt(i), , "replace;7=UP", , , tlForceFlow) 'transfer_to_forceflow
            TheExec.flow.TestLimit resultVal:=DC_Input_CLK_UP.Multiply(100), Tname:=TestNameInput, ForceResults:=tlForceFlow, scaletype:=scaleNoScaling, formatStr:="%1.2f", unit:=unitCustom, customUnit:="%" 'transfer_to_forceflow
            TestNameInput = Report_TName_From_Instance("Calc", SplitCH(i + 1) & SplitDQ(j + 1), "DC_Input_CLK", CInt(i), , "replace;7=NODCC", , , tlForceFlow) 'transfer_to_forceflow
            TheExec.flow.TestLimit resultVal:=DC_Input_CLK_NO_DCC.Multiply(100), Tname:=TestNameInput, ForceResults:=tlForceFlow, scaletype:=scaleNoScaling, formatStr:="%1.2f", unit:=unitCustom, customUnit:="%" 'transfer_to_forceflow
            TestNameInput = Report_TName_From_Instance("Calc", SplitCH(i + 1) & SplitDQ(j + 1), "DC_Input_CLK", CInt(i), , "replace;7=DOWN", , , tlForceFlow) 'transfer_to_forceflow
            TheExec.flow.TestLimit resultVal:=DC_Input_CLK_DOWN.Multiply(100), Tname:=TestNameInput, ForceResults:=tlForceFlow, scaletype:=scaleNoScaling, formatStr:="%1.2f", unit:=unitCustom, customUnit:="%" 'transfer_to_forceflow
            
            TestNameInput = Report_TName_From_Instance("Calc", SplitCH(i + 1) & SplitDQ(j + 1), "DC_Skew_Input_CLK", CInt(i), , "replace;7=UP", , , tlForceFlow) 'transfer_to_forceflow
            TheExec.flow.TestLimit resultVal:=DC_Skew_Input_CLK_UP.Multiply(100), Tname:=TestNameInput, ForceResults:=tlForceFlow, scaletype:=scaleNoScaling, formatStr:="%1.2f", unit:=unitCustom, customUnit:="%" 'transfer_to_forceflow
            TestNameInput = Report_TName_From_Instance("Calc", SplitCH(i + 1) & SplitDQ(j + 1), "DC_Skew_Input_CLK", CInt(i), , "replace;7=NODCC", , , tlForceFlow) 'transfer_to_forceflow
            TheExec.flow.TestLimit resultVal:=DC_Skew_Input_CLK_NO_DCC.Multiply(100), Tname:=TestNameInput, ForceResults:=tlForceFlow, scaletype:=scaleNoScaling, formatStr:="%1.2f", unit:=unitCustom, customUnit:="%" 'transfer_to_forceflow
            TestNameInput = Report_TName_From_Instance("Calc", SplitCH(i + 1) & SplitDQ(j + 1), "DC_Skew_Input_CLK", CInt(i), , "replace;7=DOWN", , , tlForceFlow) 'transfer_to_forceflow
            TheExec.flow.TestLimit resultVal:=DC_Skew_Input_CLK_DOWN.Multiply(100), Tname:=TestNameInput, ForceResults:=tlForceFlow, scaletype:=scaleNoScaling, formatStr:="%1.2f", unit:=unitCustom, customUnit:="%" 'transfer_to_forceflow
            
            TestNameInput = Report_TName_From_Instance("Calc", SplitCH(i + 1) & SplitDQ(j + 1), "DCC_RANGE_UP", CInt(i), , , , , tlForceFlow) 'transfer_to_forceflow
            TheExec.flow.TestLimit resultVal:=DCC_RANGE_UP.Multiply(100), Tname:=TestNameInput, ForceResults:=tlForceFlow, scaletype:=scaleNoScaling, formatStr:="%1.2f", unit:=unitCustom, customUnit:="%" 'transfer_to_forceflow
            TestNameInput = Report_TName_From_Instance("Calc", SplitCH(i + 1) & SplitDQ(j + 1), "DCC_RANGE_DOWN", CInt(i), , , , , tlForceFlow) 'transfer_to_forceflow
            TheExec.flow.TestLimit resultVal:=DCC_RANGE_DOWN.Multiply(100), Tname:=TestNameInput, ForceResults:=tlForceFlow, scaletype:=scaleNoScaling, formatStr:="%1.2f", unit:=unitCustom, customUnit:="%" 'transfer_to_forceflow
        
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
        BinaryWave = GetStoreDataAllType(FormatSplit(0))
        
        For Each site In TheExec.sites
            For j = 0 To BinaryWave.SampleSize - 1
                AssembleStr(i) = CStr(BinaryWave(site).Element(j)) & AssembleStr(i)
            Next j
            BitPosition = 0
            For j = 0 To UBound(RegSplit)
               
                SplitDspWave(j).CreateConstant 0, 1, DspLong
                BinaryWave = BinaryWave(site).ConvertDataTypeTo(DspLong)
                CalDSPWave = BinaryWave(site).Select(CLng(BitPosition), 1, CLng(RegSplit(j))).COPY
                SplitDspWave(j) = CalDSPWave.ConvertStreamTo(tldspParallel, CLng(RegSplit(j)), 0, Bit0IsMsb)
                BitPosition = BitPosition + CLng(RegSplit(j))
            Next j
            duty_cycle(i).Element(0) = (SplitDspWave(0).Element(0) * 2) / (SplitDspWave(1).Element(0) + 2 ^ 17)
            duty_cycle(i).Element(0) = duty_cycle(i).Element(0) * 100
        Next site
        TheExec.Datalog.WriteComment FormatSplit(0) & " Binary Value : " & AssembleStr(i)
    Next i
    For i = 0 To argc - 1
        TheExec.flow.TestLimit resultVal:=duty_cycle(i).Element(0), Tname:=FormatSplit(0), ForceResults:=tlForceFlow
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
                new_DSP = RAW_DSP.Select(0, sgmt_size, ADC_bits).COPY
            Else
                new_DSP = new_DSP.Concatenate(RAW_DSP.Select(i * ADC_bits * sgmt_size + j, sgmt_size, ADC_bits).COPY)
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
        DSPWave_Ori(i) = GetStoreDataAllType(key(i))
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
            CalcResult(i).pins("Mean").value(site) = mean
            CalcResult(i).pins("Variance").value(site) = STDEV * STDEV
            CalcResult(i).pins("MaxErr").value(site) = ADC_result(i).CalcMaximumValue - ADC_result(i).CalcMinimumValue
        Next i
    Next site
    
    k = 0
    For i = 0 To GroupCount - 1
        For Each site In TheExec.sites.Active
            If True Then
                For j = 0 To ADC_result(i).SampleSize - 1
                    TheExec.flow.TestLimit ADC_result(i).data(j), Limit_Val(i) - Limit_Dev, Limit_Val(i) + Limit_Dev, _
                    Tname:=Report_TName_From_Instance("C", "X", key(i), Instance_Data.TestSeqNum + CInt(k)), ForceResults:=tlForceFlow 'transfer_to_forceflow
                    k = k + 1
                Next j
            End If
            TheExec.flow.TestLimit resultVal:=CalcResult(i).pins("Mean").value(site), _
            Tname:=Report_TName_From_Instance("C", "X", key(i), Instance_Data.TestSeqNum + CInt(k + 1), , "replace;7=Mean"), ForceResults:=tlForceFlow 'transfer_to_forceflow
            TheExec.flow.TestLimit resultVal:=CalcResult(i).pins("Variance").value(site), _
            Tname:=Report_TName_From_Instance("C", "X", key(i), Instance_Data.TestSeqNum + CInt(k + 2), , "replace;7=Variance"), ForceResults:=tlForceFlow 'transfer_to_forceflow
            TheExec.flow.TestLimit resultVal:=CalcResult(i).pins("MaxErr").value(site), _
            Tname:=Report_TName_From_Instance("C", "X", key(i), Instance_Data.TestSeqNum + CInt(k + 3), , "replace;7=MaxErr"), ForceResults:=tlForceFlow 'transfer_to_forceflow
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
        
        PartialWave = InWave.Select(i * sgmt_size, 1, sgmt_size).COPY
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
    
    InWave = GetStoreDataAllType(argv(0))
    InWave = InWave.ConvertDataTypeTo(DspLong)
    
    For i = 0 To 2
        k = 0
        'Chk_wave(i) = InWave.Select(13 * (i + 1) - 1, 39, 256).Copy
        Chk_wave(i) = InWave.Select(3328 * i - 1 + 13, 13, 256).COPY
        Set ADC_wave(i) = Nothing
        ADC_wave(i) = ADC_wave(i).ConvertDataTypeTo(DspLong)
        For j = 0 To 255
'            ADC_wave(i) = ADC_wave(i).Concatenate(InWave.Select(39 * j + 13 * i, 1, 12).Copy)
             ADC_wave(i) = ADC_wave(i).Concatenate(InWave.Select(13 * j + 3328 * i, 1, 12).COPY)
        Next j
        
        ADC_wave(i) = ADC_wave(i).ConvertStreamTo(tldspParallel, 12, 0, Bit0IsMsb)
        
        For j = 0 To 255
            TheExec.flow.TestLimit ADC_wave(i).data(j), _
            Tname:=Report_TName_From_Instance("C", "X", "SDM" & (5 + i), CInt(k)), ForceResults:=tlForceFlow 'transfer_to_forceflow
            TheExec.flow.TestLimit Chk_wave(i).data(j), _
            Tname:=Report_TName_From_Instance("C", "X", "SDM" & (5 + i), CInt(k), , "replace;7=Chk"), ForceResults:=tlForceFlow 'transfer_to_forceflow
            k = k + 1
        Next j
        
        mean = ADC_wave(i).CalcMeanWithStdDev(STDEV)
        STDEV = STDEV * STDEV
        MaxErr = ADC_wave(i).CalcMaximumValue - ADC_wave(i).CalcMinimumValue
                    
        TheExec.flow.TestLimit resultVal:=mean, Tname:=Report_TName_From_Instance("C", "X", "SDM" & (5 + i), CInt(k), , "replace;7=Mean"), ForceResults:=tlForceFlow 'transfer_to_forceflow
        TheExec.flow.TestLimit resultVal:=STDEV, Tname:=Report_TName_From_Instance("C", "X", "SDM" & (5 + i), CInt(k + 1), , "replace;7=Variance"), ForceResults:=tlForceFlow 'transfer_to_forceflow
        TheExec.flow.TestLimit resultVal:=MaxErr, Tname:=Report_TName_From_Instance("C", "X", "SDM" & (5 + i), CInt(k + 2), , "replace;7=MaxErr"), ForceResults:=tlForceFlow 'transfer_to_forceflow
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
    
    
    
    FirstLoop = CStr(TheExec.flow.var("SrcCodeIndx").value)
    SecondLoop = CStr(TheExec.flow.var("SrcCodeIndx1").value)
''''''''''    Debug.Print "SrcCodeIndx Value : " & FirstLoop
''''''''''    Debug.Print "SrcCodeIndx1 Value : " & SecondLoop
    
    
''''''''''    LoopNum = SecondLoop = SrcCodeIndx1
''''''''''    LoopNum1 = FirstLoop = SrcCodeIndx
    
    For i = 0 To argc - 1
        SplitCalStr = Split(argv(i), "@")
        DictValue_DSPWave = GetStoreDataAllType(SplitCalStr(0))
        For Each site In TheExec.sites.Active
            DictValue_DSPWave = DictValue_DSPWave.ConvertDataTypeTo(DspLong)
            DictValue_DSPWave = DictValue_DSPWave.ConvertStreamTo(tldspParallel, DictValue_DSPWave.SampleSize, 0, Bit0IsMsb)
        Next site
        Call StoreDataAllType(SplitCalStr(0) & "_" & SecondLoop, DictValue_DSPWave)
               
        If SecondLoop = 0 Then
            SaveDeltaValue_DSPWave.CreateConstant 0, CLng(SplitCalStr(1)), DspLong
            SaveMeasNum_DSPWave.CreateConstant 0, CLng(SplitCalStr(1)) + 1, DspLong
            
            If FirstLoop = 0 Then
                SaveDCK_DSPWave.CreateConstant 0, CLng(SplitCalStr(1)) + 1, DspLong
            Else
                SaveDCK_DSPWave = GetStoreDataAllType("SaveDCK_DSPWaveData")
            End If
            For Each site In TheExec.sites.Active
                SaveDCK_DSPWave.Element(FirstLoop) = DictValue_DSPWave.Element(0)
                SaveMeasNum_DSPWave.Element(SecondLoop) = DictValue_DSPWave.Element(0)
            Next site
            Call StoreDataAllType("SaveDCK_DSPWaveData", SaveDCK_DSPWave)
            Call StoreDataAllType("SaveMeasNum_DSPWaveData", SaveMeasNum_DSPWave)
            Call StoreDataAllType("SaveDeltaValue_DSPWaveData", SaveDeltaValue_DSPWave)
            
        ElseIf SecondLoop <> 64 Then
            SaveMeasNum_DSPWave = GetStoreDataAllType("SaveMeasNum_DSPWaveData")
            SaveDeltaValue_DSPWave = GetStoreDataAllType("SaveDeltaValue_DSPWaveData")
            PreValue_DSPWave = GetStoreDataAllType(SplitCalStr(0) & "_" & SecondLoop - 1)
            
            For Each site In TheExec.sites.Active
                SaveDeltaValue_DSPWave.Element(SecondLoop - 1) = Abs(DictValue_DSPWave.Element(0) - PreValue_DSPWave.Element(0))
                SaveMeasNum_DSPWave.Element(SecondLoop) = DictValue_DSPWave.Element(0)
            Next site
            Call StoreDataAllType("SaveMeasNum_DSPWaveData", SaveMeasNum_DSPWave)
            Call StoreDataAllType("SaveDeltaValue_DSPWaveData", SaveDeltaValue_DSPWave)
            
        ElseIf SecondLoop = 64 Then
            SaveMeasNum_DSPWave = GetStoreDataAllType("SaveMeasNum_DSPWaveData")
            SaveDeltaValue_DSPWave = GetStoreDataAllType("SaveDeltaValue_DSPWaveData")
            
            For Each site In TheExec.sites.Active
                ValueMax = SaveMeasNum_DSPWave.CalcMaximumValue(IndexOfMaximumValue)
                ValueMin = SaveMeasNum_DSPWave.CalcMinimumValue(IndexOfMinimumValue)
                Idsvalue = (ValueMax - ValueMin) / (CLng(SplitCalStr(1)) + 1)
''''''''''                TheExec.Datalog.WriteComment "MAX_Value:  " & SaveMeasNum_DSPWave.CalcMaximumValue(IndexOfMaximumValue)
''''''''''                TheExec.Datalog.WriteComment "MIN_Value:  " & SaveMeasNum_DSPWave.CalcMinimumValue(IndexOfMinimumValue)
''''''''''                TheExec.Datalog.WriteComment "FinalIds value : " & Idsvalue
            Next site
            TestNameInput = Report_TName_From_Instance("C", "X", "CK" & Format(FirstLoop, "00") & "-START", CInt(i))
            TheExec.flow.TestLimit resultVal:=ValueMax, Tname:=TestNameInput, ForceResults:=tlForceFlow
            TestNameInput = Report_TName_From_Instance("C", "X", "CK" & Format(FirstLoop, "00") & "-END", CInt(i))
            TheExec.flow.TestLimit resultVal:=ValueMin, Tname:=TestNameInput, ForceResults:=tlForceFlow
            TestNameInput = Report_TName_From_Instance("C", "X", "DCK" & Format(FirstLoop, "00") & "-IDEAL", CInt(i))
            TheExec.flow.TestLimit resultVal:=Idsvalue, Tname:=TestNameInput, ForceResults:=tlForceFlow

            DNLValue_DSPWave.CreateConstant 0, CLng(SplitCalStr(1)), DspDouble
            For j = 0 To CLng(SplitCalStr(1)) - 1
                For Each site In TheExec.sites.Active
                    DNLValue_DSPWave.Element(j) = (SaveDeltaValue_DSPWave.Element(j) - Idsvalue) / Idsvalue
''''''''''                    TheExec.Datalog.WriteComment "CK Delta value" & j & " = " & CStr(SaveDeltaValue_DSPWave.Element(j))
''''''''''                    TheExec.Datalog.WriteComment "DNL value on point" & j & " = " & CStr(DNLValue_DSPWave.Element(j))
                Next site
                TestNameInput = Report_TName_From_Instance("C", "X", "CKDeltavalue" & Format(j, "00"), CInt(i))
                TheExec.flow.TestLimit resultVal:=SaveDeltaValue_DSPWave.Element(j), Tname:=TestNameInput, ForceResults:=tlForceFlow
                TestNameInput = Report_TName_From_Instance("C", "X", "CKDNLvalue" & Format(j, "00"), CInt(i))
                TheExec.flow.TestLimit resultVal:=DNLValue_DSPWave.Element(j), Tname:=TestNameInput, ForceResults:=tlForceFlow
            Next j
            
            For Each site In TheExec.sites.Active
                MaxCKDNL = DNLValue_DSPWave.CalcMaximumValue(IndexOfMaximumValue)
                MinCKDNL = DNLValue_DSPWave.CalcMinimumValue(IndexOfMinimumValue)
''''''''''                TheExec.Datalog.WriteComment "MAX DNL:" & MaxCKDNL
''''''''''                TheExec.Datalog.WriteComment "MIN DNL:" & MinCKDNL
            Next site
            TestNameInput = Report_TName_From_Instance("C", "X", "CK" & Format(FirstLoop, "00") & "-MAX-DNL", CInt(i))
            TheExec.flow.TestLimit resultVal:=MaxCKDNL, Tname:=TestNameInput, ForceResults:=tlForceFlow
            TestNameInput = Report_TName_From_Instance("C", "X", "CK" & Format(FirstLoop, "00") & "-MIN-DNL", CInt(i))
            TheExec.flow.TestLimit resultVal:=MinCKDNL, Tname:=TestNameInput, ForceResults:=tlForceFlow
            TestNameInput = Report_TName_From_Instance("C", "X", "CK" & Format(FirstLoop, "00") & "-MAXStepDelta", CInt(i))
            For Each site In TheExec.sites.Active
                TheExec.flow.TestLimit resultVal:=CStr(SaveDeltaValue_DSPWave.CalcMaximumValue(IndexOfMaximumValue)), Tname:=TestNameInput, ForceResults:=tlForceFlow
                TheExec.flow.TestLimitIndex = TheExec.flow.TestLimitIndex - 1
            Next site
            TheExec.flow.TestLimitIndex = TheExec.flow.TestLimitIndex + 1
            
            TestNameInput = Report_TName_From_Instance("C", "X", "CK" & Format(FirstLoop, "00") & "-MINStepDelta", CInt(i))
            For Each site In TheExec.sites.Active
                TheExec.flow.TestLimit resultVal:=CStr(SaveDeltaValue_DSPWave.CalcMinimumValue(IndexOfMinimumValue)), Tname:=TestNameInput, ForceResults:=tlForceFlow
                TheExec.flow.TestLimitIndex = TheExec.flow.TestLimitIndex - 1
            Next site
            TheExec.flow.TestLimitIndex = TheExec.flow.TestLimitIndex + 1
        End If
            
        If SecondLoop = 64 And FirstLoop = 64 Then
            SaveDCK_DSPWave = GetStoreDataAllType("SaveDCK_DSPWaveData")
            
            For Each site In TheExec.sites.Active
                DCKMAX = SaveDCK_DSPWave.CalcMaximumValue(IndexOfMaximumValue)
                DCKMin = SaveDCK_DSPWave.CalcMinimumValue(IndexOfMinimumValue)
                DCKIdsvalue = (DCKMAX - DCKMin) / (CLng(SplitCalStr(1)) + 1)
            Next site
            TheExec.Datalog.WriteComment "------------------------------------------------- SDLL_DCK summary--------------------------------------------------------------"
            TestNameInput = Report_TName_From_Instance("C", "X", "DCK" & Format(FirstLoop, "00") & "-START", CInt(i))
            TheExec.flow.TestLimit resultVal:=DCKMAX, Tname:=TestNameInput, ForceResults:=tlForceFlow
            TestNameInput = Report_TName_From_Instance("C", "X", "DCK" & Format(FirstLoop, "00") & "-END", CInt(i))
            TheExec.flow.TestLimit resultVal:=DCKMin, Tname:=TestNameInput, ForceResults:=tlForceFlow
            TestNameInput = Report_TName_From_Instance("C", "X", "DCK" & Format(FirstLoop, "00") & "-IDEAL", CInt(i))
            TheExec.flow.TestLimit resultVal:=DCKIdsvalue, Tname:=TestNameInput, ForceResults:=tlForceFlow

            DeltaDCK_DSPWave.CreateConstant 0, CLng(SplitCalStr(1)), DspDouble
            DNLDCKValue_DSPWave.CreateConstant 0, CLng(SplitCalStr(1)), DspDouble

            For j = 1 To CLng(SplitCalStr(1))
                For Each site In TheExec.sites.Active
                    DeltaDCK_DSPWave.Element(j - 1) = SaveDCK_DSPWave.Element(j - 1) - SaveDCK_DSPWave.Element(j)
                    DNLDCKValue_DSPWave.Element(j - 1) = (DeltaDCK_DSPWave.Element(j - 1) - DCKIdsvalue) / DCKIdsvalue
                Next site
                TestNameInput = Report_TName_From_Instance("C", "X", "DCKDeltavalue" & Format(j - 1, "00"), CInt(i))
                TheExec.flow.TestLimit resultVal:=SaveDCK_DSPWave.Element(j - 1), Tname:=TestNameInput, ForceResults:=tlForceFlow
                TestNameInput = Report_TName_From_Instance("C", "X", "DCKDNLvalue" & Format(j - 1, "00"), CInt(i))
                TheExec.flow.TestLimit resultVal:=DNLDCKValue_DSPWave.Element(j - 1), Tname:=TestNameInput, ForceResults:=tlForceFlow
            Next j
            
            TestNameInput = Report_TName_From_Instance("C", "X", "DCK" & Format(FirstLoop, "00") & "-MAX-DNL", CInt(i))
            For Each site In TheExec.sites.Active
                TheExec.flow.TestLimit resultVal:=DNLDCKValue_DSPWave.CalcMaximumValue(IndexOfMaximumValue), Tname:=TestNameInput, ForceResults:=tlForceFlow
                TheExec.flow.TestLimitIndex = TheExec.flow.TestLimitIndex - 1
            Next site
            TheExec.flow.TestLimitIndex = TheExec.flow.TestLimitIndex + 1
            
            TestNameInput = Report_TName_From_Instance("C", "X", "DCK" & Format(FirstLoop, "00") & "-MIN-DNL", CInt(i))
            For Each site In TheExec.sites.Active
                TheExec.flow.TestLimit resultVal:=DNLDCKValue_DSPWave.CalcMinimumValue(IndexOfMinimumValue), Tname:=TestNameInput, ForceResults:=tlForceFlow
                TheExec.flow.TestLimitIndex = TheExec.flow.TestLimitIndex - 1
            Next site
            TheExec.flow.TestLimitIndex = TheExec.flow.TestLimitIndex + 1
                
            TestNameInput = Report_TName_From_Instance("C", "X", "DCK" & Format(FirstLoop, "00") & "-MAXStepDelta", CInt(i))
            For Each site In TheExec.sites.Active
                TheExec.flow.TestLimit resultVal:=DeltaDCK_DSPWave.CalcMaximumValue(IndexOfMaximumValue), Tname:=TestNameInput, ForceResults:=tlForceFlow
                TheExec.flow.TestLimitIndex = TheExec.flow.TestLimitIndex - 1
            Next site
            TheExec.flow.TestLimitIndex = TheExec.flow.TestLimitIndex + 1
            
            TestNameInput = Report_TName_From_Instance("C", "X", "DCK" & Format(FirstLoop, "00") & "-MINStepDelta", CInt(i))
            For Each site In TheExec.sites.Active
                TheExec.flow.TestLimit resultVal:=DeltaDCK_DSPWave.CalcMinimumValue(IndexOfMinimumValue), Tname:=TestNameInput, ForceResults:=tlForceFlow
                TheExec.flow.TestLimitIndex = TheExec.flow.TestLimitIndex - 1
            Next site
            TheExec.flow.TestLimitIndex = TheExec.flow.TestLimitIndex + 1
            
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
        Mdll_lockvalue = GetStoreDataAllType(argv(4))
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
            DSPWaveTemp = GetStoreDataAllType(strTemp(j))
            If CLng(SweepConterStr) <> CLng(CounterByStart) Then
                DataSite(j) = GetStoreDataAllType(strTemp(j) & "_" & "AssemblyStr")
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
            Call StoreDataAllType(strTemp(j) & "_" & "AssemblyStr", DataSite(j))
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
                        
                   
                       
                        TheExec.flow.TestLimit lowVal:=mdll_low(0), hiVal:=mdll_high(0), resultVal:=PrintEye(z + k * CounterByWidth) * EyeDivide, formatStr:="%i", Tname:=TestNameInput, ForceResults:=tlForceFlow, scaletype:=scaleNoScaling 'transfer_to_forceflow
                        
                        
                        TestNameInputeye = Report_TName_From_Instance("calc", "x", TestNameInputeye, x, 0)
                        
                        Eye_precent = (PrintEye(z + k * CounterByWidth) * EyeDivide) / Mdll_lock(site).Element(0)
                        
                        TheExec.flow.TestLimit lowVal:=25, hiVal:=50, resultVal:=Format(Eye_precent * 100, 0), formatStr:="%i", Tname:=TestNameInputeye, ForceResults:=tlForceFlow, scaletype:=scaleNoScaling 'transfer_to_forceflow
                        
                        
                        
                        
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
        Mdll_lockvalue = GetStoreDataAllType(argv(4))
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
            DSPWaveTemp = GetStoreDataAllType(strTemp(j))
            If CLng(SweepConterStr) <> CLng(CounterByStart) Then
                DataSite(j) = GetStoreDataAllType(strTemp(j) & "_" & "AssemblyStr")
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
            Call StoreDataAllType(strTemp(j) & "_" & "AssemblyStr", DataSite(j))
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
                        
                   
                       
                        'TheExec.Flow.TestLimit LowVal:=mdll_low(0), HiVal:=mdll_high(0), resultVal:=PrintEye(z + k * CounterByWidth) * EyeDivide, FormatStr:="%i", TName:=TestNameInput, ForceResults:=tlForceNone, ScaleType:=scaleNoScaling
                        TheExec.flow.TestLimit lowVal:=mdll_low(0), hiVal:=mdll_high(0), resultVal:=PrintEye(z + k * CounterByWidth) * EyeDivide, formatStr:="%i", Tname:=TestNameInput, ForceResults:=tlForceFlow, scaletype:=scaleNoScaling 'transfer_to_forceflow
                        
                        TestNameInputeye = Report_TName_From_Instance("calc", "x", TestNameInputeye, x, 0)
                        
                        Eye_precent = (PrintEye(z + k * CounterByWidth) * EyeDivide) / Mdll_lock(site).Element(0)
                        
                        TheExec.flow.TestLimit lowVal:=25, hiVal:=50, resultVal:=Format(Eye_precent * 100, 0), formatStr:="%i", Tname:=TestNameInputeye, ForceResults:=tlForceFlow, scaletype:=scaleNoScaling 'transfer_to_forceflow
                        
                        
                        
                        
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

Public Function Find_transition_point(argc As Integer, argv() As String) As Long
On Error GoTo errHandler 'Add ErrHandler 2023/05/29
    ''Find_transition_point(x_sweep,start,end,step,Pin)

    Dim TestSize As Long
    Dim TestPoint As Long
    Dim site As Variant
    Dim i As Long
    Dim index As New DSPWave
    Dim CurrentPoint As Long
    CurrentPoint = val(TheExec.flow.var(argv(0)).value)
    Dim result As New SiteDouble
    Dim pin As String
    pin = argv(4)
    
    Dim DicTransitionKey As String
    Dim DicTransitionKey_Cond As String
    Dim DSP_Transition As New DSPWave
    Dim DSP_Transition_Cond As New DSPWave
    Dim TestNameInput As String
    Dim N_idx As Long
        
    DicTransitionKey = "Dic_" & TheExec.DataManager.instancename
    DicTransitionKey_Cond = DicTransitionKey & "_Cond"
    
    If Not gDictDSPWaves.Exists(LCase(DicTransitionKey)) Then 'lcase for all dict name
        TestSize = Abs((argv(2) - argv(1)) / argv(3)) + 1
        DSP_Transition.CreateConstant 2, TestSize, DspLong
        DSP_Transition_Cond.CreateConstant 0, TestSize, DspDouble
        
        For i = 0 To TestSize - 1
            DSP_Transition_Cond.Element(i) = argv(1) + i * argv(3)
        Next
        For Each site In TheExec.sites
            DSP_Transition.Element(0) = TheHdw.Digital.Patgen.PatternBurstPassed
        Next site
        Call StoreDataAllType(DicTransitionKey, DSP_Transition)
        Call StoreDataAllType(DicTransitionKey_Cond, DSP_Transition_Cond)
    End If
    
    If CurrentPoint = argv(1) Then                                  '1st point
        'do nothing
    Else                                                            'other point
        Set DSP_Transition = GetStoreDataAllType(DicTransitionKey)
        Set DSP_Transition_Cond = GetStoreDataAllType(DicTransitionKey_Cond)
        index = DSP_Transition_Cond.FindIndices(EqualTo, CurrentPoint)
        For Each site In TheExec.sites
'            If TheHdw.Digital.Patgen.PatternBurstPassed = False Then
'                Stop
'            End If
            
            DSP_Transition.Element(index.Element(0)) = TheHdw.Digital.Patgen.PatternBurstPassed
        Next
        
        If CurrentPoint = argv(2) Then                              'last point
            'DSP_Transition_Cond = DSP_Transition_Cond.Divide(1000).Subtract(0.2)
             DSP_Transition_Cond = DSP_Transition_Cond.divide(1000)
            For Each site In TheExec.sites
                index = DSP_Transition.FindIndices(EqualTo, -1)
                N_idx = index.SampleSize
                If index.SampleSize = 0 Then
                    result = -9999
                ElseIf index.SampleSize = DSP_Transition.SampleSize Then
                    result = DSP_Transition_Cond.Element(index.Element(0))
                Else
                    If argv(1) < argv(2) Then
                        'result = DSP_Transition_Cond.Element(index.Element(N_idx - 1))
                         result = DSP_Transition_Cond.Element(index.Element(0))
                    Else
                        'result = DSP_Transition_Cond.Element(index.Element(0))
                        result = DSP_Transition_Cond.Element(index.Element(N_idx - 1))
                    End If
                End If
            Next
            TestNameInput = Report_TName_From_Instance("V", pin, , 0, 0)
            TheExec.flow.TestLimit result, , , , , , unitVolt, , Tname:=TestNameInput, ForceResults:=tlForceFlow
        End If
        
        
        '''backup
'''''''''        If CurrentPoint = argv(2) Then                              'last point
'''''''''            DSP_Transition_Cond = DSP_Transition_Cond.Divide(1000).Subtract(0.2)
'''''''''            For Each site In theexec.sites
'''''''''                index = DSP_Transition.FindIndices(EqualTo, -1)
'''''''''                N_idx = index.SampleSize
'''''''''                If index.SampleSize = 0 Then
'''''''''                    result = -9999
'''''''''                Else
'''''''''                    If argv(1) < argv(2) Then
'''''''''                        result = DSP_Transition_Cond.Element(index.Element(0))
'''''''''                    Else
'''''''''                        result = DSP_Transition_Cond.Element(index.Element(N_idx - 1))
'''''''''                    End If
'''''''''                End If
'''''''''            Next
'''''''''            TestNameInput = Report_TName_From_Instance("V", Pin, , 0, 0)
'''''''''            theexec.Flow.TestLimit result, , , , , , unitVolt, , Tname:=TestNameInput, ForceResults:=tlForceFlow
'''''''''        End If
        
        
        
    End If
Exit Function 'Add ErrHandler 2023/05/29
errHandler: 'Add ErrHandler 2023/05/29
    Call Print_Error_Message(Error_Info, "LIB_HardIP_Calc", "Find_transition_point") 'Add ErrHandler 2023/05/29
    If AbortTest Then Exit Function Else Resume Next 'Add ErrHandler 2023/05/29

End Function

Public Function LB_Error_Count(argc As Integer, argv() As String) As Long
    
    Dim TestNameInput As String
    Dim site As Variant
    Dim LPK_result As New SiteDouble
    Dim ERR_CNT As New DSPWave
    Dim ParallelStream As New DSPWave
    Dim ErrorStr As String
    Dim i As Long
    
    ERR_CNT = GetStoreDataAllType(argv(0))
    For Each site In TheExec.sites
        LPK_result = ERR_CNT.Element(ERR_CNT.SampleSize - 1)
        ParallelStream = ERR_CNT.ConvertStreamTo(tldspParallel, 4, 0, Bit0IsMsb)
    Next site
    
    TestNameInput = Report_TName_From_Instance("Calc", "ERR_CNT", , 0, 0, , , , tlForceFlow)
    TheExec.flow.TestLimit LPK_result, 1, 1, , , , , "0.0f", Tname:=TestNameInput, ForceResults:=tlForceFlow 'transfer_to_forceflow

    For Each site In TheExec.sites
        ErrorStr = vbNullString
        For i = (ParallelStream.SampleSize - 1) To 0 Step -1
            ErrorStr = ErrorStr & Hex(ParallelStream.Element(i))
        Next i
        TheExec.Datalog.WriteComment "Site " & site & ":Error Count on (" & TheExec.DataManager.instancename & "),Hex =" & ErrorStr
    Next site
    
End Function

Public Function Calc_AUS_DCD(argc As Integer, argv() As String) As Long
On Error GoTo errHandler 'Add ErrHandler 2023/05/29

    Dim TestNameInput As String
    Dim RAW_DATA As New DSPWave
    Dim Result_DATA As New DSPWave
    'Dim target As Long
    Dim i As Long
    Dim site As Variant
    Dim TempVal() As New SiteDouble
    
    'target = argv(0)
    ReDim TempVal(argc - 1)
    RAW_DATA.CreateConstant 0, argc, DspDouble
    
    For Each site In TheExec.sites
        For i = 0 To argc - 1
            TempVal(i) = GetStoreDataAllType(argv(i) & "_para")
            RAW_DATA.Element(i) = IIf(TempVal(i) >= 512, 512 - TempVal(i), TempVal(i))
        Next i
    ''' Update for Donan for % Unit issue @William 230308
                'Result_DATA = RAW_DATA.divide(1024).Multiply(100)
        Result_DATA = RAW_DATA.divide(1024)
    Next site
    
    For i = 0 To argc - 1
        TestNameInput = Report_TName_From_Instance("Calc", "DCD", , CInt(i), 0, , , , tlForceFlow)
        TheExec.flow.TestLimit Result_DATA.Element(i), , , , , , , , Tname:=TestNameInput, ForceResults:=tlForceFlow
    Next i
    
Exit Function 'Add ErrHandler 2023/05/29
errHandler: 'Add ErrHandler 2023/05/29
    Call Print_Error_Message(Error_Info, "LIB_HardIP_Calc", "Calc_AUS_DCD") 'Add ErrHandler 2023/05/29
    If AbortTest Then Exit Function Else Resume Next 'Add ErrHandler 2023/05/29
End Function
Public Function Calc_GRPPercentage(argc As Integer, argv() As String) As Long
On Error GoTo errHandler 'Add ErrHandler 2023/05/29
   Dim i As Integer
    Dim site As Variant
    Dim SplitStrAry() As String
    
    
    Dim DSP_LoLimit As New DSPWave
    
    
''''''''''    Dim DSP_Temp As New DSPWave
''''''''''    Dim DSP_Summary As New DSPWave
    Dim DSP_EyeWidth As New DSPWave
    Dim DSP_Percentage As New DSPWave
    Dim TestNameInput As String
    Dim Percentage_bysite() As New SiteDouble
    
    
    Dim DSP_Temp() As New DSPWave
    Dim DSP_Summary() As New DSPWave
    ReDim DSP_Temp(argc - 1)
    ReDim DSP_Summary(argc - 1)
    ReDim Percentage_bysite(argc - 1)
    
    For i = 0 To argc - 1
        SplitStrAry = Split(argv(i), "&")
        DSP_Summary(i) = GetStoreDataAllType("Summary" & SplitStrAry(0))
        DSP_Temp(i) = GetStoreDataAllType(SplitStrAry(1))
       
        For Each site In TheExec.sites
            DSP_EyeWidth = DSP_Temp(i).ConvertStreamTo(tldspParallel, DSP_Temp(i).SampleSize, 0, Bit0IsMsb)
            DSP_Percentage = DSP_EyeWidth.Multiply(100).divide(DSP_Summary(i).Multiply(2))
            Percentage_bysite(i) = DSP_Percentage.Element(0)
        Next site
        
        
        TestNameInput = Report_TName_From_Instance("C", SplitStrAry(1), "Percentage" & i, i)
        For Each site In TheExec.sites
            TheExec.flow.TestLimit DSP_EyeWidth.Element(0), DSP_Summary(i).Element(0) / 2, DSP_Summary(i).Element(0) * 2, , , , , , Tname:=TestNameInput, ForceResults:=tlForceFlow
            TheExec.flow.TestLimitIndex = TheExec.flow.TestLimitIndex - 1
        Next site
        TheExec.flow.TestLimitIndex = TheExec.flow.TestLimitIndex + 1
        
        
    Next i
    
    For i = 0 To argc - 1
        SplitStrAry = Split(argv(i), "&")
        TestNameInput = Report_TName_From_Instance("C", SplitStrAry(1), "Percentage" & i, i, , , , , tlForceNone)
        'For Each site In TheExec.sites
            TheExec.flow.TestLimit Percentage_bysite(i), 25, 100, , , , , , Tname:=TestNameInput, ForceResults:=tlForceNone 'Un-Used_
       ' Next site
    Next i
    
    
    
    
    
''''''''''    For i = 0 To argc - 1
''''''''''        SplitStrAry = Split(argv(i), "&")
''''''''''        DSP_Summary = GetStoreDataAllType("Summary" & SplitStrAry(0))
''''''''''        DSP_Temp = GetStoreDataAllType(SplitStrAry(1))
''''''''''
''''''''''
''''''''''
''''''''''        For Each site In TheExec.sites
''''''''''            DSP_EyeWidth = DSP_Temp.ConvertStreamTo(tldspParallel, DSP_Temp.SampleSize, 0, Bit0IsMsb)
''''''''''            DSP_Percentage = DSP_EyeWidth.Multiply(100).Divide(DSP_Summary.Multiply(2))
''''''''''            Percentage_bysite = DSP_Percentage.Element(0)
''''''''''        Next site
''''''''''
''''''''''
''''''''''        'TestNameInput = Report_TName_From_Instance("C", SplitStrAry(1), "EyeWIDTH" & i, i)
''''''''''        'TheExec.Flow.TestLimit DSP_EyeWidth.Element(0), DSP_Summary.Element(0) / 2, DSP_Summary.Element(0) * 2, , , , , , Tname:=TestNameInput, ForceResults:=tlForceFlow
''''''''''        TestNameInput = Report_TName_From_Instance("C", SplitStrAry(1), "Percentage" & i, i)
''''''''''        For Each site In TheExec.sites
''''''''''        'TestNameInput = Report_TName_From_Instance("C", SplitStrAry(1), "Percentage" & i, i)
''''''''''        TheExec.Flow.TestLimit Percentage_bysite, DSP_Summary.Element(0) / 2, DSP_Summary.Element(0), , , , , , Tname:=TestNameInput, ForceResults:=tlForceFlow
''''''''''        TheExec.Flow.TestLimitIndex = TheExec.Flow.TestLimitIndex - 1
''''''''''        Next site
'''''''''''        TheExec.Flow.TestLimit DSP_EyeWidth.Element(0), 2, 4, , , , , , Tname:=TestNameInput, ForceResults:=tlForceNone
''''''''''        TheExec.Flow.TestLimitIndex = TheExec.Flow.TestLimitIndex + 1
''''''''''
''''''''''
''''''''''    Next i
Exit Function 'Add ErrHandler 2023/05/29
errHandler: 'Add ErrHandler 2023/05/29
    Call Print_Error_Message(Error_Info, "LIB_HardIP_Calc", "Calc_GRPPercentage") 'Add ErrHandler 2023/05/29
    If AbortTest Then Exit Function Else Resume Next 'Add ErrHandler 2023/05/29

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
            eqc_ctrl = GetStoreDataAllType(argv(0))
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
        StoreDataAllType argv(1), EQCWave
    
        
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
    'TheExec.Flow.TestLimit resultVal:=EQC_temp, Tname:=TestNameInput, ForceResults:=tlForceNone
   
   
   
   'i_min = GetStoreDataAllType("aus16pll_fcal_bypass_code")
   
   
    TestNameInput = Report_TName_From_Instance(CalcC, "X", argv(1) & "_EQCalWave", CInt(i))
    
    TheExec.flow.TestLimit resultVal:=EQC_temp, Tname:=TestNameInput, ForceResults:=tlForceFlow

End Function
Public Function Calc_EyeWidthLimitForEachMode(argc As Integer, argv() As String) As Long
On Error GoTo errHandler 'Add ErrHandler 2023/05/29
    Dim i As Integer
    Dim site As Variant
    Dim TestNameInput As String
    Dim SplitStrAry() As String
    Dim PercentageTemp() As New SiteDouble
    Dim SiteDouble_Dcode As New SiteDouble
    Dim SiteDouble_EyeWidth As New SiteDouble
    
    
    ReDim PercentageTemp(argc - 1)
    
    
    For i = 0 To argc - 1
        SplitStrAry = Split(argv(i), "&")
        SiteDouble_Dcode = GetStoreDataAllType(SplitStrAry(0) & "_para")
        SiteDouble_EyeWidth = GetStoreDataAllType(SplitStrAry(1) & "_para")
        
        '------------ For avoid Capture 0 error --------------
        For Each site In TheExec.sites
            If SiteDouble_Dcode = 0 Then
               SiteDouble_Dcode = 0.01
               TheExec.Datalog.WriteComment SplitStrAry(0) & " Capture 0 so give defult value"
            End If
            If SiteDouble_EyeWidth = 0 Then
               SiteDouble_EyeWidth = 0.01
               TheExec.Datalog.WriteComment SplitStrAry(1) & " Capture 0 so give defult value"
            End If
        Next site
        '-----------------------------------------------
        
        TestNameInput = Report_TName_From_Instance("C", vbNullString, ForceResult:=tlForceFlow)
        For Each site In TheExec.sites
            TheExec.flow.TestLimit SiteDouble_EyeWidth * 100, SiteDouble_Dcode / 8, SiteDouble_Dcode, , , , , , Tname:=TestNameInput, ForceResults:=tlForceFlow
            TheExec.flow.TestLimitIndex = TheExec.flow.TestLimitIndex - 1
        Next site
        TheExec.flow.TestLimitIndex = TheExec.flow.TestLimitIndex + 1
        
        PercentageTemp(i) = SiteDouble_EyeWidth.divide(SiteDouble_Dcode).Multiply(100)
    Next i
    
    
    For i = 0 To argc - 1
        SplitStrAry = Split(argv(i), "&")
        TestNameInput = Report_TName_From_Instance("C", SplitStrAry(1), "Percentage" & i, i, , , , , tlForceNone_CZ)
        TheExec.flow.TestLimit PercentageTemp(i), 25, 100, , , , , , Tname:=TestNameInput, ForceResults:=tlForceNone_CZ
    Next i
    
    
Exit Function 'Add ErrHandler 2023/05/29
errHandler: 'Add ErrHandler 2023/05/29
    Call Print_Error_Message(Error_Info, "LIB_HardIP_Calc", "Calc_EyeWidthLimitForEachMode") 'Add ErrHandler 2023/05/29
    If AbortTest Then Exit Function Else Resume Next 'Add ErrHandler 2023/05/29


End Function

Public Function Calc_DCSkew_IN_CLK(argc As Integer, argv() As String) As Long
    'Calc_DCSkew_IN_CLK(CNT_100_0,CNT_INPUT)
    Dim CNT_100_0 As New DSPWave: CNT_100_0 = GetStoreDataAllType(argv(0))
    Dim CNT_INPUT As New DSPWave: CNT_INPUT = GetStoreDataAllType(argv(1))
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
    TheExec.flow.TestLimit resultVal:=DC_InClk, Tname:=TestNameInput, ForceResults:=tlForceFlow

End Function

Public Function Functional_Parametric_LoopFunction(argc As Integer, argv() As String) As Long

    Dim site As Variant
    Dim IndexName As String
    Dim PassFailEachSite As New SiteLong
    Dim PassFailBoolean As New SiteBoolean
    
    IndexName = CStr(argv(0))
    PassFailEachSite = TheExec.flow.LastFlowStepResult
    
    For Each site In TheExec.sites.Active
        If PassFailEachSite = 1 Then                     ' 1 = tlResultPass
            PassFailBoolean = True
        ElseIf PassFailEachSite = 0 Then                 ' 0 = tlResultFail
            PassFailBoolean = False
        End If
    Next site
    
    If PassFailBoolean.All(True) Then
        TheExec.flow.var(IndexName).value = 100
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
        DAC_5_4(i) = GetStoreDataAllType(DicName_5_4)
        DAC_3_0(i) = GetStoreDataAllType(DicName_3_0)
        For Each site In TheExec.sites
            DAC_temp = DAC_3_0(i).Concatenate(DAC_5_4(i))
            DAC_result(i) = DAC_temp.ConvertStreamTo(tldspParallel, 6, 0, Bit0IsMsb)
        Next site
        
        TestNameInput = Report_TName_From_Instance(CalcC, vbNullString, , , i)
        TheExec.flow.TestLimit resultVal:=DAC_result(i).Element(0), Tname:=TestNameInput, ForceResults:=tlForceFlow
    Next i
End Function

Public Function Calc_DCC_Skew_Range_DSP_Long(argc As Integer, argv() As String) As Long
On Error GoTo errHandler 'Add ErrHandler 2023/05/29
 
    ''''Demo String : GRP@GRP0@GRP1@GRP2@GRP3@GRP4@GRP5@GRP6@GRP7,CH@CH0@CH1,DQ@DQ0@DQ1,CountIN@0x1F@0x0@1Fx0,Count1000@0x1F@0x0@1Fx0,SkewFactor@0.5,InputFactor@1.5, PatternBit@13
    Dim i, j, k, y, g As Long
    Dim SplitGRP() As String
    Dim SplitCH() As String
    Dim SplitDQ() As String
    Dim SplitCountIN() As String
    Dim SplitCount1000() As String
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
    
    
    SplitGRP = Split(argv(0), "@")  'GRP@GRP0@GRP1@GRP2@GRP3@GRP4@GRP5@GRP6@GRP7
    SplitCH = Split(argv(1), "@")   'CH@CH0@CH1
    SplitDQ = Split(argv(2), "@")   'DQ@DQ0@DQ1
    SplitCountIN = Split(argv(3), "@")   'CountIN@0x1F@0x0@1Fx0
    SplitCount1000 = Split(argv(4), "@")   'Count1000@0x1F@0x0@1Fx0
    SkewFactor = Split(argv(5), "@")(1) 'SkewFactor@0.5
    InputFactor = Split(argv(6), "@")(1) 'InputFactor@1.5
    DSPWaveTemp.CreateConstant 0, Split(argv(7), "@")(1), DspLong 'PatternBit@13
    DSPWaveTemp1.CreateConstant 0, Split(argv(7), "@")(1), DspLong 'PatternBit@13
    DSPWaveDec.CreateConstant 0, 1
    DSPWaveDec1.CreateConstant 0, 1
    ReDim DC_Skew_Input_Array(UBound(SplitCountIN) - 1)
    
    For g = 0 To (UBound(SplitGRP) - 1)
        For i = 0 To (UBound(SplitCH) - 1)
            For j = 0 To (UBound(SplitDQ) - 1)
                For Each site In TheExec.sites.Active
                    For k = 0 To (UBound(SplitCountIN) - 1)
    ''''''''''                    DSPWaveTemp = GetStoreDataAllType(SplitCH(i + 1) & SplitDQ(j + 1) & "x" & SplitCountIN(k + 1) & SplitCountIN(0))
    ''''''''''                    DSPWaveTemp1 = GetStoreDataAllType(SplitCH(i + 1) & SplitDQ(j + 1) & "x" & SplitCount1000(k + 1) & SplitCount1000(0))
'                        DSPWaveDec = GetStoreDataAllType("2SDEC_" & SplitGRP(g + 1) & SplitCH(i + 1) & SplitDQ(j + 1) & SplitCountIN(k + 1) & "dcciocountout" & SplitCountIN(0)) 'DCCIOCOUNTOUT
'                        DSPWaveDec1 = GetStoreDataAllType("2SDEC_" & SplitGRP(g + 1) & SplitCH(i + 1) & SplitDQ(j + 1) & SplitCount1000(k + 1) & "dcciocountout" & SplitCount1000(0))
                         DSPWaveDec = GetStoreDataAllType("2SDEC_" & SplitGRP(g + 1) & SplitCH(i + 1) & SplitDQ(j + 1) & SplitCountIN(k + 1) & SplitCountIN(0)) 'DCCIOCOUNTOUT
                         DSPWaveDec1 = GetStoreDataAllType("2SDEC_" & SplitGRP(g + 1) & SplitCH(i + 1) & SplitDQ(j + 1) & SplitCount1000(k + 1) & SplitCount1000(0))


                        
    ''''''''''                    DSPWaveDec = DSPWaveTemp.ConvertStreamTo(tldspParallel, Split(argv(6), "@")(1), 0, Bit0IsMsb)
    ''''''''''                    DSPWaveDec1 = DSPWaveTemp1.ConvertStreamTo(tldspParallel, Split(argv(6), "@")(1), 0, Bit0IsMsb)
                        If DSPWaveDec1.Element(0) = 0 Then
                            TheExec.Datalog.WriteComment ("Can't divide by 0")
                        Else
                            DC_Skew_Input_Array(k) = (DSPWaveDec.Element(0) / DSPWaveDec1.Element(0)) * SkewFactor
                        End If
                    Next k
                    DC_Input_CLK_UP = DC_Skew_Input_Array(0) + InputFactor
                    DC_Input_CLK_NO_DCC = DC_Skew_Input_Array(1) + InputFactor
                    DC_Input_CLK_DOWN = DC_Skew_Input_Array(2) + InputFactor
                    
                    DC_Skew_Input_CLK_UP = DC_Skew_Input_Array(0)
                    DC_Skew_Input_CLK_NO_DCC = DC_Skew_Input_Array(1)
                    DC_Skew_Input_CLK_DOWN = DC_Skew_Input_Array(2)
                    DCC_RANGE_UP = DC_Skew_Input_CLK_UP - DC_Skew_Input_CLK_NO_DCC
                    DCC_RANGE_DOWN = DC_Skew_Input_CLK_DOWN - DC_Skew_Input_CLK_NO_DCC
    ''''''''''                TheExec.Datalog.WriteComment ("Site " & Site & " : " & SplitCH(i + 1) & "_" & SplitDQ(j + 1) & "_" & "DCC_RANGE_UP" & " = " & DCC_RANGE_UP)
    ''''''''''                TheExec.Datalog.WriteComment ("Site " & Site & " : " & SplitCH(i + 1) & "_" & SplitDQ(j + 1) & "_" & "DCC_RANGE_DOWN" & " = " & DCC_RANGE_DOWN)
                Next site
                
                TestNameInput = Report_TName_From_Instance("Calc", SplitGRP(g + 1) & SplitCH(i + 1) & SplitDQ(j + 1), "DCInputCLKUP" & CInt(i), CInt(i), , , , , tlForceNone_CZ)
                TheExec.flow.TestLimit resultVal:=DC_Input_CLK_UP.Multiply(100), Tname:=TestNameInput, ForceResults:=tlForceNone_CZ, scaletype:=scaleNoScaling, formatStr:="%1.2f", unit:=unitCustom, customUnit:="%"
                TestNameInput = Report_TName_From_Instance("Calc", SplitGRP(g + 1) & SplitCH(i + 1) & SplitDQ(j + 1), "DCInputCLKNODCC" & CInt(i), CInt(i), , , , , tlForceNone_CZ)
                TheExec.flow.TestLimit resultVal:=DC_Input_CLK_NO_DCC.Multiply(100), Tname:=TestNameInput, ForceResults:=tlForceNone_CZ, scaletype:=scaleNoScaling, formatStr:="%1.2f", unit:=unitCustom, customUnit:="%"
                TestNameInput = Report_TName_From_Instance("Calc", SplitGRP(g + 1) & SplitCH(i + 1) & SplitDQ(j + 1), "DCInputCLKDOWN" & CInt(i), CInt(i), , , , , tlForceNone_CZ)
                TheExec.flow.TestLimit resultVal:=DC_Input_CLK_DOWN.Multiply(100), Tname:=TestNameInput, ForceResults:=tlForceNone_CZ, scaletype:=scaleNoScaling, formatStr:="%1.2f", unit:=unitCustom, customUnit:="%"
                
                TestNameInput = Report_TName_From_Instance("Calc", SplitGRP(g + 1) & SplitCH(i + 1) & SplitDQ(j + 1), "DCSkewInputCLKUP" & CInt(i), CInt(i), , , , , tlForceNone_CZ)
                TheExec.flow.TestLimit resultVal:=DC_Skew_Input_CLK_UP.Multiply(100), Tname:=TestNameInput, ForceResults:=tlForceNone_CZ, scaletype:=scaleNoScaling, formatStr:="%1.2f", unit:=unitCustom, customUnit:="%"
                TestNameInput = Report_TName_From_Instance("Calc", SplitGRP(g + 1) & SplitCH(i + 1) & SplitDQ(j + 1), "DCSkewInputCLKNODCC" & CInt(i), CInt(i), , , , , tlForceNone_CZ)
                TheExec.flow.TestLimit resultVal:=DC_Skew_Input_CLK_NO_DCC.Multiply(100), Tname:=TestNameInput, ForceResults:=tlForceNone_CZ, scaletype:=scaleNoScaling, formatStr:="%1.2f", unit:=unitCustom, customUnit:="%"
                TestNameInput = Report_TName_From_Instance("Calc", SplitGRP(g + 1) & SplitCH(i + 1) & SplitDQ(j + 1), "DCSkewInputCLKDOWN" & CInt(i), CInt(i), , , , , tlForceNone_CZ)
                TheExec.flow.TestLimit resultVal:=DC_Skew_Input_CLK_DOWN.Multiply(100), Tname:=TestNameInput, ForceResults:=tlForceNone_CZ, scaletype:=scaleNoScaling, formatStr:="%1.2f", unit:=unitCustom, customUnit:="%"
                
                TestNameInput = Report_TName_From_Instance("Calc", SplitGRP(g + 1) & SplitCH(i + 1) & SplitDQ(j + 1), "DCCRANGEUP" & CInt(i), CInt(i), , , , , tlForceNone_CZ)
                TheExec.flow.TestLimit resultVal:=DCC_RANGE_UP.Multiply(100), Tname:=TestNameInput, ForceResults:=tlForceNone_CZ, scaletype:=scaleNoScaling, formatStr:="%1.2f", unit:=unitCustom, customUnit:="%"
                TestNameInput = Report_TName_From_Instance("Calc", SplitGRP(g + 1) & SplitCH(i + 1) & SplitDQ(j + 1), "DCCRANGEDOWN" & CInt(i), CInt(i), , , , , tlForceNone_CZ)
                TheExec.flow.TestLimit resultVal:=DCC_RANGE_DOWN.Multiply(100), Tname:=TestNameInput, ForceResults:=tlForceNone_CZ, scaletype:=scaleNoScaling, formatStr:="%1.2f", unit:=unitCustom, customUnit:="%"
            
            Next j
        Next i
    Next g
Exit Function 'Add ErrHandler 2023/05/29
errHandler: 'Add ErrHandler 2023/05/29
    Call Print_Error_Message(Error_Info, "LIB_HardIP_Calc", "Calc_DCC_Skew_Range_DSP_Long") 'Add ErrHandler 2023/05/29
    If AbortTest Then Exit Function Else Resume Next 'Add ErrHandler 2023/05/29
End Function

Public Function Calc_GRPPercentageF1(argc As Integer, argv() As String) As Long
On Error GoTo errHandler 'Add ErrHandler 2023/05/29
   Dim i As Integer
    Dim site As Variant
    Dim SplitStrAry() As String
    
    
    Dim DSP_LoLimit As New DSPWave
    
    
''''''''''    Dim DSP_Temp As New DSPWave
''''''''''    Dim DSP_Summary As New DSPWave
    Dim DSP_EyeWidth As New DSPWave
    Dim DSP_Percentage As New DSPWave
    Dim TestNameInput As String
    Dim Percentage_bysite() As New SiteDouble
    
    
    Dim DSP_Temp() As New DSPWave
''    Dim DSP_Summary() As New DSPWave
    Dim DSP_Summary() As New SiteDouble 'Upadte variable type due to new "Calc_MDLL_Monotonicity_DevideBlock_SEGTTR" @CW 211004 by Carter
    ReDim DSP_Temp(argc - 1)
    ReDim DSP_Summary(argc - 1)
    ReDim Percentage_bysite(argc - 1)
    
    For i = 0 To argc - 1
        SplitStrAry = Split(argv(i), "&")
''        DSP_Summary(i) = GetStoreDataAllType("Summary" & SplitStrAry(0))
        DSP_Summary(i) = GetStoreDataAllType("Summary" & SplitStrAry(0)) 'Upadte variable type due to new "Calc_MDLL_Monotonicity_DevideBlock_SEGTTR" @CW 211004 by Carter
        DSP_Temp(i) = GetStoreDataAllType(SplitStrAry(1))
       
        For Each site In TheExec.sites
            DSP_EyeWidth = DSP_Temp(i).ConvertStreamTo(tldspParallel, DSP_Temp(i).SampleSize, 0, Bit0IsMsb)
'            DSP_Percentage = DSP_EyeWidth.Multiply(100).Divide(DSP_Summary(i).Multiply(2.1))
            DSP_Percentage = DSP_EyeWidth.Multiply(100).divide(DSP_Summary(i).Multiply(2.25))
            Percentage_bysite(i) = DSP_Percentage.Element(0)
        Next site
        
        
        'TestNameInput = Report_TName_From_Instance("C",  SplitStrAry(1), "Percentage" & i, i)
        TestNameInput = Report_TName_From_Instance("CalcC", vbNullString, vbNullString, i)
        For Each site In TheExec.sites
            'theexec.Flow.TestLimit DSP_EyeWidth.Element(0), DSP_Summary(i).Element(0) / 2, DSP_Summary(i).Element(0) * 2.1, , , , , , Tname:=TestNameInput, ForceResults:=tlForceFlow
            'theexec.Flow.TestLimit DSP_EyeWidth.Element(0), DSP_Summary(i).Element(0) / 2, DSP_Summary(i).Element(0) * 2.25, , , , , , Tname:=TestNameInput, ForceResults:=tlForceFlow
                        TheExec.flow.TestLimit DSP_EyeWidth.Element(0), DSP_Summary(i) / 2, DSP_Summary(i) * 2.3, , , , , , PinName:="X", Tname:=TestNameInput, ForceResults:=tlForceFlow
            TheExec.flow.TestLimitIndex = TheExec.flow.TestLimitIndex - 1
        Next site
        TheExec.flow.TestLimitIndex = TheExec.flow.TestLimitIndex + 1
        
        
    Next i
    
    For i = 0 To argc - 1
        SplitStrAry = Split(argv(i), "&")
        'TestNameInput = Report_TName_From_Instance("C", SplitStrAry(1), "Percentage" & i, i, , , , , tlForceNone)
        TestNameInput = Report_TName_From_Instance("CalcC", vbNullString, vbNullString, i)
        'For Each site In TheExec.sites
            TheExec.flow.TestLimit Percentage_bysite(i), 25, 100, , , , , , PinName:="X", Tname:=TestNameInput, ForceResults:=tlForceFlow
       ' Next site
    Next i

Exit Function 'Add ErrHandler 2023/05/29
errHandler: 'Add ErrHandler 2023/05/29
    Call Print_Error_Message(Error_Info, "LIB_HardIP_Calc", "Calc_GRPPercentageF1") 'Add ErrHandler 2023/05/29
    If AbortTest Then Exit Function Else Resume Next 'Add ErrHandler 2023/05/29
End Function

Public Function Calc_GRPPercentageF2(argc As Integer, argv() As String) As Long
On Error GoTo errHandler 'Add ErrHandler 2023/05/29
   Dim i As Integer
    Dim site As Variant
    Dim SplitStrAry() As String
    
    
    Dim DSP_LoLimit As New DSPWave
    
    
''''''''''    Dim DSP_Temp As New DSPWave
''''''''''    Dim DSP_Summary As New DSPWave
    Dim DSP_EyeWidth As New DSPWave
    Dim DSP_Percentage As New DSPWave
    Dim TestNameInput As String
    Dim Percentage_bysite() As New SiteDouble
    
    
    Dim DSP_Temp() As New DSPWave
''    Dim DSP_Summary() As New DSPWave
    Dim DSP_Summary() As New SiteDouble 'Upadte variable type due to new "Calc_MDLL_Monotonicity_DevideBlock_SEGTTR" @CW 211004 by Carter
    ReDim DSP_Temp(argc - 1)
    ReDim DSP_Summary(argc - 1)
    ReDim Percentage_bysite(argc - 1)
    
    For i = 0 To argc - 1
        SplitStrAry = Split(argv(i), "&")
''        DSP_Summary(i) = GetStoreDataAllType("Summary" & SplitStrAry(0))
        DSP_Summary(i) = GetStoreDataAllType("Summary" & SplitStrAry(0)) 'Upadte variable type due to new "Calc_MDLL_Monotonicity_DevideBlock_SEGTTR" @CW 211004 by Carter
        DSP_Temp(i) = GetStoreDataAllType(SplitStrAry(1))
       
        For Each site In TheExec.sites
            DSP_EyeWidth = DSP_Temp(i).ConvertStreamTo(tldspParallel, DSP_Temp(i).SampleSize, 0, Bit0IsMsb)
            DSP_Percentage = DSP_EyeWidth.Multiply(100).divide(DSP_Summary(i).Multiply(2))
            'DSP_Percentage = DSP_EyeWidth.Multiply(100).Divide(DSP_Summary(i).Multiply(2.3))
            Percentage_bysite(i) = DSP_Percentage.Element(0)
        Next site
        
        
        'TestNameInput = Report_TName_From_Instance("C",  SplitStrAry(1), "Percentage" & i, i)
        
        '@210707 CW TTR eye_width
        '--------------------------
'        TestNameInput = Report_TName_From_Instance("CalcC", "", "", i)
'        For Each Site In TheExec.sites
'            'theexec.Flow.TestLimit DSP_EyeWidth.Element(0), DSP_Summary(i).Element(0) / 2, DSP_Summary(i).Element(0) * 2.1, , , , , , Tname:=TestNameInput, ForceResults:=tlForceFlow
'            'theexec.Flow.TestLimit DSP_EyeWidth.Element(0), DSP_Summary(i).Element(0) / 2, DSP_Summary(i).Element(0) * 2.3, , , , , , Tname:=TestNameInput, ForceResults:=tlForceFlow
'
'            If TheExec.Flow.EnableWord("AMPLP5_BinCut_Enable_Flag") = True Then
'                TheExec.Flow.TestLimit DSP_EyeWidth.Element(0), -9999, 9999, , , , , , PinName:="X", Tname:=TestNameInput, ForceResults:=tlForceFlow
'            Else
'                TheExec.Flow.TestLimit DSP_EyeWidth.Element(0), -9999, 9999, , , , , , PinName:="X", Tname:=TestNameInput, ForceResults:=tlForceFlow
'            End If
'
'
'            TheExec.Flow.TestLimitIndex = TheExec.Flow.TestLimitIndex - 1
'        Next Site
'        TheExec.Flow.TestLimitIndex = TheExec.Flow.TestLimitIndex + 1
        '------------------------------
        
    Next i
    
    For i = 0 To argc - 1
        SplitStrAry = Split(argv(i), "&")
        'TestNameInput = Report_TName_From_Instance("C", SplitStrAry(1), "Percentage" & i, i, , , , , tlForceNone)
        TestNameInput = Report_TName_From_Instance("CalcC", vbNullString, vbNullString, i)
        'For Each site In TheExec.sites
            TheExec.flow.TestLimit Percentage_bysite(i), 25, 120, , , , , , PinName:="X", Tname:=TestNameInput, ForceResults:=tlForceFlow
       ' Next site
    Next i
Exit Function 'Add ErrHandler 2023/05/29
errHandler: 'Add ErrHandler 2023/05/29
    Call Print_Error_Message(Error_Info, "LIB_HardIP_Calc", "Calc_GRPPercentageF2") 'Add ErrHandler 2023/05/29
    If AbortTest Then Exit Function Else Resume Next 'Add ErrHandler 2023/05/29

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
        TheExec.flow.TestLimit 9999, 0, 1, , , , , , PinName:="X", Tname:=TestNameInput, ForceResults:=tlForceFlow 'transfer_to_forceflow
        Exit Function
    End If
    
    For i = 0 To argc - 2   ' org : argc-1
        SplitStrAry = Split(argv(i), "&")
'       DSP_Summary(i) = GetStoreDataAllType("Summary" & SplitStrAry(0))
        DSP_Summary(i) = GetStoreDataAllType("Summary" & SplitStrAry(0)) 'Upadte variable type due to new "Calc_MDLL_Monotonicity_DevideBlock_SEGTTR" @CW 211004 by Carter
        DSP_Temp(i) = GetStoreDataAllType(SplitStrAry(1))
       
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
        'TestNameInput = Report_TName_From_Instance("C", SplitStrAry(1), "Percentage" & i, i, , , , , tlForceNone)
        TestNameInput = Report_TName_From_Instance("CalcC", vbNullString, vbNullString, i)
        'TheExec.Flow.TestLimit Percentage_bysite(i), 25, 120, , , , , , PinName:="X", Tname:=TestNameInput, ForceResults:=tlForceFlow
        TheExec.flow.TestLimit Percentage_bysite(i), , , , , , , , PinName:="X", Tname:=TestNameInput, ForceResults:=tlForceFlow
    Next i
    
End Function

Public Function CalcDutyDelay_CapturedFreq(argc As Integer, argv() As String) As Long
On Error GoTo errHandler 'Add ErrHandler 2023/05/29

'"PILINEARF1" as example
'obsolete=>Alg::CalcDutyDelay_M9_PI(SrcCodeIndx,FREQ_RO_DDR0BYTE,112,PILINEARF1)

'New=>Alg::CalcDutyDelay_CapturedFreq(176,200,MD005)

    Dim TestNameInput As String
    
    Dim i As Long, j As Long, k As Long, m As Long, p As Long, q As Long
    Dim site As Variant

    Dim b_DivideZeroError As New SiteBoolean
    b_DivideZeroError = False
    
    Dim MaxNumOfDuty As Long
    Dim StartNumOfDuty As Long
    StartNumOfDuty = argv(0)
    MaxNumOfDuty = argv(1)
    Dim CalcDutyVal() As New PinListData
    ReDim CalcDutyVal(MaxNumOfDuty - StartNumOfDuty) As New PinListData
    
    'Oscar
    Dim tempStoreAllDuty As New PinListData
    Dim EmptyPLD As New PinListData
    Dim CategoryCheckDict As New Dictionary
    Dim DutyNumber As Long: DutyNumber = 0
    Dim GetCapStrArr() As String
    If InStr(LCase(Instance_Data.CUS_Str_DigCapData), "reg_assign") Then
        Dim SplitInstance() As String
        SplitInstance = Split(Instance_Data.CUS_Str_DigCapData, ":")
        Instance_Data.CUS_Str_DigCapData = Replace(Instance_Data.CUS_Str_DigCapData, SplitInstance(0) & ":" & SplitInstance(1), RegDict(LCase("DigCapData_" & SplitInstance(1))))
    End If
    
    GetCapStrArr = Split(Instance_Data.CUS_Str_DigCapData, ",")
    
    For i = 1 To UBound(GetCapStrArr)
    
        GetCapStrArr(i) = Split(GetCapStrArr(i), ":")(2)
    
    Next i
    
    
    Dim DictNum As Long
    Dim Temp_DictNum As Long
    Dim GRP_Num As Long
    Dim CategoryName As String
    Dim keyword As String: keyword = "-sdlinecode"
    
    Dim DutyName As String
    Dim NeedSortData As New SiteLong
    Dim PerfMode As String: PerfMode = argv(2)
    Dim Para As Double
    Dim TCLK As Double
    Dim LSB As Double
    Dim Oct As Double
    
    Dim DNL_MaxMin_HighLimit As Double
    Dim DNL_MaxMin_LowLimit As Double
    Dim PE_MaxMin_HighLimit As Double
    Dim PE_MaxMin_LowLimit As Double
    
       
    Select Case LCase(PerfMode)
    
        Case "md001"
                Para = 22158535
                TCLK = 1250 * ps

''                DNL_MaxMin_HighLimit = 2.5
''                DNL_MaxMin_LowLimit = -1
''                PE_MaxMin_HighLimit = 45
''                PE_MaxMin_LowLimit = -45
        Case "md002"
                Para = 14863473
                TCLK = 625 * ps

''                DNL_MaxMin_HighLimit = 2
''                DNL_MaxMin_LowLimit = -1
''                PE_MaxMin_HighLimit = 30
''                PE_MaxMin_LowLimit = -30
        Case "md003"
                Para = 12958125
                TCLK = 469 * ps

''                DNL_MaxMin_HighLimit = 1.5
''                DNL_MaxMin_LowLimit = -1
''                PE_MaxMin_HighLimit = 25
''                PE_MaxMin_LowLimit = -25
        Case "md004"
                Para = 12228792
                TCLK = 364 * ps

''                DNL_MaxMin_HighLimit = 1.5
''                DNL_MaxMin_LowLimit = -1
''                PE_MaxMin_HighLimit = 25
''                PE_MaxMin_LowLimit = -25
        Case "md005"
                Para = 11861250
                TCLK = 313 * ps

''                DNL_MaxMin_HighLimit = 1.5
''                DNL_MaxMin_LowLimit = -1
''                PE_MaxMin_HighLimit = 20
''                PE_MaxMin_LowLimit = -20
        Case Else
    End Select
    
    LSB = TCLK / 128
    Oct = TCLK / 8
    
    TheExec.Datalog.WriteComment "Mode = " & PerfMode
    TheExec.Datalog.WriteComment "Tclk = " & TCLK / ps & "ps"
    TheExec.Datalog.WriteComment "LSB = " & TCLK / ps & " / 128"
    TheExec.Datalog.WriteComment "Ideal_Oct = " & TCLK / ps & " / 8"
''    TheExec.Datalog.WriteComment "For n = 176 to n = 200"
''    TheExec.Datalog.WriteComment "  DNL0=((Delay[n+1] - Delay[n])/LSB-1)"
''    TheExec.Datalog.WriteComment "Next n"
    
    
    For DictNum = 1 To UBound(GetCapStrArr)

                DutyName = Split(GetCapStrArr(DictNum), keyword)(1)
                CategoryName = Split(GetCapStrArr(DictNum), keyword)(0)
                'GetCapStrArr(DictNum) = Split(GetCapStrArr(DictNum), ":")(2)
                
                NeedSortData = GetStoreDataAllType(GetCapStrArr(DictNum) & "_para")
                
                If CategoryCheckDict.Exists(DutyName) Then                                  ' Save Category name by GRPXX_CH0 and 1 , each group have 32 data
                    tempStoreAllDuty = CategoryCheckDict(DutyName)
                    tempStoreAllDuty.AddPin CategoryName
                    tempStoreAllDuty.pins(CategoryName) = NeedSortData
                    CategoryCheckDict.Remove (DutyName)
                    'CategoryCheckDict.Item(DutyName) = tempStoreAllDuty.Copy
                    CategoryCheckDict.Add DutyName, tempStoreAllDuty
                    Call CategoryCheckDict.Remove(DutyName)
                    Call CategoryCheckDict.Add(DutyName, tempStoreAllDuty.COPY)
                Else
                    tempStoreAllDuty.AddPin CategoryName
                    tempStoreAllDuty.pins(CategoryName) = NeedSortData
                    Call CategoryCheckDict.Add(DutyName, tempStoreAllDuty.COPY)
                End If
                Set tempStoreAllDuty = Nothing

    Next DictNum
    Dim CalcDutyValName As String
    CalcDutyValName = vbNullString
    'Oscar
        
    Dim PinName As String
    For i = 0 To MaxNumOfDuty - StartNumOfDuty
        CalcDutyVal(i) = CategoryCheckDict.Items(i)
        If TheExec.TesterMode = testModeOffline Then
            For j = 0 To CalcDutyVal(i).pins.Count - 1
                CalcDutyVal(i).pins(j) = 1000000 - 1000 * j - i * 2000
            Next j
            ''change simulation data
            If (i Mod 2) = 0 Then
                CalcDutyVal(i) = 100000000
            Else
                CalcDutyVal(i) = 50000000
            End If
        End If
        For j = 0 To CalcDutyVal(i).pins.Count - 1
            'If InStr(UCase(CalcDutyVal(i).Pins(j)), "_N") <> 0 Then 'Oscar
                For Each site In TheExec.sites
                    If CalcDutyVal(i).pins(j).value(site) = 0 Then
                        b_DivideZeroError(site) = True
                        TheExec.Datalog.WriteComment ("Site " & site & " Freq Meas 0 Hz , No CalcDutyDelay ")
                        CalcDutyVal(i).pins(j).value = 1
                    End If
                Next site
                
                CalcDutyVal(i).pins(j).value = CalcDutyVal(i).pins(j).divide(Para)  'Oscar  ''''''Freq = Capture_code/Parameter
                
                CalcDutyVal(i).pins(j).value = CalcDutyVal(i).pins(j).Multiply(2).Invert    ''''''Delay calculation
                                    
                For Each site In TheExec.sites
                    If b_DivideZeroError(site) = True Then
''                        TheExec.Flow.TestLimit resultVal:=-999, ScaleType:=scalePico, Tname:="Delay" & "_" & TestNameInput, ForceResults:=tlForceNone
                        CalcDutyVal(i).pins(j).value = -999
                    End If
                Next site
                 
                TestNameInput = Report_TName_From_Instance("F", CalcDutyVal(i).pins(j), "Delay_Jitter", 0, , , , , tlForceNone_CZ)
                TheExec.flow.TestLimit resultVal:=CalcDutyVal(i).pins(j), Tname:=TestNameInput
                
            'End If 'Oscar
        Next j
        
        'Call StoreDataAllType(LCase(argv(i) & "_" & CStr(i)), CalcDutyVal(i))
        
    Next i
    
    ''''''''DNL0, DNL1 calculation''''''''''''''''''''''''''''''''
        'ReDim CalcDutyVal(MaxNumOfDuty + 1) As New PinListData
        Dim DeltaDelayVal() As New PinListData
        ReDim DeltaDelayVal(MaxNumOfDuty - StartNumOfDuty - 1) As New PinListData

        Dim Freq_Dll_Str As String
        Freq_Dll_Str = argv(argc - 1)

        Dim DNL_Val() As New PinListData
        ReDim DNL_Val(MaxNumOfDuty - StartNumOfDuty - 1) As New PinListData





        For k = 0 To MaxNumOfDuty - StartNumOfDuty - 1
            DeltaDelayVal(k) = CalcDutyVal(k + 1).Math.Subtract(CalcDutyVal(k))  ''''''Step size calculation
            DNL_Val(k) = CalcDutyVal(k + 1).Math.Subtract(CalcDutyVal(k)).divide(LSB).Subtract(1)


            For m = 0 To DNL_Val(k).pins.Count - 1
                PinName = DNL_Val(k).pins(m)
                '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                For Each site In TheExec.sites
                    ''20171117 add for debugging
                    If b_DivideZeroError(site) = True Then
                        DNL_Val(k).pins(m).value = -999
                    End If




                Next site
                '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\

                'Print resuls of "DNL0, DNL1 calculation"'
                TestNameInput = Report_TName_From_Instance("F", DeltaDelayVal(k).pins(m), "Delta_Delay", 0, , , , , tlForceNone_CZ)
                TheExec.flow.TestLimit resultVal:=DeltaDelayVal(k).pins(m), Tname:=TestNameInput
                

            Next m

        Next k

    
Exit Function 'Add ErrHandler 2023/05/29
errHandler: 'Add ErrHandler 2023/05/29
    Call Print_Error_Message(Error_Info, "LIB_HardIP_Calc", "CalcDutyDelay_CapturedFreq") 'Add ErrHandler 2023/05/29
    If AbortTest Then Exit Function Else Resume Next 'Add ErrHandler 2023/05/29
End Function

Public Function CalcDutyDelay_CapturedFreq_PILinear(argc As Integer, argv() As String) As Long
On Error GoTo errHandler 'Add ErrHandler 2023/05/29

'"PILINEARF1" as example
'obsolete=>Alg::CalcDutyDelay_M9_PI(SrcCodeIndx,FREQ_RO_DDR0BYTE,112,PILINEARF1)

'New=>Alg::CalcDutyDelay_JadeCdie(0,112,MD003)

    Dim TestNameInput As String
    Dim CalcDutyVal() As New PinListData
    ReDim CalcDutyVal(112) As New PinListData

    
    Dim i As Long, j As Long, k As Long, m As Long, p As Long, q As Long
    Dim site As Variant

    Dim b_DivideZeroError As New SiteBoolean
    b_DivideZeroError = False
    
    Dim MaxNumOfDuty As Long
    Dim StartNumOfDuty As Long
    StartNumOfDuty = argv(0)
    MaxNumOfDuty = argv(1)
    
    
    'Oscar
    Dim tempStoreAllDuty As New PinListData
    Dim EmptyPLD As New PinListData
    Dim CategoryCheckDict As New Dictionary
    Dim DutyNumber As Long: DutyNumber = 0
    Dim GetCapStrArr() As String
    If InStr(LCase(Instance_Data.CUS_Str_DigCapData), "reg_assign") Then
        Dim SplitInstance() As String
        SplitInstance = Split(Instance_Data.CUS_Str_DigCapData, ":")
        Instance_Data.CUS_Str_DigCapData = Replace(Instance_Data.CUS_Str_DigCapData, SplitInstance(0) & ":" & SplitInstance(1), RegDict(LCase("DigCapData_" & SplitInstance(1))))
    End If
    
    GetCapStrArr = Split(Instance_Data.CUS_Str_DigCapData, ",")
    
    For i = 1 To UBound(GetCapStrArr)
    
        GetCapStrArr(i) = Split(GetCapStrArr(i), ":")(2)
    
    Next i
    
    
    Dim DictNum As Long
    Dim CategoryName As String
    Dim keyword As String: keyword = "-picode-"
    
    Dim DutyName As String
    Dim NeedSortData As New SiteLong
    Dim PerfMode As String: PerfMode = argv(2)
    Dim Para As Double
    Dim TCLK As Double
    Dim LSB As Double
    Dim Oct As Double
    
    Dim DNL_MaxMin_HighLimit As Double
    Dim DNL_MaxMin_LowLimit As Double
    Dim PE_MaxMin_HighLimit As Double
    Dim PE_MaxMin_LowLimit As Double
    
    Select Case LCase(PerfMode)
    
        Case "md001"                                            '''''''Modified 20201215
                Para = 119512155
                TCLK = 1303.8

                DNL_MaxMin_HighLimit = 2.5
                DNL_MaxMin_LowLimit = -1
                PE_MaxMin_HighLimit = 45
                PE_MaxMin_LowLimit = -45
        Case "md002"
                Para = 63461316
                TCLK = 652.3

                DNL_MaxMin_HighLimit = 2
                DNL_MaxMin_LowLimit = -1
                PE_MaxMin_HighLimit = 30
                PE_MaxMin_LowLimit = -30
        Case "md003"
                Para = 47880000
                TCLK = 468.8

                DNL_MaxMin_HighLimit = 1.5
                DNL_MaxMin_LowLimit = -1
                PE_MaxMin_HighLimit = 25
                PE_MaxMin_LowLimit = -25
        Case "md004"
                Para = 39713928
                TCLK = 365.9

                DNL_MaxMin_HighLimit = 1.5
                DNL_MaxMin_LowLimit = -1
                PE_MaxMin_HighLimit = 25
                PE_MaxMin_LowLimit = -25
        Case "md005"
                Para = 35552500
                TCLK = 312.5

                DNL_MaxMin_HighLimit = 1.5
                DNL_MaxMin_LowLimit = -1
                PE_MaxMin_HighLimit = 20
                PE_MaxMin_LowLimit = -20
        Case Else
    End Select
    
    LSB = TCLK / 128
    Oct = TCLK / 8
    
    TheExec.Datalog.WriteComment "Mode = " & PerfMode
    TheExec.Datalog.WriteComment "Tclk = " & TCLK & "ps"
    TheExec.Datalog.WriteComment "LSB = " & TCLK & " / 128"
    TheExec.Datalog.WriteComment "Ideal_Oct = " & TCLK & " / 8"
    TheExec.Datalog.WriteComment "For n = 0 to n = 111"
    TheExec.Datalog.WriteComment "  DNL0=((Delay[n+1] - Delay[n])/LSB-1)"
    TheExec.Datalog.WriteComment "Next n"
    
    
    For DictNum = 1 To UBound(GetCapStrArr)
    
        DutyName = Split(GetCapStrArr(DictNum), keyword)(1)
        CategoryName = Split(GetCapStrArr(DictNum), keyword)(0)
        
        'GetCapStrArr(DictNum) = Split(GetCapStrArr(DictNum), ":")(2)
        
        NeedSortData = GetStoreDataAllType(GetCapStrArr(DictNum) & "_para")
        
        If CategoryCheckDict.Exists(DutyName) Then
            tempStoreAllDuty = CategoryCheckDict(DutyName)
            tempStoreAllDuty.AddPin CategoryName
            tempStoreAllDuty.pins(CategoryName) = NeedSortData
            CategoryCheckDict.Remove (DutyName)
           ' CategoryCheckDict.Item(DutyName) = tempStoreAllDuty.Copy  'update 20200831
            CategoryCheckDict.Add DutyName, tempStoreAllDuty
            Call CategoryCheckDict.Remove(DutyName)
            Call CategoryCheckDict.Add(DutyName, tempStoreAllDuty.COPY)
        Else
            tempStoreAllDuty.AddPin CategoryName
            tempStoreAllDuty.pins(CategoryName) = NeedSortData
            Call CategoryCheckDict.Add(DutyName, tempStoreAllDuty.COPY)
        End If
        Set tempStoreAllDuty = Nothing
    Next DictNum
    Dim CalcDutyValName As String
    CalcDutyValName = vbNullString
    'Oscar
    
    Dim PinName As String
    For i = StartNumOfDuty To MaxNumOfDuty
        CalcDutyVal(i) = CategoryCheckDict.Items(i)
        If TheExec.TesterMode = testModeOffline Then
            For j = 0 To CalcDutyVal(i).pins.Count - 1
                CalcDutyVal(i).pins(j) = 1000000 - 1000 * j - i * 2000
            Next j
            ''change simulation data
            If (i Mod 2) = 0 Then
                CalcDutyVal(i) = 100000000
            Else
                CalcDutyVal(i) = 50000000
            End If
        End If
        For j = 0 To CalcDutyVal(i).pins.Count - 1
            'If InStr(UCase(CalcDutyVal(i).Pins(j)), "_N") <> 0 Then 'Oscar
                For Each site In TheExec.sites
                    If CalcDutyVal(i).pins(j).value(site) = 0 Then
                        b_DivideZeroError(site) = True
                        TheExec.Datalog.WriteComment ("Site " & site & " Freq Meas 0 Hz , No CalcDutyDelay ")
                        CalcDutyVal(i).pins(j).value = 1
                    End If
                Next site
                
                CalcDutyVal(i).pins(j).value = CalcDutyVal(i).pins(j).divide(Para) 'Oscar   ''''''Freq = Capture_code/Parameter     Modified 20201215
                
                CalcDutyVal(i).pins(j).value = CalcDutyVal(i).pins(j).Multiply(2).Invert    ''''''Delay calculation
                    
                For Each site In TheExec.sites
                    If b_DivideZeroError(site) = True Then
''                        TheExec.Flow.TestLimit resultVal:=-999, ScaleType:=scalePico, Tname:="Delay" & "_" & TestNameInput, ForceResults:=tlForceNone
                        CalcDutyVal(i).pins(j).value = -999
                    End If
                Next site
                 
                TestNameInput = Report_TName_From_Instance("C", CalcDutyVal(i).pins(j), "Delay[" & i & "]", 0, , , , , tlForceNone_CZ)
                ''''TheExec.Flow.TestLimit resultVal:=CalcDutyVal(i).Pins(j), lowVal:=DNL_MaxMin_LowLimit, hiVal:=DNL_MaxMin_HighLimit, Tname:=TestNameInput
                TheExec.flow.TestLimit resultVal:=CalcDutyVal(i).pins(j), Tname:=TestNameInput
            'End If 'Oscar
        Next j
        
        'Call StoreDataAllType(LCase(argv(i) & "_" & CStr(i)), CalcDutyVal(i))
        
    Next i
    
    ''''''''DNL0, DNL1 calculation''''''''''''''''''''''''''''''''
        'ReDim CalcDutyVal(MaxNumOfDuty + 1) As New PinListData
        Dim DeltaDelayVal() As New PinListData
        ReDim DeltaDelayVal(MaxNumOfDuty + 1) As New PinListData
    
        Dim Freq_Dll_Str As String
        Freq_Dll_Str = argv(argc - 1)

        
        Dim DNL_Val_Max As New PinListData
        Dim DNL_Val_Min As New PinListData
        Dim DNL_Val_Max_DigSrcCode() As New SiteLong
        Dim DNL_Val_Min_DigSrcCode() As New SiteLong
        
        ReDim DNL_Val_Max_DigSrcCode(CalcDutyVal(0).pins.Count - 1) As New SiteLong
        ReDim DNL_Val_Min_DigSrcCode(CalcDutyVal(0).pins.Count - 1) As New SiteLong
        
        
        

        For k = 0 To MaxNumOfDuty - 1
            DeltaDelayVal(k) = CalcDutyVal(k + 1).Math.Subtract(CalcDutyVal(k)).divide(LSB).Subtract(1)     '''''DNL

            If k = 0 Then  'Initialize Max/Min DNL and DigSrcCode values
                DNL_Val_Max = DeltaDelayVal(0)
                DNL_Val_Min = DeltaDelayVal(0)
                For m = 0 To DeltaDelayVal(k).pins.Count - 1
                    DNL_Val_Max_DigSrcCode(m) = 0
                    DNL_Val_Min_DigSrcCode(m) = 0
                Next m
            End If
            
            For m = 0 To DeltaDelayVal(k).pins.Count - 1
                PinName = DeltaDelayVal(k).pins(m)
                '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                For Each site In TheExec.sites
                    ''20171117 add for debugging
                    If b_DivideZeroError(site) = True Then
                        DeltaDelayVal(k).pins(m).value = -999
                    End If
                    
                    'Find MAX/MIN of DNL0/1 for DigSrc Code[0:111]
                    If DeltaDelayVal(k).pins(m).value > DNL_Val_Max.pins(m).value Then
                        DNL_Val_Max.pins(m).value = DeltaDelayVal(k).pins(m).value
                        DNL_Val_Max_DigSrcCode(m) = k
                    End If
                    If DeltaDelayVal(k).pins(m).value < DNL_Val_Min.pins(m).value Then
                        DNL_Val_Min.pins(m).value = DeltaDelayVal(k).pins(m).value
                        DNL_Val_Min_DigSrcCode(m) = k
                    End If
                    
                Next site
                '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                
                'Print resuls of "DNL0, DNL1 calculation"'
                TestNameInput = Report_TName_From_Instance("C", DeltaDelayVal(k).pins(m), "DNL[" & k & "]", 0, , , , , tlForceNone_CZ)
                TheExec.flow.TestLimit resultVal:=DeltaDelayVal(k).pins(m), Tname:=TestNameInput      '''''DNL


            Next m
                       
        Next k
        
        
        'Print resuls of finding Max/Min of "DNL0, DNL1" value and DigSrc code'
        For p = 0 To DNL_Val_Max.pins.Count - 1
            TestNameInput = Report_TName_From_Instance("C", DNL_Val_Max.pins(p), "Max-Delta-Delay", 0, , , , , tlForceNone_CZ)
            TheExec.flow.TestLimit resultVal:=DNL_Val_Max.pins(p), lowVal:=DNL_MaxMin_LowLimit, hiVal:=DNL_MaxMin_HighLimit, Tname:=TestNameInput, ForceResults:=tlForceNone_CZ
            TestNameInput = Report_TName_From_Instance("C", "X", "Max-Delta-Delay-DigSrcCode", 0, , , , , tlForceNone_CZ)
            TheExec.flow.TestLimit resultVal:=DNL_Val_Max_DigSrcCode(p), Tname:=TestNameInput
        Next
        
        For p = 0 To DNL_Val_Max.pins.Count - 1
            TestNameInput = Report_TName_From_Instance("C", DNL_Val_Min.pins(p), "Min-Delta-Delay", 0, , , , , tlForceNone_CZ)
            TheExec.flow.TestLimit resultVal:=DNL_Val_Min.pins(p), lowVal:=DNL_MaxMin_LowLimit, hiVal:=DNL_MaxMin_HighLimit, Tname:=TestNameInput, ForceResults:=tlForceNone_CZ
            TestNameInput = Report_TName_From_Instance("C", "X", "Min-Delta-Delay-DigSrcCode", 0, , , , , tlForceNone_CZ)
            TheExec.flow.TestLimit resultVal:=DNL_Val_Min_DigSrcCode(p), Tname:=TestNameInput
        Next
        
             
    '''''PE0 to PE7 calculation''''''''''
        Dim OctantIndex As Long
        Dim OctantMaxNum As Long
        'OctantIndex = 0
        OctantMaxNum = 7
        'ReDim CalcDutyVal(CInt(MaxDigSrcCode)) As New PinListData
        Dim Octant_Val() As New PinListData
        ReDim Octant_Val(OctantMaxNum) As New PinListData
        Dim n As Long
        
        Dim PE_Val_Max As New PinListData
        Dim PE_Val_Min As New PinListData
        Dim PE_Val_Max_DigSrcCode() As New SiteLong
        Dim PE_Val_Min_DigSrcCode() As New SiteLong
        
        ReDim PE_Val_Max_DigSrcCode(CalcDutyVal(0).pins.Count - 1) As New SiteLong
        ReDim PE_Val_Min_DigSrcCode(CalcDutyVal(0).pins.Count - 1) As New SiteLong
        
        TheExec.Datalog.WriteComment "***PE0 to PE7 calculation for" & PerfMode & " ***"

        For n = 0 To OctantMaxNum
            If n <> 7 Then
                Octant_Val(n) = CalcDutyVal((n + 1) * 16).Math.Subtract(CalcDutyVal(n * 16)).Subtract(Oct)
                '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                ''20171117 add for debugging
                For m = 0 To Octant_Val(n).pins.Count - 1
                    For Each site In TheExec.sites
                        If b_DivideZeroError(site) = True Then
                            Octant_Val(n).pins(m).value = -999
                        End If
                    Next site
                Next m
                '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
            Else
                Octant_Val(n) = CalcDutyVal(n * 0).Math.Subtract(CalcDutyVal(n * 16)).Add(TCLK).Subtract(Oct)
                '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                ''20171117 add for debugging
                For m = 0 To Octant_Val(n).pins.Count - 1
                    For Each site In TheExec.sites
                        If b_DivideZeroError(site) = True Then
                            Octant_Val(n).pins(m).value = -999
                        End If
                    Next site
                Next m
                '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
            End If
            'TestNameInput = Report_TName_From_Instance("C", "X", "PE" & n & "Delta_Delay", 0, , , , , tlForceNone)
            'TheExec.Flow.TestLimit resultVal:=Octant_Val(n), Tname:=TestNameInput           '''''Modified 20201215
            
            For m = 0 To Octant_Val(n).pins.Count - 1
                TestNameInput = Report_TName_From_Instance("C", Octant_Val(n).pins(m), "PE" & n & "-Delta-Delay", 0, , , , , tlForceNone_CZ)
                TheExec.flow.TestLimit resultVal:=Octant_Val(n).pins(m), Tname:=TestNameInput           '''''Modified 20201215
            Next m
        
        
        'Find MAX/MIN of PE0 to PE7         add 20201215
            
            If n = 0 Then  'Initialize Max/Min PE and DigSrcCode values
                PE_Val_Max = Octant_Val(0)
                PE_Val_Min = Octant_Val(0)
                For m = 0 To Octant_Val(n).pins.Count - 1
                    PE_Val_Max_DigSrcCode(m) = 0
                    PE_Val_Min_DigSrcCode(m) = 0
                Next m
            End If
        
            For m = 0 To Octant_Val(n).pins.Count - 1

                '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                For Each site In TheExec.sites
                    If b_DivideZeroError(site) = True Then
                        Octant_Val(n).pins(m).value = -999
                    End If

                    If Octant_Val(n).pins(m).value > PE_Val_Max.pins(m).value Then
                        PE_Val_Max.pins(m).value = Octant_Val(n).pins(m).value
                        PE_Val_Max_DigSrcCode(m) = n
                    End If
                    If DeltaDelayVal(n).pins(m).value < PE_Val_Min.pins(m).value Then
                        PE_Val_Min.pins(m).value = Octant_Val(n).pins(m).value
                        PE_Val_Min_DigSrcCode(m) = n
                    End If
                    
                Next site
                '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
      
            Next m
            
            
        Next n
        
        For p = 0 To PE_Val_Max.pins.Count - 1
            TestNameInput = Report_TName_From_Instance("C", PE_Val_Max.pins(p), "Max-PE", 0, , , , , tlForceNone_CZ)
            TheExec.flow.TestLimit resultVal:=PE_Val_Max.pins(p), lowVal:=PE_MaxMin_LowLimit, hiVal:=PE_MaxMin_HighLimit, Tname:=TestNameInput, ForceResults:=tlForceNone_CZ
            TestNameInput = Report_TName_From_Instance("C", "X", "Max-PE-DigSrcCode", 0, , , , , tlForceNone_CZ)
            TheExec.flow.TestLimit resultVal:=PE_Val_Max_DigSrcCode(p), Tname:=TestNameInput
        Next
        
        For p = 0 To PE_Val_Max.pins.Count - 1
            TestNameInput = Report_TName_From_Instance("C", PE_Val_Min.pins(p), "Min-PE", 0, , , , , tlForceNone_CZ)
            TheExec.flow.TestLimit resultVal:=PE_Val_Min.pins(p), lowVal:=PE_MaxMin_LowLimit, hiVal:=PE_MaxMin_HighLimit, Tname:=TestNameInput, ForceResults:=tlForceNone_CZ
            TestNameInput = Report_TName_From_Instance("C", "X", "Min-PE-DigSrcCode", 0, , , , , tlForceNone_CZ)
            TheExec.flow.TestLimit resultVal:=PE_Val_Min_DigSrcCode(p), Tname:=TestNameInput
        Next
    
    
    
Exit Function 'Add ErrHandler 2023/05/29
errHandler: 'Add ErrHandler 2023/05/29
    Call Print_Error_Message(Error_Info, "LIB_HardIP_Calc", "CalcDutyDelay_CapturedFreq_PILinear") 'Add ErrHandler 2023/05/29
    If AbortTest Then Exit Function Else Resume Next 'Add ErrHandler 2023/05/29
End Function

Public Function CalcDutyDelay_CapturedFreq_PILinear_Backup(argc As Integer, argv() As String) As Long
On Error GoTo errHandler 'Add ErrHandler 2023/05/29
''' Update for Crete MN pattern @William
'20210225 change for Staten @CW
'"PILINEARF1" as example
'obsolete=>Alg::CalcDutyDelay_M9_PI(SrcCodeIndx,FREQ_RO_DDR0BYTE,112,PILINEARF1)

'New=>Alg::CalcDutyDelay_JadeCdie(0,112,MD003)

    Dim TestNameInput As String
    Dim CalcDutyVal() As New PinListData
    'ReDim CalcDutyVal(112) As New PinListData

    
    Dim i As Long, j As Long, k As Long, m As Long, p As Long, q As Long
    Dim site As Variant

    Dim Start_i As Long, End_i As Long, z As Long: z = 0
    Dim b_DivideZeroError As New SiteBoolean
    b_DivideZeroError = False
    
    Dim MaxNumOfDuty As Long
    Dim StartNumOfDuty As Long
    StartNumOfDuty = argv(0)
    MaxNumOfDuty = argv(1)
    
    ReDim CalcDutyVal(MaxNumOfDuty - StartNumOfDuty) As New PinListData ' @CW
    
    'Oscar
    Dim tempStoreAllDuty As New PinListData
    Dim EmptyPLD As New PinListData
    Dim CategoryCheckDict As New Dictionary
    Dim DutyNumber As Long: DutyNumber = 0
    Dim GetCapStrArr() As String
    Dim GetCapStrArr_temp() As String
    If InStr(LCase(Instance_Data.CUS_Str_DigCapData), "reg_assign") Then
        Dim SplitInstance() As String
        SplitInstance = Split(Instance_Data.CUS_Str_DigCapData, ":")
        Instance_Data.CUS_Str_DigCapData = Replace(Instance_Data.CUS_Str_DigCapData, SplitInstance(0) & ":" & SplitInstance(1), RegDict(LCase("DigCapData_" & SplitInstance(1))))
    End If
    
    GetCapStrArr_temp = Split(Instance_Data.CUS_Str_DigCapData, ",")
    
    
    'CW
    If argc > 3 Then
        For i = 1 To UBound(GetCapStrArr_temp)
            If LCase(argv(3)) = LCase(Split(GetCapStrArr_temp(i), ":")(2)) Then Start_i = i
            If LCase(argv(4)) = LCase(Split(GetCapStrArr_temp(i), ":")(2)) Then End_i = i
        Next i
    Else
        Start_i = 1
        End_i = UBound(GetCapStrArr_temp)
    End If
    
    ReDim GetCapStrArr(1 To End_i - Start_i + 1)
    
    For i = Start_i To End_i
        z = z + 1
        GetCapStrArr(z) = Split(GetCapStrArr_temp(i), ":")(2)
    
    Next i
    
    
    Dim DictNum As Long
    Dim CategoryName As String
    Dim keyword As String: keyword = "__pi_" '"-picode-"
    
    Dim DutyName As String
    Dim NeedSortData As New SiteLong
    Dim PerfMode As String: PerfMode = argv(2)
    Dim Para As Double
    Dim TCLK As Double
    Dim LSB As Double
    Dim Oct As Double
    
    Dim DNL_MaxMin_HighLimit As Double
    Dim DNL_MaxMin_LowLimit As Double
    Dim PE_MaxMin_HighLimit As Double
    Dim PE_MaxMin_LowLimit As Double
    
    Select Case LCase(PerfMode)
    
        Case "md001"                                            '''''''Modified 20201215
                Para = 119512155
                TCLK = 1303.8

                DNL_MaxMin_HighLimit = 2.5
                DNL_MaxMin_LowLimit = -1
                PE_MaxMin_HighLimit = 45
                PE_MaxMin_LowLimit = -45
        Case "md002"
                Para = 63461316
                TCLK = 652.3

                DNL_MaxMin_HighLimit = 2
                DNL_MaxMin_LowLimit = -1
                PE_MaxMin_HighLimit = 30
                PE_MaxMin_LowLimit = -30
        Case "md003"
                Para = 47880000
                TCLK = 468.8

                DNL_MaxMin_HighLimit = 1.5
                DNL_MaxMin_LowLimit = -1
                PE_MaxMin_HighLimit = 25
                PE_MaxMin_LowLimit = -25
        Case "md004"
                Para = 39713928
                TCLK = 365.9

                DNL_MaxMin_HighLimit = 1.5
                DNL_MaxMin_LowLimit = -1
                PE_MaxMin_HighLimit = 25
                PE_MaxMin_LowLimit = -25
        Case "md005"
                Para = 35552500
                TCLK = 312.5

                DNL_MaxMin_HighLimit = 1.5
                DNL_MaxMin_LowLimit = -1
                PE_MaxMin_HighLimit = 20
                PE_MaxMin_LowLimit = -20
        Case Else
    End Select
    
    LSB = TCLK / 128
    Oct = TCLK / 8
    
    TheExec.Datalog.WriteComment "Mode = " & PerfMode
    TheExec.Datalog.WriteComment "Tclk = " & TCLK & "ps"
    TheExec.Datalog.WriteComment "LSB = " & TCLK & " / 128"
    TheExec.Datalog.WriteComment "Ideal_Oct = " & TCLK & " / 8"
    TheExec.Datalog.WriteComment "For n = 0 to n = 111"
    TheExec.Datalog.WriteComment "  DNL0=((Delay[n+1] - Delay[n])/LSB-1)"
    TheExec.Datalog.WriteComment "Next n"
    
    
    For DictNum = 1 To UBound(GetCapStrArr)
    
        DutyName = Split(GetCapStrArr(DictNum), keyword)(1)
        CategoryName = Split(GetCapStrArr(DictNum), keyword)(0)
        
        'GetCapStrArr(DictNum) = Split(GetCapStrArr(DictNum), ":")(2)
        
        NeedSortData = GetStoreDataAllType(GetCapStrArr(DictNum) & "_para")
        
        If CategoryCheckDict.Exists(DutyName) Then
            tempStoreAllDuty = CategoryCheckDict(DutyName)
            tempStoreAllDuty.AddPin CategoryName
            tempStoreAllDuty.pins(CategoryName) = NeedSortData
            CategoryCheckDict.Remove (DutyName)
           ' CategoryCheckDict.Item(DutyName) = tempStoreAllDuty.Copy  'update 20200831
            CategoryCheckDict.Add DutyName, tempStoreAllDuty
            Call CategoryCheckDict.Remove(DutyName)
            Call CategoryCheckDict.Add(DutyName, tempStoreAllDuty.COPY)
        Else
            tempStoreAllDuty.AddPin CategoryName
            tempStoreAllDuty.pins(CategoryName) = NeedSortData
            Call CategoryCheckDict.Add(DutyName, tempStoreAllDuty.COPY)
        End If
        Set tempStoreAllDuty = Nothing
    Next DictNum
    Dim CalcDutyValName As String
    CalcDutyValName = vbNullString
    'Oscar
    
    Dim PinName As String
    'For i = StartNumOfDuty To MaxNumOfDuty
    For i = 0 To MaxNumOfDuty - StartNumOfDuty '@CW
        CalcDutyVal(i) = CategoryCheckDict.Items(i)
        If TheExec.TesterMode = testModeOffline Then
            For j = 0 To CalcDutyVal(i).pins.Count - 1
                CalcDutyVal(i).pins(j) = 1000000 - 1000 * j - i * 2000
            Next j
            ''change simulation data
            If (i Mod 2) = 0 Then
                CalcDutyVal(i) = 100000000
            Else
                CalcDutyVal(i) = 50000000
            End If
        End If
        For j = 0 To CalcDutyVal(i).pins.Count - 1
            'If InStr(UCase(CalcDutyVal(i).Pins(j)), "_N") <> 0 Then 'Oscar
                For Each site In TheExec.sites
                    If CalcDutyVal(i).pins(j).value(site) = 0 Then
                        b_DivideZeroError(site) = True
                        TheExec.Datalog.WriteComment ("Site " & site & " Freq Meas 0 Hz , No CalcDutyDelay ")
                        CalcDutyVal(i).pins(j).value = 1
                    End If
                Next site
                
                CalcDutyVal(i).pins(j).value = CalcDutyVal(i).pins(j).divide(Para) 'Oscar   ''''''Freq = Capture_code/Parameter     Modified 20201215
                
                CalcDutyVal(i).pins(j).value = CalcDutyVal(i).pins(j).Multiply(2).Invert    ''''''Delay calculation
                    
                For Each site In TheExec.sites
                    If b_DivideZeroError(site) = True Then
''                        TheExec.Flow.TestLimit resultVal:=-999, ScaleType:=scalePico, Tname:="Delay" & "_" & TestNameInput, ForceResults:=tlForceNone
                        CalcDutyVal(i).pins(j).value = -999
                    End If
                Next site
                 
                'TestNameInput = Report_TName_From_Instance("F", CalcDutyVal(i).Pins(j), CalcDutyVal(i).Pins(j) & "__Jitter-" & CStr(i + StartNumOfDuty), 0, , , , , tlForceNone)
                TestNameInput = Report_TName_From_Instance("C", CalcDutyVal(i).pins(j), "Delay[" & i & "]", 0, , , , , tlForceFlow) 'transfer_to_forceflow
                ''''TheExec.Flow.TestLimit resultVal:=CalcDutyVal(i).Pins(j), lowVal:=DNL_MaxMin_LowLimit, hiVal:=DNL_MaxMin_HighLimit, Tname:=TestNameInput
                TheExec.flow.TestLimit resultVal:=CalcDutyVal(i).pins(j), Tname:=TestNameInput
                
                
            'End If 'Oscar
        Next j
        
        'Call StoreDataAllType(LCase(argv(i) & "_" & CStr(i)), CalcDutyVal(i))
        
    Next i
    
    ''''''''DNL0, DNL1 calculation''''''''''''''''''''''''''''''''
        'ReDim CalcDutyVal(MaxNumOfDuty + 1) As New PinListData
        Dim DeltaDelayVal() As New PinListData
        'ReDim DeltaDelayVal(MaxNumOfDuty + 1) As New PinListData
        ReDim DeltaDelayVal(MaxNumOfDuty - StartNumOfDuty - 1) As New PinListData '@CW
        
        Dim Freq_Dll_Str As String
        Freq_Dll_Str = argv(argc - 1)

        
        Dim DNL_Val_Max As New PinListData
        Dim DNL_Val_Min As New PinListData
        Dim DNL_Val_Max_DigSrcCode() As New SiteLong
        Dim DNL_Val_Min_DigSrcCode() As New SiteLong
        
        ReDim DNL_Val_Max_DigSrcCode(CalcDutyVal(0).pins.Count - 1) As New SiteLong
        ReDim DNL_Val_Min_DigSrcCode(CalcDutyVal(0).pins.Count - 1) As New SiteLong
        
        
        

        'For k = 0 To MaxNumOfDuty - 1
        For k = 0 To MaxNumOfDuty - StartNumOfDuty - 1 '@CW
            DeltaDelayVal(k) = CalcDutyVal(k + 1).Math.Subtract(CalcDutyVal(k)).divide(LSB).Subtract(1)     '''''DNL

            If k = 0 Then  'Initialize Max/Min DNL and DigSrcCode values
                DNL_Val_Max = DeltaDelayVal(0)
                DNL_Val_Min = DeltaDelayVal(0)
                For m = 0 To DeltaDelayVal(k).pins.Count - 1
                    DNL_Val_Max_DigSrcCode(m) = 0
                    DNL_Val_Min_DigSrcCode(m) = 0
                Next m
            End If
            
            For m = 0 To DeltaDelayVal(k).pins.Count - 1
                PinName = DeltaDelayVal(k).pins(m)
                '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                For Each site In TheExec.sites
                    ''20171117 add for debugging
                    If b_DivideZeroError(site) = True Then
                        DeltaDelayVal(k).pins(m).value = -999
                    End If
                    
                    'Find MAX/MIN of DNL0/1 for DigSrc Code[0:111]
                    If DeltaDelayVal(k).pins(m).value > DNL_Val_Max.pins(m).value Then
                        DNL_Val_Max.pins(m).value = DeltaDelayVal(k).pins(m).value
                        DNL_Val_Max_DigSrcCode(m) = k
                    End If
                    If DeltaDelayVal(k).pins(m).value < DNL_Val_Min.pins(m).value Then
                        DNL_Val_Min.pins(m).value = DeltaDelayVal(k).pins(m).value
                        DNL_Val_Min_DigSrcCode(m) = k
                    End If
                    
                Next site
                '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                
                'Print resuls of "DNL0, DNL1 calculation"'
                TestNameInput = Report_TName_From_Instance("C", DeltaDelayVal(k).pins(m), "DNL[" & k & "]", 0, , , , , tlForceNone)
                TheExec.flow.TestLimit resultVal:=DeltaDelayVal(k).pins(m), Tname:=TestNameInput      '''''DNL

                'TestNameInput = Report_TName_From_Instance("F", CalcDutyVal(k).Pins(m), "", 0, , , , , tlForceFlow)
                'TheExec.Flow.TestLimit resultVal:=DeltaDelayVal(k).Pins(m), Tname:=TestNameInput, ForceResults:=tlForceFlow

            Next m
                       
        Next k
        
        
        'Print resuls of finding Max/Min of "DNL0, DNL1" value and DigSrc code'
        For p = 0 To DNL_Val_Max.pins.Count - 1
            TestNameInput = Report_TName_From_Instance("C", DNL_Val_Max.pins(p), "Max-Delta-Delay", 0, , , , , tlForceNone)
            TheExec.flow.TestLimit resultVal:=DNL_Val_Max.pins(p), lowVal:=DNL_MaxMin_LowLimit, hiVal:=DNL_MaxMin_HighLimit, Tname:=TestNameInput, ForceResults:=tlForceNone 'Un-Used_
            TestNameInput = Report_TName_From_Instance("C", DNL_Val_Max.pins(p), "Max-Delta-Delay-DigSrcCode", 0, , , , , tlForceNone)
            TheExec.flow.TestLimit resultVal:=DNL_Val_Max_DigSrcCode(p), PinName:=DNL_Val_Max.pins(p), Tname:=TestNameInput
        Next
        
        For p = 0 To DNL_Val_Max.pins.Count - 1
            TestNameInput = Report_TName_From_Instance("C", DNL_Val_Min.pins(p), "Min-Delta-Delay", 0, , , , , tlForceNone)
            TheExec.flow.TestLimit resultVal:=DNL_Val_Min.pins(p), lowVal:=DNL_MaxMin_LowLimit, hiVal:=DNL_MaxMin_HighLimit, Tname:=TestNameInput, ForceResults:=tlForceNone 'Un-Used_
            TestNameInput = Report_TName_From_Instance("C", DNL_Val_Min.pins(p), "Min-Delta-Delay-DigSrcCode", 0, , , , , tlForceNone)
            TheExec.flow.TestLimit resultVal:=DNL_Val_Min_DigSrcCode(p), PinName:=DNL_Val_Min.pins(p), Tname:=TestNameInput
        Next
        
        Exit Function 'CW
    '''''PE0 to PE7 calculation''''''''''
        Dim OctantIndex As Long
        Dim OctantMaxNum As Long
        'OctantIndex = 0
        OctantMaxNum = 7
        'ReDim CalcDutyVal(CInt(MaxDigSrcCode)) As New PinListData
        Dim Octant_Val() As New PinListData
        ReDim Octant_Val(OctantMaxNum) As New PinListData
        Dim n As Long
        
        Dim PE_Val_Max As New PinListData
        Dim PE_Val_Min As New PinListData
        Dim PE_Val_Max_DigSrcCode() As New SiteLong
        Dim PE_Val_Min_DigSrcCode() As New SiteLong
        
        ReDim PE_Val_Max_DigSrcCode(CalcDutyVal(0).pins.Count - 1) As New SiteLong
        ReDim PE_Val_Min_DigSrcCode(CalcDutyVal(0).pins.Count - 1) As New SiteLong
        
        TheExec.Datalog.WriteComment "***PE0 to PE7 calculation for" & PerfMode & " ***"

        For n = 0 To OctantMaxNum
            If n <> 7 Then
                Octant_Val(n) = CalcDutyVal((n + 1) * 16).Math.Subtract(CalcDutyVal(n * 16)).Subtract(Oct)
                '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                ''20171117 add for debugging
                For m = 0 To Octant_Val(n).pins.Count - 1
                    For Each site In TheExec.sites
                        If b_DivideZeroError(site) = True Then
                            Octant_Val(n).pins(m).value = -999
                        End If
                    Next site
                Next m
                '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
            Else
                Octant_Val(n) = CalcDutyVal(n * 0).Math.Subtract(CalcDutyVal(n * 16)).Add(TCLK).Subtract(Oct)
                '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                ''20171117 add for debugging
                For m = 0 To Octant_Val(n).pins.Count - 1
                    For Each site In TheExec.sites
                        If b_DivideZeroError(site) = True Then
                            Octant_Val(n).pins(m).value = -999
                        End If
                    Next site
                Next m
                '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
            End If
            'TestNameInput = Report_TName_From_Instance("C", "X", "PE" & n & "Delta_Delay", 0, , , , , tlForceNone)
            'TheExec.Flow.TestLimit resultVal:=Octant_Val(n), Tname:=TestNameInput           '''''Modified 20201215
            
            For m = 0 To Octant_Val(n).pins.Count - 1
                TestNameInput = Report_TName_From_Instance("C", Octant_Val(n).pins(m), "PE" & n & "-Delta-Delay", 0, , , , , tlForceNone)
                TheExec.flow.TestLimit resultVal:=Octant_Val(n).pins(m), Tname:=TestNameInput           '''''Modified 20201215
            Next m
        
        
        'Find MAX/MIN of PE0 to PE7         add 20201215
            
            If n = 0 Then  'Initialize Max/Min PE and DigSrcCode values
                PE_Val_Max = Octant_Val(0)
                PE_Val_Min = Octant_Val(0)
                For m = 0 To Octant_Val(n).pins.Count - 1
                    PE_Val_Max_DigSrcCode(m) = 0
                    PE_Val_Min_DigSrcCode(m) = 0
                Next m
            End If
        
            For m = 0 To Octant_Val(n).pins.Count - 1

                '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                For Each site In TheExec.sites
                    If b_DivideZeroError(site) = True Then
                        Octant_Val(n).pins(m).value = -999
                    End If

                    If Octant_Val(n).pins(m).value > PE_Val_Max.pins(m).value Then
                        PE_Val_Max.pins(m).value = Octant_Val(n).pins(m).value
                        PE_Val_Max_DigSrcCode(m) = n
                    End If
                    If DeltaDelayVal(n).pins(m).value < PE_Val_Min.pins(m).value Then
                        PE_Val_Min.pins(m).value = Octant_Val(n).pins(m).value
                        PE_Val_Min_DigSrcCode(m) = n
                    End If
                    
                Next site
                '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
      
            Next m
            
            
        Next n
        
        For p = 0 To PE_Val_Max.pins.Count - 1
            TestNameInput = Report_TName_From_Instance("C", PE_Val_Max.pins(p), "Max-PE", 0, , , , , tlForceNone)
            TheExec.flow.TestLimit resultVal:=PE_Val_Max.pins(p), lowVal:=PE_MaxMin_LowLimit, hiVal:=PE_MaxMin_HighLimit, Tname:=TestNameInput, ForceResults:=tlForceNone 'Un-Used_
            TestNameInput = Report_TName_From_Instance("C", "X", "Max-PE-DigSrcCode", 0, , , , , tlForceNone)
            TheExec.flow.TestLimit resultVal:=PE_Val_Max_DigSrcCode(p), Tname:=TestNameInput
        Next
        
        For p = 0 To PE_Val_Max.pins.Count - 1
            TestNameInput = Report_TName_From_Instance("C", PE_Val_Min.pins(p), "Min-PE", 0, , , , , tlForceNone)
            TheExec.flow.TestLimit resultVal:=PE_Val_Min.pins(p), lowVal:=PE_MaxMin_LowLimit, hiVal:=PE_MaxMin_HighLimit, Tname:=TestNameInput, ForceResults:=tlForceNone 'Un-Used_
            TestNameInput = Report_TName_From_Instance("C", "X", "Min-PE-DigSrcCode", 0, , , , , tlForceNone)
            TheExec.flow.TestLimit resultVal:=PE_Val_Min_DigSrcCode(p), Tname:=TestNameInput
        Next
    
    
    
Exit Function 'Add ErrHandler 2023/05/29
errHandler: 'Add ErrHandler 2023/05/29
    Call Print_Error_Message(Error_Info, "LIB_HardIP_Calc", "CalcDutyDelay_CapturedFreq_PILinear_Backup") 'Add ErrHandler 2023/05/29
    If AbortTest Then Exit Function Else Resume Next 'Add ErrHandler 2023/05/29
End Function

Public Function Calc_EquationTrim(argc As Integer, argv() As String) As Long
On Error GoTo errHandler 'Add ErrHandler 2023/05/29

    'argv(0) : Target       ex:5600
    'argv(1) : MeasOrCap    ex:Meas or Cap
    'argv(2) : InputDicKey  ex:DigCap1+DigCap2
    'argv(3) : FuseDicKey   ex:FuseName1+FuseName2
    'argv(4) : FuseBitNum   ex:6
    'argv(5) : TrimCodeSize ex:63
    'argv(6) : TrimCodeEq   ex:0.894+3.04E-3+1.17E-5+1.72E-7
    '                          from x^0 to x^3
    '                          (1.72 * 10 ^ -7 * i ^ 3) + (1.17 * 10 ^ -5 * i ^ 2) + (3.04 * 10 ^ -3 * i) + 0.894
    
    '        Multiply_Fcount(i) = (2.19 * 10 ^ -9 * i ^ 4) - (1.08 * 10 ^ -7 * i ^ 3) + _
                             (2.2 * 10 ^ -5 * i ^ 2) + (2.76 * 10 ^ -3 * i) + 0.855
    
    Dim vsite As Variant
    Dim TestNameInput As String
    Dim i As Long
    Dim j As Long
    
    Dim TrimTarget As Double
    Dim InputDicKeyArr() As String
    Dim MeasOrCap As String
    Dim FuseDicKeyArr() As String
    Dim FuseBitNum As Long
    Dim TrimCodeSize As Long
    Dim TrimCodeEqArr() As String
    Dim Result_TempStr As String
    
    Dim InputDicVal() As New SiteDouble
    Dim CalculatedDsp() As New DSPWave
    Dim TrimCodeEqDsp As New DSPWave
    Dim TrimCodeEqDsp_Coef As New DSPWave
    Dim tempLong As Long
    Dim Result_DiffVal() As New SiteDouble
    Dim Result_Code() As New SiteLong
    Dim Result_CodeDecDsp() As New DSPWave
    Dim Result_CodeBinDsp() As New DSPWave
    Dim Result_CodeBinStr As String
    Dim TempLog As String
    
    Const cSplitSymbol = "+"
    Const cPrintInputVal = True
    Const cPrintAllCodeVal = False
    Const cPrintFuseBin = False
    
    'init
    TrimTarget = CDbl(argv(0))
    MeasOrCap = argv(1)
    InputDicKeyArr = Split(argv(2), cSplitSymbol)
    FuseDicKeyArr = Split(argv(3), cSplitSymbol)
    FuseBitNum = CLng(argv(4))
    TrimCodeSize = CLng(argv(5))
    TrimCodeEqArr = Split(argv(6), cSplitSymbol)
    ReDim InputDicVal(UBound(InputDicKeyArr))
    ReDim CalculatedDsp(UBound(InputDicKeyArr))
    ReDim Result_DiffVal(UBound(InputDicKeyArr))
    ReDim Result_Code(UBound(InputDicKeyArr))
    ReDim Result_CodeDecDsp(UBound(InputDicKeyArr))
    ReDim Result_CodeBinDsp(UBound(InputDicKeyArr))
    
    'equation
    TrimCodeEqDsp_Coef.CreateConstant 0, UBound(TrimCodeEqArr) + 1, DspDouble
    For i = 0 To UBound(TrimCodeEqArr)
        TrimCodeEqDsp_Coef.Element(i) = CDbl(TrimCodeEqArr(i))
    Next i
'    TrimCodeEqDsp.CreatePolynomial TrimCodeEqDsp_Coef, TrimCodeSize + 1
    
    'trim calc
    For i = 0 To UBound(InputDicKeyArr)
        If MeasOrCap = "Meas" Then
            InputDicVal(i) = GetStoreDataAllType(InputDicKeyArr(i))
        ElseIf MeasOrCap = "Cap" Then
            InputDicVal(i) = GetStoreDataAllType(InputDicKeyArr(i) & "_para")
        Else
            TheExec.Datalog.WriteComment "Calc_EquationTrim : input dic is not correct."
        End If
        CalculatedDsp(i).CreateConstant 0, TrimCodeSize + 1, DspDouble
        Result_CodeDecDsp(i).CreateConstant 0, 1, DspLong
        If True Then
            Call rundsp.Calc_EquationTrim_DSP(Result_Code(i), Result_DiffVal(i), Result_CodeBinDsp(i), Result_CodeDecDsp(i), CalculatedDsp(i), TrimTarget, InputDicVal(i), FuseBitNum, TrimCodeSize, TrimCodeEqDsp_Coef)
        Else
            TrimCodeEqDsp_Coef.CreateConstant 0, UBound(TrimCodeEqArr) + 1, DspDouble
            TrimCodeEqDsp.CreatePolynomial TrimCodeEqDsp_Coef, TrimCodeSize + 1
            For Each vsite In TheExec.sites.Selected
                CalculatedDsp(i) = TrimCodeEqDsp.Multiply(CDbl(InputDicVal(i)(vsite))).Subtract(TrimTarget).Abs.COPY
                Result_DiffVal(i)(vsite) = CalculatedDsp(i).CalcMinimumValue(tempLong)
                Result_Code(i)(vsite) = tempLong
                Result_CodeDecDsp(i).Element(0) = Result_Code(i)(vsite)
                Result_CodeBinDsp(i) = Result_CodeDecDsp(i).ConvertStreamTo(tldspSerial, FuseBitNum, 0, Bit0IsMsb)
            Next vsite
        End If
    Next i
    
    'add fuse
    If UBound(InputDicKeyArr) <> UBound(FuseDicKeyArr) Then
        TheExec.Datalog.WriteComment "Calc_EquationTrim : input is not match."
    Else
        For i = 0 To UBound(InputDicKeyArr)
            Call StoreDataAllType(FuseDicKeyArr(i), Result_CodeBinDsp(i))
        Next i
    End If
    
    'log
    If cPrintInputVal Then
        TheExec.Datalog.WriteComment "------------------------------------" & "Input and Trim Code Result" & "------------------------------------"
        For Each vsite In TheExec.sites.Selected
            TempLog = vbNullString
            For i = 0 To UBound(InputDicKeyArr)
                TempLog = TempLog & "site[" & vsite & "]" & " InputDicValue =" & CStr(InputDicVal(i)(vsite))
                TempLog = TempLog & ", FinalTrimCode = " & CStr(Result_Code(i)(vsite))
                TempLog = TempLog & ", FinalDiffVal = " & CStr(Result_DiffVal(i)(vsite)) & vbCrLf
            Next i
            TheExec.Datalog.WriteComment TempLog
        Next vsite
    End If
    
    'log
    If cPrintAllCodeVal Then
        TheExec.Datalog.WriteComment "------------------------------------" & "Target Trim for " & TrimTarget & "------------------------------------"
        For Each vsite In TheExec.sites.Selected
            For i = 0 To UBound(InputDicKeyArr)
                TempLog = "Start Find Trim Value For " & InputDicKeyArr(i) & vbCrLf
                For j = 0 To TrimCodeSize
                    If j = Result_Code(i)(vsite) Then
                        TempLog = TempLog & "site[" & vsite & "]" & " trim_code" & j & ": " & CalculatedDsp(i).Element(j) & "<------- Closet to zero" & vbCrLf
                    Else
                        TempLog = TempLog & "site[" & vsite & "]" & " trim_code" & j & ": " & CalculatedDsp(i).Element(j) & vbCrLf
                    End If
                Next j
                TheExec.Datalog.WriteComment TempLog
             Next i
        Next vsite
    End If
    
    'log
    If cPrintFuseBin Then
        For Each vsite In TheExec.sites.Selected
            For i = 0 To UBound(FuseDicKeyArr)
                Result_TempStr = vbNullString
                For j = 0 To FuseBitNum - 1
                    Result_TempStr = Result_TempStr & Result_CodeBinDsp(i).Element(j)
                Next j
                TheExec.Datalog.WriteComment "site[" & vsite & "] " & FuseDicKeyArr(i) & " [LSB to MSB]: " & Result_TempStr
            Next i
        Next vsite
    End If
    
    'test limit
    For i = 0 To UBound(InputDicKeyArr)
       TestNameInput = Report_TName_From_Instance(CalcC, TrimTarget & InputDicKeyArr(i), , , i)
       TheExec.flow.TestLimit resultVal:=Result_Code(i), Tname:=TestNameInput, ForceResults:=tlForceFlow
    Next i

Exit Function 'Add ErrHandler 2023/05/29
errHandler: 'Add ErrHandler 2023/05/29
    Call Print_Error_Message(Error_Info, "LIB_HardIP_Calc", "Calc_EquationTrim") 'Add ErrHandler 2023/05/29
    If AbortTest Then Exit Function Else Resume Next 'Add ErrHandler 2023/05/29
End Function

Public Function Calc_DigCap_Dec_Avg(argc As Integer, argv() As String) As Long
On Error GoTo errHandler 'Add ErrHandler 2023/05/29

    Dim Dictionary_index As Integer
    Dim Dictionart_name() As String
    Dim Store_Dictionary_name As String
    Dim Sum_value_DEC As New SiteDouble
    Dim Temp_Sum_value_DEC As New SiteDouble
    Dim Avg_value_DEC As New SiteDouble
    Dim DSP_Avg_value_DEC_efuse As New DSPWave
    Dim TestNameInput As String
    Dim i As Integer
    
    Dictionart_name = Split(argv(0), "+")
    Dictionary_index = argv(1)
    Store_Dictionary_name = argv(2)
    
    DSP_Avg_value_DEC_efuse.CreateConstant 0, 1, DspLong
    
    
    For i = 0 To UBound(Dictionart_name)
        
        Temp_Sum_value_DEC = GetStoreDataAllType(Dictionart_name(i) & "_para")
        Sum_value_DEC = Sum_value_DEC.Add(Temp_Sum_value_DEC)
    
    Next i
    
    Avg_value_DEC = Sum_value_DEC.divide(Dictionary_index)
    
    Dim FuseDSP As New DSPWave
    
    
    For Each site In TheExec.sites.Active
        DSP_Avg_value_DEC_efuse.Element(0) = Format(Avg_value_DEC(site), 0)
        DSP_Avg_value_DEC_efuse.ConvertDataTypeTo (DspLong)
        
        FuseDSP = DSP_Avg_value_DEC_efuse.ConvertStreamTo(tldspSerial, 8, 0, Bit0IsMsb)
    Next site
    
    TestNameInput = Report_TName_From_Instance("CalcC", vbNullString)
    TheExec.flow.TestLimit resultVal:=DSP_Avg_value_DEC_efuse.Element(0), Tname:=TestNameInput, ForceResults:=tlForceFlow, scaletype:=scaleNoScaling
    
    Call StoreDataAllType(Store_Dictionary_name, FuseDSP)

Exit Function
errHandler:
    Call Print_Error_Message(Error_Info, "LIB_HardIP_Calc", "Calc_DigCap_Dec_Avg") 'Add ErrHandler 2023/05/29
    If AbortTest Then Exit Function Else Resume Next
End Function

Public Function P2PBundle_Unflipeye_NewCal(argc As Integer, argv() As String) As Long
On Error GoTo errHandler 'Add ErrHandler 2023/05/29

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
        Mdll_lockvalue = GetStoreDataAllType(argv(4))
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
            DSPWaveTemp = GetStoreDataAllType(strTemp(j))
            If CLng(SweepConterStr) <> CLng(CounterByStart) Then
                DataSite(j) = GetStoreDataAllType(strTemp(j) & "_" & "AssemblyStr")
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
            Call StoreDataAllType(strTemp(j) & "_" & "AssemblyStr", DataSite(j))
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
                                    Else
                                    End If
                                    Exit For
                                End If
                            Next m
                        Next t
                        
                        
                        
                        TheExec.flow.TestLimit lowVal:=LowLimitValue, hiVal:=HighLimitValue, resultVal:=PrintEye(z + k * CounterByWidth) * EyeDivide, formatStr:="%i", Tname:=TestNameInput, ForceResults:=tlForceNone, scaletype:=scaleNoScaling 'eng_forceflow_transfer
                        '---------------------------------------------------------------------------------------'
                   
                        'TheExec.Flow.TestLimit LowVal:=mdll_low(0), HiVal:=mdll_high(0), resultVal:=PrintEye(z + k * CounterByWidth) * EyeDivide, FormatStr:="%i", TName:=TestNameInput, ForceResults:=tlForceNone, ScaleType:=scaleNoScaling
''''''''''                        TheExec.Flow.TestLimit lowVal:=mdll_low(0), hiVal:=mdll_high(0), resultVal:=PrintEye(z + k * CounterByWidth) * EyeDivide, formatStr:="%i", Tname:=TestNameInput, ForceResults:=tlForceNone, scaletype:=scaleNoScaling
                        
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
                                        Else
                                        End If
                                        Exit For
                                    End If
                                Next m
                            Next t
                        End If
                        
                        TheExec.flow.TestLimit lowVal:=25, hiVal:=100, resultVal:=Format(Eye_precent * 100, 0), formatStr:="%i", Tname:=TestNameInputeye, ForceResults:=tlForceNone, scaletype:=scaleNoScaling 'eng_forceflow_transfer
                        
                        
                        
                        
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
Exit Function 'Add ErrHandler 2023/05/29
errHandler: 'Add ErrHandler 2023/05/29
    Call Print_Error_Message(Error_Info, "LIB_HardIP_Calc", "P2PBundle_Unflipeye_NewCal") 'Add ErrHandler 2023/05/29
    If AbortTest Then Exit Function Else Resume Next 'Add ErrHandler 2023/05/29

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

Check_pinvalue = TheHdw.DCVS.pins(Check_pinname).Voltage.Main.ValuePerSite
TestNameInput = Report_TName_From_Instance(CalcC, Check_pinname, "CheckPowerRail", CInt(0))
TheExec.flow.TestLimit resultVal:=Check_pinvalue, Tname:=TestNameInput, ForceResults:=tlForceFlow

For Each site In TheExec.sites
    If Check_pinvalue > 0.805 And Check_pinvalue < 0.84 Then
        Save_firstDSPvalue.Element(0) = First_fuseName(1)
        Save_SecondDSPvalue.Element(0) = Second_fuseName(1)
    Else
        Save_firstDSPvalue.Element(0) = 0
        Save_SecondDSPvalue.Element(0) = 0

    End If
Next site
    
Call StoreDataAllType(CStr(First_fuseName(0)), Save_firstDSPvalue)
Call StoreDataAllType(CStr(Second_fuseName(0)), Save_SecondDSPvalue)



End Function

Public Function Calc_LPDPRX_RX12_Freq(argc As Integer, argv() As String) As Long
On Error GoTo errHandler 'Add ErrHandler 2023/05/29
                                                                                                                          
    Dim i As Long
    Dim TestNameInput As String
    Dim Dict_Name As String
    Dim site As Variant
    Dim meas_val1 As New PinListData
    Dim meas_val2 As New PinListData
     Dim meas_val3 As New PinListData
    Set meas_val1 = Nothing
    meas_val1 = GetStoreDataAllType(argv(0))
    meas_val2 = GetStoreDataAllType(argv(1))
    meas_val3 = GetStoreDataAllType(argv(2))


    For i = 3 To argc - 1 Step 3
   ' Dict_Name = argv(i)
   ' meas_val1 = meas_val1.Math.Subtract(GetStoreDataAllType(argv(i))).Multiply(80)
   ' meas_val2 = meas_val2.Math.Subtract(GetStoreDataAllType(argv(i + 1))).Multiply(80)
   ' meas_val3 = meas_val3.Math.Subtract(GetStoreDataAllType(argv(i + 2))).Multiply(80)


    meas_val1 = GetStoreDataAllType(argv(i)).Math.Subtract(meas_val1).Multiply(80)
    meas_val2 = GetStoreDataAllType(argv(i + 1)).Math.Subtract(meas_val2).Multiply(80)
    meas_val3 = GetStoreDataAllType(argv(i + 2)).Math.Subtract(meas_val3).Multiply(80)

                                                                                                                          
      TestNameInput = Report_TName_From_Instance(CalcF, vbNullString)
        TheExec.flow.TestLimit resultVal:=meas_val1, Tname:=TestNameInput, ForceResults:=tlForceFlow, scaletype:=scalePico

      TestNameInput = Report_TName_From_Instance(CalcF, vbNullString)
        TheExec.flow.TestLimit resultVal:=meas_val2, Tname:=TestNameInput, ForceResults:=tlForceFlow, scaletype:=scalePico

      TestNameInput = Report_TName_From_Instance(CalcF, vbNullString)
        TheExec.flow.TestLimit resultVal:=meas_val3, Tname:=TestNameInput, ForceResults:=tlForceFlow, scaletype:=scalePico



    meas_val1 = GetStoreDataAllType(argv(i))
    meas_val2 = GetStoreDataAllType(argv(i + 1))
    meas_val3 = GetStoreDataAllType(argv(i + 2))


                                                                                                                          
    'meas_val_now.Pins(i).Value
      Next i
Exit Function 'Add ErrHandler 2023/05/29
errHandler: 'Add ErrHandler 2023/05/29
    Call Print_Error_Message(Error_Info, "LIB_HardIP_Calc", "Calc_LPDPRX_RX12_Freq") 'Add ErrHandler 2023/05/29
    If AbortTest Then Exit Function Else Resume Next 'Add ErrHandler 2023/05/29
End Function


Public Function Calc_RX12_Freq(argc As Integer, argv() As String) As Long
On Error GoTo errHandler 'Add ErrHandler 2023/05/29
                                                                            
    Dim i As Long
    Dim TestNameInput As String
    Dim Dict_Name As String
    Dim site As Variant
    Dim meas_val1 As New PinListData
    Dim meas_val2 As New PinListData
                                                                            
    Set meas_val1 = Nothing
    For i = 2 To argc - 1 Step 2
   ' Dict_Name = argv(i)
    meas_val1 = GetStoreDataAllType(argv(i + 1))
    meas_val2 = GetStoreDataAllType(argv(i))
    
    meas_val2 = meas_val2.Math.Subtract(meas_val1).Multiply(80)
    
        TestNameInput = Report_TName_From_Instance(CalcF, vbNullString)
        TheExec.flow.TestLimit resultVal:=meas_val2.pins(argv(0)), Tname:=TestNameInput, ForceResults:=tlForceFlow, scaletype:=scalePico
        TestNameInput = Report_TName_From_Instance(CalcF, vbNullString)
        TheExec.flow.TestLimit resultVal:=meas_val2.pins(argv(1)), Tname:=TestNameInput, ForceResults:=tlForceFlow, scaletype:=scalePico
    'meas_val_now.Pins(i).Value
      Next i
Exit Function 'Add ErrHandler 2023/05/29
errHandler: 'Add ErrHandler 2023/05/29
    Call Print_Error_Message(Error_Info, "LIB_HardIP_Calc", "Calc_RX12_Freq") 'Add ErrHandler 2023/05/29
    If AbortTest Then Exit Function Else Resume Next 'Add ErrHandler 2023/05/29

End Function

Public Function Calc_DigCap_Offset_Store_old(argc As Integer, argv() As String) As Long
        
    Dim i As Long
    Dim lVoltage_Low As Double
    Dim lVoltage_High As Double
    Dim lYintercept_spec As Double
    Dim DSPWave_LowBin As New DSPWave
    Dim DSPWave_HighBin As New DSPWave
    Dim DSPWave_OffsetBin As New DSPWave
    
    Dim DSPWave_LowDec As New DSPWave
    Dim DSPWave_HighDec As New DSPWave
    Dim DSPWave_SlopeDec As New DSPWave
    Dim DSPWave_YaxisBDec As New DSPWave
    Dim DSPWave_InterpolateDec As New DSPWave
    Dim DSPWave_OffsetDec As New DSPWave
    
    Dim SL_BitWidth As New SiteLong
    
    Dim site As Variant
    Dim TestNameInput As String


    ''''argv(1) = aneivdm_trim_low1
    ''''argv(2) = 0.75
    ''''argv(3) = aneivdm_trim_high1
    ''''argv(4) = 1.1
    ''''argv(5) = Y_intercept_spec
    ''''argv(argc-1) = aneivdm_trim_offset1
    ''''------example------
    ''''y1 = a * x1 + b
    ''''(x1,y1) = (0.75,aneivdm_trim_low1) and (x2,y2) = (1.1,aneivdm_trim_high1)
    ''''a = (y1 - y2) / (x1 - x2)
    ''''b = y1 - a * x1 = Y_intercept_measured
    ''''Offset = Y_intercept_spec - b
    ''''------example------

    DSPWave_LowBin = GetStoreDataAllType(argv(0))
    lVoltage_Low = CDbl(argv(1))
    DSPWave_HighBin = GetStoreDataAllType(argv(2))
    lVoltage_High = CDbl(argv(3))
    lYintercept_spec = CDbl(argv(4))
    
    ''20220115 CW add to print lYintercept_spec in datalogs
    TheExec.Datalog.WriteComment "***IVDM-Yintercept_spec = " & lYintercept_spec
    
    Call rundsp.BinToDec(DSPWave_LowBin, DSPWave_LowDec)
    Call rundsp.BinToDec(DSPWave_HighBin, DSPWave_HighDec)


    DSPWave_SlopeDec.CreateConstant 0, 1, DspDouble
    DSPWave_YaxisBDec.CreateConstant 0, 1, DspDouble
    DSPWave_InterpolateDec.CreateConstant 0, 1, DspDouble
    DSPWave_OffsetDec.CreateConstant 0, 1, DspDouble
    For Each site In TheExec.sites
        SL_BitWidth = DSPWave_LowBin.SampleSize
        DSPWave_SlopeDec = DSPWave_HighDec.Subtract(DSPWave_LowDec).divide(lVoltage_High - lVoltage_Low)
        DSPWave_YaxisBDec = DSPWave_LowDec.Subtract(DSPWave_SlopeDec.Multiply(lVoltage_Low))
        DSPWave_InterpolateDec = DSPWave_YaxisBDec.Add(DSPWave_SlopeDec.Multiply(0.95))
        DSPWave_OffsetDec = DSPWave_InterpolateDec.Negate.Add(lYintercept_spec)
    Next site
    
    For Each site In TheExec.sites
        TheExec.Datalog.WriteComment "Site : " & site & ", Offset result:" & DSPWave_OffsetDec(site).Element(0)
        DSPWave_OffsetDec(site).Element(0) = FormatNumber(DSPWave_OffsetDec(site).Element(0), 0)
    Next site

    'Call rundsp.DSPWf_Dec2Binary(DSPWave_OffsetDec, SL_BitWidth, DSPWave_OffsetBin)
    'Call StoreDataAllType(argv(argc - 1), DSPWave_AverageDec)
    TestNameInput = Report_TName_From_Instance(CalcC, "X", , CInt(i), , , , , tlForceFlow)

    TheExec.flow.TestLimit resultVal:=DSPWave_OffsetDec.Element(0), Tname:=TestNameInput, ForceResults:=tlForceFlow  'tlForceFlow
    
    For Each site In TheExec.sites
        If DSPWave_OffsetDec(site).Element(0) < 0 Then
            DSPWave_OffsetDec = DSPWave_OffsetDec.Subtract(128).Abs
        End If
    Next site
    
    Call rundsp.DSPWf_Dec2Binary(DSPWave_OffsetDec, SL_BitWidth, DSPWave_OffsetBin)
    Call StoreDataAllType(argv(argc - 1), DSPWave_OffsetBin)




End Function

Public Function Calc_DigCap_Offset_Store(argc As Integer, argv() As String) As Long
On Error GoTo errHandler 'Add ErrHandler 2023/05/29
        'Update for IVDM function TTR (DSP process from local to DSPPC )-- 20220105
    Dim i As Long
    Dim lVoltage_Low As Double
    Dim lVoltage_High As Double
    Dim lYintercept_spec As Double
    Dim DSPWave_LowBin As New DSPWave
    Dim DSPWave_HighBin As New DSPWave
    Dim DSPWave_OffsetBin As New DSPWave
    
    Dim DSPWave_LowDec As New DSPWave
    Dim DSPWave_HighDec As New DSPWave
    Dim DSPWave_SlopeDec As New DSPWave
    Dim DSPWave_YaxisBDec As New DSPWave
    Dim DSPWave_InterpolateDec As New DSPWave
    Dim DSPWave_OffsetDec As New DSPWave

    Dim site As Variant
    Dim TestNameInput As String
    Dim Sample_Size As Long: Sample_Size = 0
    
    Dim serial_data As Long
    Dim Input_Arry() As String
    Dim sLen As Integer
    Dim Pasing_data_Arry() As String
    Dim Input_Str As String
    Input_Str = vbNullString
    For i = 0 To argc - 1
        If i <> argc - 1 Then
            Input_Str = Input_Str & argv(i) & ","
        Else
            Input_Str = Input_Str & argv(i)
        End If
    Next i
    Input_Arry = Split(Input_Str, "@")
    
    TheHdw.dsp.ExecutionMode = tlDSPModeAutomatic
    
    ''''argv(1) = aneivdm_trim_low1
    ''''argv(2) = 0.75
    ''''argv(3) = aneivdm_trim_high1
    ''''argv(4) = 1.1
    ''''argv(5) = Y_intercept_spec
    ''''argv(argc-1) = aneivdm_trim_offset1
    ''''------example------
    ''''y1 = a * x1 + b
    ''''(x1,y1) = (0.75,aneivdm_trim_low1) and (x2,y2) = (1.1,aneivdm_trim_high1)
    ''''a = (y1 - y2) / (x1 - x2)
    ''''b = y1 - a * x1 = Y_intercept_measured
    ''''Offset = Y_intercept_spec - b
    ''''------example------
    DSPWave_SlopeDec.CreateConstant 0, 1, DspDouble
    DSPWave_YaxisBDec.CreateConstant 0, 1, DspDouble
    DSPWave_InterpolateDec.CreateConstant 0, 1, DspDouble
    DSPWave_OffsetDec.CreateConstant 0, 1, DspDouble
    DSPWave_LowBin.CreateConstant 0, 8, DspLong
    DSPWave_HighBin.CreateConstant 0, 8, DspLong
    DSPWave_OffsetBin.CreateConstant 0, 8, DspLong

    For serial_data = 0 To UBound(Input_Arry)
    
        Pasing_data_Arry = Split(Input_Arry(serial_data), ",")
        sLen = UBound(Pasing_data_Arry)
       
        DSPWave_LowBin = GetStoreDataAllType(Pasing_data_Arry(0))
        lVoltage_Low = CDbl(Pasing_data_Arry(1))
        DSPWave_HighBin = GetStoreDataAllType(Pasing_data_Arry(2))
        lVoltage_High = CDbl(Pasing_data_Arry(3))
        lYintercept_spec = CDbl(Pasing_data_Arry(4))
        
        'Update for IVDM function TTR (DSP process from local to DSPPC )
        TheHdw.dsp.ExecutionMode = tlDSPModeForceAutomatic

        Call rundsp.DSPWF_OFFSET(DSPWave_LowBin, lVoltage_Low, DSPWave_HighBin, lVoltage_High, lYintercept_spec, DSPWave_SlopeDec, DSPWave_YaxisBDec, _
                                 DSPWave_InterpolateDec, DSPWave_OffsetDec, DSPWave_OffsetBin, True)

        For Each site In TheExec.sites
            TheExec.Datalog.WriteComment "Site : " & site & ", Offset result:" & DSPWave_OffsetDec(site).Element(0) & ", Convert to 2S Complement"
            DSPWave_OffsetDec(site).Element(0) = FormatNumber(DSPWave_OffsetDec(site).Element(0), 0)
        Next site
    
        TestNameInput = Report_TName_From_Instance(CalcC, "X", , , , , , , tlForceFlow)
        TheExec.flow.TestLimit resultVal:=DSPWave_OffsetDec.Element(0), Tname:=TestNameInput, ForceResults:=tlForceFlow  'tlForceFlow
        
        'TestNameInput = Report_TName_From_Instance(CalcC, "X", , , , , , , tlForceFlow)
        'TheExec.Flow.TestLimit resultVal:=DSPWave_SlopeDec.Element(0), Tname:=TestNameInput, ForceResults:=tlForceFlow  'tlForceFlow
        
        'TestNameInput = Report_TName_From_Instance(CalcC, "X", , , , , , , tlForceFlow)
        'TheExec.Flow.TestLimit resultVal:=DSPWave_InterpolateDec.Element(0), Tname:=TestNameInput, ForceResults:=tlForceFlow  'tlForceFlow
        
        Call StoreDataAllType(Pasing_data_Arry(sLen), DSPWave_OffsetBin)
    Next serial_data
Exit Function 'Add ErrHandler 2023/05/29
errHandler: 'Add ErrHandler 2023/05/29
    Call Print_Error_Message(Error_Info, "LIB_HardIP_Calc", "Calc_DigCap_Offset_Store") 'Add ErrHandler 2023/05/29
    If AbortTest Then Exit Function Else Resume Next 'Add ErrHandler 2023/05/29
End Function

Public Function Calc_MDLL_Monotonicity_DevideBlock_SEGTTR(argc As Integer, argv() As String) As Long
On Error GoTo errHandler 'Add ErrHandler 2023/05/29
'Update from Staten @William 211101
'TTR,20200423, Oscar
    On Error GoTo errHandler
    
''    Dim LIB_HardIP_Calc_ProfileMark_12721 As Long: LIB_HardIP_Calc_ProfileMark_12721 = ProfileMarkEnter(2, "Calc_MDLL_Monotonicity_DevideBlock_SEGTTR-1")    ' Profile Mark
        
    Dim funcName As String:: funcName = "Calc_MDLL_Monotonicity_DevideBlock_SEGTTR"
''    Exit Function
    Dim i As Long, j As Long
''    Dim l_LimitIndex As Long
    Dim InputKey As String
    Dim TestNameInput As String
    
    Dim site As Variant
    
    Dim DSP_Input As New DSPWave
    Dim ConcatenateDSP_BeforeSort As New DSPWave
    Dim ConcatenateDSP_AfterSort As New DSPWave
''    l_LimitIndex = TheExec.Flow.TestLimitIndex
''Calc_MDLL_MONOTONICITY_DEVIDEBLOCK_SEG(
''GRP13-----CH0-DQ0-STS-SDLL-CODE-F1,
''GRP13-----CH0-DQ0-STS-SDLL-CODE-GRP1---w210,
''GRP13-----CH0-DQ0-STS-SDLL-CODE-GRP2---w543,
''GRP13-----CH0-DQ0-STS-SDLL-CODE-GRP3---w76,
''GRP13-----CH0-DQ1-STS-SDLL-CODE-F1,
''GRP13-----CH0-DQ1-STS-SDLL-CODE-GRP1---w210,
''GRP13-----CH0-DQ1-STS-SDLL-CODE-GRP2---w543,
''GRP13-----CH0-DQ1-STS-SDLL-CODE-GRP3---w76)
    ConcatenateDSP_BeforeSort.CreateConstant 0, argc - (argc \ 4), DspLong
''    ConcatenateDSP_BeforeSort.CreateConstant 0, (argc * 2 \ 4) + 1, DspLong
    Dim SL_MDLL As New SiteLong
    Dim MDLLindex As Long: MDLLindex = 0
    For i = 0 To argc - 1
        If i Mod 4 <> 0 Then
            InputKey = LCase(argv(i))
''            Set DSP_Input = Nothing
''            DSP_Input = GetStoreDataAllType(InputKey)
            
            SL_MDLL = GetStoreDataAllType(InputKey & "_para") ''Decimal with 32 bits
            If TheExec.TesterMode = testModeOffline Then
                If i = 1 Then SL_MDLL = 256
                If i = 2 Then SL_MDLL = 1024
                If i = 3 Then SL_MDLL = 16777216
                If i = 5 Then SL_MDLL = 65536
                If i = 6 Then SL_MDLL = 4096
                If i = 7 Then SL_MDLL = 256
            End If

            For Each site In TheExec.sites
                ConcatenateDSP_BeforeSort.ElementLite(MDLLindex) = SL_MDLL ''''Decimal with 96 bits (32 bits * 3 GRP)
            Next site
            MDLLindex = MDLLindex + 1
        End If
     Next i
''ProfileMarkLeave LIB_HardIP_Calc_ProfileMark_12721    ' Profile Mark
''Dim LIB_HardIP_Calc_ProfileMark_12715 As Long: LIB_HardIP_Calc_ProfileMark_12715 = ProfileMarkEnter(2, "Calc_MDLL_Monotonicity_DevideBlock_SEGTTR-2")    ' Profile Mark
    
    Dim sl_Temp_Sum As New DSPWave
    Dim sl_Sum_Val As New DSPWave
    Dim sl_Mean_Val As New DSPWave
    Dim sl_Diff_MaxMin As New DSPWave
    Dim sl_MDLL_UniqueDirection As New DSPWave
    Dim sl_MDLL_DecreaseDirection As New DSPWave
    
    TheHdw.dsp.ExecutionMode = tlDSPModeForceAutomatic
    
    sl_Temp_Sum.CreateConstant 0, 1, DspLong
    sl_Diff_MaxMin.CreateConstant 0, argc \ 4
    sl_MDLL_UniqueDirection.CreateConstant 1, argc \ 4
    sl_MDLL_DecreaseDirection.CreateConstant 1, argc \ 4
    sl_Sum_Val.CreateConstant 1, argc \ 4
    sl_Mean_Val.CreateConstant 1, argc \ 4
''ProfileMarkLeave LIB_HardIP_Calc_ProfileMark_12715    ' Profile Mark
''Dim LIB_HardIP_Calc_ProfileMark_12717 As Long: LIB_HardIP_Calc_ProfileMark_12717 = ProfileMarkEnter(2, "Calc_MDLL_Monotonicity_DevideBlock_SEGTTR-3")    ' Profile Mark

    Call rundsp.DSP_CalcMDLLMonotonicityDevideBlockSEG(ConcatenateDSP_BeforeSort, ConcatenateDSP_AfterSort, sl_Diff_MaxMin, sl_MDLL_DecreaseDirection, sl_Mean_Val, sl_MDLL_UniqueDirection, sl_Sum_Val)
    
''ProfileMarkLeave LIB_HardIP_Calc_ProfileMark_12717    ' Profile Mark

''Dim LIB_HardIP_Calc_ProfileMark_12745 As Long: LIB_HardIP_Calc_ProfileMark_12745 = ProfileMarkEnter(2, "Calc_MDLL_Monotonicity_DevideBlock_SEGTTR-4")    ' Profile Mark
    Dim TempResult As New SiteLong
    Dim SD_tempResult As New SiteDouble
    Dim temp(0) As Long
    For j = 0 To argc \ 4 - 1
    
        TempResult = sl_Mean_Val.Element(j)
        TestNameInput = Report_TName_From_Instance(CalcC, "X", , 0, 0, ForceResult:=tlForceFlow)
        TheExec.flow.TestLimit resultVal:=TempResult, Tname:=TestNameInput, ForceResults:=tlForceFlow, PinName:="X"


        TempResult = sl_Diff_MaxMin.Element(j)
        TestNameInput = Report_TName_From_Instance(CalcC, "X", , 0, 0, ForceResult:=tlForceFlow)
        TheExec.flow.TestLimit resultVal:=TempResult, Tname:=TestNameInput, ForceResults:=tlForceFlow, PinName:="X"

        TempResult = sl_MDLL_UniqueDirection.Element(j)
        TestNameInput = Report_TName_From_Instance(CalcC, "X", , 0, 0, ForceResult:=tlForceFlow)
        TheExec.flow.TestLimit resultVal:=TempResult, Tname:=TestNameInput, ForceResults:=tlForceFlow, PinName:="X"


        TempResult = sl_MDLL_DecreaseDirection.Element(j)
        TestNameInput = Report_TName_From_Instance(CalcC, "X", , 0, 0, ForceResult:=tlForceFlow)
        TheExec.flow.TestLimit resultVal:=TempResult, Tname:=TestNameInput, ForceResults:=tlForceFlow, PinName:="X"


        For Each site In TheExec.sites
            temp(0) = sl_Sum_Val.Element(j)
            sl_Temp_Sum.data = temp
        Next site
        Call StoreDataAllType("Summary" & argv(j * 4), sl_Temp_Sum)

'        Call StoreDataAllType("Summary" & argv(j * 4), sl_Sum_Val.Element(j))
    Next j
    
    
''ProfileMarkLeave LIB_HardIP_Calc_ProfileMark_12745    ' Profile Mark

Exit Function

errHandler:
    Call Print_Error_Message(Error_Info, "LIB_HardIP_Calc", "Calc_MDLL_Monotonicity_DevideBlock_SEGTTR") 'Add ErrHandler 2023/05/29
    If AbortTest Then Exit Function Else Resume Next
End Function

Public Function Calc_EyeWidthLimitForEachMode_TTR(argc As Integer, argv() As String) As Long
On Error GoTo errHandler 'Add ErrHandler 2023/05/29
    Dim i As Integer
    Dim site As Variant
    Dim TestNameInput As String
    Dim SplitStrAry() As String
    Dim PercentageTemp() As New SiteDouble
    Dim SiteDouble_Dcode As New SiteDouble
    Dim SiteDouble_EyeWidth As New SiteDouble
    
    
    ReDim PercentageTemp(argc - 1)
    
    
    For i = 0 To argc - 1
        SplitStrAry = Split(argv(i), "&")
        SiteDouble_Dcode = GetStoreDataAllType(SplitStrAry(0) & "_para")
        SiteDouble_EyeWidth = GetStoreDataAllType(SplitStrAry(1) & "_para")
        
        '------------ For avoid Capture 0 error --------------
        For Each site In TheExec.sites
            If SiteDouble_Dcode = 0 Then
               SiteDouble_Dcode = 0.01
               TheExec.Datalog.WriteComment SplitStrAry(0) & " Capture 0 so give defult value"
            End If
            If SiteDouble_EyeWidth = 0 Then
               SiteDouble_EyeWidth = 0.01
               TheExec.Datalog.WriteComment SplitStrAry(1) & " Capture 0 so give defult value"
            End If
        Next site
        '-----------------------------------------------
        
        '@210707 CW TTR eye_width
        '----------------------------------
'        TestNameInput = Report_TName_From_Instance("C", "", ForceResult:=tlForceFlow)
'        For Each Site In TheExec.sites
'            If TheExec.Flow.EnableWord("AMPLP5_BinCut_Enable_Flag") = True Then
'                TheExec.Flow.TestLimit SiteDouble_EyeWidth, , , , , , , , PinName:="X", Tname:=TestNameInput, ForceResults:=tlForceFlow
'            Else
'                TheExec.Flow.TestLimit SiteDouble_EyeWidth, SiteDouble_Dcode / 8, SiteDouble_Dcode, , , , , , PinName:="X", Tname:=TestNameInput, ForceResults:=tlForceFlow
'            End If
'            TheExec.Flow.TestLimitIndex = TheExec.Flow.TestLimitIndex - 1
'        Next Site
'        TheExec.Flow.TestLimitIndex = TheExec.Flow.TestLimitIndex + 1
        '----------------------------------
        
        PercentageTemp(i) = SiteDouble_EyeWidth.divide(SiteDouble_Dcode).Multiply(200)
    Next i
    
    
    For i = 0 To argc - 1
        SplitStrAry = Split(argv(i), "&")
        'TestNameInput = Report_TName_From_Instance("C", SplitStrAry(1), "Percentage" & i, i, , , , , tlForceNone)
        TestNameInput = Report_TName_From_Instance("C", "X", , , , , , , ForceResult:=tlForceFlow)
        TheExec.flow.TestLimit PercentageTemp(i), 25, 100, , , , , , PinName:="X", Tname:=TestNameInput, ForceResults:=tlForceFlow
    Next i
Exit Function 'Add ErrHandler 2023/05/29
errHandler: 'Add ErrHandler 2023/05/29
    Call Print_Error_Message(Error_Info, "LIB_HardIP_Calc", "Calc_EyeWidthLimitForEachMode_TTR") 'Add ErrHandler 2023/05/29
    If AbortTest Then Exit Function Else Resume Next 'Add ErrHandler 2023/05/29

End Function

Public Function Calc_ADC2AnalogVoltage(argc As Integer, argv() As String) As Long '0428wade
    Dim Str_Vref As String
    Dim Str_Mode As String
    Dim SiteDbl_Vref As New SiteDouble
    Dim Voltage As Double
    Dim Analog_Voltage() As New SiteDouble
    ReDim Analog_Voltage(((argc - 1) / 2) - 1) As New SiteDouble
    'argv()= Dict1_cal,mode1,Dict2_cal,mode2...,Dictk_cal,mode k
    Dim i As Integer: i = 0 'Dict counter
    Dim k As Integer: k = 0 'Total number of Tname have *_cal
    Dim TestNameInput As String

    Voltage = TheHdw.DCVS.pins(argv(0)).Voltage.value

      For i = 1 To UBound(argv) - 1 Step 2
        Str_Vref = argv(i)                'Get Dict_XXXX from HardIP calc function
        'SiteDbl_Vref = GetStoreDataAllType(Str_Vref & "_para")
        SiteDbl_Vref = GetStoreDataAllType(Str_Vref & "_para")
        If TheExec.TesterMode = testModeOffline Then SiteDbl_Vref = 230
        Str_Mode = argv(i + 1)            'Get Dict_XXXX "mode" from HardIP calc function
            For Each site In TheExec.sites
                If UCase(Str_Mode) = "LOW" Then
                    Analog_Voltage(k) = (SiteDbl_Vref.Subtract(128).divide(256).Multiply(Voltage * 0.5)) + Voltage * 0.25
                ElseIf UCase(Str_Mode) = "MID" Then
                    Analog_Voltage(k) = (SiteDbl_Vref.Subtract(128).divide(256).Multiply(Voltage * 0.5)) + Voltage * 0.5
                ElseIf UCase(Str_Mode) = "HIGH" Then
                    Analog_Voltage(k) = (SiteDbl_Vref.Subtract(128).divide(256).Multiply(Voltage * 0.5)) + Voltage * 0.75
                End If
            Next site

            k = k + 1
      Next i

    For k = 0 To (argc - 1) / 2 - 1 '@220101 updated by Walker
'    For k = 0 To (argc - 1 / 2) - 1
        TestNameInput = Report_TName_From_Instance("CalcV", "", , 0) 'Tname:=FlowTname(k),TestSeqNum:=k,ForceResult:=tlForceNone
        TheExec.flow.TestLimit resultVal:=Analog_Voltage(k), Tname:=TestNameInput, PinName:="", ForceResults:=tlForceFlow
    Next k

Exit Function
errHandler:
    TheExec.Datalog.WriteComment "error in Calc_ADC2AnalogVoltage"
    If AbortTest Then Exit Function Else Resume Next
End Function
Public Function Calc_ADC_Voltage2Code(argc As Integer, argv() As String) As Long
    '---- New Function : Use for SEPF2M --
    'Alg::Calc_ADC_Voltage2Code(
    'Alg::Calc_ADC_Voltage2Code(vref_pre_trim,trim_ref_ate,8)
    'Formula : Code = 63-4((7*vref_pre_trim)/(1.2373-2*vref_pre_trim)-7)

    Dim VREF_Pre_Trim_Value As New SiteDouble
    Dim VREF_Pre_Trim_Value_ENG As New SiteDouble
    
    Dim OUTPUT_DICT_NAME As String
    Dim CTRL_REF_DICT_NAME As String
    Dim CODE_DEC As New SiteDouble
    Dim DSPWF_CODE_DEC As New DSPWave
    DSPWF_CODE_DEC.CreateConstant 0, 1, DspDouble
    
    Dim DSPWF_BIT_SIZE As Long
    Dim DSPWF_CTRL_BIT_SIZE As Long
    Dim DSPWF_CODE_BIN As New DSPWave
    Dim ctrl_ref_BIN As New DSPWave
    Dim vsite As Variant
    Dim TestNameInput As String
    Dim ENG_String As String
    'Dim DSPWave_Dict As New DSPWave
    Dim k As Double
    Dim ctrl_ref As Double
    Dim vest As Double
    Dim p1 As Double
    Dim ctrl_ref_dec As New DSPWave
    ctrl_ref_dec.CreateConstant 0, 1, DspDouble
    On Error GoTo errHandler
    
    If TheExec.TesterMode = testModeOffline Then
        For Each vsite In TheExec.sites
            VREF_Pre_Trim_Value_ENG = 0.5
        Next vsite
        ENG_String = argv(0)
        Call StoreDataAllType(ENG_String, VREF_Pre_Trim_Value_ENG)
     End If
    
    VREF_Pre_Trim_Value = GetStoreDataAllType(argv(0))
    OUTPUT_DICT_NAME = argv(1)
    CTRL_REF_DICT_NAME = argv(2)
    DSPWF_BIT_SIZE = CDbl(argv(3))
    DSPWF_CTRL_BIT_SIZE = CDbl(argv(4))
    
    'Formula : Code = 63-4((7*vref_pre_trim)/(1.2373-2*vref_pre_trim)-7)
    ''' Update for Hidra. @William 240223
    For Each vsite In TheExec.sites
       'CODE_DEC(vsite) = 63 - 4 * (((7 * VREF_Pre_Trim_Value(vsite)) / (1.2373 - (2 * VREF_Pre_Trim_Value(vsite)))) - 7)
       'Coll test plan #102 220627
'       CODE_DEC(vsite) = 63 - 4 * (((11 * VREF_Pre_Trim_Value(vsite)) / (1.3729 - (2 * VREF_Pre_Trim_Value(vsite)))) - 7)
       
'       k = Int((40.5 * (1 - 0.485 / VREF_Pre_Trim_Value(vsite))) + 0.5)
'       k = Int((21.75 * (1 - 0.485 / VREF_Pre_Trim_Value(vsite))) + 0.5) '20230116 Alfred modify for Donan new operation
        k = Int((21.75 * (1 - 0.34758 / VREF_Pre_Trim_Value(vsite))) + 0.5) 'Update for Donan V09A Change List @William 230320
        If k < 0 Then
        
            TheExec.Datalog.WriteComment ("site" & vsite & " k=" & k & "  ,then set k = 0")
            k = 0
            
        ElseIf k > 7 Then
        
            TheExec.Datalog.WriteComment ("site" & vsite & " k=" & k & "  ,then set k = 7")
            k = 7
        
        End If
        
       ctrl_ref = ((7 - k) * 8) + 3
       
'       vest = VREF_Pre_Trim_Value(vsite) - (VREF_Pre_Trim_Value(vsite) * k / 40.5)
'       vest = VREF_Pre_Trim_Value(vsite) - (VREF_Pre_Trim_Value(vsite) * k / 21.75) '20230116 Alfred modify for Donan new operation
       vest = (87 / 63) * (VREF_Pre_Trim_Value(vsite) - (VREF_Pre_Trim_Value(vsite) * k / 21.75)) 'Update for Donan V09A Change List @William 230320

'       p1 = ((7 - k) / 2) + 2
       p1 = 7 - k '20230116 Alfred modify for Donan new operation
       
'       CODE_DEC(vsite) = 63 - 4 * (((2 * p1 * vest) / (1 + (0.06779 * p1) - (2 * vest))) - 7)
       CODE_DEC(vsite) = 63 - 4 * (((2.0835 * p1 * vest) / (1 + (0.06779 * p1) - (2.0835 * vest))) - 7) 'Update for Donan V09A Change List @William 230320

       
       If CODE_DEC(vsite) < 0 Then
            CODE_DEC(vsite) = 0
            TheExec.Datalog.WriteComment ("trim value < 0")
       ElseIf CODE_DEC(vsite) > 63 Then
            CODE_DEC(vsite) = 63
            TheExec.Datalog.WriteComment ("trim value > 63")
       End If
       
       CODE_DEC(vsite) = Int(CODE_DEC(vsite) + 0.5)
       
       DSPWF_CODE_DEC.Element(0) = CODE_DEC(vsite)
       ctrl_ref_dec.Element(0) = ctrl_ref
    
    TheExec.Datalog.WriteComment ("site" & vsite & " k=" & k & ", ctrl_ref =" & ctrl_ref & ", vest =" & vest & ", p1 =" & p1 & " trim_ref =" & CODE_DEC(vsite))
     
    
    Next vsite
    
    'Print out datalog
    TestNameInput = Report_TName_From_Instance(CalcC, "X", , , , , , , tlForceFlow)
    TheExec.flow.TestLimit resultVal:=DSPWF_CODE_DEC.Element(0), PinName:="X", Tname:=TestNameInput, scaletype:=scaleNone, ForceResults:=tlForceFlow
    TestNameInput = Report_TName_From_Instance(CalcC, "X", , , , , , , tlForceFlow)
    TheExec.flow.TestLimit resultVal:=ctrl_ref_dec.Element(0), PinName:="X", Tname:=TestNameInput, scaletype:=scaleNone, ForceResults:=tlForceFlow
    
    Call HardIP_Dec2Bin(DSPWF_CODE_BIN, DSPWF_CODE_DEC, DSPWF_BIT_SIZE)
    Call HardIP_Dec2Bin(ctrl_ref_BIN, ctrl_ref_dec, DSPWF_CTRL_BIT_SIZE)
    Call StoreDataAllType(LCase(OUTPUT_DICT_NAME), DSPWF_CODE_BIN)
    Call StoreDataAllType(LCase(CTRL_REF_DICT_NAME), ctrl_ref_BIN)
   'DSPWave_Dict = GetStoreDataAllType(LCase(OUTPUT_DICT_NAME))
   
    Exit Function

errHandler:
    TheExec.Datalog.WriteComment "error in Calc_ADC_Voltage2Code"
    If AbortTest Then Exit Function Else Resume Next
End Function
Public Function Calc_UCSDM(argc As Integer, argv() As String) As Long
On Error GoTo errHandler 'Add ErrHandler 2023/05/29

''Calc_UCSDM(DOUT1_avg,DOUT2_avg,DOUT3_avg)
    Dim i As Long
    
    Dim D1_ideal As Double
    Dim D2_ideal As Double
    Dim D3_ideal As Double
    Dim D_ideal() As Double
    Dim DOUT_avg() As New SiteDouble
    Dim DOUT1_avg As New SiteDouble
    Dim DOUT2_avg As New SiteDouble
    Dim DOUT3_avg As New SiteDouble
    
    Dim DSP_DOUT_avg() As New DSPWave
    
    Dim dConstant As Double
    Dim dAccuracy As Double
    Dim dMidpoint As Double
    ReDim D_ideal(argc - 2)
    ReDim DOUT_avg(argc - 2)
    ReDim DSP_DOUT_avg(argc - 2)
    
    dAccuracy = 0
    dConstant = 12 - 2 * dAccuracy
    dMidpoint = ((2 ^ dConstant) - 2) / 2
    D_ideal(0) = (0.325 / 1.3) * ((2 ^ dConstant) - 2)
    D_ideal(1) = (0.65 / 1.3) * ((2 ^ dConstant) - 2)
    D_ideal(2) = (0.975 / 1.3) * ((2 ^ dConstant) - 2)
''    D1_ideal = (0.05 / 1.3) * ((2 ^ dconstant) - 2)
''    D2_ideal = (0.5 / 1.3) * ((2 ^ dconstant) - 2)
''    D3_ideal = (1.25 / 1.3) * ((2 ^ dconstant) - 2)
    
    For i = 0 To argc - 2
''        DOUT_avg(i) = GetStoreDataAllType(argv(i))
        DSP_DOUT_avg(i) = GetStoreDataAllType(argv(i))
        DOUT_avg(i) = DSP_DOUT_avg(i).Element(0)
    Next i
    
    Dim site As Variant
    Dim vresult() As Variant
    Dim svresult() As New SiteVariant
    
    Dim dGainval As New SiteDouble
    Dim dOffsetval As New SiteDouble
    Dim dCalcOffsetval As New SiteDouble
    Dim dGainTrim As New SiteDouble
    Dim dOffsetTrim As New SiteDouble
    
    Dim dTemp() As Double
    ReDim dTemp(argc - 2)
    
    
    For Each site In TheExec.sites
        If TheExec.TesterMode = testModeOffline Then
            dTemp(0) = 162.5
            dTemp(1) = 2032.375
            dTemp(2) = 3952
        Else
            For i = 0 To UBound(DOUT_avg)
                dTemp(i) = DOUT_avg(i)
            Next i
        End If
        vresult = Application.WorksheetFunction.LinEst(dTemp, D_ideal)
        dGainval = Application.WorksheetFunction.index(vresult, 1) ''   CDbl(vresult(1))
        dOffsetval = Application.WorksheetFunction.index(vresult, 2) ''CDbl(vresult(2))
        
        If site = 0 Then dGainval(site) = 0
        If site = 2 Then dGainval(site) = 0
        
        'Add Error message when dGainval = 0 -- 20220525
        If dGainval(site) = 0 Then
            dGainval = dGainval.Maximum(0.000000001)
            TheExec.Datalog.WriteComment "*** Error : " & "Site[" & site & "]" & " dGainval = 0 !! The Calc_UCSDM will get Infinity value!!***"
        End If
    Next site
        
    dCalcOffsetval = dGainval.Multiply(dMidpoint).Add(dOffsetval).Subtract(dMidpoint)
    dGainTrim = dGainval.Invert.Multiply(1).Negate.Add(1).Multiply(2 ^ 11).Add(2 ^ 6)
    For Each site In TheExec.sites
        dGainTrim = FormatNumber(dGainTrim(site), 0)
    Next site
    dOffsetTrim = dCalcOffsetval.Truncate.Add(1).Add(2 ^ 6)
    
    Dim DSPGainTrim As New DSPWave
    Dim DSPOffsetTrim As New DSPWave
    
    Dim DSPOUTPUTGainTrim As New DSPWave
    Dim DSPOUTPUTOffsetTrim As New DSPWave
    DSPGainTrim.CreateConstant 0, 1, DspDouble
    DSPOffsetTrim.CreateConstant 0, 1, DspDouble
    For Each site In TheExec.sites
        DSPGainTrim.Element(0) = dGainTrim
        DSPOffsetTrim.Element(0) = dOffsetTrim
    Next site
    
    Dim TestNameInput As String
    TestNameInput = Report_TName_From_Instance("CalcC", "X", , 0, 0)
    TheExec.flow.TestLimit resultVal:=DSPGainTrim.Element(0), Tname:=TestNameInput, ForceResults:=tlForceFlow, scaletype:=scaleNoScaling, formatStr:="%.3f"
    TestNameInput = Report_TName_From_Instance("CalcC", "X", , 0, 0)
    TheExec.flow.TestLimit resultVal:=DSPOffsetTrim.Element(0), Tname:=TestNameInput, ForceResults:=tlForceFlow, scaletype:=scaleNoScaling, formatStr:="%.3f"
    
    For Each site In TheExec.sites
        If (DSPGainTrim.Element(0) > 127 Or DSPGainTrim.Element(0) < 0) Then DSPGainTrim.Element(0) = 0
        If (DSPOffsetTrim.Element(0) > 127 Or DSPOffsetTrim.Element(0) < 0) Then DSPOffsetTrim.Element(0) = 0
    Next site
   
    Call rundsp.DSPWf_Dec2Binary(DSPGainTrim, 7, DSPOUTPUTGainTrim)
    Call rundsp.DSPWf_Dec2Binary(DSPOffsetTrim, 7, DSPOUTPUTOffsetTrim)
    
    Call StoreDataAllType(argv(UBound(argv)) & "_Dict_GAIN_TRIM", DSPOUTPUTGainTrim)
    Call StoreDataAllType(argv(UBound(argv)) & "_Dict_OFST_TRIM", DSPOUTPUTOffsetTrim)
Exit Function 'Add ErrHandler 2023/05/29
errHandler: 'Add ErrHandler 2023/05/29
    Call Print_Error_Message(Error_Info, "LIB_HardIP_Calc", "Calc_UCSDM") 'Add ErrHandler 2023/05/29
    If AbortTest Then Exit Function Else Resume Next 'Add ErrHandler 2023/05/29
End Function
Public Function Calc_Current_Beta(argc As Integer, argv() As String) As Long
    
    ''''argv(0) = IB2 'GetDictName
    ''''argv(1) = MTR_TD_B 'PinName
    ''''argv(2) = 0.00024
    ''''argv(3) = BETA2 'Dict_Store_Name for eFuse
    
    Dim l_Divide As Double
    Dim pl_MeasureDict As New PinListData
    Dim sd_MeasurePin As New SiteDouble
    Dim sd_CalcPinValue As New SiteDouble
    Dim TnameInput As String
    Dim ds_CalcPinValue As New DSPWave
    Dim ds_CalcPinValue_Fuse As New DSPWave
    
    ds_CalcPinValue.CreateConstant 0, 1, DspDouble
    pl_MeasureDict = GetStoreDataAllType(argv(0))
    sd_MeasurePin = pl_MeasureDict.pins(argv(1)) 'IB2
    l_Divide = argv(2) 'IE2
    sd_MeasurePin = sd_MeasurePin.Abs '|IB2|
    l_Divide = l_Divide '|IE2|
    'sd_CalcPinValue = sd_MeasurePin.Subtract(l_Divide).divide(sd_MeasurePin)
    sd_CalcPinValue = sd_MeasurePin.Negate.Add(l_Divide).divide(sd_MeasurePin)

    TnameInput = Report_TName_From_Instance(CalcI, "X", , , , , , , tlForceFlow)
    TheExec.flow.TestLimit sd_CalcPinValue, , , , , , , , TnameInput, , "X", , , , , tlForceFlow
    
    sd_CalcPinValue = sd_CalcPinValue.divide(10)
    sd_CalcPinValue = sd_CalcPinValue.Abs
    
    For Each site In TheExec.sites.Active
        ds_CalcPinValue.Element(0) = sd_CalcPinValue.Multiply(2 ^ 8)
    Next site
    sd_CalcPinValue = sd_CalcPinValue.Multiply(2 ^ 8)
    TheExec.flow.TestLimit sd_CalcPinValue, , , , , scaleNoScaling, , , "HAC_CalcI_3_TDIODE_MTR_X_mtr-tdiode-beta_x_x_0", , "X", , , , , tlForceFlow
    Call HardIP_Dec2Bin(ds_CalcPinValue_Fuse, ds_CalcPinValue, 8)
    Call StoreDataAllType(argv(3), ds_CalcPinValue_Fuse)
''    Call StoreDataAllType(argv(3) & "_para", sd_CalcPinValue)
    
End Function

Public Function Calc_Voltage_Delta(argc As Integer, argv() As String) As Long
    
    
    ''''argv(0) = VBE2 'GetDictName
    ''''argv(1) = MTR_TD_ES-MTR_TD_BS 'PinName
    ''''argv(2) = VBE1 'GetDictName
    ''''argv(3) = MTR_TD_ES-MTR_TD_BS 'PinName
    ''''argv(4) = Delta_VBE 'Dict_Store_Name for eFuse

    Dim pl_MeasureDict1 As New PinListData
    Dim pl_MeasureDict2 As New PinListData
    Dim sd_MeasurePin1 As New SiteDouble
    Dim sd_MeasurePin2 As New SiteDouble
    Dim sd_CalcPinValue As New SiteDouble
    Dim TnameInput As String
    Dim ds_CalcPinValue As New DSPWave
    Dim ds_CalcPinValue_Fuse As New DSPWave
    
    ds_CalcPinValue.CreateConstant 0, 1, DspDouble
    pl_MeasureDict1 = GetStoreDataAllType(argv(0))
    sd_MeasurePin1 = pl_MeasureDict1.pins(argv(1)) 'VBE2
    pl_MeasureDict2 = GetStoreDataAllType(argv(2))
    sd_MeasurePin2 = pl_MeasureDict2.pins(argv(3)) 'VBE1
    sd_CalcPinValue = sd_MeasurePin1.Subtract(sd_MeasurePin2)
    'sd_CalcPinValue = sd_CalcPinValue.Abs
    'sd_CalcPinValue = sd_CalcPinValue.Multiply(1000) ''unit: mV
    'sd_CalcPinValue = sd_CalcPinValue.divide(200)
    
    TnameInput = Report_TName_From_Instance(CalcV, "X", , , , , , , tlForceFlow)
    TheExec.flow.TestLimit sd_CalcPinValue, , , , , , , , TnameInput, , "X", , , , , tlForceFlow
    
    sd_CalcPinValue = sd_CalcPinValue.divide(200)
    sd_CalcPinValue = sd_CalcPinValue.Abs
    sd_CalcPinValue = sd_CalcPinValue.Multiply(1000)
    For Each site In TheExec.sites.Active
        ds_CalcPinValue.Element(0) = sd_CalcPinValue.Multiply(2 ^ 16)
    Next site
    sd_CalcPinValue = sd_CalcPinValue.Multiply(2 ^ 16)
    TheExec.flow.TestLimit sd_CalcPinValue, , , , , scaleNoScaling, , , "HAC_CalcI_3_TDIODE_MTR_X_mtr-tdiode-delta-vbe_x_x_0", , "X", , , , , tlForceFlow
    Call HardIP_Dec2Bin(ds_CalcPinValue_Fuse, ds_CalcPinValue, 16)
    Call StoreDataAllType(argv(4), ds_CalcPinValue_Fuse)
''    Call StoreDataAllType(argv(4) & "_para", sd_CalcPinValue)

End Function
Public Function Calc_CIO_Freq_Delta(argc As Integer, argv() As String) As Long
    
    'From T-Col
    ''''argv(0) = ln0_freq_dco 'GetDictName
    ''''argv(1) = MTR_TD_ES-MTR_TD_BS 'PinName
    ''''argv(2) = VBE1 'GetDictName
    ''''argv(3) = MTR_TD_ES-MTR_TD_BS 'PinName
    ''''

    Dim pl_MeasureDict1 As New PinListData
    Dim pl_MeasureDict2 As New PinListData
    Dim sd_MeasurePin1 As New SiteDouble
    Dim sd_MeasurePin2 As New SiteDouble
    Dim sd_CalcPinValue As New SiteDouble
    Dim TnameInput As String
    Dim ds_CalcPinValue As New DSPWave
    Dim ds_CalcPinValue_Fuse As New DSPWave
    Dim index As Long
    
    If gl_Loop_count = gl_Loop_Max Then
        For index = 0 To UBound(glb_src_dec_value) - 1
            pl_MeasureDict1 = GetStoredMeasurement(argv(0) & "_" & CStr(glb_src_dec_value(index)))
            pl_MeasureDict2 = GetStoredMeasurement(argv(0) & "_" & CStr(glb_src_dec_value(index + 1)))
            'GetStoredMeasurement (argv(0) & "_" & CStr(glb_src_dec_value(gl_Loop_count)))
            'ds_CalcPinValue.CreateConstant 0, 1, DspDouble
            'pl_MeasureDict1 = GetStoredMeasurement(argv(0))
            sd_MeasurePin1 = pl_MeasureDict1.pins(argv(1)) 'VBE2
            'pl_MeasureDict2 = GetStoredMeasurement(argv(2))
            sd_MeasurePin2 = pl_MeasureDict2.pins(argv(1)) 'VBE1
            sd_CalcPinValue = sd_MeasurePin1.Subtract(sd_MeasurePin2)
            'sd_CalcPinValue = sd_CalcPinValue.Abs
            'sd_CalcPinValue = sd_CalcPinValue.Multiply(1000) ''unit: mV
            'sd_CalcPinValue = sd_CalcPinValue.divide(200)
        
            TnameInput = Report_TName_From_Instance(CalcF, "X", , , , , , , tlForceFlow)
            TheExec.flow.TestLimit sd_CalcPinValue, , , , , , , , TnameInput, , "X", , , , , tlForceFlow
            
        Next index
    End If
'    sd_CalcPinValue = sd_CalcPinValue.divide(200)
'    sd_CalcPinValue = sd_CalcPinValue.Abs
'    sd_CalcPinValue = sd_CalcPinValue.Multiply(1000)
'    For Each site In theexec.sites.Active
'        ds_CalcPinValue.Element(0) = sd_CalcPinValue.Multiply(2 ^ 16)
'    Next site
'    sd_CalcPinValue = sd_CalcPinValue.Multiply(2 ^ 16)
'    theexec.Flow.TestLimit sd_CalcPinValue, , , , , scaleNoScaling, , , "HAC_CalcI_3_TDIODE_MTR_X_mtr-tdiode-delta-vbe_x_x_0", , "X", , , , , tlForceFlow
'    Call HardIP_Dec2Bin(ds_CalcPinValue_Fuse, ds_CalcPinValue, 16)
'    Call AddStoredCaptureData(argv(4), ds_CalcPinValue_Fuse)
''    Call AddStoredData(argv(4) & "_para", sd_CalcPinValue)

End Function


Public Function Calc_Current_Ratio(argc As Integer, argv() As String) As Long
    
    
    ''''argv(0) = IB2 'GetDictName
    ''''argv(1) = MTR_TD_B-MTR_TD_BS 'PinName
    ''''argv(2) = IB1 'GetDictName
    ''''argv(3) = MTR_TD_B-MTR_TD_BS 'PinName
    ''''argv(4) = IC_RATIO 'Dict_Store_Name for eFuse

    Dim pl_MeasureDict1 As New PinListData
    Dim pl_MeasureDict2 As New PinListData
    Dim sd_MeasurePin1 As New SiteDouble
    Dim sd_MeasurePin2 As New SiteDouble
    Dim sd_CalcPinValue As New SiteDouble
    Dim TnameInput As String
    Dim ds_CalcPinValue As New DSPWave
    Dim ds_CalcPinValue_Fuse As New DSPWave
    
    ds_CalcPinValue.CreateConstant 0, 1, DspDouble
    pl_MeasureDict1 = GetStoreDataAllType(argv(0))
    sd_MeasurePin1 = pl_MeasureDict1.pins(argv(1)).Abs 'IB2
    pl_MeasureDict2 = GetStoreDataAllType(argv(2))
    sd_MeasurePin2 = pl_MeasureDict2.pins(argv(3)).Abs 'IB1
    'sd_CalcPinValue = sd_MeasurePin2.Negate.Add(0.00024).divide(sd_MeasurePin1.Negate.Add(0.00003))
    sd_CalcPinValue = sd_MeasurePin1.Negate.Add(0.00024).divide(sd_MeasurePin2.Negate.Add(0.00003))
    
    TnameInput = Report_TName_From_Instance(CalcI, "X", , , , , , , tlForceFlow)
    TheExec.flow.TestLimit sd_CalcPinValue, , , , , , , , TnameInput, , "X", , , , , tlForceFlow
    
    sd_CalcPinValue = sd_CalcPinValue.divide(24)
    sd_CalcPinValue = sd_CalcPinValue.Abs
    
    For Each site In TheExec.sites.Active
        ds_CalcPinValue.Element(0) = sd_CalcPinValue.Multiply(2 ^ 16)
    Next site
    sd_CalcPinValue = sd_CalcPinValue.Multiply(2 ^ 16)
    TheExec.flow.TestLimit sd_CalcPinValue, , , , , scaleNoScaling, , , "HAC_CalcI_3_TDIODE_MTR_X_mtr-tdiode-ic-ratio_x_x_0", , "X", , , , , tlForceFlow
    Call HardIP_Dec2Bin(ds_CalcPinValue_Fuse, ds_CalcPinValue, 16)
    Call StoreDataAllType(argv(4), ds_CalcPinValue_Fuse)
''    Call StoreDataAllType(argv(4) & "_para", sd_CalcPinValue)

End Function

Public Function Calc_Judge_Dram_Type(argc As Integer, argv() As String) As Long
    Dim i As Integer
    Dim site As Variant
    Dim MRR_vender As New SiteDouble
    Dim MRR_size As New SiteDouble
    Dim First_run As Boolean: First_run = False
    Dim First_site_vender_data As Double
    Dim First_site_size_data As Double
    
    MRR_vender = GetStoreDataAllType(argv(0) & "_para")
    MRR_size = GetStoreDataAllType(argv(1) & "_para")
    

    TheExec.flow.enableWord("DRAM_Hynix8GB") = False
    TheExec.flow.enableWord("DRAM_Hynix16GB") = False
    TheExec.flow.enableWord("DRAM_Hynix24GB") = False
    TheExec.flow.enableWord("DRAM_Micron8GB") = False
    TheExec.flow.enableWord("DRAM_Micron16GB") = False
    TheExec.flow.enableWord("DRAM_Micron24GB") = False

    For Each site In TheExec.sites
        If MRR_vender = 6 Then
            TheExec.Datalog.WriteComment "Site: " & site & " is DRAM type = Hynix"
        ElseIf MRR_vender = 255 Then
            TheExec.Datalog.WriteComment "Site: " & site & " is DRAM type = Micron"
        Else
            TheExec.Datalog.WriteComment "Site: " & site & " Can't Judge the DRAM TYPE"
        End If


        If MRR_size = 16 Then
            TheExec.Datalog.WriteComment "Site: " & site & " DRAM Size = 8G"
        ElseIf MRR_size = 80 Then
            TheExec.Datalog.WriteComment "Site: " & site & " DRAM Size = 16G"
        ElseIf MRR_size = 84 Then
            TheExec.Datalog.WriteComment "Site: " & site & " DRAM Size = 24G"
        ElseIf MRR_vender = 255 And MRR_size = 85 Then
            TheExec.Datalog.WriteComment "Site: " & site & " DRAM Size = 24G"
        Else
            TheExec.Datalog.WriteComment "Site: " & site & " Can't Judge the Size"
        End If
        
        If First_run = False Then
            First_site_vender_data = MRR_vender(site)
            First_site_size_data = MRR_size(site)
            First_run = True
        Else
            If First_site_vender_data <> MRR_vender(site) Or First_site_size_data <> MRR_size(site) Then
                TheExec.sites.item(site).FlagState("F_HardIP_DRAM_MATCH") = logicTrue
                TheExec.Datalog.WriteComment "DRAM not match for all sites"
            End If
        End If
    
    Next site
    
    For Each site In TheExec.sites
        If MRR_vender = 6 And MRR_size = 16 Then
            TheExec.flow.enableWord("DRAM_Hynix8GB") = True
        ElseIf MRR_vender = 6 And MRR_size = 80 Then
            TheExec.flow.enableWord("DRAM_Hynix16GB") = True
        ElseIf MRR_vender = 6 And MRR_size = 84 Then
            TheExec.flow.enableWord("DRAM_Hynix24GB") = True
        ElseIf MRR_vender = 255 And MRR_size = 16 Then
            TheExec.flow.enableWord("DRAM_Micron8GB") = True
        ElseIf MRR_vender = 255 And MRR_size = 80 Then
            TheExec.flow.enableWord("DRAM_Micron16GB") = True
        ElseIf MRR_vender = 255 And MRR_size = 84 Then
            TheExec.flow.enableWord("DRAM_Micron24GB") = True
        ElseIf MRR_vender = 255 And MRR_size = 85 Then
            TheExec.flow.enableWord("DRAM_Micron24GB") = True
        End If
        
        Exit For
    Next site

End Function

Public Function Calc_ADC2AnalogVoltage_MIPI(argc As Integer, argv() As String) As Long '0428wade
    Dim InWf As New DSPWave
    Dim OutWf As New DSPWave
    Dim SiteDbl_Vref As New SiteDouble
    Dim Voltage As Double
    Dim V_scale As Double
    Dim Analog_Voltage(0) As New SiteDouble
    'Alg::Calc_ADC2AnalogVoltage_MIPI(VDD12_PLL_DDR,adccode1lsb,adccode1msb,....,1024)
    Dim i As Integer 'Dict counter
    Dim TestNameInput As String
    
    Voltage = TheHdw.DCVS.pins(argv(0)).Voltage.value
    V_scale = CDbl(argv(UBound(argv)))
    OutWf.CreateConstant 0, 1, DspDouble
    
    For Each site In TheExec.sites.Active
        InWf.Clear
        For i = 1 To argc - 2
            If i = 1 Then
                InWf = GetStoreDataAllType(argv(i))
            Else
                InWf = InWf.Concatenate(GetStoreDataAllType(argv(i)))
            End If
        Next i
        Call HardIP_Bin2Dec(OutWf, InWf)
        Analog_Voltage(0) = OutWf.Element(0) * Voltage / V_scale
    Next site
          
        TestNameInput = Report_TName_From_Instance("CalcV", vbNullString, , 0) 'Tname:=FlowTname(k),TestSeqNum:=k,ForceResult:=tlForceNone
        TheExec.flow.TestLimit resultVal:=Analog_Voltage(0), Tname:=TestNameInput, PinName:=vbNullString, ForceResults:=tlForceFlow

Exit Function
errHandler:
    TheExec.Datalog.WriteComment "error in Calc_ADC2AnalogVoltage_MIPI"
    If AbortTest Then Exit Function Else Resume Next
End Function

Public Function Calc_ABSMIN(argc As Integer, argv() As String) As Long
    Dim InWf As New DSPWave
    Dim InWf_SiteDouble() As New SiteDouble: ReDim InWf_SiteDouble(UBound(Split(argv(0), "+")))
    Dim InWf_Array() As Double: ReDim InWf_Array(UBound(Split(argv(0), "+")))
    Dim InWf_Split() As String: InWf_Split = Split(argv(0), "+")
    
    Dim Offset_Data As New DSPWave
    Dim Offset_Data_Array() As Long
    Dim Offset_LSB_Data() As New DSPWave
    Dim Offset_LSB_Array(0) As Long
    Dim glsweepnumber As Integer: glsweepnumber = 0
    Dim srcsweepname As String
    
    Dim DSP_Gain_Mean_Final As New DSPWave
    Dim DSP_Gain_Mean_Final_Array(0) As Double
    Dim TestNameInput As String
    Dim i As Long
    Dim Temp_data As New DSPWave

    srcsweepname = Sweepnameforsweep(srcnameindex)
        ReDim Offset_LSB_Data(UBound(argv))
        For i = 0 To UBound(argv)
            'InWf_SiteDouble(i) = GetStoreDataAllType(argv(i) & "_para")
            Offset_Data = GetStoreDataAllType(argv(i))
            For Each site In TheExec.sites.Active
                If TheExec.TesterMode = testModeOffline Then
                    Offset_Data.CreateRandom 0, 1, 10, DataType:=DspLong
                    Offset_LSB_Array(0) = Offset_Data.data(0)
                Else
                    Offset_LSB_Array(0) = Offset_Data.data(0)
                End If
                Offset_LSB_Data(i).data = Offset_LSB_Array
            Next site
            
''            For Each site In TheExec.sites.Active
''                Offset_Data = GetStoreDataAllType(argv(i))
''                Offset_Data_Array = Offset_Data.data
''                Offset_LSB_Array(0) = Offset_Data_Array(0)
''                Offset_LSB_Data.data = Offset_LSB_Array
''            Next site
            Call StoreDataAllType(argv(i) & "_" & CStr(srcsweepname), Offset_LSB_Data(i))
            Temp_data = GetStoreDataAllType(argv(i) & "_" & CStr(srcsweepname))
        Next i

    If InStr(srcsweepname, "111") Then
        For i = 0 To 4 ''For each Instance A,B,C0,D0,D1
            For Each site In TheExec.sites.Active
            TestNameInput = Report_TName_From_Instance("CalcC", "", "")
            TestNameInput = Replace(TestNameInput, "Input6to4V111", "x")

            If GetStoreDataAllType(argv(2 * i) & "_Input_6_to_4=000").Element(0) = 1 Then
                If GetStoreDataAllType(argv(2 * i) & "_Input_6_to_4=100").Element(0) = 0 Then
                    If GetStoreDataAllType(argv(2 * i + 1) & "_Input_6_to_4=000").Element(0) = 1 Then
                        TheExec.flow.TestLimit resultVal:="-1", Tname:=TestNameInput, ForceResults:=tlForceFlow, scaletype:=scaleNoScaling
                    Else
                        TheExec.flow.TestLimit resultVal:="0", Tname:=TestNameInput, ForceResults:=tlForceFlow, scaletype:=scaleNoScaling
                    End If
                    'TheExec.Flow.TestLimit resultVal:="-1", Tname:=TestNameInput, ForceResults:=tlForceFlow, scaletype:=scaleNoScaling
                Else
                    If GetStoreDataAllType(argv(2 * i) & "_Input_6_to_4=010").Element(0) = 0 Then
                        TheExec.flow.TestLimit resultVal:="-2", Tname:=TestNameInput, ForceResults:=tlForceFlow, scaletype:=scaleNoScaling
                    Else
                        If GetStoreDataAllType(argv(2 * i) & "_Input_6_to_4=110").Element(0) = 0 Then
                            TheExec.flow.TestLimit resultVal:="-3", Tname:=TestNameInput, ForceResults:=tlForceFlow, scaletype:=scaleNoScaling
                        Else
                            If GetStoreDataAllType(argv(2 * i) & "_Input_6_to_4=001").Element(0) = 0 Then
                                TheExec.flow.TestLimit resultVal:="-4", Tname:=TestNameInput, ForceResults:=tlForceFlow, scaletype:=scaleNoScaling
                            Else
                                If GetStoreDataAllType(argv(2 * i) & "_Input_6_to_4=101").Element(0) = 0 Then
                                    TheExec.flow.TestLimit resultVal:="-5", Tname:=TestNameInput, ForceResults:=tlForceFlow, scaletype:=scaleNoScaling
                                Else
                                    If GetStoreDataAllType(argv(2 * i) & "_Input_6_to_4=011").Element(0) = 0 Then
                                        TheExec.flow.TestLimit resultVal:="-6", Tname:=TestNameInput, ForceResults:=tlForceFlow, scaletype:=scaleNoScaling
                                    Else
                                        If GetStoreDataAllType(argv(2 * i) & "_Input_6_to_4=111").Element(0) = 0 Then
                                            TheExec.flow.TestLimit resultVal:="-7", Tname:=TestNameInput, ForceResults:=tlForceFlow, scaletype:=scaleNoScaling
                                        Else
                                            TheExec.flow.TestLimit resultVal:="-8", Tname:=TestNameInput, ForceResults:=tlForceFlow, scaletype:=scaleNoScaling
                                        End If
                                    End If
                                End If
                            End If
                        End If
                    End If
                End If
            End If

            If GetStoreDataAllType(argv(2 * i) & "_Input_6_to_4=000").Element(0) = 0 Then
                If GetStoreDataAllType(argv(2 * i + 1) & "_Input_6_to_4=000").Element(0) = 1 Then
                    TheExec.flow.TestLimit resultVal:="0", Tname:=TestNameInput, ForceResults:=tlForceFlow, scaletype:=scaleNoScaling
                Else
                    If GetStoreDataAllType(argv(2 * i + 1) & "_Input_6_to_4=100").Element(0) = 1 Then
                        TheExec.flow.TestLimit resultVal:="1", Tname:=TestNameInput, ForceResults:=tlForceFlow, scaletype:=scaleNoScaling
                    Else
                        If GetStoreDataAllType(argv(2 * i + 1) & "_Input_6_to_4=010").Element(0) = 1 Then
                            TheExec.flow.TestLimit resultVal:="2", Tname:=TestNameInput, ForceResults:=tlForceFlow, scaletype:=scaleNoScaling
                        Else
                            If GetStoreDataAllType(argv(2 * i + 1) & "_Input_6_to_4=110").Element(0) = 1 Then
                                TheExec.flow.TestLimit resultVal:="3", Tname:=TestNameInput, ForceResults:=tlForceFlow, scaletype:=scaleNoScaling
                            Else
                                If GetStoreDataAllType(argv(2 * i + 1) & "_Input_6_to_4=001").Element(0) = 1 Then
                                    TheExec.flow.TestLimit resultVal:="4", Tname:=TestNameInput, ForceResults:=tlForceFlow, scaletype:=scaleNoScaling
                                Else
                                    If GetStoreDataAllType(argv(2 * i + 1) & "_Input_6_to_4=101").Element(0) = 1 Then
                                        TheExec.flow.TestLimit resultVal:="5", Tname:=TestNameInput, ForceResults:=tlForceFlow, scaletype:=scaleNoScaling
                                    Else
                                        If GetStoreDataAllType(argv(2 * i + 1) & "_Input_6_to_4=011").Element(0) = 1 Then
                                            TheExec.flow.TestLimit resultVal:="6", Tname:=TestNameInput, ForceResults:=tlForceFlow, scaletype:=scaleNoScaling
                                        Else
                                            If GetStoreDataAllType(argv(2 * i + 1) & "_Input_6_to_4=111").Element(0) = 1 Then
                                                TheExec.flow.TestLimit resultVal:="7", Tname:=TestNameInput, ForceResults:=tlForceFlow, scaletype:=scaleNoScaling
                                            Else
                                                TheExec.flow.TestLimit resultVal:="8", Tname:=TestNameInput, ForceResults:=tlForceFlow, scaletype:=scaleNoScaling
                                            End If
                                        End If
                                    End If
                                End If
                            End If
                        End If
                    End If
                End If
            End If
            TheExec.flow.TestLimitIndex = TheExec.flow.TestLimitIndex - 1
            Next site
            TheExec.flow.TestLimitIndex = TheExec.flow.TestLimitIndex + 1
        Next i
        'Stop
    End If
    
    Exit Function
errHandler:
    TheExec.Datalog.WriteComment "error in Calc_ABSMIN"
    If AbortTest Then Exit Function Else Resume Next
End Function


Public Function Calc_ABSMIN_DD(argc As Integer, argv() As String) As Long
    Dim InWf As New DSPWave
    Dim InWf_SiteDouble() As New SiteDouble: ReDim InWf_SiteDouble(UBound(Split(argv(0), "+")))
    Dim InWf_Array() As Double: ReDim InWf_Array(UBound(Split(argv(0), "+")))
    Dim InWf_Split() As String: InWf_Split = Split(argv(0), "+")
    
    Dim Offset_Data As New DSPWave
    Dim Offset_Data_Array() As Long
    Dim Offset_LSB_Data() As New DSPWave
    Dim Offset_LSB_Array(0) As Long
    Dim glsweepnumber As Integer: glsweepnumber = 0
    Dim srcsweepname As String
    
    Dim DSP_Gain_Mean_Final As New DSPWave
    Dim DSP_Gain_Mean_Final_Array(0) As Double
    Dim TestNameInput As String
    Dim i As Long
    Dim Temp_data As New DSPWave
    
'    srcsweepname = Sweepnameforsweep(srcnameindex)
'        ReDim Offset_LSB_Data(UBound(argv))
'        For i = 0 To UBound(argv)
'            'InWf_SiteDouble(i) = GetStoreDataAllType(argv(i) & "_para")
'            Offset_Data = GetStoreDataAllType(argv(i))
'            For Each site In TheExec.sites.Active
'                If TheExec.TesterMode = testModeOffline Then
'                    Offset_Data.CreateRandom 0, 1, 10, DataType:=DspLong
'                    Offset_LSB_Array(0) = Offset_Data.data(0)
'                Else
'                    Offset_LSB_Array(0) = Offset_Data.data(0)
'                End If
'                Offset_LSB_Data(i).data = Offset_LSB_Array
'            Next site
'
'''            For Each site In TheExec.sites.Active
'''                Offset_Data = GetStoreDataAllType(argv(i))
'''                Offset_Data_Array = Offset_Data.data
'''                Offset_LSB_Array(0) = Offset_Data_Array(0)
'''                Offset_LSB_Data.data = Offset_LSB_Array
'''            Next site
'            Call StoreDataAllType(argv(i) & "_" & CStr(srcsweepname), Offset_LSB_Data(i))
'            temp_data = GetStoreDataAllType(argv(i) & "_" & CStr(srcsweepname))
'        Next i
        
    'If InStr(srcsweepname, "111") Then
    'If InStr(UCase(TheExec.DataManager.instancename), "ABSMIN_DDCPUOFFS2") <> 0 Then
     '        TNameSeg(8) = Replace("MULTI&" & CStr(Evaluate("0.4+" + CStr(TheExec.Flow.var(gl_Sweep_Name).value) + "*0.2")), ".", "p")
    'End If
    If UCase(TheExec.DataManager.instancename) Like "ABSMIN_CZCPUOFFS2*" Then
        For i = 0 To (argc / 2) - 1 ''For each Instance A,B,C0,D0,D1
            For Each site In TheExec.sites.Active
            TestNameInput = Report_TName_From_Instance("CalcC", "", "")
            'TestNameInput = Replace(TestNameInput, "Input6to4V111", "x")

            If GetStoreDataAllType(argv(2 * i) & "_Input_6_to_4_DD=000").Element(0) = 1 Then
                If GetStoreDataAllType(argv(2 * i) & "_Input_6_to_4_DD=100").Element(0) = 0 Then
                    If GetStoreDataAllType(argv(2 * i + 1) & "_Input_6_to_4_DD=000").Element(0) = 1 Then
                        TheExec.flow.TestLimit resultVal:="-1", Tname:=TestNameInput, ForceResults:=tlForceFlow, scaletype:=scaleNoScaling
                    Else
                        TheExec.flow.TestLimit resultVal:="0", Tname:=TestNameInput, ForceResults:=tlForceFlow, scaletype:=scaleNoScaling
                    End If
                    'TheExec.Flow.TestLimit resultVal:="-1", Tname:=TestNameInput, ForceResults:=tlForceFlow, scaletype:=scaleNoScaling
                Else
                    If GetStoreDataAllType(argv(2 * i) & "_Input_6_to_4_DD=010").Element(0) = 0 Then
                        TheExec.flow.TestLimit resultVal:="-2", Tname:=TestNameInput, ForceResults:=tlForceFlow, scaletype:=scaleNoScaling
                    Else
                        If GetStoreDataAllType(argv(2 * i) & "_Input_6_to_4_DD=110").Element(0) = 0 Then
                            TheExec.flow.TestLimit resultVal:="-3", Tname:=TestNameInput, ForceResults:=tlForceFlow, scaletype:=scaleNoScaling
                        Else
                            If GetStoreDataAllType(argv(2 * i) & "_Input_6_to_4_DD=001").Element(0) = 0 Then
                                TheExec.flow.TestLimit resultVal:="-4", Tname:=TestNameInput, ForceResults:=tlForceFlow, scaletype:=scaleNoScaling
                            Else
                                If GetStoreDataAllType(argv(2 * i) & "_Input_6_to_4_DD=101").Element(0) = 0 Then
                                    TheExec.flow.TestLimit resultVal:="-5", Tname:=TestNameInput, ForceResults:=tlForceFlow, scaletype:=scaleNoScaling
                                Else
                                    If GetStoreDataAllType(argv(2 * i) & "_Input_6_to_4_DD=011").Element(0) = 0 Then
                                        TheExec.flow.TestLimit resultVal:="-6", Tname:=TestNameInput, ForceResults:=tlForceFlow, scaletype:=scaleNoScaling
                                    Else
                                        If GetStoreDataAllType(argv(2 * i) & "_Input_6_to_4_DD=111").Element(0) = 0 Then
                                            TheExec.flow.TestLimit resultVal:="-7", Tname:=TestNameInput, ForceResults:=tlForceFlow, scaletype:=scaleNoScaling
                                        Else
                                            TheExec.flow.TestLimit resultVal:="-8", Tname:=TestNameInput, ForceResults:=tlForceFlow, scaletype:=scaleNoScaling
                                        End If
                                    End If
                                End If
                            End If
                        End If
                    End If
                End If
            End If

            If GetStoreDataAllType(argv(2 * i) & "_Input_6_to_4_DD=000").Element(0) = 0 Then
                If GetStoreDataAllType(argv(2 * i + 1) & "_Input_6_to_4_DD=000").Element(0) = 1 Then
                    TheExec.flow.TestLimit resultVal:="0", Tname:=TestNameInput, ForceResults:=tlForceFlow, scaletype:=scaleNoScaling
                Else
                    If GetStoreDataAllType(argv(2 * i + 1) & "_Input_6_to_4_DD=100").Element(0) = 1 Then
                        TheExec.flow.TestLimit resultVal:="1", Tname:=TestNameInput, ForceResults:=tlForceFlow, scaletype:=scaleNoScaling
                    Else
                        If GetStoreDataAllType(argv(2 * i + 1) & "_Input_6_to_4_DD=010").Element(0) = 1 Then
                            TheExec.flow.TestLimit resultVal:="2", Tname:=TestNameInput, ForceResults:=tlForceFlow, scaletype:=scaleNoScaling
                        Else
                            If GetStoreDataAllType(argv(2 * i + 1) & "_Input_6_to_4_DD=110").Element(0) = 1 Then
                                TheExec.flow.TestLimit resultVal:="3", Tname:=TestNameInput, ForceResults:=tlForceFlow, scaletype:=scaleNoScaling
                            Else
                                If GetStoreDataAllType(argv(2 * i + 1) & "_Input_6_to_4_DD=001").Element(0) = 1 Then
                                    TheExec.flow.TestLimit resultVal:="4", Tname:=TestNameInput, ForceResults:=tlForceFlow, scaletype:=scaleNoScaling
                                Else
                                    If GetStoreDataAllType(argv(2 * i + 1) & "_Input_6_to_4_DD=101").Element(0) = 1 Then
                                        TheExec.flow.TestLimit resultVal:="5", Tname:=TestNameInput, ForceResults:=tlForceFlow, scaletype:=scaleNoScaling
                                    Else
                                        If GetStoreDataAllType(argv(2 * i + 1) & "_Input_6_to_4_DD=011").Element(0) = 1 Then
                                            TheExec.flow.TestLimit resultVal:="6", Tname:=TestNameInput, ForceResults:=tlForceFlow, scaletype:=scaleNoScaling
                                        Else
                                            If GetStoreDataAllType(argv(2 * i + 1) & "_Input_6_to_4_DD=111").Element(0) = 1 Then
                                                TheExec.flow.TestLimit resultVal:="7", Tname:=TestNameInput, ForceResults:=tlForceFlow, scaletype:=scaleNoScaling
                                            Else
                                                TheExec.flow.TestLimit resultVal:="8", Tname:=TestNameInput, ForceResults:=tlForceFlow, scaletype:=scaleNoScaling
                                            End If
                                        End If
                                    End If
                                End If
                            End If
                        End If
                    End If
                End If
            End If
            TheExec.flow.TestLimitIndex = TheExec.flow.TestLimitIndex - 1
            Next site
            TheExec.flow.TestLimitIndex = TheExec.flow.TestLimitIndex + 1
        Next i
        'Stop
    Else
        For Each site In TheExec.sites.Active
            TestNameInput = Report_TName_From_Instance("CalcC", vbNullString, vbNullString)
            'TestNameInput = Replace(TestNameInput, "Input6to4V111", "x")

            If GetStoreDataAllType(argv(2 * i) & "_Input_6_to_4_DD=000").Element(0) = 1 Then
                If GetStoreDataAllType(argv(2 * i) & "_Input_6_to_4_DD=100").Element(0) = 0 Then
                    If GetStoreDataAllType(argv(2 * i + 1) & "_Input_6_to_4_DD=000").Element(0) = 1 Then
                        TheExec.flow.TestLimit resultVal:="-1", Tname:=TestNameInput, ForceResults:=tlForceFlow, scaletype:=scaleNoScaling
                    Else
                        TheExec.flow.TestLimit resultVal:="0", Tname:=TestNameInput, ForceResults:=tlForceFlow, scaletype:=scaleNoScaling
                    End If
                    'TheExec.Flow.TestLimit resultVal:="-1", Tname:=TestNameInput, ForceResults:=tlForceFlow, scaletype:=scaleNoScaling
                Else
                    If GetStoreDataAllType(argv(2 * i) & "_Input_6_to_4_DD=010").Element(0) = 0 Then
                        TheExec.flow.TestLimit resultVal:="-2", Tname:=TestNameInput, ForceResults:=tlForceFlow, scaletype:=scaleNoScaling
                    Else
                        If GetStoreDataAllType(argv(2 * i) & "_Input_6_to_4_DD=110").Element(0) = 0 Then
                            TheExec.flow.TestLimit resultVal:="-3", Tname:=TestNameInput, ForceResults:=tlForceFlow, scaletype:=scaleNoScaling
                        Else
                            If GetStoreDataAllType(argv(2 * i) & "_Input_6_to_4_DD=001").Element(0) = 0 Then
                                TheExec.flow.TestLimit resultVal:="-4", Tname:=TestNameInput, ForceResults:=tlForceFlow, scaletype:=scaleNoScaling
                            Else
                                If GetStoreDataAllType(argv(2 * i) & "_Input_6_to_4_DD=101").Element(0) = 0 Then
                                    TheExec.flow.TestLimit resultVal:="-5", Tname:=TestNameInput, ForceResults:=tlForceFlow, scaletype:=scaleNoScaling
                                Else
                                    If GetStoreDataAllType(argv(2 * i) & "_Input_6_to_4_DD=011").Element(0) = 0 Then
                                        TheExec.flow.TestLimit resultVal:="-6", Tname:=TestNameInput, ForceResults:=tlForceFlow, scaletype:=scaleNoScaling
                                    Else
                                        If GetStoreDataAllType(argv(2 * i) & "_Input_6_to_4_DD=111").Element(0) = 0 Then
                                            TheExec.flow.TestLimit resultVal:="-7", Tname:=TestNameInput, ForceResults:=tlForceFlow, scaletype:=scaleNoScaling
                                        Else
                                            TheExec.flow.TestLimit resultVal:="-8", Tname:=TestNameInput, ForceResults:=tlForceFlow, scaletype:=scaleNoScaling
                                        End If
                                    End If
                                End If
                            End If
                        End If
                    End If
                End If
            End If

            If GetStoreDataAllType(argv(2 * i) & "_Input_6_to_4_DD=000").Element(0) = 0 Then
                If GetStoreDataAllType(argv(2 * i + 1) & "_Input_6_to_4_DD=000").Element(0) = 1 Then
                    TheExec.flow.TestLimit resultVal:="0", Tname:=TestNameInput, ForceResults:=tlForceFlow, scaletype:=scaleNoScaling
                Else
                    If GetStoreDataAllType(argv(2 * i + 1) & "_Input_6_to_4_DD=100").Element(0) = 1 Then
                        TheExec.flow.TestLimit resultVal:="1", Tname:=TestNameInput, ForceResults:=tlForceFlow, scaletype:=scaleNoScaling
                    Else
                        If GetStoreDataAllType(argv(2 * i + 1) & "_Input_6_to_4_DD=010").Element(0) = 1 Then
                            TheExec.flow.TestLimit resultVal:="2", Tname:=TestNameInput, ForceResults:=tlForceFlow, scaletype:=scaleNoScaling
                        Else
                            If GetStoreDataAllType(argv(2 * i + 1) & "_Input_6_to_4_DD=110").Element(0) = 1 Then
                                TheExec.flow.TestLimit resultVal:="3", Tname:=TestNameInput, ForceResults:=tlForceFlow, scaletype:=scaleNoScaling
                            Else
                                If GetStoreDataAllType(argv(2 * i + 1) & "_Input_6_to_4_DD=001").Element(0) = 1 Then
                                    TheExec.flow.TestLimit resultVal:="4", Tname:=TestNameInput, ForceResults:=tlForceFlow, scaletype:=scaleNoScaling
                                Else
                                    If GetStoreDataAllType(argv(2 * i + 1) & "_Input_6_to_4_DD=101").Element(0) = 1 Then
                                        TheExec.flow.TestLimit resultVal:="5", Tname:=TestNameInput, ForceResults:=tlForceFlow, scaletype:=scaleNoScaling
                                    Else
                                        If GetStoreDataAllType(argv(2 * i + 1) & "_Input_6_to_4_DD=011").Element(0) = 1 Then
                                            TheExec.flow.TestLimit resultVal:="6", Tname:=TestNameInput, ForceResults:=tlForceFlow, scaletype:=scaleNoScaling
                                        Else
                                            If GetStoreDataAllType(argv(2 * i + 1) & "_Input_6_to_4_DD=111").Element(0) = 1 Then
                                                TheExec.flow.TestLimit resultVal:="7", Tname:=TestNameInput, ForceResults:=tlForceFlow, scaletype:=scaleNoScaling
                                            Else
                                                TheExec.flow.TestLimit resultVal:="8", Tname:=TestNameInput, ForceResults:=tlForceFlow, scaletype:=scaleNoScaling
                                            End If
                                        End If
                                    End If
                                End If
                            End If
                        End If
                    End If
                End If
            End If
            TheExec.flow.TestLimitIndex = TheExec.flow.TestLimitIndex - 1
            Next site
            TheExec.flow.TestLimitIndex = TheExec.flow.TestLimitIndex + 1
    End If
    'End If
    
    Exit Function
errHandler:
    TheExec.Datalog.WriteComment "error in Calc_ABSMIN"
    If AbortTest Then Exit Function Else Resume Next
End Function



Public Function Calc_SDLL_PI_Linearity(argc As Integer, argv() As String) As Long
    '---- New Function : SDLL /PI Linearity Calc -- Neil,20220406
    Dim T_Period As Double
    Dim Counter_duration As Double
    Dim T_duration As Double
    Dim PI_Start As Long            ' PI_Start Index ( SDLL:1 / PI:0)
    Dim PI_End As Long              ' PI_End Index (PI Index size)
    Dim Inst_index, DQ_index, PI_index As Long
    Dim decValue As New SiteDouble
    Dim Trosc_DecValue As New SiteDouble
    Dim Trosc_step_DecVal As New SiteDouble
    Dim temp_str As String
    Dim TestNameInput As String
    Dim Splt_str() As String
    Dim Splt_str2() As String
    Dim Dic_str(2) As String
    Dim Dic_str_Name As String
    Dim i As Long
    
    On Error GoTo errHandler

    Splt_str = Split(argv(0), "code")
    Splt_str2 = Split(Splt_str(1), "_")
    Dic_str(0) = Splt_str(0)
    For i = 1 To UBound(Splt_str2)
        Dic_str(1) = Dic_str(1) & "_" & Splt_str2(i)
    Next i
    
    T_Period = CDbl(argv(1)) * 0.000000000001 'Unit:ps
    Counter_duration = CDbl(argv(2))
    T_duration = T_Period * Counter_duration
    PI_Start = CDbl(argv(3))
    PI_End = CDbl(argv(4))
    ReDim Trosc(PI_End - PI_Start + 1) As New SiteDouble
        
    'Print out datalog
    Dim TName_Ary() As String
    Splt_str = Split(argv(0), "__")
    TestNameInput = Report_TName_From_Instance(CalcC, "X", , , , , , , tlForceNone)
    Splt_str2 = Split(TestNameInput, "_")
    Splt_str2(7) = Replace(UCase(Splt_str(0)), "_", "-") ' DDR0_CH0_DQ0 -> DDR0-CH0-DQ0
    TestNameInput = Join(Splt_str2, "_")
    
    For PI_index = PI_Start To PI_End
        decValue = GetStoreDataAllType(Dic_str(0) & "code" & PI_index & Dic_str(1) & "_para")
        Trosc(PI_index) = decValue.Invert.Multiply(T_duration)      'Trosc  = Tduration/ro_count_out   ; Unit pS
        
        'Print out datalog
        'TestNameInput = Report_TName_From_Instance(CalcC, "X", "Trosc-" & Dic_str(0) & "code" & PI_index & Dic_str(1), , , , , , tlForceNone)
        Splt_str2(6) = "picode" & PI_index & "-Trosc"
        TestNameInput = Join(Splt_str2, "_")
        TheExec.flow.TestLimit resultVal:=Trosc(PI_index), PinName:="X", Tname:=TestNameInput, scaletype:=scaleNone, ForceResults:=tlForceNone
    Next PI_index
    
    For PI_index = PI_Start To PI_End - 1
        Trosc_step_DecVal = Trosc(PI_index + 1).Subtract(Trosc(PI_index))
        
        'Print out datalog
        'TestNameInput = Report_TName_From_Instance(CalcC, "X", "Trosc-step-size-" & PI_index + 1 & "-" & PI_index, , , , , , tlForceNone)
        Splt_str2(6) = "picode" & PI_index & "-Trosc-step-size"
        TestNameInput = Join(Splt_str2, "_")
        TheExec.flow.TestLimit resultVal:=Trosc_step_DecVal, PinName:="X", Tname:=TestNameInput, scaletype:=scaleNone, ForceResults:=tlForceNone
    Next PI_index
    
    Exit Function

errHandler:
    TheExec.Datalog.WriteComment "error in Calc_SDLL_PI_Linearity"
    If AbortTest Then Exit Function Else Resume Next
End Function

Public Function Calc_PI_DCD_Duty_Cycle(argc As Integer, argv() As String) As Long
    '---- New Function : Library function to calculate dcd duty cycle -- Neil, 20220406
    'Note :
    '1. Capture data must be convert to 2SCOMPLEMENT first
    '2. Define 2SCOMPLEMENT in "CUS_Str_MainProgram" argurment
    '3. Calculate the following parameters for each Instance and byte
    '   m1 = (step3_dcm_count_out ? step2_dcm_count_out )/100/6
    '   m2 = (step0_dcm_count_out ? step3_dcm_count_out )/100/6
    '   m3 = (step1_dcm_count_out ? step0_dcm_count_out )/100/6
    '   m4 = (step1_dcm_count_out ? step3_dcm_count_out )/100/6
    '   m = (m1 + m2 + m3 + m4) / 4
    '   k = step3_dcm_count_out - m * 50
    '   dc1 = (step4_dcm_count_out - k) / m
    '   dc2 = (step5_dcm_count_out - k) / m
    '
    
    Dim site As Variant
    Dim TestNameInput As String
    Dim Splt_str() As String
    Dim Splt_str2() As String

    Dim m1 As New SiteDouble
    Dim m2 As New SiteDouble
    Dim m3 As New SiteDouble
    Dim m4 As New SiteDouble
    Dim m As New SiteDouble
    Dim k As New SiteDouble
    Dim dc1 As New SiteDouble
    Dim dc2 As New SiteDouble
    
    Dim step0_dcm_count_out As New SiteDouble
    Dim step1_dcm_count_out As New SiteDouble
    Dim step2_dcm_count_out As New SiteDouble
    Dim step3_dcm_count_out As New SiteDouble
    Dim step4_dcm_count_out As New SiteDouble
    Dim step5_dcm_count_out As New SiteDouble

    On Error GoTo errHandler

    step0_dcm_count_out = GetStoreDataAllType(argv(0) & "_para")
    step1_dcm_count_out = GetStoreDataAllType(argv(1) & "_para")
    step2_dcm_count_out = GetStoreDataAllType(argv(2) & "_para")
    step3_dcm_count_out = GetStoreDataAllType(argv(3) & "_para")
    step4_dcm_count_out = GetStoreDataAllType(argv(4) & "_para")
    step5_dcm_count_out = GetStoreDataAllType(argv(5) & "_para")
       
    m1 = step3_dcm_count_out.Subtract(step2_dcm_count_out).divide(100 / 6)
    m2 = step0_dcm_count_out.Subtract(step3_dcm_count_out).divide(100 / 6)
    m3 = step1_dcm_count_out.Subtract(step0_dcm_count_out).divide(100 / 3)
    m4 = step1_dcm_count_out.Subtract(step3_dcm_count_out).divide(100 / 2)
    m = m1.Add(m2).Add(m3).Add(m4)
    m = m.divide(4)
    k = step3_dcm_count_out.Subtract(m.Multiply(50))
    dc1 = step4_dcm_count_out.Subtract(k).divide(m)
    dc2 = step5_dcm_count_out.Subtract(k).divide(m)

    'Print out datalog
    Dim i As Long
    Dim TName_Ary() As String
    Splt_str = Split(argv(0), "__")      ' Use for combin TName [ Format : DDR0_CH0_DQ0__step0_dcm_count_out ]
    
    TestNameInput = Report_TName_From_Instance(CalcC, "X", , , , , , , tlForceNone)
    Splt_str2 = Split(TestNameInput, "_")
    Splt_str2(7) = Replace(UCase(Splt_str(0)), "_", "-") ' DDR0_CH0_DQ0 -> DDR0-CH0-DQ0
    TestNameInput = Join(Splt_str2, "_")
    
    Splt_str2(6) = "m1"
    TestNameInput = Join(Splt_str2, "_")
    TheExec.flow.TestLimit resultVal:=m1, PinName:="X", Tname:=TestNameInput, scaletype:=scaleNone, ForceResults:=tlForceNone

    Splt_str2(6) = "m2"
    TestNameInput = Join(Splt_str2, "_")
    TheExec.flow.TestLimit resultVal:=m2, PinName:="X", Tname:=TestNameInput, scaletype:=scaleNone, ForceResults:=tlForceNone

    Splt_str2(6) = "m3"
    TestNameInput = Join(Splt_str2, "_")
    TheExec.flow.TestLimit resultVal:=m3, PinName:="X", Tname:=TestNameInput, scaletype:=scaleNone, ForceResults:=tlForceNone
    
    Splt_str2(6) = "m4"
    TestNameInput = Join(Splt_str2, "_")
    TheExec.flow.TestLimit resultVal:=m4, PinName:="X", Tname:=TestNameInput, scaletype:=scaleNone, ForceResults:=tlForceNone
    
    Splt_str2(6) = "m"
    TestNameInput = Join(Splt_str2, "_")
    TheExec.flow.TestLimit resultVal:=m, PinName:="X", Tname:=TestNameInput, scaletype:=scaleNone, ForceResults:=tlForceNone

    Splt_str2(6) = "k"
    TestNameInput = Join(Splt_str2, "_")
    TheExec.flow.TestLimit resultVal:=k, PinName:="X", Tname:=TestNameInput, scaletype:=scaleNone, ForceResults:=tlForceNone
    
    Splt_str2(6) = "dc1"
    TestNameInput = Join(Splt_str2, "_")
    TheExec.flow.TestLimit resultVal:=dc1, PinName:="X", Tname:=TestNameInput, scaletype:=scaleNone, ForceResults:=tlForceNone
    
    Splt_str2(6) = "dc2"
    TestNameInput = Join(Splt_str2, "_")
    TheExec.flow.TestLimit resultVal:=dc2, PinName:="X", Tname:=TestNameInput, scaletype:=scaleNone, ForceResults:=tlForceNone
   
    Exit Function
    
errHandler:
    TheExec.Datalog.WriteComment "error in Calc_PI_DCD_Duty_Cycle"
    If AbortTest Then Exit Function Else Resume Next
End Function

Public Function Calc_PI_DCC_RANGE(argc As Integer, argv() As String) As Long
    '---- New Function : Library function to calculate dcd duty cycle Rnage F1 to F5 -- Neil, 20220406
    'Note :
    '1. Capture data must be convert to 2SCOMPLEMENT first
    '2. Define 2SCOMPLEMENT in "CUS_Str_MainProgram" argurment
    '3.Calculate the following parameters for each Instance and byte
    '    DutyCylce -max = DDR0_CH0_DQ0__dcdfree0_dcdup0_dcddelay63_wck_dccio_count_out / DDR0_CH0_DQ0__dcdfree1_dcdup0_dcddelay0_wck_dccio_count_out
    '    DutyCylce -min = DDR0_CH0_DQ0__dcdfree0_dcdup1_dcddelay63_wck_dccio_count_out / DDR0_CH0_DQ0__dcdfree1_dcdup0_dcddelay0_wck_dccio_count_out
    '    DutyCylce -No - DCD = DDR0_CH0_DQ0__dcdfree0_dcdup1_dcddelay63_wck_dccio_count_out / DDR0_CH0_DQ0__dcdfree1_dcdup0_dcddelay0_wck_dccio_count_out

    Dim site As Variant
    Dim TestNameInput As String
    Dim Splt_str() As String
    Dim Splt_str2() As String

    Dim DutyCylce_Max As New SiteDouble
    Dim DutyCylce_Min As New SiteDouble
    Dim DutyCylce_No_DCD As New SiteDouble

    Dim Free1_Up0_Dly0  As New SiteDouble
    Dim Free0_Up0_Dly63 As New SiteDouble
    Dim Free0_Up1_Dly63 As New SiteDouble
    Dim Free0_Up0_Dly0  As New SiteDouble
  
    On Error GoTo errHandler

    Free1_Up0_Dly0 = GetStoreDataAllType(argv(0) & "_para")
    Free0_Up0_Dly63 = GetStoreDataAllType(argv(1) & "_para")
    Free0_Up1_Dly63 = GetStoreDataAllType(argv(2) & "_para")
    Free0_Up0_Dly0 = GetStoreDataAllType(argv(3) & "_para")

    DutyCylce_Max = Free0_Up0_Dly63.divide(Free1_Up0_Dly0).Multiply(50).Add(50)
    DutyCylce_Min = Free0_Up1_Dly63.divide(Free1_Up0_Dly0).Multiply(50).Add(50)
    DutyCylce_No_DCD = Free0_Up0_Dly0.divide(Free1_Up0_Dly0).Multiply(50).Add(50)
    
    'Print out datalog
    Dim i As Long
    Dim TName_Ary() As String
    
    Splt_str = Split(argv(0), "__")      ' Use for combin TName [ Format : ddr0_dq0_ch0_free1_up0_dly0_cnt0 ]
    TestNameInput = Report_TName_From_Instance(CalcC, "X", , , , , , , tlForceNone)
    Splt_str2 = Split(TestNameInput, "_")
    
    Splt_str2(7) = Replace(UCase(Splt_str(0)), "_", "-") ' DDR0_CH0_DQ0 -> DDR0-CH0-DQ0
    TestNameInput = Join(Splt_str2, "_")
    
    Splt_str2(6) = "DutyCylce-Max"
    TestNameInput = Join(Splt_str2, "_")
    TheExec.flow.TestLimit resultVal:=DutyCylce_Max, PinName:="X", Tname:=TestNameInput, scaletype:=scaleNone, ForceResults:=tlForceNone
    
    Splt_str2(6) = "DutyCylce-Min"
    TestNameInput = Join(Splt_str2, "_")
    TheExec.flow.TestLimit resultVal:=DutyCylce_Min, PinName:="X", Tname:=TestNameInput, scaletype:=scaleNone, ForceResults:=tlForceNone
    
    Splt_str2(6) = "DutyCylce-No-DCD"
    TestNameInput = Join(Splt_str2, "_")
    TheExec.flow.TestLimit resultVal:=DutyCylce_No_DCD, PinName:="X", Tname:=TestNameInput, scaletype:=scaleNone, ForceResults:=tlForceNone

    Exit Function
    
errHandler:
    TheExec.Datalog.WriteComment "error in Calc_PI_DCC_RANGE"
    If AbortTest Then Exit Function Else Resume Next
End Function

Public Function Calc_Current_VTHSENSOR(argc As Integer, argv() As String) As Long
    'argv(0): start point
    'argv(1): step first point
    'argv(2): step end point
    'argv(3): measure store name
    'argv(4): step size
    
    Dim loop_start, loop_now, loop_end, i As Long
    Dim pl_MeasureDict As New PinListData
    Dim sd_MeasureData As Double
    Dim sd_DeltaValue As Double
    Dim temp_High As Double
    Dim temp_Low  As Double
    Dim Final_High As New SiteLong
    Dim Final_Low As New SiteLong
    Dim Final_High_Voltage As New SiteDouble
    Dim Final_Low_Voltage As New SiteDouble
    Dim Final_High_Current As New SiteDouble
    Dim Final_Low_Current As New SiteDouble
    Dim TnameInput As String
    'Dim site As New SiteVariant

    
    loop_now = gl_sweepVoltage_Loop_Indx       'Get Loop Now
    loop_start = CInt(argv(1))      'Get Loop Start
    loop_end = CInt(argv(2))      'Get Loop End
    pl_MeasureDict = GetStoreDataAllType(argv(3))      'Get Stored Current Data
    Call StoreDataAllType(argv(3) & "_" & gl_sweepVoltage_Loop_Indx, pl_MeasureDict)     'Add Stored Current Data with "_SrcCodeIndex"
    
    'By site to find closest to target -1.92uA in 31 force condition
    If gl_sweepVoltage_Loop_Indx = loop_end Then
        Dim Setup_string_ary() As String
        Dim SplitLeftBigColon() As String
        Dim Setup_string As String
        Setup_string_ary = Split(Instance_Data.Interpose_PreMeas, ";")
        For i = 0 To UBound(Setup_string_ary)
            'If InStr(Setup_string_ary(i), "[") > 0 Then     '(Ex: PinA:V:0.5*[SrcCodeIndx]+0.001
                SplitLeftBigColon = Split(Setup_string_ary(i), ":")
                Setup_string = SplitLeftBigColon(UBound(SplitLeftBigColon))
                Exit For
            'End If
        Next i
        For Each site In TheExec.sites.Active
            
            temp_High = 9999
            temp_Low = 9999
            Final_High = 0
            Final_Low = 0
            Final_High_Voltage = 0
            Final_Low_Voltage = 0
            Final_High_Current = 0
            Final_Low_Current = 0
            
            For i = loop_start To loop_end
                Set pl_MeasureDict = Nothing
                pl_MeasureDict = GetStoreDataAllType(argv(3) & "_" & i)     'Get Stored Current Data
                sd_MeasureData = pl_MeasureDict.pins(0).value
                'High Search (<0uA) --> Small than -1.92uA (Ex: -2.07uA)
                If sd_MeasureData < 0 And sd_MeasureData < -1.92 * 0.000001 Then
                    sd_DeltaValue = Abs(sd_MeasureData - (-1.92 * 0.000001))
                    If sd_DeltaValue < temp_High Then
                        temp_High = sd_DeltaValue
                        Final_High = i
                        Final_High_Current = sd_MeasureData
                        Final_High_Voltage = argv(0) + Final_High * argv(4)
                    End If
                End If
                'Low Search (<0uA) --> Larger than -1.92uA (Ex: -1.05uA)
                If sd_MeasureData < 0 And sd_MeasureData > -1.92 * 0.000001 Then
                    sd_DeltaValue = Abs(sd_MeasureData - (-1.92 * 0.000001))
                    If sd_DeltaValue < temp_Low Then
                        temp_Low = sd_DeltaValue
                        Final_Low = i
                        Final_Low_Current = sd_MeasureData
                        Final_Low_Voltage = argv(0) + Final_Low * argv(4)
                    End If
                End If
                
            Next i
            
       If TheExec.TesterMode = testModeOffline Then
          Final_High_Current(site) = 5
          Final_Low_Current(site) = 4
       End If
            
       Dim m As New SiteDouble
       Dim b As New SiteDouble
       Dim Vt As New SiteDouble
       Dim T2NMOSVt As New SiteDouble
       Dim T3PMOSVt As New SiteDouble

            
       If (Final_High_Current(site) - Final_Low_Current(site)) = 0 Then
        Final_High_Current(site) = Final_High_Current(site) + 0.00000001
       End If
            m(site) = (Final_High_Voltage(site) - Final_Low_Voltage(site)) / (Abs(Final_High_Current(site) * 1000000#) - Abs(Final_Low_Current(site) * 1000000#))
            
            b(site) = Final_Low_Voltage(site) - m(site) * Abs(Final_Low_Current(site)) * 1000000
            
            Vt(site) = m(site) * 1.93 + b(site)
            
            T2NMOSVt(site) = Vt(site)
            
            T3PMOSVt(site) = 0.75 - Vt(site)
                  
            
        Next site
       

    
        Dim TName_Ary() As String
        'High Voltage Result
        TnameInput = Report_TName_From_Instance(CalcV, "X", , , , , , , tlForceNone)
        TName_Ary = Split(TnameInput, "_")
        TName_Ary(8) = "High"
        TnameInput = Join(TName_Ary, "_")   'EX: HAC_CalcV_N_CZT2PCORE_VTHSENSOR_X_X_dpcsocchpcmwrapctrlvthsnscfgV00010100&dpcsocchpcmwrapctrlvthsnsnostressV0_High_0
        TheExec.flow.TestLimit Final_High_Voltage, 0, , , , , unitVolt, , TnameInput, , SplitLeftBigColon(0), , , , , tlForceNone       'Search Nothing would '0V' --> Lolimit
        'High Current Result
        TName_Ary(1) = "CalcI"
        TnameInput = Join(TName_Ary, "_")   'EX: HAC_CalcI_N_CZT2PCORE_VTHSENSOR_X_X_dpcsocchpcmwrapctrlvthsnscfgV00010100&dpcsocchpcmwrapctrlvthsnsnostressV0_High_0
        TheExec.flow.TestLimit Final_High_Current, , 0, , , , unitAmp, , TnameInput, , pl_MeasureDict.pins(0).name, , , , , tlForceNone        'Search Nothing would '0A' --> Hilimit
        'Low Voltage Result
        TName_Ary(1) = "CalcV"
        TName_Ary(8) = "Low"
        TnameInput = Join(TName_Ary, "_")   'EX: HAC_CalcV_N_CZT2PCORE_VTHSENSOR_X_X_dpcsocchpcmwrapctrlvthsnscfgV00010100&dpcsocchpcmwrapctrlvthsnsnostressV0_Low_0
        TheExec.flow.TestLimit Final_Low_Voltage, 0, , , , , unitVolt, , TnameInput, , SplitLeftBigColon(0), , , , , tlForceNone
        'Low Current Result
        TName_Ary(1) = "CalcI"
        TnameInput = Join(TName_Ary, "_")   'EX: HAC_CalcI_N_CZT2PCORE_VTHSENSOR_X_X_dpcsocchpcmwrapctrlvthsnscfgV00010100&dpcsocchpcmwrapctrlvthsnsnostressV0_Low_0
        TheExec.flow.TestLimit Final_Low_Current, , 0, , , , unitAmp, , TnameInput, , pl_MeasureDict.pins(0).name, , , , , tlForceNone
    
    
    
        For Each site In TheExec.sites.Active
            TheExec.Datalog.WriteComment "***site(" & site & "),m= " & m & " Volt ***"
            TheExec.Datalog.WriteComment "***site(" & site & "),b= " & b & " Volt ***"
        Next site
    
  
        'T2 Voltage Result
        If Split(TheExec.DataManager.instancename, "_")(1) Like "*T2*" Then
            TName_Ary(1) = "CalcV"
            TName_Ary(8) = "T2-NMOS-Vt"
            TnameInput = Join(TName_Ary, "_")   'EX: HAC_CalcV_N_CZT2PCORE_VTHSENSOR_X_X_dpcsocchpcmwrapctrlvthsnscfgV00010100&dpcsocchpcmwrapctrlvthsnsnostressV0_Low_0
            TheExec.flow.TestLimit T2NMOSVt, 0, , , , , , , TnameInput, , SplitLeftBigColon(0), , , , , tlForceNone
        'T3 Voltage Result
        ElseIf Split(TheExec.DataManager.instancename, "_")(1) Like "*T3*" Then
            TName_Ary(1) = "CalcV"
            TName_Ary(8) = "T3-PMOS-Vt"
            TnameInput = Join(TName_Ary, "_")   'EX: HAC_CalcV_N_CZT2PCORE_VTHSENSOR_X_X_dpcsocchpcmwrapctrlvthsnscfgV00010100&dpcsocchpcmwrapctrlvthsnsnostressV0_Low_0
            TheExec.flow.TestLimit T3PMOSVt, 0, , , , , , , TnameInput, , SplitLeftBigColon(0), , , , , tlForceNone
        End If


    
    
    End If
    
    
    
End Function
Public Function Calc_DigCap_Avg_Store_STDEV_IVDM(argc As Integer, argv() As String) As Long
    Dim i As Long
    'Dim site As Variant
    '''Syntax: Alg::Calc_DigCap_Avg_Store(aneivdm_1,...,aneivdm_30, ("2SCOMPLEMENT"), aneivdm_trim_low1)
        '''                     aneivdm_trim_low1 : Dict Store Name used for Calc_Eqn "Calc_DigCap_Offset_Store"
        Dim DSPWave_Bin() As New DSPWave
    Dim DSPWave_Dec() As New DSPWave
    Dim DSPWave_Bin_MSB1st() As New DSPWave
    ReDim DSPWave_Bin(argc - 3) As New DSPWave
    ReDim DSPWave_Dec(argc - 3) As New DSPWave
    ReDim DSPWave_Bin_MSB1st(argc - 3) As New DSPWave
    'DSPWave_Dec(0).CreateConstant 0, 1, DspDouble
    Dim DSPWave_AverageDec As New DSPWave
    DSPWave_AverageDec.CreateConstant 0, 1, DspDouble
    
    Dim DSPWave_SumDec As New DSPWave
    DSPWave_SumDec.CreateConstant 0, 1, DspLong
    
    Dim DSPWave_AverageBin As New DSPWave
    'DSPWave_AverageBin.CreateConstant 0, 18
    
    '======@YM 20221117 IVDM_TTR======
    Dim IVDM_DEC_STDEV As New PinListData
    Dim IVDM_DEC As New SiteDouble
    Dim IVDM_STDEV As New SiteDouble
    Dim IVDM_Mean As New SiteDouble
    '======@YM 20221117 IVDM_TTR======
    Dim TestNameInput As String
    
    Dim SL_BitWidth As New SiteLong
    
    Dim Is_2sComplement As Boolean: Is_2sComplement = False
    TheHdw.dsp.ExecutionMode = tlDSPModeHostDebug
    
    If UCase(argv(argc - 2)) = "2SCOMPLEMENT" Then Is_2sComplement = True
    
    For i = 0 To argc - 4
        DSPWave_Bin(i) = GetStoreDataAllType(argv(i))
        '======@YM 20221117 IVDM_TTR======
        IVDM_DEC_STDEV.AddPin (argv(i) & "_para")
        IVDM_DEC_STDEV.pins(argv(i) & "_para").value = GetStoreDataAllType(argv(i) & "_para")
        '======@YM 20221117 IVDM_TTR======
        For Each site In TheExec.sites
            SL_BitWidth(site) = DSPWave_Bin(i)(site).SampleSize
        Next site
        
        If UCase(Instance_Data.CUS_Str_MainProgram) = UCase("DigCap_LSBtoMSB") Then
            rundsp.Split_Dspwave_Reverse DSPWave_Bin(i), DSPWave_Bin_MSB1st(i)
            DSPWave_Bin(i) = DSPWave_Bin_MSB1st(i)
        End If
        
        DSPWave_Dec(i).CreateConstant 0, 1, DspLong
        
        If Is_2sComplement = True Then
            Call rundsp.DSP_2S_Complement_To_SignDec(DSPWave_Bin(i), SL_BitWidth, DSPWave_Dec(i))
        Else
            Call rundsp.BinToDec(DSPWave_Bin(i), DSPWave_Dec(i))
        End If
        
        For Each site In TheExec.sites
            DSPWave_Dec(i)(site) = DSPWave_Dec(i)(site).ConvertDataTypeTo(DspLong)
        Next site
        Call rundsp.DSP_Add(DSPWave_SumDec, DSPWave_Dec(i))
    Next i
    For Each site In TheExec.sites
        DSPWave_AverageDec = DSPWave_SumDec.ConvertDataTypeTo(DspDouble)
    Next site
    Call rundsp.DSP_DivideConstant(DSPWave_AverageDec, argc - 2)

    For Each site In TheExec.sites
        TheExec.Datalog.WriteComment "Site : " & site & ", Average result:" & DSPWave_AverageDec(site).Element(0)
        DSPWave_AverageDec(site).Element(0) = FormatNumber(DSPWave_AverageDec(site).Element(0), 0)
    Next site
    
    Call rundsp.DSPWf_Dec2Binary(DSPWave_AverageDec, SL_BitWidth, DSPWave_AverageBin)
    'Call StoreDataAllType(argv(argc - 1), DSPWave_AverageDec)

    TestNameInput = Report_TName_From_Instance(CalcC, "X", , CInt(i), , , , , tlForceFlow)

    TheExec.flow.TestLimit resultVal:=DSPWave_AverageDec.Element(0), Tname:=TestNameInput, ForceResults:=tlForceFlow  'tlForceFlow
    Call StoreDataAllType(argv(argc - 2), DSPWave_AverageBin)
    '======@YM 20221117 IVDM_TTR======
    IVDM_STDEV = IVDM_DEC_STDEV.Analyze.stdDev
    TestNameInput = Report_TName_From_Instance(CalcC, "X", , CInt(i), , , , , tlForceFlow)
    TheExec.flow.TestLimit resultVal:=IVDM_STDEV, Tname:=TestNameInput, ForceResults:=tlForceFlow  'tlForceFlow
    '======@YM 20221117 IVDM_TTR======
    TheHdw.dsp.ExecutionMode = tlDSPModeAutomatic
End Function

Public Function Calc_ACAL_TCOMP(argc As Integer, argv() As String) As Long
    'From T-Col
    '---- New Function : Use for RLXOSC --
    'MeasF Pin = I2S0_DIN (F1)"LPRO_FOUT_F1"

    'Alg::Calc_ACAL_TCOMP(DICT_F1,TEMP_FT1,TEMP_FT2)
    'Ex. Alg::Calc_ACAL_TCOMP(F1,25.3,105.3)
    'Formula : TCOMP = ROUND{5-(F1-1MHz)*150/(Temp(FT2)-Temp(FT1))/(1MHz*0.001)}

    Dim FREQ_Value As New SiteDouble
    'Dim FREQ_Value_ENG As New SiteDouble
    Dim FREQ_Value_ENG As New PinListData
    Dim TEMP_FT1 As Double
    Dim TEMP_FT2 As Double
    Dim TEMP_DIFF As Double
    
    Dim TCOMP_Value As New SiteDouble
    Dim vsite As Variant
    Dim TestNameInput As String
    Dim ENG_String As String
    'Dim DSPWave_Dict As New DSPWave
    
    On Error GoTo errHandler
    
    If TheExec.TesterMode = testModeOffline Then
        For Each vsite In TheExec.sites
            FREQ_Value_ENG.AddPin ("dummy") ''''''''''RRRRRRRRRRRR
            FREQ_Value_ENG.pins("dummy").value = 976000# 'Typ=1MHz
        Next vsite
        ENG_String = argv(0)
        Call AddStoredMeasurement(ENG_String, FREQ_Value_ENG)
     End If
    
    FREQ_Value = GetStoredMeasurement(argv(0))
    TEMP_FT1 = CDbl(argv(1))
    TEMP_FT2 = CDbl(argv(2))
    TEMP_DIFF = TEMP_FT2 - TEMP_FT1
    
    'Formula : TCOMP = ROUND{5-(F1-1MHz)*150/(Temp(FT2)-Temp(FT1))/(1MHz*0.001)}
    If TEMP_DIFF <> 0 Then
        'TCOMP_Value(Vsite) = ((FREQ_Value(Vsite) - 1000000#) * 150) / TEMP_DIFF
        FREQ_Value = FREQ_Value.Subtract(1000000#)
        FREQ_Value = FREQ_Value.Multiply(150)
        FREQ_Value = FREQ_Value.divide(TEMP_DIFF)
        FREQ_Value = FREQ_Value.divide(1000000# * 0.001)
        For Each vsite In TheExec.sites
            TCOMP_Value(vsite) = Round(5 - FREQ_Value(vsite))
        Next vsite
    Else
        TheExec.Datalog.WriteComment ("Error! <<TEMP_FT2-TEMP_FT1 = 0>>, Divide 0.")
        TCOMP_Value(vsite) = 9999
    End If

    'Print out datalog
    TestNameInput = Report_TName_From_Instance(CalcC, "X", , , , , , , tlForceNone)
    TheExec.flow.TestLimit resultVal:=TCOMP_Value, PinName:="X", Tname:=TestNameInput, scaletype:=scaleNone, ForceResults:=tlForceNone
    
    Exit Function

errHandler:
    TheExec.Datalog.WriteComment "error in Calc_ACAL_TCOMP"
    If AbortTest Then Exit Function Else Resume Next
End Function
'Alg::Calc_DSP_Dictionary_Process(111&Dict_A:0:2,Dict_B)
Public Function Calc_DSP_Dictionary_Process(argc As Integer, argv() As String) As Long
On Error GoTo errHandler
'Alg::Calc_DSP_Dictionary_Process(111&Dict_A:0:2,Dict_B)
    'New Calc for T-Don --20230118
    'Reference DigSrc Setup syntax to convert DSPWave data for eFuse
    'Case support Dictionary, "&", ":", "copy"
    'Ex. DicA&DictA:19:19:copy:5
    '    0111&DictA:19:19:copy:5
    
    Dim i, j, k As Long
    Dim vsite As Variant
    Dim Input_Str As String
    Dim Input_StoreName As String
    Dim Output_StoreName As String
    Dim SplitByAnd() As String
    Dim RdIn() As String
    Dim b_WithDictionary As Boolean
    Dim DSPWF_DictData_Ary() As New DSPWave
    Dim DSPWF_DictData As New DSPWave
    Dim DSPWF_TempData As New DSPWave
    Dim DSPWF_TempDataForCopy As New DSPWave
    Dim DSPWF_OutputData As New DSPWave
    
    Dim Str_TempData As String
    Dim StartBit As String
    Dim StopBit As String
    Dim Format As String
    Dim CopyTimes As String
    
    
    On Error GoTo errHandler
    Set DSPWF_DictData = Nothing
    Set DSPWF_TempData = Nothing
    Set DSPWF_TempDataForCopy = Nothing
    Set DSPWF_OutputData = Nothing
    
    TheExec.Datalog.WriteComment "-----     Calc_DSP_Dictionary_Process Start     -----"
    
    Input_Str = argv(0)
    Output_StoreName = CStr(argv(1))
    
    
'    If TheExec.TesterMode = testModeOffline Then
'        Dim Dummy_Str As String: Dummy_Str = "0001110111001110"
'        Dim Dummy_DSPWF As New DSPWave
'        Dummy_DSPWF.CreateConstant 0, Len(Dummy_Str), DspDouble
'
'        For Each site In TheExec.sites
'            For i = 0 To Len(Dummy_Str) - 1
'                Dummy_DSPWF.Element(i) = CDbl(mid(Dummy_Str, i + 1, 1))
'            Next i
'        Next site
'        Call StoreDataAllType("Dict_A", Dummy_DSPWF)
'    End If
    If InStr(Input_Str, "&") <> 0 Or InStr(Input_Str, "@") <> 0 Then
        SplitByAnd = Split(Input_Str, "&")
        
        For i = 0 To UBound(SplitByAnd)
            RdIn = Split(SplitByAnd(i), "@")
            b_WithDictionary = Checker_WithDictionary(RdIn(0))
                
            If b_WithDictionary Then
                DSPWF_DictData = GetStoreDataAllType(RdIn(0))
            Else
                DSPWF_TempData.CreateConstant 0, Len(RdIn(0)), DspLong
                For Each vsite In TheExec.sites
                    For j = 0 To Len(RdIn(0)) - 1
                        DSPWF_TempData.Element(j) = CDbl(mid(RdIn(0), j + 1, 1))
                    Next j
                Next vsite
            End If
       
            '==== Copy Format ====
            If UBound(RdIn) = 4 Then
                StartBit = RdIn(1)
                StopBit = RdIn(2)
                Format = RdIn(3)
                CopyTimes = RdIn(4)
    
                DSPWF_TempDataForCopy.CreateConstant 0, CDbl(StopBit - StartBit) + 1, DspDouble
                
                For Each vsite In TheExec.sites
                    For j = StartBit To StopBit
                        DSPWF_TempDataForCopy.Element(j - CDbl(StartBit)) = DSPWF_DictData.Element(j)
                    Next j
                
                    For k = 1 To CopyTimes
                        If k = 1 Then
                            DSPWF_TempData = DSPWF_TempDataForCopy.COPY
                        Else
                            DSPWF_TempData = DSPWF_TempData.Concatenate(DSPWF_TempDataForCopy)
                        End If
                    Next k
                Next vsite
            ElseIf b_WithDictionary Then
                For Each vsite In TheExec.sites
                    DSPWF_TempData = DSPWF_DictData.COPY
                Next vsite
            Else
            End If
            
            For Each vsite In TheExec.sites
                If i = 0 Then
                    DSPWF_OutputData = DSPWF_TempData.COPY
                Else
                    DSPWF_OutputData = DSPWF_OutputData.Concatenate(DSPWF_TempData)
                End If
            Next vsite
            
            Set DSPWF_DictData = Nothing
            Set DSPWF_TempData = Nothing
            Set DSPWF_TempDataForCopy = Nothing
        Next i
    ElseIf InStr(Input_Str, "=") <> 0 Then
        RdIn = Split(Input_Str, "=")
        b_WithDictionary = Checker_WithDictionary(RdIn(0))
            
        If b_WithDictionary Then
            DSPWF_DictData = GetStoreDataAllType(RdIn(0))
            DSPWF_TempData.CreateConstant 0, Len(RdIn(1)), DspLong
            
            For Each vsite In TheExec.sites
                For i = 0 To Len(RdIn(1)) - 1
                    DSPWF_TempData.Element(i) = CDbl(mid(RdIn(1), i + 1, 1))
                Next i
            Next vsite
            For Each vsite In TheExec.sites
                DSPWF_OutputData = DSPWF_TempData.COPY
            Next vsite
            
        Else
            TheExec.Datalog.WriteComment "Error! Calc_DSP_Dictionary_Process : Format is incorrect"
            Exit Function
        End If
        Set DSPWF_DictData = Nothing
        Set DSPWF_TempData = Nothing
    
    ElseIf InStr(Input_Str, ":") <> 0 Then
        '20240131 Hidra by YM
        'Alg::Calc_DSP_Dictionary_Process(Dict_A:0:2,Dict_B)
        RdIn = Split(Input_Str, ":")
        Dim start_bit As New SiteLong
        Dim length As New SiteLong
        
        start_bit = CLng(RdIn(1))
        length = (CLng(RdIn(2)) - CLng(RdIn(1)) + 1)
        
        b_WithDictionary = Checker_WithDictionary(RdIn(0))
        If b_WithDictionary Then
            DSPWF_DictData = GetStoreDataAllType(RdIn(0))
        End If
        rundsp.DSP_SelectBits DSPWF_DictData, DSPWF_OutputData, start_bit, length
        Set DSPWF_DictData = Nothing
        Set DSPWF_TempData = Nothing
    Else
        If IsNumeric(Input_Str) = True Then
        DSPWF_TempData.CreateConstant 0, Len(Input_Str), DspLong
            
        For Each vsite In TheExec.sites
            For i = 0 To Len(Input_Str) - 1
                DSPWF_TempData.Element(i) = CDbl(mid(Input_Str, i + 1, 1))
            Next i
        Next vsite
        Else
            DSPWF_TempData = GetStoreDataAllType(Input_Str)
        End If
        For Each vsite In TheExec.sites
            DSPWF_OutputData = DSPWF_TempData.COPY
        Next vsite
        Set DSPWF_DictData = Nothing
        Set DSPWF_TempData = Nothing
    End If
    
    If gl_Disable_HIP_debug_log = False Then
        Dim s_result As String
        TheExec.Datalog.WriteComment "Output DSP Process String [ LSB(L) ==> MSB(R) ]:"
        For Each site In TheExec.sites.Active
            s_result = ""
            For i = 0 To DSPWF_OutputData.SampleSize - 1
                s_result = s_result & CStr(DSPWF_OutputData.Element(i))
            Next i
            TheExec.Datalog.WriteComment "Site [" & site & "] " & "Process DSP : " & Output_StoreName & " = " & Input_Str & " = " & s_result
        Next site
    End If
    If UCase(TheExec.DataManager.instancename) Like "*NTS_NTSREADT1P1*" Then
        TheExec.flow.TestLimit resultVal:=1, Tname:="CHK_" & Output_StoreName, ForceResults:=tlForceFlow
        Else
        TheExec.flow.TestLimit resultVal:=1, lowVal:=1, hiVal:=1, Tname:="CHK_" & Output_StoreName, ForceResults:=tlForceNone
    End If
    '===== Store Final Data to Dictionary ======
    Call StoreDataAllType(Output_StoreName, DSPWF_OutputData)
    
    Exit Function
    
errHandler:
    TheExec.Datalog.WriteComment "error in Calc_DSP_Dictionary_Process"
    If AbortTest Then Exit Function Else Resume Next
End Function




Public Function Calc_MetrologyBTS_Temperature_T5P4(argc As Integer, argv() As String) As Long

    Dim Temperature As New SiteDouble
    Dim Temperature_Sensor() As String: Temperature_Sensor = Split(argv(0), "+")
    Dim TemperatureTemp As Double
    Dim Temperature_Array_1() As New SiteDouble: ReDim Temperature_Array_1(UBound(Split(argv(0), "+")))
    Dim Temperature_Array_2() As New SiteDouble: ReDim Temperature_Array_2(UBound(Split(argv(0), "+")))
    Dim Temperature_Array_3() As New SiteDouble: ReDim Temperature_Array_3(UBound(Split(argv(0), "+")))
'    Dim Temperature_Array_SiteDouble() As New SiteDouble: ReDim Temperature_Array_SiteDouble(UBound(Split(argv(0), "+")))
    Dim VoltageValue As Double
    Dim DivValue As New SiteDouble
'    Dim DSP_Temperature() As New DSPWave: ReDim DSP_Temperature(UBound(Temperature_Sensor))
'    Dim DSP_MTRSNS_Temperature() As New DSPWave: ReDim DSP_MTRSNS_Temperature(UBound(Temperature_Sensor))
'    Dim DSP_MTRSNS_Temperature_eFuse() As New DSPWave: ReDim DSP_MTRSNS_Temperature_eFuse(UBound(Temperature_Sensor))
    Dim i As Long
    Dim j As Long
    Dim site As Variant
    Dim TestNameInput As String
'    Dim Temperature_Dictionary() As String
'    Dim Sensor_Num() As String
    
'    Dim DSP_MTRSNS_Temperature_Average As New DSPWave
'    Dim DSP_MTRSNS_Temperature_Maximum As New DSPWave
'    Dim DSP_MTRSNS_Temperature_Minimum As New DSPWave
'    Dim DSP_MTRSNS_Temperature_Maximum_2 As New DSPWave
    
'    DSP_MTRSNS_Temperature_Average.CreateConstant 0, UBound(Temperature_Sensor) + 1
'    DSP_MTRSNS_Temperature_Maximum.CreateConstant 0, UBound(Temperature_Sensor) + 1
'    DSP_MTRSNS_Temperature_Minimum.CreateConstant 0, UBound(Temperature_Sensor) + 1
'    DSP_MTRSNS_Temperature_Maximum_2.CreateConstant 0, UBound(Temperature_Sensor) + 1
'    For Each Site In TheExec.sites
        For j = 1 To 3
            For i = 0 To UBound(Temperature_Sensor)
                Temperature = GetStoreDataAllType(Temperature_Sensor(i) & "_" & j & "_para")
                For Each site In TheExec.sites
                
'                    If Temperature >= 65536 Then
'                        TemperatureTemp = Abs((Temperature - 131072) / 131072)
'                    Else
'                        TemperatureTemp = Abs(Temperature / 131072)
'                    End If
                    
                    If Temperature >= 65536 Then
                        TemperatureTemp = ((Temperature - 131072) / 131072) * -1
                    Else
                        TemperatureTemp = (Temperature / 131072) * -1
                    End If
                    
                    
                    
                    If LCase(Temperature_Sensor(i) & "_" & j) Like "*_tp*" Or LCase(Temperature_Sensor(i) & "_" & j) Like "*_te*" Then
                        VoltageValue = TheHdw.DCVS.pins("VDD_FIXED_PCPU_MTR").Voltage.value
                        TemperatureTemp = TemperatureTemp * VoltageValue
                    ElseIf LCase(Temperature_Sensor(i) & "_" & j) Like "*_ts*" Or LCase(Temperature_Sensor(i) & "_" & j) Like "*_tg*" Then
                        VoltageValue = TheHdw.DCVS.pins("VDD_SRAM_SOC").Voltage.value
                        TemperatureTemp = TemperatureTemp * VoltageValue
                    ElseIf LCase(Temperature_Sensor(i) & "_" & j) Like "*_ta*" Then
                        VoltageValue = TheHdw.DCVS.pins("VDD_FIXED_GRP").Voltage.value
                        TemperatureTemp = TemperatureTemp * VoltageValue
                    End If
                        
                    If j = 1 Then
                        Temperature_Array_1(i) = TemperatureTemp
                    ElseIf j = 2 Then
                        Temperature_Array_2(i) = TemperatureTemp
                    ElseIf j = 3 Then
                        Temperature_Array_3(i) = TemperatureTemp
                    End If
                Next site
            Next i
        Next j
'    Next Site
        For i = 0 To UBound(Temperature_Sensor)
                TestNameInput = Report_TName_From_Instance("CalcC", "")
                TheExec.flow.TestLimit resultVal:=Temperature_Array_1(i), Tname:=TestNameInput, ForceResults:=tlForceFlow
        Next i
        For j = 1 To 2
            For i = 0 To UBound(Temperature_Sensor)
            For Each site In TheExec.sites
                    If j = 1 Then
                            DivValue = Temperature_Array_2(i) - Temperature_Array_1(i)
                    ElseIf j = 2 Then
                            DivValue = Temperature_Array_3(i) - Temperature_Array_1(i)
                    End If
            Next site
                    TestNameInput = Report_TName_From_Instance("CalcC", "")
                    TheExec.flow.TestLimit resultVal:=DivValue, Tname:=TestNameInput, ForceResults:=tlForceFlow
            Next i
        Next j
    
    
Exit Function
errHandler:
    TheExec.Datalog.WriteComment "error in Calc_MetrologyBTS_Temperature"
    If AbortTest Then Exit Function Else Resume Next

End Function

Public Function Calc_VDAC(argc As Integer, argv() As String) As Long
    'From T-Col

    On Error GoTo errHandler
    
    Dim TestNameInput As String
    Dim i, Resolution, ratio As Double
    Dim VDAC_Value As New SiteDouble
    
    Resolution = argv(argc - 2)
    ratio = argv(argc - 1)
    
    For i = 0 To argc - 3
        VDAC_Value = GetStoreDataAllType(argv(i) & "_para")
        VDAC_Value = VDAC_Value.divide(Resolution).Multiply(ratio)
        
        
        'TestNameInput = Report_TName_From_Instance(CalcC, "", argv(i), , , , , , tlForceFlow)
        TestNameInput = Report_TName_From_Instance(CalcC, "", , , , , , , tlForceFlow)
        TheExec.flow.TestLimit resultVal:=VDAC_Value, PinName:="", Tname:=TestNameInput, scaletype:=scaleNone, ForceResults:=tlForceFlow
    Next i
        
    Exit Function

errHandler:
    TheExec.Datalog.WriteComment "error in Calc_VDAC"
    If AbortTest Then Exit Function Else Resume Next
End Function
Public Function Calc_MDLL_DCD(argc As Integer, argv() As String) As Long
    'argv(0) = capture data of dout1
    'argv(1) = capture data of dout2
    'argv(2) = factor
    'argv(3) = Tper
    
    Dim dout1 As New SiteDouble
    Dim dout2 As New SiteDouble
    Dim dout1_4094 As New SiteDouble
    Dim dout2_4094 As New SiteDouble
    Dim dout1_abs As New SiteDouble
    Dim dout2_abs As New SiteDouble
    Dim dllpulse_err As New SiteDouble
    Dim dcAvg As New SiteDouble
    Dim Tper As New SiteDouble
    Dim LowLimit As Long: LowLimit = 0
    Dim hiLimit As Long: hiLimit = 0
    Dim TestNameInput As String
    Dim Splt_str() As String
    Dim Splt_str2() As String
    Dim tmp_str() As String
     
     
    dout1 = GetStoreDataAllType(argv(0) & "_para")
    dout2 = GetStoreDataAllType(argv(1) & "_para")
    
    dout1_4094 = dout1.Negate.Add(4094)
    dout2_4094 = dout2.Negate.Add(4094)
    
    dout1_abs = dout1.Maximum(dout1_4094)
    dout2_abs = dout2.Maximum(dout2_4094)
    
    Tper = CDbl(argv(3)) * 10 ^ -12
    
    dcAvg = dout1_abs.Add(dout2_abs).Multiply(0.5).divide(4094).Negate.Add(1)
    
    dllpulse_err = dcAvg.Multiply(argv(2)).Subtract(1).Multiply(argv(3))
    For Each site In TheExec.sites
        TheExec.Datalog.WriteComment "Site: " & site & " The DC_AVG is: " & dcAvg(site)
    Next site
    Splt_str = Split(argv(0), "__")      ' Use for combin TName [ Format : ddr0__ch0_dq0_rdmdll_selftest_dout1 ]
    
    TestNameInput = Report_TName_From_Instance(CalcC, "X", , , , , , , tlForceNone)
    Splt_str2 = Split(TestNameInput, "_")
    tmp_str = Split(Splt_str(1), "_")
    Splt_str2(6) = UCase(tmp_str(UBound(tmp_str) - 1)) & "-err"
    Splt_str2(7) = Replace(UCase(Splt_str(0)), "_", "-")
    'Splt_str2(7) = Replace(UCase(Splt_str(1)), "_", "-") ' ch0_dq0_rdmdll_selftest_dout1 -> ch0-dq0-rdmdll-selftest-dout1
    TestNameInput = Join(Splt_str2, "_")
    'TestNameInput = Report_TName_From_Instance(CalcC, "X", , , , , , , tlForceNone)
    
    If LCase(TheExec.DataManager.instancename) Like "*cz*md1*" Then
        LowLimit = -220
        hiLimit = 220
    ElseIf LCase(TheExec.DataManager.instancename) Like "*cz*md2*" Then
        LowLimit = -110
        hiLimit = 110
    ElseIf LCase(TheExec.DataManager.instancename) Like "*cz*md3*" Then
        LowLimit = -100
        hiLimit = 100
    ElseIf LCase(TheExec.DataManager.instancename) Like "*cz*md4*" Then
        LowLimit = -80
        hiLimit = 80
    ElseIf LCase(TheExec.DataManager.instancename) Like "*cz*md5*" Then
        LowLimit = -70
        hiLimit = 70
    ElseIf LCase(TheExec.DataManager.instancename) Like "*cz*md6*" Then
        LowLimit = -60
        hiLimit = 60
    End If
    
    TheExec.flow.TestLimit resultVal:=dllpulse_err, lowVal:=LowLimit, hiVal:=hiLimit, PinName:="X", Tname:=TestNameInput, scaletype:=scaleNone, ForceResults:=tlForceNone
    
End Function

Public Function Calc_PI_DCD_Duty_Cycle_AMPLP5(argc As Integer, argv() As String) As Long
    '---- New Function : Library function to calculate dcd duty cycle -- Neil, 20220406
    'Note :
    '1. Capture data must be convert to 2SCOMPLEMENT first
    '2. Define 2SCOMPLEMENT in "CUS_Str_MainProgram" argurment
    '3. Calculate the following parameters for each Instance and byte
    '   m1 = (step3_dcm_count_out ? step2_dcm_count_out )/100/6
    '   m2 = (step0_dcm_count_out ? step3_dcm_count_out )/100/6
    '   m3 = (step1_dcm_count_out ? step0_dcm_count_out )/100/6
    '   m4 = (step1_dcm_count_out ? step3_dcm_count_out )/100/6
    '   m = (m1 + m2 + m3 + m4) / 4
    '   k = step3_dcm_count_out - m * 50
    '   dc1 = (step4_dcm_count_out - k) / m
    '   dc2 = (step5_dcm_count_out - k) / m
    '
    
    Dim site As Variant
    Dim TestNameInput As String
    Dim Splt_str() As String
    Dim Splt_str2() As String

    Dim m1 As New SiteDouble
    Dim m2 As New SiteDouble
    Dim m3 As New SiteDouble
    Dim m4 As New SiteDouble
    Dim m As New SiteDouble
    Dim k As New SiteDouble
    Dim dc1 As New SiteDouble
    Dim dc2 As New SiteDouble
    Dim dc3 As New SiteDouble
    
    Dim step0_dcm_count_out As New SiteDouble
    Dim step1_dcm_count_out As New SiteDouble
    Dim step2_dcm_count_out As New SiteDouble
    Dim step3_dcm_count_out As New SiteDouble
    Dim step4_dcm_count_out As New SiteDouble
    Dim step5_dcm_count_out As New SiteDouble
    Dim step6_dcm_count_out As New SiteDouble

    On Error GoTo errHandler

    step0_dcm_count_out = GetStoreDataAllType(argv(0) & "_para")
    step1_dcm_count_out = GetStoreDataAllType(argv(1) & "_para")
    step2_dcm_count_out = GetStoreDataAllType(argv(2) & "_para")
    step3_dcm_count_out = GetStoreDataAllType(argv(3) & "_para")
    step4_dcm_count_out = GetStoreDataAllType(argv(4) & "_para")
    step5_dcm_count_out = GetStoreDataAllType(argv(5) & "_para")
    step6_dcm_count_out = GetStoreDataAllType(argv(6) & "_para")
       
    m1 = step3_dcm_count_out.Subtract(step2_dcm_count_out).divide(100 / 6)
    m2 = step0_dcm_count_out.Subtract(step3_dcm_count_out).divide(100 / 6)
    m3 = step1_dcm_count_out.Subtract(step0_dcm_count_out).divide(100 / 3)
    m4 = step1_dcm_count_out.Subtract(step3_dcm_count_out).divide(100 / 2)
    m = m1.Add(m2).Add(m3).Add(m4)
    m = m.divide(4)
    k = step3_dcm_count_out.Subtract(m.Multiply(50))
    dc1 = step4_dcm_count_out.Subtract(k).divide(m)
    dc2 = step5_dcm_count_out.Subtract(k).divide(m)
    dc3 = step6_dcm_count_out.Subtract(k).divide(m)

    'Print out datalog
    Dim i As Long
    Dim TName_Ary() As String
    Splt_str = Split(argv(0), "__")      ' Use for combin TName [ Format : DDR0_CH0_DQ0__step0_dcm_count_out ]
    
    TestNameInput = Report_TName_From_Instance(CalcC, "X", , , , , , , tlForceNone)
    Splt_str2 = Split(TestNameInput, "_")
    Splt_str2(7) = Replace(UCase(Splt_str(0)), "_", "-") ' DDR0_CH0_DQ0 -> DDR0-CH0-DQ0
    TestNameInput = Join(Splt_str2, "_")
    
    Splt_str2(6) = "m1"
    TestNameInput = Join(Splt_str2, "_")
    TheExec.flow.TestLimit resultVal:=m1, PinName:="X", Tname:=TestNameInput, scaletype:=scaleNone, ForceResults:=tlForceNone

    Splt_str2(6) = "m2"
    TestNameInput = Join(Splt_str2, "_")
    TheExec.flow.TestLimit resultVal:=m2, PinName:="X", Tname:=TestNameInput, scaletype:=scaleNone, ForceResults:=tlForceNone

    Splt_str2(6) = "m3"
    TestNameInput = Join(Splt_str2, "_")
    TheExec.flow.TestLimit resultVal:=m3, PinName:="X", Tname:=TestNameInput, scaletype:=scaleNone, ForceResults:=tlForceNone
    
    Splt_str2(6) = "m4"
    TestNameInput = Join(Splt_str2, "_")
    TheExec.flow.TestLimit resultVal:=m4, PinName:="X", Tname:=TestNameInput, scaletype:=scaleNone, ForceResults:=tlForceNone
    
    Splt_str2(6) = "m"
    TestNameInput = Join(Splt_str2, "_")
    TheExec.flow.TestLimit resultVal:=m, PinName:="X", Tname:=TestNameInput, scaletype:=scaleNone, ForceResults:=tlForceNone

    Splt_str2(6) = "k"
    TestNameInput = Join(Splt_str2, "_")
    TheExec.flow.TestLimit resultVal:=k, PinName:="X", Tname:=TestNameInput, scaletype:=scaleNone, ForceResults:=tlForceNone
    
    Splt_str2(6) = "dc1"
    TestNameInput = Join(Splt_str2, "_")
    TheExec.flow.TestLimit resultVal:=dc1, PinName:="X", Tname:=TestNameInput, scaletype:=scaleNone, ForceResults:=tlForceNone
    
    Splt_str2(6) = "dc2"
    TestNameInput = Join(Splt_str2, "_")
    TheExec.flow.TestLimit resultVal:=dc2, PinName:="X", Tname:=TestNameInput, scaletype:=scaleNone, ForceResults:=tlForceNone

    Splt_str2(6) = "dc3"
    TestNameInput = Join(Splt_str2, "_")
    TheExec.flow.TestLimit resultVal:=dc3, PinName:="X", Tname:=TestNameInput, scaletype:=scaleNone, ForceResults:=tlForceNone
    Exit Function
    
errHandler:
    TheExec.Datalog.WriteComment "error in Calc_PI_DCD_Duty_Cycle"
    If AbortTest Then Exit Function Else Resume Next
End Function
Public Function Calc_VDAC_Delta(argc As Integer, argv() As String) As Long
    'From T-Col
    '20240131 Hidra by YM
    'Add Delta function , Save Dictionary
    
    On Error GoTo errHandler
    
    Dim TestNameInput As String
    Dim i, j, Resolution, ratio As Double
    Dim VDAC_Value As New SiteDouble
    Dim Delta_Value As New SiteDouble
    Dim argu_ary() As String
    Dim storename As String
    Dim temp As String
    
    For i = 0 To argc - 1
    
        temp = argv(i)
        storename = ""
        'Calc_VDAC_Delta(ucsdm0_ts3__dout_1+4094+1.32:StoreA,ucsdm1_ts3__dout_1+4094+1.32:StoreB)
        If InStr(argv(i), ":") > 0 Then
            storename = Split(argv(i), ":")(1)
            temp = Split(argv(i), ":")(0)
        End If
        
        argu_ary = Split(temp, "+")
        Resolution = argu_ary(1)
        ratio = argu_ary(2)
        
        'Calc_VDAC_Delta(ucsdm0_ts3__dout_1+4094+1.32,ucsdm1_ts3__dout_1+4094+1.32)
        If UBound(argu_ary) = 2 Then
            VDAC_Value = GetStoreDataAllType(argu_ary(0) & "_para")
            VDAC_Value = VDAC_Value.divide(Resolution).Multiply(ratio)
            'TestNameInput = Report_TName_From_Instance(CalcC, "", argv(i), , , , , , tlForceFlow)
            TestNameInput = Report_TName_From_Instance(CalcC, "", , , , , , , tlForceFlow)
            TheExec.flow.TestLimit resultVal:=VDAC_Value, PinName:="", Tname:=TestNameInput, scaletype:=scaleNone, ForceResults:=tlForceFlow
        
        'Calc_VDAC_Delta(ucsdm0_ts3__dout_1+4094+1.32+0.66,ucsdm1_ts3__dout_1+4094+1.32+0.66)
        ElseIf UBound(argu_ary) = 3 Then
            VDAC_Value = GetStoreDataAllType(argu_ary(0) & "_para")
            VDAC_Value = VDAC_Value.divide(Resolution).Multiply(ratio)
            Delta_Value = VDAC_Value.Subtract(CDbl(argu_ary(3)))
            
            TestNameInput = Report_TName_From_Instance(CalcC, "", , , , , , , tlForceFlow)
            TheExec.flow.TestLimit resultVal:=VDAC_Value, PinName:="", Tname:=TestNameInput, scaletype:=scaleNone, ForceResults:=tlForceFlow
            
            TestNameInput = Report_TName_From_Instance(CalcC, "", , , , , , , tlForceFlow)
            TheExec.flow.TestLimit resultVal:=Delta_Value, PinName:="", Tname:=TestNameInput, scaletype:=scaleNone, ForceResults:=tlForceFlow
        End If
    
        If storename <> "" Then
            Call StoreDataAllType(storename, VDAC_Value)
        End If
        
    Next i
    
        
    Exit Function

errHandler:
    TheExec.Datalog.WriteComment "error in Calc_VDAC_Delta"
    If AbortTest Then Exit Function Else Resume Next
End Function




Public Function Calc_ADC_LDOT3P285C2C3_Store(argc As Integer, argv() As String) As Long

Dim DSPWave_Dict_soc_calculation As New DSPWave
Dim DSPWave_Dict_soc_C2 As New SiteDouble: DSPWave_Dict_soc_C2 = GetStoreDataAllType(argv(1) & "_para")
Dim DSPWave_Dict_soc_C3 As New SiteDouble: DSPWave_Dict_soc_C3 = GetStoreDataAllType(argv(2) & "_para")

Dim DSPWave_Dict_accp_calculation As New DSPWave
Dim DSPWave_Dict_accp_C2 As New SiteDouble: DSPWave_Dict_accp_C2 = GetStoreDataAllType(argv(4) & "_para")
Dim DSPWave_Dict_accp_C3 As New SiteDouble: DSPWave_Dict_accp_C3 = GetStoreDataAllType(argv(5) & "_para")

Dim DSPWave_Dict_acce_calculation As New DSPWave
Dim DSPWave_Dict_acce_C2 As New SiteDouble: DSPWave_Dict_acce_C2 = GetStoreDataAllType(argv(7) & "_para")
Dim DSPWave_Dict_acce_C3 As New SiteDouble: DSPWave_Dict_acce_C3 = GetStoreDataAllType(argv(8) & "_para")

Dim trim_code_size As Long

trim_code_size = CLng(argv(9))
DSPWave_Dict_soc_calculation.CreateConstant 0, 1, DspLong
DSPWave_Dict_accp_calculation.CreateConstant 0, 1, DspLong
DSPWave_Dict_acce_calculation.CreateConstant 0, 1, DspLong

For Each site In TheExec.sites

    DSPWave_Dict_soc_calculation.Element(0) = CLng(Format(((DSPWave_Dict_soc_C2 + DSPWave_Dict_soc_C3) / 2), "0"))
    If DSPWave_Dict_soc_calculation.Element(0) < 0 Then DSPWave_Dict_soc_calculation.Element(0) = DSPWave_Dict_soc_calculation.Element(0) + (2 ^ trim_code_size)
    
    DSPWave_Dict_accp_calculation.Element(0) = CLng(Format(((DSPWave_Dict_accp_C2 + DSPWave_Dict_accp_C3) / 2), "0"))
    If DSPWave_Dict_accp_calculation.Element(0) < 0 Then DSPWave_Dict_accp_calculation.Element(0) = DSPWave_Dict_accp_calculation.Element(0) + (2 ^ trim_code_size)
    
    DSPWave_Dict_acce_calculation.Element(0) = CLng(Format(((DSPWave_Dict_acce_C2 + DSPWave_Dict_acce_C3) / 2), "0"))
    If DSPWave_Dict_acce_calculation.Element(0) < 0 Then DSPWave_Dict_acce_calculation.Element(0) = DSPWave_Dict_acce_calculation.Element(0) + (2 ^ trim_code_size)

Next site

Call rundsp.DSPWaveDecToBinary(DSPWave_Dict_soc_calculation, 4, DSPWave_Dict_soc_calculation)
Call rundsp.DSPWaveDecToBinary(DSPWave_Dict_accp_calculation, 4, DSPWave_Dict_accp_calculation)
Call rundsp.DSPWaveDecToBinary(DSPWave_Dict_acce_calculation, 4, DSPWave_Dict_acce_calculation)

Call AddStoredCaptureData(argv(0), DSPWave_Dict_soc_calculation)
Call AddStoredCaptureData(argv(3), DSPWave_Dict_accp_calculation)
Call AddStoredCaptureData(argv(6), DSPWave_Dict_acce_calculation)

End Function

Public Function Calc_Eye_PCIE5(argc As Integer, argv() As String) As Long
'Alg::Calc_Eye_PCIE5(pcie1__eyescan_vth_at_vp_ln0+pcie1__eyescan_vth_at_vn_ln0+pcie1__eyescan_pi_at_hr_ln0+pcie1__eyescan_pi_at_hl_ln0+GEN3,
'pcie1__eyescan_vth_at_vp_ln1+pcie1__eyescan_vth_at_vn_ln1+pcie1__eyescan_pi_at_hr_ln1+pcie1__eyescan_pi_at_hl_ln1+GEN3)

On Error GoTo errHandler 'Add ErrHandler 2023/05/29
    Dim i As Integer
    Dim site As Variant
    Dim TestNameInput As String
    Dim SplitStrAry() As String
    Dim EH As New SiteDouble
    Dim EH_OS As New SiteDouble
    Dim EW As New SiteDouble
    Dim vth_at_vp As New SiteDouble
    Dim vth_at_vn As New SiteDouble
    Dim pi_at_hr As New SiteDouble
    Dim pi_at_hl As New SiteDouble
    Dim GEN As String
    Dim PI_STEP_SIZE As Double
    
    For i = 0 To argc - 1
        SplitStrAry = Split(argv(i), "+")
        vth_at_vp = GetStoreDataAllType(SplitStrAry(0) & "_para")
        vth_at_vn = GetStoreDataAllType(SplitStrAry(1) & "_para")
        pi_at_hr = GetStoreDataAllType(SplitStrAry(2) & "_para")
        pi_at_hl = GetStoreDataAllType(SplitStrAry(3) & "_para")
        GEN = UCase(SplitStrAry(4))
        
        '==========================
        '1.04ps  for Gen 3  -> PH3B
        '1.041ps for Gen 4  -> PH4B
        '0.521ps for Gen 5  -> PH5B
        '==========================
        Select Case GEN
            Case "GEN3"
                PI_STEP_SIZE = 1.04 * 10 ^ -12
            Case "GEN4"
                PI_STEP_SIZE = 1.041 * 10 ^ -12
            Case "GEN5"
                PI_STEP_SIZE = 0.521 * 10 ^ -12
        End Select
        
        TheExec.Datalog.WriteComment GEN & " >>> PI_STEP_SIZE = " & PI_STEP_SIZE
        
        'EH = (vth-at-vp - vth-at-vn) *1.75mV -> NOTE: VTH-at-vn is 2's complement. Need to convert this from readbaack value to decimal. Example: Readback 325 is actually -187
        EH = vth_at_vp.Subtract(vth_at_vn).Multiply(0.00175)
        
        TestNameInput = Report_TName_From_Instance("C", vbNullString, ForceResult:=tlForceFlow)
        TheExec.flow.TestLimit resultVal:=EH, Tname:=TestNameInput, ForceResults:=tlForceFlow
        
        'EH_OS = 0.5*(vth-at-vp + vth-at-vn) *1.75mV -> NOTE: VTH-at-vn is 2's complement
        EH_OS = vth_at_vp.Add(vth_at_vn).Multiply(0.00175).Multiply(0.5)
        
        TestNameInput = Report_TName_From_Instance("C", vbNullString, ForceResult:=tlForceFlow)
        TheExec.flow.TestLimit resultVal:=EH_OS, Tname:=TestNameInput, ForceResults:=tlForceFlow
        
        'EW = (pi-at-hr - pi-at-hl) * PI_STEP_SIZE; PI_STEP_SIZE for various gen is as follows:
        EW = pi_at_hr.Subtract(pi_at_hl).Multiply(PI_STEP_SIZE)
        
        TestNameInput = Report_TName_From_Instance("C", vbNullString, ForceResult:=tlForceFlow)
        TheExec.flow.TestLimit resultVal:=EW, Tname:=TestNameInput, ForceResults:=tlForceFlow
        
    Next i
    
Exit Function 'Add ErrHandler 2023/05/29
errHandler: 'Add ErrHandler 2023/05/29
    Call Print_Error_Message(Error_Info, "LIB_HardIP_Calc", "Calc_Eye_PCIE5") 'Add ErrHandler 2023/05/29
    If AbortTest Then Exit Function Else Resume Next 'Add ErrHandler 2023/05/29


End Function


Public Function Calc_NJP_AMP_Voltage(argc As Integer, argv() As String) As Long
    Dim site As Variant
    Dim i, j As Long
    Dim input_1 As New PinListData
    Dim output_1 As New PinListData
    Dim output_2 As New PinListData
    Dim output_2_split5 As New PinListData
    Dim output_2_split1 As New PinListData
    Dim bitsize As Integer
    Dim output_condition() As String
    Dim trim_tname As String
    Dim calc_result As New SiteDouble
    Dim calc_result_split5 As New SiteDouble
    Dim calc_result_split1 As New SiteDouble
    
    Dim input_2 As New SiteDouble
    Dim Parameter_1, Parameter_2, Parameter_3, Parameter_4 As New SiteDouble
    Dim npjs_vvdd_trim As String
    Dim Final_Calc As New SiteDouble
    Dim TestNameInput As String
    Dim round_flag As Long
    Dim cap_output As New DSPWave
    Dim cap_output_ldo As New DSPWave
    Dim cap_output_ldo_spare9 As New DSPWave
    Dim cap_output_dec2 As New DSPWave
    Dim cap_output_dec2_spare9 As New DSPWave
    Dim cap_output_dec As New DSPWave

    
    
    cap_output_dec.CreateConstant 0, 1, DspDouble
    cap_output_dec2.CreateConstant 0, 1, DspDouble
    cap_output_dec2_spare9.CreateConstant 0, 1, DspDouble

    input_1 = GetStoredMeasurement(argv(0))
    
    Parameter_1 = CDbl(argv(1))
    Parameter_2 = CDbl(argv(2))
    output_condition = Split(argv(3), "@")
    bitsize = CInt(output_condition(0))

   ' For Each site In theexec.sites
    
        For i = 0 To input_1.pins.Count - 1
            
            For Each site In TheExec.sites
                calc_result = Round((Parameter_1 / input_1.pins(i).value(site)) - Parameter_2)
                
            Next
            'calc_result = (Parameter_1 / input_1.Pins(i).value) - Parameter_2
           ' calc_result = input_1.Pins(i).Invert.Multiply(Parameter_1).Subtract(Parameter_2)
            
            
            
            
            If bitsize = 6 Then
            
'                If calc_result > 63 Then
'                    theexec.Datalog.WriteComment "Warning!! calculate result = " & calc_result & "> 63(6 bits), please check!!"
'                    calc_result = 63
'
'                End If
                
                For Each site In TheExec.sites
                    cap_output_dec.Element(0) = calc_result
                    If calc_result(site) > 63 Then
                        TheExec.Datalog.WriteComment "Warning!! calculate result" & "(" & site & ") =" & calc_result(site) & "> 63(6 bits), please check!!"
                        calc_result(site) = 63
                    ElseIf calc_result(site) < 0 Then
                        TheExec.Datalog.WriteComment "Warning!! calculate result" & "(" & site & ") =" & calc_result(site) & "< 0(6 bits), please check!!"
                        calc_result(site) = 63
                    End If
                    If calc_result(site) Mod 2 = 0 Then
                        cap_output_dec2(site).Element(0) = (calc_result) / 2
                        calc_result_split5(site) = (calc_result) / 2
                        cap_output_dec2_spare9(site).Element(0) = 0
                        calc_result_split1(site) = 0
                        
                    Else
                        cap_output_dec2(site).Element(0) = (calc_result - 1) / 2
                        calc_result_split5(site) = (calc_result - 1) / 2
                        cap_output_dec2_spare9(site).Element(0) = 1
                        calc_result_split1(site) = 1
                    End If
                    Call HardIP_Dec2Bin(cap_output, cap_output_dec, bitsize)
                    Call HardIP_Dec2Bin(cap_output_ldo, cap_output_dec2, bitsize - 1)
                    Call HardIP_Dec2Bin(cap_output_ldo_spare9, cap_output_dec2_spare9, 1)
                Next
'                If calc_result Mod 2 = 0 Then
'                    cap_output_dec2.Element(0) = (calc_result) / 2
'                    cap_output_dec2_spare9.Element(0) = 0
'                Else
'                    cap_output_dec2.Element(0) = (calc_result - 1) / 2
'                    cap_output_dec2_spare9.Element(0) = 1
'                End If

                
                'reDim cap_output As New DSPWave

                


                Call AddStoredCaptureData(input_1.pins(i) & "_" & output_condition(2), cap_output_ldo_spare9)
                Call AddStoredCaptureData(input_1.pins(i) & "_" & output_condition(1), cap_output_ldo)
                'Call GetStoreDataAllType(input_1.Pins(i) & "_" & output_condition(1))
                'Call GetStoreDataAllType(input_1.Pins(i) & "_" & output_condition(2))
                'trim_tname = Replace(output_condition(1), "amptrim", "trim")
                
                output_2.AddPin (input_1.pins(i))
                output_2_split5.AddPin (input_1.pins(i))
                output_2_split1.AddPin (input_1.pins(i))
                output_2.pins(input_1.pins(i)).value = calc_result
                output_2_split5.pins(input_1.pins(i)).value = calc_result_split5
                output_2_split1.pins(input_1.pins(i)).value = calc_result_split1
                trim_tname = Replace(output_condition(1), "amptrim", "trim")
                
                TestNameInput = Report_TName_From_Instance("CalcC", input_1.pins(i), output_condition(1), , , , , , ForceResult:=tlForceFlow)
                TheExec.flow.TestLimit resultVal:=output_2_split5, PinName:=input_1.pins(i), Tname:=TestNameInput, ForceResults:=tlForceFlow
                trim_tname = Replace(output_condition(1), "ldo", "spar9")
                TestNameInput = Report_TName_From_Instance("CalcC", input_1.pins(i), trim_tname, , , , , , ForceResult:=tlForceFlow)
                TheExec.flow.TestLimit resultVal:=output_2_split1, PinName:=input_1.pins(i), Tname:=TestNameInput, ForceResults:=tlForceFlow
            ElseIf bitsize = 5 Then

                'cap_output_dec.Element(0) = calc_result
                
                For Each site In TheExec.sites
                     If calc_result(site) > 31 Then
                        TheExec.Datalog.WriteComment "Warning!! calculate result" & "(" & site & ") =" & calc_result(site) & "> 31(5 bits), please check!!"
                        calc_result(site) = 31
                    ElseIf calc_result(site) < 0 Then
                        TheExec.Datalog.WriteComment "Warning!! calculate result" & "(" & site & ") =" & calc_result(site) & "< 0(5 bits), please check!!"
                        calc_result(site) = 31
                    End If
                    cap_output_dec.Element(0) = calc_result(site)
                    Call HardIP_Dec2Bin(cap_output, cap_output_dec, bitsize)
                Next site
                
                trim_tname = output_condition(1)
                           
                Call AddStoredCaptureData(input_1.pins(i) & "_" & output_condition(1), cap_output)
                'Call GetStoreDataAllType(input_1.Pins(i) & "_" & output_condition(1))
                output_2.AddPin (input_1.pins(i))
                output_2.pins(input_1.pins(i)).value = calc_result
                TestNameInput = Report_TName_From_Instance("CalcC", input_1.pins(i), trim_tname, , , , , , ForceResult:=tlForceFlow)
    
                TheExec.flow.TestLimit resultVal:=calc_result, PinName:=input_1.pins(i), Tname:=TestNameInput, ForceResults:=tlForceFlow
                    
            End If
            

        Next i
  '  Next site
    
    
'    If bitsize = 6 Then
'        For i = 0 To input_1.Pins.Count - 1
'            output_2.AddPin (input_1.Pins(i))
'            output_2.Pins(input_1.Pins(i)).value = calc_result
'            TestNameInput = Report_TName_From_Instance("CalcC", input_1.Pins(i), trim_tname, , , , , , ForceResult:=tlForceFlow)
'            theexec.Flow.TestLimit resultVal:=calc_result, PinName:=input_1.Pins(i), Tname:=TestNameInput, ForceResults:=tlForceFlow
'
'        Next i
'
'    Else
'        For i = 0 To input_1.Pins.Count - 1
'            output_2.AddPin (input_1.Pins(i))
'            output_2.Pins(input_1.Pins(i)).value = calc_result
'            TestNameInput = Report_TName_From_Instance("CalcC", input_1.Pins(i), trim_tname, , , , , , ForceResult:=tlForceFlow)
'
'            theexec.Flow.TestLimit resultVal:=calc_result, PinName:=input_1.Pins(i), Tname:=TestNameInput, ForceResults:=tlForceFlow
'
'        Next i
'
'    End If
    
    

    
End Function



Public Function Calc_VDAC_DD(argc As Integer, argv() As String) As Long
    'From T-Col

    On Error GoTo errHandler
    
    Dim TestNameInput As String
    Dim i, Resolution, ratio As Double
    Dim VDAC_Value As New SiteDouble
    Dim VDAC_Value2 As New SiteDouble
    Dim VDAC_Value3 As New SiteDouble
    Resolution = argv(argc - 2)
    ratio = argv(argc - 1)
    
    
    
    
    
    For i = 1 To 3
        
        If i = 3 Then
            VDAC_Value = GetStoreDataAllType(argv(1) & "_para")
            VDAC_Value = VDAC_Value.divide(Resolution).Multiply(ratio)
            VDAC_Value2 = GetStoreDataAllType(argv(2) & "_para")
            VDAC_Value2 = VDAC_Value2.divide(Resolution).Multiply(ratio)
            
            VDAC_Value3 = VDAC_Value.Subtract(VDAC_Value2)
            
            TestNameInput = Report_TName_From_Instance(CalcC, "", , , , , , , tlForceFlow)
            TheExec.flow.TestLimit resultVal:=VDAC_Value3, PinName:="", Tname:=TestNameInput, scaletype:=scaleNone, ForceResults:=tlForceFlow
        Else
            VDAC_Value = GetStoreDataAllType(argv(i) & "_para")
            VDAC_Value = VDAC_Value.divide(Resolution).Multiply(ratio)
            TestNameInput = Report_TName_From_Instance(CalcC, "", , , , , , , tlForceFlow)
            TheExec.flow.TestLimit resultVal:=VDAC_Value, PinName:="", Tname:=TestNameInput, scaletype:=scaleNone, ForceResults:=tlForceFlow
        End If
        
        
        'TestNameInput = Report_TName_From_Instance(CalcC, "", argv(i), , , , , , tlForceFlow)
        
        

    Next i
        
    Exit Function

errHandler:
    TheExec.Datalog.WriteComment "error in Calc_VDAC"
    If AbortTest Then Exit Function Else Resume Next
End Function

Public Function Calc_PU_PD_Resistance_By_MeasV(argc As Integer, argv() As String) As Long
    'From T-Col

    On Error GoTo errHandler
    
    Dim TestNameInput As String
    Dim i As Long
    Dim PLD_Data As New PinListData
    Dim PLD_Res As New PinListData
    Dim VDD_Voltage As New SiteDouble
    Dim VDD_Pin As String
    Dim Force As Double
    Dim Direction As String
    Dim ary() As String
    Dim pinType As String
    For i = 0 To argc - 1
        ary = Split(argv(i), "+")
        PLD_Data = GetStoreDataAllType(ary(0))
        Direction = ary(1)
        Force = CDbl(ary(2))
        
        If Direction = "PD" Then
            PLD_Res = PLD_Data.Math.divide(Force)
        ElseIf Direction = "PU" Then
            VDD_Pin = ary(3)
            pinType = SortPinInstrument(ary(3))
            If pinType = "DC-07" Or pinType = "DC-30" Or pinType = "DC-75" Then     'DCVIThen
               For Each site In TheExec.sites.Active
                    VDD_Voltage = TheHdw.DCVI.pins(VDD_Pin).Voltage.value
                Next site
            Else
                VDD_Voltage = TheHdw.DCVS.pins(VDD_Pin).Voltage.Main.ValuePerSite
                
            End If
            PLD_Res = PLD_Data.Math.Negate.Add(VDD_Voltage).divide(Force)
        End If
        TestNameInput = Report_TName_From_Instance(CalcC, "", , , , , , , tlForceFlow)
        TheExec.flow.TestLimit resultVal:=PLD_Res, Tname:=TestNameInput, scaletype:=scaleNone, ForceResults:=tlForceFlow
    Next i
      
    Exit Function

errHandler:
    TheExec.Datalog.WriteComment "Calc_PU_PD_Resistance_By_MeasV"
    If AbortTest Then Exit Function Else Resume Next
End Function
Public Function Calc_MeasR_AVG(argc As Integer, argv() As String) As Long
    On Error GoTo errHandler
    
    Dim R_Value As Double
    Dim i As Integer
    Dim tmp As New PinListData
    
    tmp = GetStoreDataAllType(argv(0))
    
    For Each site In TheExec.sites.Active
        For i = 0 To tmp.pins.Count - 1
            R_Value = R_Value + (tmp.pins(i).value)
        Next i
        R_Value = R_Value / (tmp.pins.Count)
        TheExec.flow.TestLimit resultVal:=R_Value, Tname:="Calc_MeasR_AVG" & argv(0) & "_Site" & site, ForceResults:=tlForceNone, unit:=unitOhm
        R_Value = 0
    Next site
    R_Value = 0
   
Exit Function
errHandler:
    TheExec.Datalog.WriteComment "error in Calc_MeasR_AVG"
    If AbortTest Then Exit Function Else Resume Next
End Function

Public Function Calc_DigCap_Frequency_Transfer_DSG(argc As Integer, argv() As String) As Long

    Dim i As Long
    Dim l_Num_DictName As Long
    Dim l_Ratio As Double
    'Dim l_multiplex As Double
    'Dim sd_Input() As New SiteDouble
    Dim sd_Input() As New SiteDouble
    Dim s_DictName As String
    Dim s_TestNameInput As String
    Dim s_DictName_Ary() As String
    Dim dsp_Output As New DSPWave
    Dim dsp_ConcatenateDSP As New DSPWave
    Dim TempResult As New SiteDouble
    Dim site As Variant
    
    On Error GoTo errHandler
    
    l_Ratio = argv(0)
    'l_multiplex = argv(1)
    s_DictName = argv(1)
    If InStr(s_DictName, "&") <> 0 Then
        s_DictName_Ary = Split(s_DictName, "&")
    Else
        ReDim s_DictName_Ary(0)
        s_DictName_Ary(0) = s_DictName
    End If
    
    l_Num_DictName = UBound(s_DictName_Ary)
    ReDim sd_Input(l_Num_DictName)
    'sd_Input = UBound(s_DictName_Ary)
    dsp_ConcatenateDSP.CreateConstant 0, l_Num_DictName + 1, DspDouble
    dsp_Output.CreateConstant 0, l_Num_DictName + 1, DspDouble
    
    For i = 0 To l_Num_DictName
        sd_Input(i) = GetStoreDataAllType(s_DictName_Ary(i) & "_para")
        
        
        For Each site In TheExec.sites
            dsp_ConcatenateDSP.ElementLite(i) = sd_Input(i)
        Next site
    Next i
    
    TheHdw.dsp.ExecutionMode = tlDSPModeForceAutomatic
    rundsp.DSP_DigCap_Frequency_Transfer dsp_ConcatenateDSP, l_Ratio, dsp_Output
       
    For i = 0 To l_Num_DictName
        TempResult = dsp_Output.Element(i)
        s_TestNameInput = Report_TName_From_Instance(CalcC, "X", , 0, 0, ForceResult:=tlForceFlow)
        TheExec.flow.TestLimit resultVal:=TempResult, Tname:=s_TestNameInput, ForceResults:=tlForceFlow, PinName:="X"
    Next i
    
    Exit Function
    
errHandler:
    TheExec.Datalog.WriteComment "error in Calc_DSG_CaptureC_Frequency_Transfer"
    If AbortTest Then Exit Function Else Resume Next

End Function
