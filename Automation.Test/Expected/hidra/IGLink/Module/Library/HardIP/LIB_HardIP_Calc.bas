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
Dim site As Variant 'Carter, 20240304

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
            Temp_target(i) = Target_Step_Binarry.Select(i * Fuse_Digsrc_size, 1, Fuse_Digsrc_size).Copy
            
            Call StoreDataAllType(ADCLDO_Fusearray(i), Temp_target(i))
        Next i

    Next site
    
'    For Each site In TheExec.sites.Active
        For z = 0 To UBound(InWf_Split)
           TestNameInput = Report_TName_From_Instance(CalcC, InWf_Split(z), , , z)
           TheExec.Flow.TestLimit resultVal:=Trim_code(z), Tname:=TestNameInput, ForceResults:=tlForceFlow
        Next z
'     Next site
 
 
Exit Function 'Add ErrHandler 2023/05/29
errHandler: 'Add ErrHandler 2023/05/29
    Call Print_Error_Message(Error_Info, "LIB_HardIP_Calc", "Calc_ADC_LDO_TRIM_TTR") 'Add ErrHandler 2023/05/29
    If AbortTest Then Exit Function Else Resume Next 'Add ErrHandler 2023/05/29
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
            For i = 0 To meas_val_now.Pins.Count - 1
                If meas_val_now.Pins(i).value = 0 Then
                    meas_val_now.Pins(i).value = 0.0000000001
                End If
            Next i
            For i = 0 To meas_val_before.Pins.Count - 1
                If meas_val_before.Pins(i).value = 0 Then
                    meas_val_before.Pins(i).value = 0.0000000001
                End If
            Next i
        Next site
        '=================prevent divide 0==============
        meas_val_now = meas_val_now.Math.Invert.Subtract(meas_val_before.Math.Invert).Multiply(0.5)
        Dim PLD_For_TestLimit As New PinListData
        For i = 0 To meas_val_now.Pins.Count - 1
            If UCase(meas_val_now.Pins(i)) Like "*DQS_P*" Then
                PLD_For_TestLimit.AddPin (meas_val_now.Pins(i))
                For Each site In TheExec.sites
                    PLD_For_TestLimit.Pins(meas_val_now.Pins(i)).value = meas_val_now.Pins(i).value
                Next site
            End If
        Next i
        TestNameInput = Report_TName_From_Instance(CalcC, vbNullString)
        TheExec.Flow.TestLimit resultVal:=PLD_For_TestLimit, Tname:=TestNameInput, ForceResults:=tlForceFlow, scaletype:=scalePico
        meas_val_before = meas_val
    End If

    meas_val_delay_instance_name = TheExec.DataManager.instancename
    
Exit Function 'Add ErrHandler 2023/05/29
errHandler: 'Add ErrHandler 2023/05/29
    Call Print_Error_Message(Error_Info, "LIB_HardIP_Calc", "Calc_delay_Sicily") 'Add ErrHandler 2023/05/29
    If AbortTest Then Exit Function Else Resume Next 'Add ErrHandler 2023/05/29
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
    
    For Each PinName In Meas_I1_PLD.Pins
        
        'If PinName is not exist then add new one to Global PinListData
        For Each PinName_Glb_PLD In R_Path_PLD.Pins
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
            
            Dim i1 As Double: i1 = Meas_I1_PLD.Pins(PinName).value(site)
            Dim i2 As Double: i2 = Meas_I2_PLD.Pins(PinName).value(site)
            Dim i3 As Double: i3 = Meas_I3_PLD.Pins(PinName).value(site)
            Dim I3_I1 As Double: I3_I1 = Meas_I3_PLD.Pins(PinName).value(site) - Meas_I1_PLD.Pins(PinName).value(site)
            Dim I3_I2 As Double: I3_I2 = Meas_I3_PLD.Pins(PinName).value(site) - Meas_I2_PLD.Pins(PinName).value(site)
            
            'Initialize the Value on PinListData to prevent any cross usage between samples
            R_Path_PLD.Pins(PinName).value(site) = 0
            R_Contact_PLD.Pins(PinName).value(site) = 0
            
            'RAK_Val = TheHdw.PPMU.ReadRakValuesByPinnames(PinName, site)
       
''            If InStr(UCase(TheExec.CurrentChanMap), "FT") <> 0 Then
''                Total_RAK_Val = RAK_Val(0) + FT_Card_RAK.Pins(PinName).Value(Site)
''            Else
''                Total_RAK_Val = RAK_Val(0) + CP_Card_RAK.Pins(PinName).Value(Site)
''            End If
            
            
            Total_RAK_Val = CurrentJob_Card_RAK.Pins(PinName).value(site)     ' Edited by Dylan 2019/12/04 (For debug running)
           
            If i1 <> 0 And i2 <> 0 And i3 <> 0 Then
                If I3_I1 > 0 And I3_I2 > 0 And (i1 * i2) > 0 Then
                    R_Path_PLD.Pins(PinName).value(site) = (Force_Cond / i3) * (1 - ((I3_I1 * I3_I2) / (i1 * i2)) ^ 0.5)
                Else
                    R_Path_PLD.Pins(PinName).value(site) = 999 ' report R= 999 when divide by 0
                    TheExec.Datalog.WriteComment (" Error : PinName " & CStr(PinName) & " , Site" & CStr(site) & " I3 should greater than I1,I2 And (I1*I2) should greater than 0!  ")
                End If
            Else
                R_Path_PLD.Pins(PinName).value(site) = 999
                TheExec.Datalog.WriteComment (" Error : PinName " & CStr(PinName) & " , Site" & CStr(site) & " Division by Zero !   ")
            End If
            
            R_Contact_PLD.Pins(PinName).value(site) = R_Path_PLD.Pins(PinName).value(site) - Total_RAK_Val
                        
            'Customize String for DDR Test only
            If Cust_Str = UCase("DDR_TEST") Then
                If i1 <> 0 And i2 <> 0 Then
                    DDR_R1.Pins(PinName).value(site) = (1 * Force_Cond / i1) - R_Path_PLD.Pins(PinName).value(site)
                    DDR_R2.Pins(PinName).value(site) = (1 * Force_Cond / i2) - R_Path_PLD.Pins(PinName).value(site)
                Else
                    DDR_R1.Pins(PinName).value(site) = 999
                    DDR_R2.Pins(PinName).value(site) = 999
                End If
            End If
                        
        Next site
        
    Next PinName
    
    Dim temp
    
    temp = TheExec.Flow.TestLimitIndex
    If EnableDigitalTestLimitTTR = True Then
        'TTR,20200423, Oscar
        TestNameInput = Report_TName_From_Instance(CalcR, "X", , 0)
        TheExec.Flow.TestLimit resultVal:=R_Contact_PLD, unit:=unitCustom, customUnit:="ohm", Tname:=TestNameInput, ForceResults:=tlForceFlow
    Else
        For Each PinName In R_Contact_PLD.Pins
                TheExec.Flow.TestLimitIndex = temp
                TestNameInput = Report_TName_From_Instance(CalcR, CStr(PinName), , 0)
                TheExec.Flow.TestLimit resultVal:=R_Contact_PLD.Pins(PinName), unit:=unitCustom, customUnit:="ohm", Tname:=TestNameInput, ForceResults:=tlForceFlow
                'TheExec.Flow.TestLimit resultVal:=R_Contact_PLD.Pins(PinName), lowval:=0, hival:=5, Unit:=unitCustom, customUnit:="ohm", TName:=TestNameInput, ForceResults:=tlForceFlow
        Next PinName
    End If
    If Cust_Str = UCase("DDR_TEST") Then
    
        If EnableDigitalTestLimitTTR = True Then
                        'TTR,20200423, Oscar
            TestNameInput = Report_TName_From_Instance(CalcR, "X", , 0)
            TheExec.Flow.TestLimit resultVal:=DDR_R1.Pins(PinName), unit:=unitCustom, customUnit:="ohm", Tname:=TestNameInput, ForceResults:=tlForceFlow
            TheExec.Flow.TestLimit resultVal:=DDR_R2.Pins(PinName), unit:=unitCustom, customUnit:="ohm", Tname:=TestNameInput, ForceResults:=tlForceFlow
        Else
            temp = TheExec.Flow.TestLimitIndex
            For Each PinName In DDR_R1.Pins
                TheExec.Flow.TestLimitIndex = temp
                TestNameInput = Report_TName_From_Instance(CalcR, CStr(PinName))
                TheExec.Flow.TestLimit resultVal:=DDR_R1.Pins(PinName), unit:=unitCustom, customUnit:="ohm", Tname:=TestNameInput, ForceResults:=tlForceFlow
            Next PinName
            
            temp = TheExec.Flow.TestLimitIndex
            For Each PinName In DDR_R2.Pins
                TheExec.Flow.TestLimitIndex = temp
                TestNameInput = Report_TName_From_Instance(CalcR, CStr(PinName))
                TheExec.Flow.TestLimit resultVal:=DDR_R2.Pins(PinName), unit:=unitCustom, customUnit:="ohm", Tname:=TestNameInput, ForceResults:=tlForceFlow
            Next PinName
        End If
    End If
   
    Exit Function

Exit Function 'Add ErrHandler 2023/05/29
errHandler: 'Add ErrHandler 2023/05/29
    Call Print_Error_Message(Error_Info, "LIB_HardIP_Calc", "Calc_R_Path_Cal") 'Add ErrHandler 2023/05/29
    If AbortTest Then Exit Function Else Resume Next 'Add ErrHandler 2023/05/29
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
            DSPWave_Combine_BIN(site) = DSPWave_Combine(i)(site).ConvertDataTypeTo(DspLong).Copy
            DSPWave_Combine_Dec(site) = DSPWave_Combine_BIN(site).ConvertStreamTo(tldspParallel, DSPWave_Combine_BIN(site).SampleSize, 0, Bit0IsMsb)
            ''===================== BinToDec (End) =====================
        Next site
        TestNameInput = Report_TName_From_Instance(CalcC, "X", "ConcatenateDSP", 0)
        
        TheExec.Flow.TestLimit resultVal:=DSPWave_Combine_Dec.Element(0), PinName:="X", Tname:=TestNameInput, ForceResults:=tlForceFlow
    Next i
    
Exit Function 'Add ErrHandler 2023/05/29
errHandler: 'Add ErrHandler 2023/05/29
    Call Print_Error_Message(Error_Info, "LIB_HardIP_Calc", "Calc_ConcatenateDSP") 'Add ErrHandler 2023/05/29
    If AbortTest Then Exit Function Else Resume Next 'Add ErrHandler 2023/05/29
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
                
        TheExec.Flow.TestLimit resultVal:=DSP_ProcessOutput_DEC.Element(0), Tname:=TestNameInput, ForceResults:=tlForceFlow
        
    Next i

Exit Function 'Add ErrHandler 2023/05/29
errHandler: 'Add ErrHandler 2023/05/29
    Call Print_Error_Message(Error_Info, "LIB_HardIP_Calc", "Calc_BitwiseDSP") 'Add ErrHandler 2023/05/29
    If AbortTest Then Exit Function Else Resume Next 'Add ErrHandler 2023/05/29
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
            
            TheExec.Flow.TestLimit resultVal:=DSP_DictKey_DEC.Element(0), Tname:=TestNameInput, ForceResults:=tlForceFlow
            
            Call Update_BC_PassFail_Flag
            
            TestNameInput = Report_TName_From_Instance(CalcC, "X", , CInt(i))
            
            TheExec.Flow.TestLimit resultVal:=DSPWave_2S_Complement(i).Element(0), Tname:=TestNameInput, ForceResults:=tlForceFlow
            
            Call Update_BC_PassFail_Flag
        Else
            
            TestNameInput = Report_TName_From_Instance(CalcC, "X", , CInt(i))
            
            TheExec.Flow.TestLimit resultVal:=DSPWave_2S_Complement(i).Element(0), Tname:=TestNameInput, ForceResults:=tlForceFlow
            
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
        TheExec.Flow.TestLimit resultVal:=Calc_DSP_DEC(i).Element(0), Tname:=TestNameInput, unit:=unitHz, ForceResults:=tlForceFlow
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

      
      DSP_Arry_Bin(4) = DSP_Captured(0).Select(0, 1, 9).Copy
      DSP_Arry_Bin(0) = DSP_Captured(0).Select(10, 1, 9).Copy
      DSP_Arry_Bin(6) = DSP_Captured(0).Select(20, 1, 9).Copy
      
      DSP_Arry_Bin(1) = DSP_Captured(1).Select(0, 1, 9).Copy
      DSP_Arry_Bin(3) = DSP_Captured(1).Select(10, 1, 9).Copy
      DSP_Arry_Bin(7) = DSP_Captured(1).Select(20, 1, 9).Copy
      
      DSP_Arry_Bin(2) = DSP_Captured(2).Select(0, 1, 9).Copy
      DSP_Arry_Bin(5) = DSP_Captured(2).Select(10, 1, 9).Copy
      
      
   
      For i = 0 To UBound(DSP_Arry_Bin)
         DSP_Arry_Bin(i) = DSP_Arry_Bin(i).ConvertDataTypeTo(DspLong)
         DSP_Arry_Dec(i) = DSP_Arry_Bin(i).ConvertStreamTo(tldspParallel, DSP_Arry_Bin(i).SampleSize, 0, Bit0IsMsb)
         'nope, only for debugging purpose
         'Report_TestLimit_by_CZ_Format resultVal:=DSP_Arry_Dec(i).Element(0), ForceResults:=tlForceFlow, UserVar6:="DSP_Arry_Dec" & i, UserVar5:=argv(0), MeasType:="C"
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
         'Do nothing
         End If
      Next i
      Max_Dec_Val(site) = DSP_Arry_Dec(0).Element(0)
   Next site
    
   Call GetFlowTName
    
   If gl_UseStandardTestName_Flag = True Then                     'Roger add
      Call Report_ALG_TName_From_Instance(OutputTname_format, "C", CStr(argv(0)) & "Max_Diff", gl_Tname_Meas_FromFlow(TheExec.Flow.TestLimitIndex))
      TestNameInput = Merge_TName(OutputTname_format)
            
   Else
      TestNameInput = TestName & "Max_Diff"
   End If
    
   TheExec.Flow.TestLimit resultVal:=Max_Dec_Val, Tname:=TestNameInput, ForceResults:=tlForceFlow
    
   If gl_UseStandardTestName_Flag = True Then                     'Roger add
      Call Report_ALG_TName_From_Instance(OutputTname_format, "C", CStr(argv(0)) & "Decrease", gl_Tname_Meas_FromFlow(TheExec.Flow.TestLimitIndex))
      TestNameInput = Merge_TName(OutputTname_format)
            
   Else
      TestNameInput = TestName & "Decrease"
   End If
    
    
   TheExec.Flow.TestLimit resultVal:=Uni_DLL_Indicator, Tname:=TestNameInput, ForceResults:=tlForceFlow
    
    
    
    
   '''    Report_TestLimit_by_CZ_Format resultVal:=Max_Dec_Val, ForceResults:=tlForceFlow, MeasType:="C"
   '''
   '''    Report_TestLimit_by_CZ_Format resultVal:=Uni_DLL_Indicator, lowVal:=1, hiVal:=2, ForceResults:=tlForceFlow, MeasType:="C"
    
Exit Function 'Add ErrHandler 2023/05/29
errHandler: 'Add ErrHandler 2023/05/29
    Call Print_Error_Message(Error_Info, "LIB_HardIP_Calc", "Calc_MDLL_Monotonicity_DevideBlock_SEG") 'Add ErrHandler 2023/05/29
    If AbortTest Then Exit Function Else Resume Next 'Add ErrHandler 2023/05/29
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
    For j = 0 To DDR0_MC_DQS_DIFFx_F(i).Pins.Count - 1
        If InStr(UCase(DDR0_MC_DQS_DIFFx_F(i).Pins(j)), "DQS_P") <> 0 Then
            For Each site In TheExec.sites
                If DDR0_MC_DQS_DIFFx_F(i).Pins(j).value = 0 Then
                    SiteDouble_Frequency(i) = 0.000000001
                    TheExec.Datalog.WriteComment "Site" & site & " : DDR F" & i & " frequency is 0"
                Else
                    SiteDouble_Frequency(i) = DDR0_MC_DQS_DIFFx_F(i).Pins(j).value
                End If
            Next site
        End If
    Next j
Next i


For i = 0 To NumberOfFreq - 1
    SiteDouble_Delay(i) = SiteDouble_Frequency(i).Multiply(2).Invert
    TestNameInput = Report_TName_From_Instance(CalcF, vbNullString, , 0)
    TheExec.Flow.TestLimit resultVal:=SiteDouble_Delay(i), Tname:=TestNameInput, ForceResults:=tlForceFlow, scaletype:=scalePico
Next i

For i = 0 To NumberOfFreq - 2
    SiteDouble_Delta(i) = SiteDouble_Delay(i + 1).Subtract(SiteDouble_Delay(i))
    TestNameInput = Report_TName_From_Instance(CalcF, vbNullString, , 0)
    TheExec.Flow.TestLimit resultVal:=SiteDouble_Delta(i), Tname:=TestNameInput, ForceResults:=tlForceFlow, scaletype:=scalePico
Next i


Exit Function 'Add ErrHandler 2023/05/29
errHandler: 'Add ErrHandler 2023/05/29
    Call Print_Error_Message(Error_Info, "LIB_HardIP_Calc", "CalcDelayDelta_Sicily") 'Add ErrHandler 2023/05/29
    If AbortTest Then Exit Function Else Resume Next 'Add ErrHandler 2023/05/29
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
            DictKey_Diff_Calc = argv(7) 'For Sicily
            'DictKey_Diff_Calc = argv(6) ' For Turks
        End If
    End If
    
    pld_V1 = GetStoreDataAllType(DictKey_V1)
    pld_V2 = GetStoreDataAllType(DictKey_V2)
    
    Dim R_Path As New SiteDouble
    'Dim R_Channel_RAK() As Double
    For Each site In TheExec.sites
        'R_Channel_RAK = TheHdw.PPMU.ReadRakValuesByPinnames(PinName, site)
        R_Path(site) = CurrentJob_Card_RAK.Pins(PinName).value(site)
    Next site
    
    Dim V_DP_K As New SiteDouble
    
    For Each site In TheExec.sites
        V_DP_K(site) = ((pld_V1.Pins(PinName).value(site) * i2 - pld_V2.Pins(PinName).value(site) * i1) * R_Term) / (pld_V1.Pins(PinName).value(site) - pld_V2.Pins(PinName).value(site) + (R_Term - R_Path(site)) * (i2 - i1))
    Next site
    Dim OutputTname_format() As String
    Dim TestNameInput As String
    
''    TheExec.Flow.TestLimit resultVal:=V_DP_K, PinName:=PinName, Tname:="Volt_meas_TX_Level", ForceResults:=tlForceFlow
    TestNameInput = Report_TName_From_Instance(CalcV, PinName, vbNullString, CInt(j))
    TheExec.Flow.TestLimit resultVal:=V_DP_K, PinName:=PinName, ForceResults:=tlForceFlow, Tname:=TestNameInput
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
            TheExec.Flow.TestLimit resultVal:=V_Diff, PinName:="P-N", ForceResults:=tlForceFlow, Tname:=TestNameInput
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
        
    
Exit Function 'Add ErrHandler 2023/05/29
errHandler: 'Add ErrHandler 2023/05/29
    Call Print_Error_Message(Error_Info, "LIB_HardIP_Calc", "TX_Level") 'Add ErrHandler 2023/05/29
    If AbortTest Then Exit Function Else Resume Next 'Add ErrHandler 2023/05/29
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
            Max_Temp.AddPin (Dict_Freq_Value(i).Pins(0).name)
            Min_Temp.AddPin (Dict_Freq_Value(i).Pins(0).name)
            Divide_Temp.AddPin (Dict_Freq_Value(i).Pins(0).name)
        End If
        For Each site In TheExec.sites
            increase_flag(site) = True
            If i = 0 Then
                Max_Temp.Pins(0).value(site) = Dict_Freq_Value(i).Pins(0).value(site)
                Min_Temp.Pins(0).value(site) = Dict_Freq_Value(i).Pins(0).value(site)
            Else
                '''''''''''''''''''''print datalog'''''''''''''''''''''''''''''''''''''
                'TheExec.Datalog.WriteComment "Site " & site & ":" & argv(i) & "-" & argv(i - 1) & "=" & Dict_Freq_Value(i).Pins(0).Value(site) - Dict_Freq_Value(i - 1).Pins(0).Value(site)
                DeltaF = Dict_Freq_Value(i).Pins(0).value(site) - Dict_Freq_Value(i - 1).Pins(0).value(site)
                '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                If Dict_Freq_Value(i).Pins(0).value(site) > Max_Temp.Pins(0).value(site) Then
                    Max_Temp.Pins(0).value(site) = Dict_Freq_Value(i).Pins(0).value(site)
                Else
                    increase_flag = False
                End If
                If Dict_Freq_Value(i).Pins(0).value(site) < Min_Temp.Pins(0).value(site) Then
                    Min_Temp.Pins(0).value(site) = Dict_Freq_Value(i).Pins(0).value(site)
                End If
            End If
        Next site
        If i <> 0 Then
            If EnableDigitalTestLimitTTR Then
                TestNameInput = Report_TName_From_Instance("Calc", "X", "Delta-" & i, CInt(i), , , , , tlForceNone)
            Else
                TestNameInput = Report_TName_From_Instance("Calc", Dict_Freq_Value(i).Pins(0).name, "Delta-" & i, CInt(i), , , , , tlForceNone)
            End If
            TheExec.Flow.TestLimit resultVal:=DeltaF, Tname:=TestNameInput, ForceResults:=tlForceFlow, PinName:=Dict_Freq_Value(i).Pins(0).name
        End If
    Next i
    
    For Each site In TheExec.sites
        If Min_Temp.Pins(0).value(site) = 0 Or increase_flag(site) = False Then
        
            Divide_Temp.Pins(0).value(site) = 999
            If Min_Temp.Pins(0).value(site) = 0 Then
                TheExec.Datalog.WriteComment ("Error! Site " & site & " Min Freq Meas(Denominator)=0 Hz ")
            Else
                TheExec.Datalog.WriteComment ("Error! Site " & site & " Not FRO0<FRO1<FRO2....<FRO23 ")
            End If
            
        Else
'            Divide_Temp.Pins(0).Value(site) = Max_Temp.Pins(0).Value(site) / Min_Temp.Pins(0).Value(site)
            Divide_Temp.Pins(0).value(site) = Dict_Freq_Value(UBound(Dict_Freq_Value)).Pins(0).value(site) / Dict_Freq_Value(0).Pins(0).value(site)
        End If
    Next site
    
    For i = 0 To Max_Temp.Pins.Count - 1

'        TestNameInput = Report_TName_From_Instance(CalcF, Max_Temp.Pins(i), "Fmax", CInt(i))
'
'        TheExec.Flow.TestLimit resultVal:=Max_Temp, Tname:=TestNameInput, ForceResults:=tlForceFlow
'
'        TestNameInput = Report_TName_From_Instance(CalcF, Max_Temp.Pins(i), "Fmin", CInt(i))
'
'        TheExec.Flow.TestLimit resultVal:=Min_Temp, Tname:=TestNameInput, ForceResults:=tlForceFlow
        
        'TestNameInput = Report_TName_From_Instance("Calc", Max_Temp.Pins(i), "FmaxDivideFmin", CInt(i))
        If EnableDigitalTestLimitTTR Then
            TestNameInput = Report_TName_From_Instance("Calc", "X", "FmaxDivideFmin", CInt(i))
        Else
            TestNameInput = Report_TName_From_Instance("Calc", Max_Temp.Pins(i), "FmaxDivideFmin", CInt(i))     '20200702 update by CT
        End If
        TheExec.Flow.TestLimit resultVal:=Divide_Temp, Tname:=TestNameInput, ForceResults:=tlForceFlow
        
    Next i
Exit Function 'Add ErrHandler 2023/05/29
errHandler: 'Add ErrHandler 2023/05/29
    Call Print_Error_Message(Error_Info, "LIB_HardIP_Calc", "Calc_Fmax_Divide_Fmin") 'Add ErrHandler 2023/05/29
    If AbortTest Then Exit Function Else Resume Next 'Add ErrHandler 2023/05/29
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
    VDD18_MIPI_value = TheHdw.DCVS.Pins("VDD12_MIPI").Voltage.Main.value
    
    If VDD18_MIPI_value = 0 Then
            VDD18_MIPI_value = 999
            TheExec.Datalog.WriteComment ("Error! Apply VDD18_MIPI=0 V  ")
    End If
    
    For Each site In TheExec.sites
        VCMTX(site).Element(0) = DSPWave_Combine_Dec(site).Element(0) / 1024 * VDD18_MIPI_value
    Next site
'    Call rundsp.DSPWaveDecToBinary(DSPWave_Combine_Dec, 10, DSPWave_Combine_verify)
    
    TestNameInput = Report_TName_From_Instance(CalcV, "X", , 0)

    TheExec.Flow.TestLimit resultVal:=VCMTX.Element(0), Tname:=TestNameInput, ForceResults:=tlForceFlow
Exit Function 'Add ErrHandler 2023/05/29
errHandler: 'Add ErrHandler 2023/05/29
    Call Print_Error_Message(Error_Info, "LIB_HardIP_Calc", "Calc_MIPI_VCMTX") 'Add ErrHandler 2023/05/29
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
    ForceVoltVal = TheHdw.DCVS.Pins(PowerPinName).Voltage.value
    
    For Each pin In StoredCurrent.Pins
        For Each site In TheExec.sites
            If StoredCurrent.Pins(pin).value(site) = 0 Then
                StoredCurrent.Pins(pin).value(site) = 0.000000000001
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
        TheExec.Flow.TestLimit CalR, , , , , , unitCustom, , TestNameInput, , , , , " ohm", , ForceResults:=tlForceFlow
    ElseIf argc <= 2 Then
        Dim Temp_index As Long
        Temp_index = TheExec.Flow.TestLimitIndex
        For i = 0 To CalR.Pins.Count - 1
            TheExec.Flow.TestLimitIndex = Temp_index
            TestNameInput = Report_TName_From_Instance(CalcR, CalR.Pins(i), , CInt(i))
'            If i = 0 Then
                TheExec.Flow.TestLimit CalR.Pins(i), , , , , , unitCustom, , TestNameInput, , , , , " ohm", , ForceResults:=tlForceFlow
'            Else
'                TheExec.Flow.TestLimit CalR.Pins(i), GetLowLimitFromFlow, GetHiLimitFromFlow, , , , unitCustom, , TestNameInput, , , , , " ohm", , ForceResults:=tlForceFlow
'            End If
        Next i
    ElseIf argv(2) = "TTR" Then
            TheExec.Flow.TestLimitIndex = TheExec.Flow.TestLimitIndex + 2   'use second test limit spec for R,first test limit for current(don't need)

        Lowlimitval_temp = GetLowLimitFromFlow
        Hilimitval_temp = GetHiLimitFromFlow
         If TheExec.enableWord("HIP_TTR_FailResultOnly") = True Then
            For Each site In TheExec.sites.Active
                For p = 0 To CalR.Pins.Count - 1
                If CalR.Pins(p).value > Hilimitval_temp Or CalR.Pins(p).value < Lowlimitval_temp Then
                    TestNameInput = Report_TName_From_Instance(CalcR, CalR.Pins(p), , CInt(p))
                    
                    'TheExec.Flow.TestLimit StoredCurrent.Pins(p), , , , , , unitAmp, , , ForceResults:=tlForceFlow
                    If gl_Disable_HIP_debug_log = False Then TheExec.Datalog.WriteComment CalR.Pins(p) & " = " & CalR.Pins(p).value(site)
                    TheExec.Flow.TestLimit CalR.Pins(p), Lowlimitval_temp, Hilimitval_temp, , , , unitCustom, , TestNameInput, , , , , " ohm", , ForceResults:=tlForceFlow

                    
                End If
                Next p
            Next site
        Else
            For p = 0 To CalR.Pins.Count - 1

                    If gl_Disable_HIP_debug_log = False Then TheExec.Datalog.WriteComment CalR.Pins(p) & " = " & CalR.Pins(p).value(site)
                    TestNameInput = Report_TName_From_Instance(CalcR, CalR.Pins(p), , CInt(p))
                    TheExec.Flow.TestLimit CalR.Pins(p), Lowlimitval_temp, Hilimitval_temp, , , , unitCustom, , TestNameInput, , , , , " ohm", , ForceResults:=tlForceFlow
            Next p
        End If
    Else
    'Do nothing
    End If
    
Exit Function 'Add ErrHandler 2023/05/29
errHandler: 'Add ErrHandler 2023/05/29
    Call Print_Error_Message(Error_Info, "LIB_HardIP_Calc", "Calc_CalR_FVMI") 'Add ErrHandler 2023/05/29
    If AbortTest Then Exit Function Else Resume Next 'Add ErrHandler 2023/05/29
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
       
        Dim tempResult As New SiteLong
        For i = 0 To 7
            tempResult = ConcatenateDSP_AfterSort.Element(i + 8 * j)
            TestNameInput = Report_TName_From_Instance(CalcC, "X", , CInt(j), ForceResult:=tlForceFlow)
            TheExec.Flow.TestLimit resultVal:=tempResult, Tname:=TestNameInput, ForceResults:=tlForceFlow
            TheExec.Flow.TestLimitIndex = TheExec.Flow.TestLimitIndex - 1
        Next i
        TheExec.Flow.TestLimitIndex = TheExec.Flow.TestLimitIndex + 1
        
        tempResult = sl_MDLL_DecreaseDirection.Element(j)
        TestNameInput = Report_TName_From_Instance(CalcC, "X", , CInt(j), ForceResult:=tlForceFlow)
        TheExec.Flow.TestLimit resultVal:=tempResult, lowVal:=1, hiVal:=1, Tname:=TestNameInput, ForceResults:=tlForceFlow
        
        
        tempResult = sl_Num_DiffVal.Element(j)
        TestNameInput = Report_TName_From_Instance(CalcC, "X", , CInt(j), ForceResult:=tlForceFlow)
        TheExec.Flow.TestLimit resultVal:=tempResult, lowVal:=1, hiVal:=2, Tname:=TestNameInput, ForceResults:=tlForceFlow
    
        tempResult = sl_Diff_MaxMin.Element(j)
        TestNameInput = Report_TName_From_Instance(CalcC, "X", , CInt(j), ForceResult:=tlForceFlow)
        TheExec.Flow.TestLimit resultVal:=tempResult, lowVal:=0, hiVal:=1, Tname:=TestNameInput, ForceResults:=tlForceFlow
    Next j

Exit Function 'Add ErrHandler 2023/05/29
errHandler: 'Add ErrHandler 2023/05/29
    Call Print_Error_Message(Error_Info, "LIB_HardIP_Calc", "Calc_MDLL_Monotonicity_DevideBlock_TTR") 'Add ErrHandler 2023/05/29
    If AbortTest Then Exit Function Else Resume Next 'Add ErrHandler 2023/05/29
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
            gl_Current_Instance_Tname_subblock = Application.Worksheets(TheExec.Flow.Raw.SheetInRun).range("AM" & CStr(TheExec.Flow.Raw.GetCurrentLineNumber + 5)).value
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
                TNameSeg(9) = TheExec.Flow.var(gl_Sweep_Name).value
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
            'Do nothing
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
        TheExec.Flow.TestLimit resultVal:=temp_RefferanceCode_DEC, ForceResults:=tlForceFlow, Tname:=TestNameInput
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
            TheExec.Flow.TestLimit resultVal:=ADC_code_average_DEC, ForceResults:=tlForceFlow, Tname:=TestNameInput
        Else
            If EnableDigitalTestLimitTTR = True Then
                TestNameInput = Report_TName_From_Instance(CalcC, "X", "_ADC")
            Else
                TestNameInput = Report_TName_From_Instance(CalcC, Tname_String, "_ADC")
            End If
            TheExec.Flow.TestLimit resultVal:=ADC_code_average_DEC, ForceResults:=tlForceFlow, Tname:=TestNameInput
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
ReDim Meas_pins(inVoh.Pins.Count - 1)
    
' 0.x get pins
For i = 0 To inVoh.Pins.Count - 1
    Meas_pins(i) = inVoh.Pins.item(i).name
    ' 0.1 add pins
    MeasureValue.AddPin (Meas_pins(i))
    Compensate_V.AddPin (Meas_pins(i))
    ' 0.2 disconect digital pins first
    TheHdw.Digital.Pins(Meas_pins(i)).Disconnect
Next i

' 1.x Force V
For Each pin In Meas_pins
        ' 1.0 Get the measured voltage firstly
        Set voh_temp = inVoh.Pins(pin)
        ' 1.1 Check if measured voltages are out of PPMU spec, if so, use 0v as default setting
        For Each site In TheExec.sites
             If voh_temp < -1 Or voh_temp > 6 Then
                    TheExec.Datalog.WriteComment "the force value " & voh_temp & "is out of PPMU range -1V ~ 6V, bypass force PPMU and set measurement result to 9999"
                    voh_temp(site) = 0
            End If
        Next site
        ' 1.2 Setup PPMU then measure current
         With TheHdw.PPMU.Pins(pin)
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
        MeasureValue.Pins(pin) = TheHdw.PPMU.Pins(pin).Read(tlPPMUReadMeasurements, pc_Def_PPMU_ReadPoint)
        ' 1.3 Set abnormal voltage to 9999
        For Each site In TheExec.sites
            If inVoh.Pins(pin).value(site) < -1 Or inVoh.Pins(pin).value(site) > 6 Then MeasureValue.Pins(pin).value(site) = 9999
        Next site
Next pin
        
' 2.x calc VOH
Dim GetRakValSite As New SiteDouble
For Each pin In Meas_pins
    ' 2.1 Get RAK value
    RAK_Pin = CStr(pin)
    GetRakValSite = CurrentJob_Card_RAK.Pins(pin)
    ' 2.2 voh = I*R + V
    Compensate_V.Pins(pin) = GetRakValSite.Multiply(MeasureValue.Pins(pin)).Abs.Add(inVoh.Pins(pin))
Next pin
      
Dim TempLimitIndex As Long
' 3.x print out datalog
For Each pin In Meas_pins
      If TheExec.TesterMode = testModeOffline Then voh_temp = 0.1
      TestNameInput = Report_TName_From_Instance("I", inVoh.Pins(pin), "OutputCurrent", CInt(i))
      'For Each site In TheExec.sites '20191231
          voh_temp = inVoh.Pins(pin)
          TempLimitIndex = TheExec.Flow.TestLimitIndex
          For Each site In TheExec.sites '20191231
              TheExec.Flow.TestLimitIndex = TempLimitIndex
              TheExec.Flow.TestLimit MeasureValue.Pins(pin), formatStr:="%.4f", Tname:=TestNameInput, ForceVal:=voh_temp, ForceUnit:=unitVolt, ForceResults:=tlForceFlow
          Next site '20191231
Next pin
'      For Each pin In Meas_pins
TestNameInput = Report_TName_From_Instance(CalcV, inVoh.Pins(pin), vbNullString, CInt(i))
TheExec.Flow.TestLimit Compensate_V, , scaletype:=scaleNone, unit:=unitVolt, formatStr:="%.4f", Tname:=TestNameInput, ForceResults:=tlForceFlow

 
' 4.x restore
Dim TestSeq As Long
For TestSeq = 0 To (inVoh.Pins.Count - 1)
      With TheHdw.PPMU.Pins(Meas_pins(TestSeq))
          .ForceI 0, 0.05
          .ForceV 0, 0.05
          .Gate = tlOff
          .Disconnect
      End With
      TheHdw.Digital.Pins(Meas_pins(TestSeq)).Connect
Next TestSeq
 
Exit Function 'Add ErrHandler 2023/05/29
errHandler: 'Add ErrHandler 2023/05/29
    Call Print_Error_Message(Error_Info, "LIB_HardIP_Calc", "compensate_Volt") 'Add ErrHandler 2023/05/29
    If AbortTest Then Exit Function Else Resume Next 'Add ErrHandler 2023/05/29
End Function

Public Function Calc_GPIO_DriverStrength(argc As Integer, argv() As String) As Long
On Error GoTo errHandler 'Add ErrHandler 2023/05/29

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
        DS_Data.Pins(Pin_Ary(i)) = Get_StoredData.Pins(Pin_Ary(i))
        For Each site In TheExec.sites
            DS_Data_DSPwave.Element(i) = DS_Data.Pins(Pin_Ary(i)).value
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
        tempTestLimitIndex = TheExec.Flow.TestLimitIndex
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
            TheExec.Flow.TestLimit DS_Data.Math.Subtract(DS_Avg.divide(1000)), hiVal:=0.005, lowVal:=-0.005, Tname:=TestNameInput, unit:=unitAmp, scaletype:=scaleMicro
        ElseIf Fuse_Bit = 8 Then 'if is DS8
            TheExec.Flow.TestLimit DS_Data.Math.Subtract(DS_Avg.divide(1000)), hiVal:=0.002, lowVal:=-0.002, Tname:=TestNameInput, unit:=unitAmp, scaletype:=scaleMicro
         Else
            TheExec.Flow.TestLimit DS_Data.Math.Subtract(DS_Avg.divide(1000)), hiVal:=0.001, lowVal:=-0.001, Tname:=TestNameInput, unit:=unitAmp, scaletype:=scaleMicro
        End If
    Else
        For j = 0 To UBound(Pin_Ary)
            'For Each site In TheExec.sites
            TheExec.Flow.TestLimitIndex = tempTestLimitIndex
                 If TestName(1) = "ioh" Then
                    TestNameInput = Report_TName_From_Instance(CalcI, CStr(Pin_Ary(j)), "CurrError" & Replace(DS_Fuse_Name, "_", vbNullString), CInt(i))
                    TestNameInput = Replace(TestNameInput, "IOL", "CurrError" & UCase(TestName(1)))
                 Else
                    TestNameInput = Report_TName_From_Instance(CalcI, CStr(Pin_Ary(j)), "CurrError" & Replace(DS_Fuse_Name, "_", vbNullString), CInt(i))
                    TestNameInput = Replace(TestNameInput, "IOL", "CurrError" & UCase(TestName(1)))
                 End If
                 If Fuse_Bit = 9 Then  'if is DS14
                    TheExec.Flow.TestLimit DSP_Result.Element(j), hiVal:=0.005, lowVal:=-0.005, Tname:=TestNameInput, PinName:=Pin_Ary(j), unit:=unitAmp, scaletype:=scaleMicro
                ElseIf Fuse_Bit = 8 Then 'if is DS8
                    TheExec.Flow.TestLimit DSP_Result.Element(j), hiVal:=0.002, lowVal:=-0.002, Tname:=TestNameInput, PinName:=Pin_Ary(j), unit:=unitAmp, scaletype:=scaleMicro
                 Else
                    TheExec.Flow.TestLimit DSP_Result.Element(j), hiVal:=0.001, lowVal:=-0.001, Tname:=TestNameInput, PinName:=Pin_Ary(j), unit:=unitAmp, scaletype:=scaleMicro
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
            TheExec.Flow.TestLimit DS_Avg.divide(1000), Tname:=TestNameInput, PinName:="CurrAvg", unit:=unitAmp, lowVal:=-0.0127, hiVal:=0
        ElseIf Fuse_Bit = 8 Then   'if is DS8
            TheExec.Flow.TestLimit DS_Avg.divide(1000), Tname:=TestNameInput, PinName:="CurrAvg", unit:=unitAmp, lowVal:=-0.0255, hiVal:=0
        Else  'if is DS14
            TheExec.Flow.TestLimit DS_Avg.divide(1000), Tname:=TestNameInput, PinName:="CurrAvg", unit:=unitAmp, lowVal:=-0.0511, hiVal:=0
        End If
    Else
        If Fuse_Bit = 7 Then  'if is DS4
            TheExec.Flow.TestLimit DS_Avg.divide(1000), Tname:=TestNameInput, PinName:="CurrAvg", unit:=unitAmp, hiVal:=0.0127, lowVal:=0
        ElseIf Fuse_Bit = 8 Then   'if is DS8
            TheExec.Flow.TestLimit DS_Avg.divide(1000), Tname:=TestNameInput, PinName:="CurrAvg", unit:=unitAmp, hiVal:=0.0255, lowVal:=0
        Else  'if is DS14
            TheExec.Flow.TestLimit DS_Avg.divide(1000), Tname:=TestNameInput, PinName:="CurrAvg", unit:=unitAmp, hiVal:=0.0511, lowVal:=0
        End If
    End If
    
Exit Function 'Add ErrHandler 2023/05/29
errHandler: 'Add ErrHandler 2023/05/29
    Call Print_Error_Message(Error_Info, "LIB_HardIP_Calc", "Calc_GPIO_DriverStrength") 'Add ErrHandler 2023/05/29
    If AbortTest Then Exit Function Else Resume Next 'Add ErrHandler 2023/05/29
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

Exit Function 'Add ErrHandler 2023/05/29
errHandler: 'Add ErrHandler 2023/05/29
    Call Print_Error_Message(Error_Info, "LIB_HardIP_Calc", "Calc_BinStr2HexStr") 'Add ErrHandler 2023/05/29
    If AbortTest Then Exit Function Else Resume Next 'Add ErrHandler 2023/05/29
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
    Dim sample_size As Long: sample_size = 0
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
                    sample_size = DSPWave_Bin(i).SampleSize
                    Exit For
                Next vsite
            End If
        Next i
        DSPWave_Split_Bit_perDic.Element(serial_data) = sample_size
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
    
    TheHdw.DSP.ExecutionMode = tlDSPModeForceAutomatic
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
        TheExec.Flow.TestLimit resultVal:=DSPWave_AverageDec.Element(serial_data), Tname:=TestNameInput, ForceResults:=tlForceFlow

        Pasing_data_Arry = Split(Input_Arry(serial_data), ",")
        sLen = UBound(Pasing_data_Arry)
        
        If Pasing_data_Arry(sLen) = True Then
            '--- Print STDEV ---- 20230116
            TestNameInput = Report_TName_From_Instance(CalcC, "X", , CInt(i), , , , , tlForceFlow)
            TheExec.Flow.TestLimit resultVal:=DSPWave_STDEV.Element(serial_data), Tname:=TestNameInput, ForceResults:=tlForceFlow
        End If

        Call StoreDataAllType(Pasing_data_Arry(sLen - 1), DSPWave_AverageDec2Bin)
    Next serial_data

Exit Function 'Add ErrHandler 2023/05/29
errHandler: 'Add ErrHandler 2023/05/29
    Call Print_Error_Message(Error_Info, "LIB_HardIP_Calc", "Calc_DigCap_Avg_Store") 'Add ErrHandler 2023/05/29
    If AbortTest Then Exit Function Else Resume Next 'Add ErrHandler 2023/05/29
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
    CurrentPoint = val(TheExec.Flow.var(argv(0)).value)
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
            TheExec.Flow.TestLimit result, , , , , , unitVolt, , Tname:=TestNameInput, ForceResults:=tlForceFlow
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
        Result_DATA = RAW_DATA.divide(1024).Multiply(100)
    Next site
    
    For i = 0 To argc - 1
        TestNameInput = Report_TName_From_Instance("Calc", "DCD", , CInt(i), 0, , , , tlForceFlow)
        TheExec.Flow.TestLimit Result_DATA.Element(i), , , , , , , , Tname:=TestNameInput, ForceResults:=tlForceFlow
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
            TheExec.Flow.TestLimit DSP_EyeWidth.Element(0), DSP_Summary(i).Element(0) / 2, DSP_Summary(i).Element(0) * 2, , , , , , Tname:=TestNameInput, ForceResults:=tlForceFlow
            TheExec.Flow.TestLimitIndex = TheExec.Flow.TestLimitIndex - 1
        Next site
        TheExec.Flow.TestLimitIndex = TheExec.Flow.TestLimitIndex + 1
        
        
    Next i
    
    For i = 0 To argc - 1
        SplitStrAry = Split(argv(i), "&")
        TestNameInput = Report_TName_From_Instance("C", SplitStrAry(1), "Percentage" & i, i, , , , , tlForceNone)
        'For Each site In TheExec.sites
            TheExec.Flow.TestLimit Percentage_bysite(i), 25, 100, , , , , , Tname:=TestNameInput, ForceResults:=tlForceNone 'Un-Used_
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
'''''''''''        TheExec.Flow.TestLimit DSP_EyeWidth.Element(0), 2, 4, , , , , , Tname:=TestNameInput, ForceResults:=tlForceFlow
''''''''''        TheExec.Flow.TestLimitIndex = TheExec.Flow.TestLimitIndex + 1
''''''''''
''''''''''
''''''''''    Next i
Exit Function 'Add ErrHandler 2023/05/29
errHandler: 'Add ErrHandler 2023/05/29
    Call Print_Error_Message(Error_Info, "LIB_HardIP_Calc", "Calc_GRPPercentage") 'Add ErrHandler 2023/05/29
    If AbortTest Then Exit Function Else Resume Next 'Add ErrHandler 2023/05/29
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
            TheExec.Flow.TestLimit SiteDouble_EyeWidth * 100, SiteDouble_Dcode / 8, SiteDouble_Dcode, , , , , , Tname:=TestNameInput, ForceResults:=tlForceFlow
            TheExec.Flow.TestLimitIndex = TheExec.Flow.TestLimitIndex - 1
        Next site
        TheExec.Flow.TestLimitIndex = TheExec.Flow.TestLimitIndex + 1
        
        PercentageTemp(i) = SiteDouble_EyeWidth.divide(SiteDouble_Dcode).Multiply(100)
    Next i
    
    
    For i = 0 To argc - 1
        SplitStrAry = Split(argv(i), "&")
        TestNameInput = Report_TName_From_Instance("C", SplitStrAry(1), "Percentage" & i, i, , , , , tlForceNone_CZ)
        TheExec.Flow.TestLimit PercentageTemp(i), 25, 100, , , , , , Tname:=TestNameInput, ForceResults:=tlForceNone_CZ
    Next i
    
    
    

Exit Function 'Add ErrHandler 2023/05/29
errHandler: 'Add ErrHandler 2023/05/29
    Call Print_Error_Message(Error_Info, "LIB_HardIP_Calc", "Calc_EyeWidthLimitForEachMode") 'Add ErrHandler 2023/05/29
    If AbortTest Then Exit Function Else Resume Next 'Add ErrHandler 2023/05/29
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
    Dim site As Variant 'Carter, 20240304
    
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
                TheExec.Flow.TestLimit resultVal:=DC_Input_CLK_UP.Multiply(100), Tname:=TestNameInput, ForceResults:=tlForceNone_CZ, scaletype:=scaleNoScaling, formatStr:="%1.2f", unit:=unitCustom, customUnit:="%"
                TestNameInput = Report_TName_From_Instance("Calc", SplitGRP(g + 1) & SplitCH(i + 1) & SplitDQ(j + 1), "DCInputCLKNODCC" & CInt(i), CInt(i), , , , , tlForceNone_CZ)
                TheExec.Flow.TestLimit resultVal:=DC_Input_CLK_NO_DCC.Multiply(100), Tname:=TestNameInput, ForceResults:=tlForceNone_CZ, scaletype:=scaleNoScaling, formatStr:="%1.2f", unit:=unitCustom, customUnit:="%"
                TestNameInput = Report_TName_From_Instance("Calc", SplitGRP(g + 1) & SplitCH(i + 1) & SplitDQ(j + 1), "DCInputCLKDOWN" & CInt(i), CInt(i), , , , , tlForceNone_CZ)
                TheExec.Flow.TestLimit resultVal:=DC_Input_CLK_DOWN.Multiply(100), Tname:=TestNameInput, ForceResults:=tlForceNone_CZ, scaletype:=scaleNoScaling, formatStr:="%1.2f", unit:=unitCustom, customUnit:="%"
                
                TestNameInput = Report_TName_From_Instance("Calc", SplitGRP(g + 1) & SplitCH(i + 1) & SplitDQ(j + 1), "DCSkewInputCLKUP" & CInt(i), CInt(i), , , , , tlForceNone_CZ)
                TheExec.Flow.TestLimit resultVal:=DC_Skew_Input_CLK_UP.Multiply(100), Tname:=TestNameInput, ForceResults:=tlForceNone_CZ, scaletype:=scaleNoScaling, formatStr:="%1.2f", unit:=unitCustom, customUnit:="%"
                TestNameInput = Report_TName_From_Instance("Calc", SplitGRP(g + 1) & SplitCH(i + 1) & SplitDQ(j + 1), "DCSkewInputCLKNODCC" & CInt(i), CInt(i), , , , , tlForceNone_CZ)
                TheExec.Flow.TestLimit resultVal:=DC_Skew_Input_CLK_NO_DCC.Multiply(100), Tname:=TestNameInput, ForceResults:=tlForceNone_CZ, scaletype:=scaleNoScaling, formatStr:="%1.2f", unit:=unitCustom, customUnit:="%"
                TestNameInput = Report_TName_From_Instance("Calc", SplitGRP(g + 1) & SplitCH(i + 1) & SplitDQ(j + 1), "DCSkewInputCLKDOWN" & CInt(i), CInt(i), , , , , tlForceNone_CZ)
                TheExec.Flow.TestLimit resultVal:=DC_Skew_Input_CLK_DOWN.Multiply(100), Tname:=TestNameInput, ForceResults:=tlForceNone_CZ, scaletype:=scaleNoScaling, formatStr:="%1.2f", unit:=unitCustom, customUnit:="%"
                
                TestNameInput = Report_TName_From_Instance("Calc", SplitGRP(g + 1) & SplitCH(i + 1) & SplitDQ(j + 1), "DCCRANGEUP" & CInt(i), CInt(i), , , , , tlForceNone_CZ)
                TheExec.Flow.TestLimit resultVal:=DCC_RANGE_UP.Multiply(100), Tname:=TestNameInput, ForceResults:=tlForceNone_CZ, scaletype:=scaleNoScaling, formatStr:="%1.2f", unit:=unitCustom, customUnit:="%"
                TestNameInput = Report_TName_From_Instance("Calc", SplitGRP(g + 1) & SplitCH(i + 1) & SplitDQ(j + 1), "DCCRANGEDOWN" & CInt(i), CInt(i), , , , , tlForceNone_CZ)
                TheExec.Flow.TestLimit resultVal:=DCC_RANGE_DOWN.Multiply(100), Tname:=TestNameInput, ForceResults:=tlForceNone_CZ, scaletype:=scaleNoScaling, formatStr:="%1.2f", unit:=unitCustom, customUnit:="%"
            
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
                        TheExec.Flow.TestLimit DSP_EyeWidth.Element(0), DSP_Summary(i) / 2, DSP_Summary(i) * 2.3, , , , , , PinName:="X", Tname:=TestNameInput, ForceResults:=tlForceFlow
            TheExec.Flow.TestLimitIndex = TheExec.Flow.TestLimitIndex - 1
        Next site
        TheExec.Flow.TestLimitIndex = TheExec.Flow.TestLimitIndex + 1
        
        
    Next i
    
    For i = 0 To argc - 1
        SplitStrAry = Split(argv(i), "&")
        'TestNameInput = Report_TName_From_Instance("C", SplitStrAry(1), "Percentage" & i, i, , , , , tlForceFlow)
        TestNameInput = Report_TName_From_Instance("CalcC", vbNullString, vbNullString, i)
        'For Each site In TheExec.sites
            TheExec.Flow.TestLimit Percentage_bysite(i), 25, 100, , , , , , PinName:="X", Tname:=TestNameInput, ForceResults:=tlForceFlow
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
        'TestNameInput = Report_TName_From_Instance("C", SplitStrAry(1), "Percentage" & i, i, , , , , tlForceFlow)
        TestNameInput = Report_TName_From_Instance("CalcC", vbNullString, vbNullString, i)
        'For Each site In TheExec.sites
            TheExec.Flow.TestLimit Percentage_bysite(i), 25, 120, , , , , , PinName:="X", Tname:=TestNameInput, ForceResults:=tlForceFlow
       ' Next site
    Next i

Exit Function 'Add ErrHandler 2023/05/29
errHandler: 'Add ErrHandler 2023/05/29
    Call Print_Error_Message(Error_Info, "LIB_HardIP_Calc", "Calc_GRPPercentageF2") 'Add ErrHandler 2023/05/29
    If AbortTest Then Exit Function Else Resume Next 'Add ErrHandler 2023/05/29
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
                TCLK = 1250 * PS

''                DNL_MaxMin_HighLimit = 2.5
''                DNL_MaxMin_LowLimit = -1
''                PE_MaxMin_HighLimit = 45
''                PE_MaxMin_LowLimit = -45
        Case "md002"
                Para = 14863473
                TCLK = 625 * PS

''                DNL_MaxMin_HighLimit = 2
''                DNL_MaxMin_LowLimit = -1
''                PE_MaxMin_HighLimit = 30
''                PE_MaxMin_LowLimit = -30
        Case "md003"
                Para = 12958125
                TCLK = 469 * PS

''                DNL_MaxMin_HighLimit = 1.5
''                DNL_MaxMin_LowLimit = -1
''                PE_MaxMin_HighLimit = 25
''                PE_MaxMin_LowLimit = -25
        Case "md004"
                Para = 12228792
                TCLK = 364 * PS

''                DNL_MaxMin_HighLimit = 1.5
''                DNL_MaxMin_LowLimit = -1
''                PE_MaxMin_HighLimit = 25
''                PE_MaxMin_LowLimit = -25
        Case "md005"
                Para = 11861250
                TCLK = 313 * PS

''                DNL_MaxMin_HighLimit = 1.5
''                DNL_MaxMin_LowLimit = -1
''                PE_MaxMin_HighLimit = 20
''                PE_MaxMin_LowLimit = -20
        Case Else
    End Select
    
    LSB = TCLK / 128
    Oct = TCLK / 8
    
    TheExec.Datalog.WriteComment "Mode = " & PerfMode
    TheExec.Datalog.WriteComment "Tclk = " & TCLK / PS & "ps"
    TheExec.Datalog.WriteComment "LSB = " & TCLK / PS & " / 128"
    TheExec.Datalog.WriteComment "Ideal_Oct = " & TCLK / PS & " / 8"
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
                    tempStoreAllDuty.Pins(CategoryName) = NeedSortData
                    CategoryCheckDict.Remove (DutyName)
                    'CategoryCheckDict.Item(DutyName) = tempStoreAllDuty.Copy
                    CategoryCheckDict.Add DutyName, tempStoreAllDuty
                    Call CategoryCheckDict.Remove(DutyName)
                    Call CategoryCheckDict.Add(DutyName, tempStoreAllDuty.Copy)
                Else
                    tempStoreAllDuty.AddPin CategoryName
                    tempStoreAllDuty.Pins(CategoryName) = NeedSortData
                    Call CategoryCheckDict.Add(DutyName, tempStoreAllDuty.Copy)
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
            For j = 0 To CalcDutyVal(i).Pins.Count - 1
                CalcDutyVal(i).Pins(j) = 1000000 - 1000 * j - i * 2000
            Next j
            ''change simulation data
            If (i Mod 2) = 0 Then
                CalcDutyVal(i) = 100000000
            Else
                CalcDutyVal(i) = 50000000
            End If
        End If
        For j = 0 To CalcDutyVal(i).Pins.Count - 1
            'If InStr(UCase(CalcDutyVal(i).Pins(j)), "_N") <> 0 Then 'Oscar
                For Each site In TheExec.sites
                    If CalcDutyVal(i).Pins(j).value(site) = 0 Then
                        b_DivideZeroError(site) = True
                        TheExec.Datalog.WriteComment ("Site " & site & " Freq Meas 0 Hz , No CalcDutyDelay ")
                        CalcDutyVal(i).Pins(j).value = 1
                    End If
                Next site
                
                CalcDutyVal(i).Pins(j).value = CalcDutyVal(i).Pins(j).divide(Para)  'Oscar  ''''''Freq = Capture_code/Parameter
                
                CalcDutyVal(i).Pins(j).value = CalcDutyVal(i).Pins(j).Multiply(2).Invert    ''''''Delay calculation
                                    
                For Each site In TheExec.sites
                    If b_DivideZeroError(site) = True Then
''                        TheExec.Flow.TestLimit resultVal:=-999, ScaleType:=scalePico, Tname:="Delay" & "_" & TestNameInput, ForceResults:=tlForceFlow
                        CalcDutyVal(i).Pins(j).value = -999
                    End If
                Next site
                 
                TestNameInput = Report_TName_From_Instance("F", CalcDutyVal(i).Pins(j), "Delay_Jitter", 0, , , , , tlForceNone_CZ)
                TheExec.Flow.TestLimit resultVal:=CalcDutyVal(i).Pins(j), Tname:=TestNameInput
                
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


            For m = 0 To DNL_Val(k).Pins.Count - 1
                PinName = DNL_Val(k).Pins(m)
                '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                For Each site In TheExec.sites
                    ''20171117 add for debugging
                    If b_DivideZeroError(site) = True Then
                        DNL_Val(k).Pins(m).value = -999
                    End If




                Next site
                '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\

                'Print resuls of "DNL0, DNL1 calculation"'
                TestNameInput = Report_TName_From_Instance("F", DeltaDelayVal(k).Pins(m), "Delta_Delay", 0, , , , , tlForceNone_CZ)
                TheExec.Flow.TestLimit resultVal:=DeltaDelayVal(k).Pins(m), Tname:=TestNameInput
                

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
            tempStoreAllDuty.Pins(CategoryName) = NeedSortData
            CategoryCheckDict.Remove (DutyName)
           ' CategoryCheckDict.Item(DutyName) = tempStoreAllDuty.Copy  'update 20200831
            CategoryCheckDict.Add DutyName, tempStoreAllDuty
            Call CategoryCheckDict.Remove(DutyName)
            Call CategoryCheckDict.Add(DutyName, tempStoreAllDuty.Copy)
        Else
            tempStoreAllDuty.AddPin CategoryName
            tempStoreAllDuty.Pins(CategoryName) = NeedSortData
            Call CategoryCheckDict.Add(DutyName, tempStoreAllDuty.Copy)
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
            For j = 0 To CalcDutyVal(i).Pins.Count - 1
                CalcDutyVal(i).Pins(j) = 1000000 - 1000 * j - i * 2000
            Next j
            ''change simulation data
            If (i Mod 2) = 0 Then
                CalcDutyVal(i) = 100000000
            Else
                CalcDutyVal(i) = 50000000
            End If
        End If
        For j = 0 To CalcDutyVal(i).Pins.Count - 1
            'If InStr(UCase(CalcDutyVal(i).Pins(j)), "_N") <> 0 Then 'Oscar
                For Each site In TheExec.sites
                    If CalcDutyVal(i).Pins(j).value(site) = 0 Then
                        b_DivideZeroError(site) = True
                        TheExec.Datalog.WriteComment ("Site " & site & " Freq Meas 0 Hz , No CalcDutyDelay ")
                        CalcDutyVal(i).Pins(j).value = 1
                    End If
                Next site
                
                CalcDutyVal(i).Pins(j).value = CalcDutyVal(i).Pins(j).divide(Para) 'Oscar   ''''''Freq = Capture_code/Parameter     Modified 20201215
                
                CalcDutyVal(i).Pins(j).value = CalcDutyVal(i).Pins(j).Multiply(2).Invert    ''''''Delay calculation
                    
                For Each site In TheExec.sites
                    If b_DivideZeroError(site) = True Then
''                        TheExec.Flow.TestLimit resultVal:=-999, ScaleType:=scalePico, Tname:="Delay" & "_" & TestNameInput, ForceResults:=tlForceFlow
                        CalcDutyVal(i).Pins(j).value = -999
                    End If
                Next site
                 
                TestNameInput = Report_TName_From_Instance("C", CalcDutyVal(i).Pins(j), "Delay[" & i & "]", 0, , , , , tlForceNone_CZ)
                ''''TheExec.Flow.TestLimit resultVal:=CalcDutyVal(i).Pins(j), lowVal:=DNL_MaxMin_LowLimit, hiVal:=DNL_MaxMin_HighLimit, Tname:=TestNameInput
                TheExec.Flow.TestLimit resultVal:=CalcDutyVal(i).Pins(j), Tname:=TestNameInput
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
        
        ReDim DNL_Val_Max_DigSrcCode(CalcDutyVal(0).Pins.Count - 1) As New SiteLong
        ReDim DNL_Val_Min_DigSrcCode(CalcDutyVal(0).Pins.Count - 1) As New SiteLong
        
        
        

        For k = 0 To MaxNumOfDuty - 1
            DeltaDelayVal(k) = CalcDutyVal(k + 1).Math.Subtract(CalcDutyVal(k)).divide(LSB).Subtract(1)     '''''DNL

            If k = 0 Then  'Initialize Max/Min DNL and DigSrcCode values
                DNL_Val_Max = DeltaDelayVal(0)
                DNL_Val_Min = DeltaDelayVal(0)
                For m = 0 To DeltaDelayVal(k).Pins.Count - 1
                    DNL_Val_Max_DigSrcCode(m) = 0
                    DNL_Val_Min_DigSrcCode(m) = 0
                Next m
            End If
            
            For m = 0 To DeltaDelayVal(k).Pins.Count - 1
                PinName = DeltaDelayVal(k).Pins(m)
                '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                For Each site In TheExec.sites
                    ''20171117 add for debugging
                    If b_DivideZeroError(site) = True Then
                        DeltaDelayVal(k).Pins(m).value = -999
                    End If
                    
                    'Find MAX/MIN of DNL0/1 for DigSrc Code[0:111]
                    If DeltaDelayVal(k).Pins(m).value > DNL_Val_Max.Pins(m).value Then
                        DNL_Val_Max.Pins(m).value = DeltaDelayVal(k).Pins(m).value
                        DNL_Val_Max_DigSrcCode(m) = k
                    End If
                    If DeltaDelayVal(k).Pins(m).value < DNL_Val_Min.Pins(m).value Then
                        DNL_Val_Min.Pins(m).value = DeltaDelayVal(k).Pins(m).value
                        DNL_Val_Min_DigSrcCode(m) = k
                    End If
                    
                Next site
                '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                
                'Print resuls of "DNL0, DNL1 calculation"'
                TestNameInput = Report_TName_From_Instance("C", DeltaDelayVal(k).Pins(m), "DNL[" & k & "]", 0, , , , , tlForceNone_CZ)
                TheExec.Flow.TestLimit resultVal:=DeltaDelayVal(k).Pins(m), Tname:=TestNameInput      '''''DNL


            Next m
                       
        Next k
        
        
        'Print resuls of finding Max/Min of "DNL0, DNL1" value and DigSrc code'
        For p = 0 To DNL_Val_Max.Pins.Count - 1
            TestNameInput = Report_TName_From_Instance("C", DNL_Val_Max.Pins(p), "Max-Delta-Delay", 0, , , , , tlForceNone_CZ)
            TheExec.Flow.TestLimit resultVal:=DNL_Val_Max.Pins(p), lowVal:=DNL_MaxMin_LowLimit, hiVal:=DNL_MaxMin_HighLimit, Tname:=TestNameInput, ForceResults:=tlForceNone_CZ
            TestNameInput = Report_TName_From_Instance("C", "X", "Max-Delta-Delay-DigSrcCode", 0, , , , , tlForceNone_CZ)
            TheExec.Flow.TestLimit resultVal:=DNL_Val_Max_DigSrcCode(p), Tname:=TestNameInput
        Next
        
        For p = 0 To DNL_Val_Max.Pins.Count - 1
            TestNameInput = Report_TName_From_Instance("C", DNL_Val_Min.Pins(p), "Min-Delta-Delay", 0, , , , , tlForceNone_CZ)
            TheExec.Flow.TestLimit resultVal:=DNL_Val_Min.Pins(p), lowVal:=DNL_MaxMin_LowLimit, hiVal:=DNL_MaxMin_HighLimit, Tname:=TestNameInput, ForceResults:=tlForceNone_CZ
            TestNameInput = Report_TName_From_Instance("C", "X", "Min-Delta-Delay-DigSrcCode", 0, , , , , tlForceNone_CZ)
            TheExec.Flow.TestLimit resultVal:=DNL_Val_Min_DigSrcCode(p), Tname:=TestNameInput
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
        
        ReDim PE_Val_Max_DigSrcCode(CalcDutyVal(0).Pins.Count - 1) As New SiteLong
        ReDim PE_Val_Min_DigSrcCode(CalcDutyVal(0).Pins.Count - 1) As New SiteLong
        
        TheExec.Datalog.WriteComment "***PE0 to PE7 calculation for" & PerfMode & " ***"

        For n = 0 To OctantMaxNum
            If n <> 7 Then
                Octant_Val(n) = CalcDutyVal((n + 1) * 16).Math.Subtract(CalcDutyVal(n * 16)).Subtract(Oct)
                '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                ''20171117 add for debugging
                For m = 0 To Octant_Val(n).Pins.Count - 1
                    For Each site In TheExec.sites
                        If b_DivideZeroError(site) = True Then
                            Octant_Val(n).Pins(m).value = -999
                        End If
                    Next site
                Next m
                '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
            Else
                Octant_Val(n) = CalcDutyVal(n * 0).Math.Subtract(CalcDutyVal(n * 16)).Add(TCLK).Subtract(Oct)
                '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                ''20171117 add for debugging
                For m = 0 To Octant_Val(n).Pins.Count - 1
                    For Each site In TheExec.sites
                        If b_DivideZeroError(site) = True Then
                            Octant_Val(n).Pins(m).value = -999
                        End If
                    Next site
                Next m
                '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
            End If
            'TestNameInput = Report_TName_From_Instance("C", "X", "PE" & n & "Delta_Delay", 0, , , , , tlForceFlow)
            'TheExec.Flow.TestLimit resultVal:=Octant_Val(n), Tname:=TestNameInput           '''''Modified 20201215
            
            For m = 0 To Octant_Val(n).Pins.Count - 1
                TestNameInput = Report_TName_From_Instance("C", Octant_Val(n).Pins(m), "PE" & n & "-Delta-Delay", 0, , , , , tlForceNone_CZ)
                TheExec.Flow.TestLimit resultVal:=Octant_Val(n).Pins(m), Tname:=TestNameInput           '''''Modified 20201215
            Next m
        
        
        'Find MAX/MIN of PE0 to PE7         add 20201215
            
            If n = 0 Then  'Initialize Max/Min PE and DigSrcCode values
                PE_Val_Max = Octant_Val(0)
                PE_Val_Min = Octant_Val(0)
                For m = 0 To Octant_Val(n).Pins.Count - 1
                    PE_Val_Max_DigSrcCode(m) = 0
                    PE_Val_Min_DigSrcCode(m) = 0
                Next m
            End If
        
            For m = 0 To Octant_Val(n).Pins.Count - 1

                '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                For Each site In TheExec.sites
                    If b_DivideZeroError(site) = True Then
                        Octant_Val(n).Pins(m).value = -999
                    End If

                    If Octant_Val(n).Pins(m).value > PE_Val_Max.Pins(m).value Then
                        PE_Val_Max.Pins(m).value = Octant_Val(n).Pins(m).value
                        PE_Val_Max_DigSrcCode(m) = n
                    End If
                    If DeltaDelayVal(n).Pins(m).value < PE_Val_Min.Pins(m).value Then
                        PE_Val_Min.Pins(m).value = Octant_Val(n).Pins(m).value
                        PE_Val_Min_DigSrcCode(m) = n
                    End If
                    
                Next site
                '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
      
            Next m
            
            
        Next n
        
        For p = 0 To PE_Val_Max.Pins.Count - 1
            TestNameInput = Report_TName_From_Instance("C", PE_Val_Max.Pins(p), "Max-PE", 0, , , , , tlForceNone_CZ)
            TheExec.Flow.TestLimit resultVal:=PE_Val_Max.Pins(p), lowVal:=PE_MaxMin_LowLimit, hiVal:=PE_MaxMin_HighLimit, Tname:=TestNameInput, ForceResults:=tlForceNone_CZ
            TestNameInput = Report_TName_From_Instance("C", "X", "Max-PE-DigSrcCode", 0, , , , , tlForceNone_CZ)
            TheExec.Flow.TestLimit resultVal:=PE_Val_Max_DigSrcCode(p), Tname:=TestNameInput
        Next
        
        For p = 0 To PE_Val_Max.Pins.Count - 1
            TestNameInput = Report_TName_From_Instance("C", PE_Val_Min.Pins(p), "Min-PE", 0, , , , , tlForceNone_CZ)
            TheExec.Flow.TestLimit resultVal:=PE_Val_Min.Pins(p), lowVal:=PE_MaxMin_LowLimit, hiVal:=PE_MaxMin_HighLimit, Tname:=TestNameInput, ForceResults:=tlForceNone_CZ
            TestNameInput = Report_TName_From_Instance("C", "X", "Min-PE-DigSrcCode", 0, , , , , tlForceNone_CZ)
            TheExec.Flow.TestLimit resultVal:=PE_Val_Min_DigSrcCode(p), Tname:=TestNameInput
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
            tempStoreAllDuty.Pins(CategoryName) = NeedSortData
            CategoryCheckDict.Remove (DutyName)
           ' CategoryCheckDict.Item(DutyName) = tempStoreAllDuty.Copy  'update 20200831
            CategoryCheckDict.Add DutyName, tempStoreAllDuty
            Call CategoryCheckDict.Remove(DutyName)
            Call CategoryCheckDict.Add(DutyName, tempStoreAllDuty.Copy)
        Else
            tempStoreAllDuty.AddPin CategoryName
            tempStoreAllDuty.Pins(CategoryName) = NeedSortData
            Call CategoryCheckDict.Add(DutyName, tempStoreAllDuty.Copy)
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
            For j = 0 To CalcDutyVal(i).Pins.Count - 1
                CalcDutyVal(i).Pins(j) = 1000000 - 1000 * j - i * 2000
            Next j
            ''change simulation data
            If (i Mod 2) = 0 Then
                CalcDutyVal(i) = 100000000
            Else
                CalcDutyVal(i) = 50000000
            End If
        End If
        For j = 0 To CalcDutyVal(i).Pins.Count - 1
            'If InStr(UCase(CalcDutyVal(i).Pins(j)), "_N") <> 0 Then 'Oscar
                For Each site In TheExec.sites
                    If CalcDutyVal(i).Pins(j).value(site) = 0 Then
                        b_DivideZeroError(site) = True
                        TheExec.Datalog.WriteComment ("Site " & site & " Freq Meas 0 Hz , No CalcDutyDelay ")
                        CalcDutyVal(i).Pins(j).value = 1
                    End If
                Next site
                
                CalcDutyVal(i).Pins(j).value = CalcDutyVal(i).Pins(j).divide(Para) 'Oscar   ''''''Freq = Capture_code/Parameter     Modified 20201215
                
                CalcDutyVal(i).Pins(j).value = CalcDutyVal(i).Pins(j).Multiply(2).Invert    ''''''Delay calculation
                    
                For Each site In TheExec.sites
                    If b_DivideZeroError(site) = True Then
''                        TheExec.Flow.TestLimit resultVal:=-999, ScaleType:=scalePico, Tname:="Delay" & "_" & TestNameInput, ForceResults:=tlForceFlow
                        CalcDutyVal(i).Pins(j).value = -999
                    End If
                Next site
                 
                'TestNameInput = Report_TName_From_Instance("F", CalcDutyVal(i).Pins(j), CalcDutyVal(i).Pins(j) & "__Jitter-" & CStr(i + StartNumOfDuty), 0, , , , , tlForceFlow)
                TestNameInput = Report_TName_From_Instance("C", CalcDutyVal(i).Pins(j), "Delay[" & i & "]", 0, , , , , tlForceNone_CZ)
                ''''TheExec.Flow.TestLimit resultVal:=CalcDutyVal(i).Pins(j), lowVal:=DNL_MaxMin_LowLimit, hiVal:=DNL_MaxMin_HighLimit, Tname:=TestNameInput
                TheExec.Flow.TestLimit resultVal:=CalcDutyVal(i).Pins(j), Tname:=TestNameInput
                
                
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
        
        ReDim DNL_Val_Max_DigSrcCode(CalcDutyVal(0).Pins.Count - 1) As New SiteLong
        ReDim DNL_Val_Min_DigSrcCode(CalcDutyVal(0).Pins.Count - 1) As New SiteLong
        
        
        

        'For k = 0 To MaxNumOfDuty - 1
        For k = 0 To MaxNumOfDuty - StartNumOfDuty - 1 '@CW
            DeltaDelayVal(k) = CalcDutyVal(k + 1).Math.Subtract(CalcDutyVal(k)).divide(LSB).Subtract(1)     '''''DNL

            If k = 0 Then  'Initialize Max/Min DNL and DigSrcCode values
                DNL_Val_Max = DeltaDelayVal(0)
                DNL_Val_Min = DeltaDelayVal(0)
                For m = 0 To DeltaDelayVal(k).Pins.Count - 1
                    DNL_Val_Max_DigSrcCode(m) = 0
                    DNL_Val_Min_DigSrcCode(m) = 0
                Next m
            End If
            
            For m = 0 To DeltaDelayVal(k).Pins.Count - 1
                PinName = DeltaDelayVal(k).Pins(m)
                '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                For Each site In TheExec.sites
                    ''20171117 add for debugging
                    If b_DivideZeroError(site) = True Then
                        DeltaDelayVal(k).Pins(m).value = -999
                    End If
                    
                    'Find MAX/MIN of DNL0/1 for DigSrc Code[0:111]
                    If DeltaDelayVal(k).Pins(m).value > DNL_Val_Max.Pins(m).value Then
                        DNL_Val_Max.Pins(m).value = DeltaDelayVal(k).Pins(m).value
                        DNL_Val_Max_DigSrcCode(m) = k
                    End If
                    If DeltaDelayVal(k).Pins(m).value < DNL_Val_Min.Pins(m).value Then
                        DNL_Val_Min.Pins(m).value = DeltaDelayVal(k).Pins(m).value
                        DNL_Val_Min_DigSrcCode(m) = k
                    End If
                    
                Next site
                '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                
                'Print resuls of "DNL0, DNL1 calculation"'
                TestNameInput = Report_TName_From_Instance("C", DeltaDelayVal(k).Pins(m), "DNL[" & k & "]", 0, , , , , tlForceNone)
                TheExec.Flow.TestLimit resultVal:=DeltaDelayVal(k).Pins(m), Tname:=TestNameInput      '''''DNL

                'TestNameInput = Report_TName_From_Instance("F", CalcDutyVal(k).Pins(m), "", 0, , , , , tlForceFlow)
                'TheExec.Flow.TestLimit resultVal:=DeltaDelayVal(k).Pins(m), Tname:=TestNameInput, ForceResults:=tlForceFlow

            Next m
                       
        Next k
        
        
        'Print resuls of finding Max/Min of "DNL0, DNL1" value and DigSrc code'
        For p = 0 To DNL_Val_Max.Pins.Count - 1
            TestNameInput = Report_TName_From_Instance("C", DNL_Val_Max.Pins(p), "Max-Delta-Delay", 0, , , , , tlForceNone)
            TheExec.Flow.TestLimit resultVal:=DNL_Val_Max.Pins(p), lowVal:=DNL_MaxMin_LowLimit, hiVal:=DNL_MaxMin_HighLimit, Tname:=TestNameInput, ForceResults:=tlForceNone 'Un-Used_
            TestNameInput = Report_TName_From_Instance("C", DNL_Val_Max.Pins(p), "Max-Delta-Delay-DigSrcCode", 0, , , , , tlForceNone)
            TheExec.Flow.TestLimit resultVal:=DNL_Val_Max_DigSrcCode(p), PinName:=DNL_Val_Max.Pins(p), Tname:=TestNameInput
        Next
        
        For p = 0 To DNL_Val_Max.Pins.Count - 1
            TestNameInput = Report_TName_From_Instance("C", DNL_Val_Min.Pins(p), "Min-Delta-Delay", 0, , , , , tlForceNone)
            TheExec.Flow.TestLimit resultVal:=DNL_Val_Min.Pins(p), lowVal:=DNL_MaxMin_LowLimit, hiVal:=DNL_MaxMin_HighLimit, Tname:=TestNameInput, ForceResults:=tlForceNone 'Un-Used_
            TestNameInput = Report_TName_From_Instance("C", DNL_Val_Min.Pins(p), "Min-Delta-Delay-DigSrcCode", 0, , , , , tlForceNone)
            TheExec.Flow.TestLimit resultVal:=DNL_Val_Min_DigSrcCode(p), PinName:=DNL_Val_Min.Pins(p), Tname:=TestNameInput
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
        
        ReDim PE_Val_Max_DigSrcCode(CalcDutyVal(0).Pins.Count - 1) As New SiteLong
        ReDim PE_Val_Min_DigSrcCode(CalcDutyVal(0).Pins.Count - 1) As New SiteLong
        
        TheExec.Datalog.WriteComment "***PE0 to PE7 calculation for" & PerfMode & " ***"

        For n = 0 To OctantMaxNum
            If n <> 7 Then
                Octant_Val(n) = CalcDutyVal((n + 1) * 16).Math.Subtract(CalcDutyVal(n * 16)).Subtract(Oct)
                '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                ''20171117 add for debugging
                For m = 0 To Octant_Val(n).Pins.Count - 1
                    For Each site In TheExec.sites
                        If b_DivideZeroError(site) = True Then
                            Octant_Val(n).Pins(m).value = -999
                        End If
                    Next site
                Next m
                '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
            Else
                Octant_Val(n) = CalcDutyVal(n * 0).Math.Subtract(CalcDutyVal(n * 16)).Add(TCLK).Subtract(Oct)
                '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                ''20171117 add for debugging
                For m = 0 To Octant_Val(n).Pins.Count - 1
                    For Each site In TheExec.sites
                        If b_DivideZeroError(site) = True Then
                            Octant_Val(n).Pins(m).value = -999
                        End If
                    Next site
                Next m
                '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
            End If
            'TestNameInput = Report_TName_From_Instance("C", "X", "PE" & n & "Delta_Delay", 0, , , , , tlForceFlow)
            'TheExec.Flow.TestLimit resultVal:=Octant_Val(n), Tname:=TestNameInput           '''''Modified 20201215
            
            For m = 0 To Octant_Val(n).Pins.Count - 1
                TestNameInput = Report_TName_From_Instance("C", Octant_Val(n).Pins(m), "PE" & n & "-Delta-Delay", 0, , , , , tlForceNone)
                TheExec.Flow.TestLimit resultVal:=Octant_Val(n).Pins(m), Tname:=TestNameInput           '''''Modified 20201215
            Next m
        
        
        'Find MAX/MIN of PE0 to PE7         add 20201215
            
            If n = 0 Then  'Initialize Max/Min PE and DigSrcCode values
                PE_Val_Max = Octant_Val(0)
                PE_Val_Min = Octant_Val(0)
                For m = 0 To Octant_Val(n).Pins.Count - 1
                    PE_Val_Max_DigSrcCode(m) = 0
                    PE_Val_Min_DigSrcCode(m) = 0
                Next m
            End If
        
            For m = 0 To Octant_Val(n).Pins.Count - 1

                '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
                For Each site In TheExec.sites
                    If b_DivideZeroError(site) = True Then
                        Octant_Val(n).Pins(m).value = -999
                    End If

                    If Octant_Val(n).Pins(m).value > PE_Val_Max.Pins(m).value Then
                        PE_Val_Max.Pins(m).value = Octant_Val(n).Pins(m).value
                        PE_Val_Max_DigSrcCode(m) = n
                    End If
                    If DeltaDelayVal(n).Pins(m).value < PE_Val_Min.Pins(m).value Then
                        PE_Val_Min.Pins(m).value = Octant_Val(n).Pins(m).value
                        PE_Val_Min_DigSrcCode(m) = n
                    End If
                    
                Next site
                '\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
      
            Next m
            
            
        Next n
        
        For p = 0 To PE_Val_Max.Pins.Count - 1
            TestNameInput = Report_TName_From_Instance("C", PE_Val_Max.Pins(p), "Max-PE", 0, , , , , tlForceNone)
            TheExec.Flow.TestLimit resultVal:=PE_Val_Max.Pins(p), lowVal:=PE_MaxMin_LowLimit, hiVal:=PE_MaxMin_HighLimit, Tname:=TestNameInput, ForceResults:=tlForceNone 'Un-Used_
            TestNameInput = Report_TName_From_Instance("C", "X", "Max-PE-DigSrcCode", 0, , , , , tlForceNone)
            TheExec.Flow.TestLimit resultVal:=PE_Val_Max_DigSrcCode(p), Tname:=TestNameInput
        Next
        
        For p = 0 To PE_Val_Max.Pins.Count - 1
            TestNameInput = Report_TName_From_Instance("C", PE_Val_Min.Pins(p), "Min-PE", 0, , , , , tlForceNone)
            TheExec.Flow.TestLimit resultVal:=PE_Val_Min.Pins(p), lowVal:=PE_MaxMin_LowLimit, hiVal:=PE_MaxMin_HighLimit, Tname:=TestNameInput, ForceResults:=tlForceNone 'Un-Used_
            TestNameInput = Report_TName_From_Instance("C", "X", "Min-PE-DigSrcCode", 0, , , , , tlForceNone)
            TheExec.Flow.TestLimit resultVal:=PE_Val_Min_DigSrcCode(p), Tname:=TestNameInput
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
                CalculatedDsp(i) = TrimCodeEqDsp.Multiply(CDbl(InputDicVal(i)(vsite))).Subtract(TrimTarget).Abs.Copy
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
       TheExec.Flow.TestLimit resultVal:=Result_Code(i), Tname:=TestNameInput, ForceResults:=tlForceFlow
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
    Dim site As Variant 'Carter, 20240304
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
    TheExec.Flow.TestLimit resultVal:=DSP_Avg_value_DEC_efuse.Element(0), Tname:=TestNameInput, ForceResults:=tlForceFlow, scaletype:=scaleNoScaling
    
    Call StoreDataAllType(Store_Dictionary_name, FuseDSP)

Exit Function
errHandler: 'Add ErrHandler 2023/05/29
    Call Print_Error_Message(Error_Info, "LIB_HardIP_Calc", "Calc_DigCap_Dec_Avg") 'Add ErrHandler 2023/05/29
    If AbortTest Then Exit Function Else Resume Next 'Add ErrHandler 2023/05/29
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
                                    'Do nothing
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
                                        Else
                                        'Do nothing
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
Exit Function 'Add ErrHandler 2023/05/29
errHandler: 'Add ErrHandler 2023/05/29
    Call Print_Error_Message(Error_Info, "LIB_HardIP_Calc", "P2PBundle_Unflipeye_NewCal") 'Add ErrHandler 2023/05/29
    If AbortTest Then Exit Function Else Resume Next 'Add ErrHandler 2023/05/29
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
        TheExec.Flow.TestLimit resultVal:=meas_val1, Tname:=TestNameInput, ForceResults:=tlForceFlow, scaletype:=scalePico

      TestNameInput = Report_TName_From_Instance(CalcF, vbNullString)
        TheExec.Flow.TestLimit resultVal:=meas_val2, Tname:=TestNameInput, ForceResults:=tlForceFlow, scaletype:=scalePico

      TestNameInput = Report_TName_From_Instance(CalcF, vbNullString)
        TheExec.Flow.TestLimit resultVal:=meas_val3, Tname:=TestNameInput, ForceResults:=tlForceFlow, scaletype:=scalePico



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
        TheExec.Flow.TestLimit resultVal:=meas_val2.Pins(argv(0)), Tname:=TestNameInput, ForceResults:=tlForceFlow, scaletype:=scalePico
        TestNameInput = Report_TName_From_Instance(CalcF, vbNullString)
        TheExec.Flow.TestLimit resultVal:=meas_val2.Pins(argv(1)), Tname:=TestNameInput, ForceResults:=tlForceFlow, scaletype:=scalePico
    'meas_val_now.Pins(i).Value
      Next i
Exit Function 'Add ErrHandler 2023/05/29
errHandler: 'Add ErrHandler 2023/05/29
    Call Print_Error_Message(Error_Info, "LIB_HardIP_Calc", "Calc_RX12_Freq") 'Add ErrHandler 2023/05/29
    If AbortTest Then Exit Function Else Resume Next 'Add ErrHandler 2023/05/29
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
    Dim sample_size As Long: sample_size = 0
    
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
    
    TheHdw.DSP.ExecutionMode = tlDSPModeAutomatic
    
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
        TheHdw.DSP.ExecutionMode = tlDSPModeForceAutomatic

        Call rundsp.DSPWF_OFFSET(DSPWave_LowBin, lVoltage_Low, DSPWave_HighBin, lVoltage_High, lYintercept_spec, DSPWave_SlopeDec, DSPWave_YaxisBDec, _
                                 DSPWave_InterpolateDec, DSPWave_OffsetDec, DSPWave_OffsetBin)

        For Each site In TheExec.sites
            TheExec.Datalog.WriteComment "Site : " & site & ", Offset result:" & DSPWave_OffsetDec(site).Element(0)
            DSPWave_OffsetDec(site).Element(0) = FormatNumber(DSPWave_OffsetDec(site).Element(0), 0)
        Next site
    
        TestNameInput = Report_TName_From_Instance(CalcC, "X", , , , , , , tlForceFlow)
        TheExec.Flow.TestLimit resultVal:=DSPWave_OffsetDec.Element(0), Tname:=TestNameInput, ForceResults:=tlForceFlow  'tlForceFlow
        
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
    
    TheHdw.DSP.ExecutionMode = tlDSPModeForceAutomatic
    
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
    Dim tempResult As New SiteLong
    Dim SD_tempResult As New SiteDouble
    Dim temp(0) As Long
    For j = 0 To argc \ 4 - 1
    
        tempResult = sl_Mean_Val.Element(j)
        TestNameInput = Report_TName_From_Instance(CalcC, "X", , 0, 0, ForceResult:=tlForceFlow)
        TheExec.Flow.TestLimit resultVal:=tempResult, Tname:=TestNameInput, ForceResults:=tlForceFlow, PinName:="X"


        tempResult = sl_Diff_MaxMin.Element(j)
        TestNameInput = Report_TName_From_Instance(CalcC, "X", , 0, 0, ForceResult:=tlForceFlow)
        TheExec.Flow.TestLimit resultVal:=tempResult, Tname:=TestNameInput, ForceResults:=tlForceFlow, PinName:="X"

        tempResult = sl_MDLL_UniqueDirection.Element(j)
        TestNameInput = Report_TName_From_Instance(CalcC, "X", , 0, 0, ForceResult:=tlForceFlow)
        TheExec.Flow.TestLimit resultVal:=tempResult, Tname:=TestNameInput, ForceResults:=tlForceFlow, PinName:="X"


        tempResult = sl_MDLL_DecreaseDirection.Element(j)
        TestNameInput = Report_TName_From_Instance(CalcC, "X", , 0, 0, ForceResult:=tlForceFlow)
        TheExec.Flow.TestLimit resultVal:=tempResult, Tname:=TestNameInput, ForceResults:=tlForceFlow, PinName:="X"


        For Each site In TheExec.sites
            temp(0) = sl_Sum_Val.Element(j)
            sl_Temp_Sum.data = temp
        Next site
        Call StoreDataAllType("Summary" & argv(j * 4), sl_Temp_Sum)
'        Call StoreDataAllType("Summary" & argv(j * 4), sl_Sum_Val.Element(j))
    Next j
    
    
''ProfileMarkLeave LIB_HardIP_Calc_ProfileMark_12745    ' Profile Mark

Exit Function 'Add ErrHandler 2023/05/29
errHandler: 'Add ErrHandler 2023/05/29
    Call Print_Error_Message(Error_Info, "LIB_HardIP_Calc", "Calc_MDLL_Monotonicity_DevideBlock_SEGTTR") 'Add ErrHandler 2023/05/29
    If AbortTest Then Exit Function Else Resume Next 'Add ErrHandler 2023/05/29
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
        'TestNameInput = Report_TName_From_Instance("C", SplitStrAry(1), "Percentage" & i, i, , , , , tlForceFlow)
        TestNameInput = Report_TName_From_Instance("C", "X", , , , , , , ForceResult:=tlForceFlow)
        TheExec.Flow.TestLimit PercentageTemp(i), 25, 100, , , , , , PinName:="X", Tname:=TestNameInput, ForceResults:=tlForceFlow
    Next i
Exit Function 'Add ErrHandler 2023/05/29
errHandler: 'Add ErrHandler 2023/05/29
    Call Print_Error_Message(Error_Info, "LIB_HardIP_Calc", "Calc_EyeWidthLimitForEachMode_TTR") 'Add ErrHandler 2023/05/29
    If AbortTest Then Exit Function Else Resume Next 'Add ErrHandler 2023/05/29
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
    TheExec.Flow.TestLimit resultVal:=DSPGainTrim.Element(0), Tname:=TestNameInput, ForceResults:=tlForceFlow, scaletype:=scaleNoScaling, formatStr:="%.3f"
    TestNameInput = Report_TName_From_Instance("CalcC", "X", , 0, 0)
    TheExec.Flow.TestLimit resultVal:=DSPOffsetTrim.Element(0), Tname:=TestNameInput, ForceResults:=tlForceFlow, scaletype:=scaleNoScaling, formatStr:="%.3f"
    
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

Public Function Calc_DSP_Dictionary_Process(argc As Integer, argv() As String) As Long
On Error GoTo errHandler
'Alg::Calc_DSP_Dictionary_Process(111&Dict_A@0@2,Dict_B)
    'New Calc for T-Don --20230118
    'Reference DigSrc Setup syntax to convert DSPWave data for eFuse
    'Case support Dictionary, "&", "@", "copy"
    'DicA = 1001001001
    'Ex. DicA&DictA@5@5@copy@5 => 100110011111
    '    0111&DictA@5@5@copy@5 => 011111111
    '    DicA@5@8  => 0100
    
    Dim i, j, k As Long
    Dim Vsite As Variant
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
    Dim site As Variant 'Carter, 20240304
    theexec.Datalog.WriteComment "-----     Calc_DSP_Dictionary_Process Start     -----"
    
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
                For Each Vsite In theexec.sites
                    For j = 0 To Len(RdIn(0)) - 1
                        DSPWF_TempData.Element(j) = CDbl(mid(RdIn(0), j + 1, 1))
                    Next j
                Next Vsite
            End If
       
            '==== Copy Format ====
            If UBound(RdIn) = 4 Then
                StartBit = RdIn(1)
                StopBit = RdIn(2)
                Format = RdIn(3)
                CopyTimes = RdIn(4)
    
                DSPWF_TempDataForCopy.CreateConstant 0, CDbl(StopBit - StartBit) + 1, DspDouble
                
                For Each Vsite In theexec.sites
                    For j = StartBit To StopBit
                        DSPWF_TempDataForCopy.Element(j - CDbl(StartBit)) = DSPWF_DictData.Element(j)
                    Next j
                
                    For k = 1 To CopyTimes
                        If k = 1 Then
                            DSPWF_TempData = DSPWF_TempDataForCopy.Copy
                        Else
                            DSPWF_TempData = DSPWF_TempData.Concatenate(DSPWF_TempDataForCopy)
                        End If
                    Next k
                Next Vsite
            ElseIf UBound(RdIn) = 2 Then
                StartBit = RdIn(1)
                StopBit = RdIn(2)
    
                DSPWF_TempDataForCopy.CreateConstant 0, CDbl(StopBit - StartBit) + 1, DspDouble
                
                For Each Vsite In theexec.sites
                    For j = StartBit To StopBit
                        DSPWF_TempDataForCopy.Element(j - CDbl(StartBit)) = DSPWF_DictData.Element(j)
                    Next j
                    DSPWF_TempData = DSPWF_TempDataForCopy.Copy
                Next Vsite
            ElseIf b_WithDictionary Then
                For Each Vsite In theexec.sites
                    DSPWF_TempData = DSPWF_DictData.Copy
                Next Vsite
            Else
            End If
            
            For Each Vsite In theexec.sites
                If i = 0 Then
                    DSPWF_OutputData = DSPWF_TempData.Copy
                Else
                    DSPWF_OutputData = DSPWF_OutputData.Concatenate(DSPWF_TempData)
                End If
            Next Vsite
            
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
            
            For Each Vsite In theexec.sites
                For i = 0 To Len(RdIn(1)) - 1
                    DSPWF_TempData.Element(i) = CDbl(mid(RdIn(1), i + 1, 1))
                Next i
            Next Vsite
            For Each Vsite In theexec.sites
                DSPWF_OutputData = DSPWF_TempData.Copy
            Next Vsite
            
        Else
            theexec.Datalog.WriteComment "Error! Calc_DSP_Dictionary_Process : Format is incorrect"
            Exit Function
        End If
        Set DSPWF_DictData = Nothing
        Set DSPWF_TempData = Nothing
    Else

        If IsNumeric(Input_Str) = True Then
            DSPWF_TempData.CreateConstant 0, Len(Input_Str), DspLong
            
            For Each Vsite In theexec.sites
                For i = 0 To Len(Input_Str) - 1
                    DSPWF_TempData.Element(i) = CDbl(mid(Input_Str, i + 1, 1))
                Next i
            Next Vsite
        Else
            DSPWF_TempData = GetStoreDataAllType(Input_Str)
        End If
        For Each Vsite In theexec.sites
            DSPWF_OutputData = DSPWF_TempData.Copy
        Next Vsite
        Set DSPWF_DictData = Nothing
        Set DSPWF_TempData = Nothing
    End If
    
    If gl_Disable_HIP_debug_log = False Then
        Dim s_result As String
        theexec.Datalog.WriteComment "Output DSP Process String [ LSB(L) ==> MSB(R) ]:"
        For Each site In theexec.sites.Active
            s_result = ""
            For i = 0 To DSPWF_OutputData.SampleSize - 1
                s_result = s_result & CStr(DSPWF_OutputData.Element(i))
            Next i
            theexec.Datalog.WriteComment "Site [" & site & "] " & "Process DSP : " & Output_StoreName & " = " & Input_Str & " = " & s_result
        Next site
    End If


    '===== Store Final Data to Dictionary ======
    Call StoreDataAllType(Output_StoreName, DSPWF_OutputData)
    
    Exit Function
    
errHandler:
    theexec.Datalog.WriteComment "error in Calc_DSP_Dictionary_Process"
    If AbortTest Then Exit Function Else Resume Next
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

Public Function Calc_VDAC(argc As Integer, argv() As String) As Long
On Error GoTo errHandler
    
    Dim TestNameInput As String
    Dim i, Resolution, ratio As Double
    Dim VDAC_Value As New SiteDouble
    
    Resolution = argv(argc - 2)
    ratio = argv(argc - 1)
    
    For i = 0 To argc - 3
        VDAC_Value = GetStoreDataAllType(argv(i) & "_para")
        VDAC_Value = VDAC_Value.divide(Resolution).Multiply(ratio)
        
        TestNameInput = Report_TName_From_Instance(CalcC, "", , , , , , , tlForceFlow)
        TheExec.Flow.TestLimit resultVal:=VDAC_Value, PinName:="", Tname:=TestNameInput, scaletype:=scaleNone, ForceResults:=tlForceFlow
    Next i
        
    Exit Function

errHandler:
    Call Print_Error_Message(Error_Info, "LIB_HardIP_Calc", "Calc_VDAC") 
    If AbortTest Then Exit Function Else Resume Next 
End Function

Public Function TX_Level_Pingroup(argc As Integer, argv() As String) As Long
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
        
        For i = 0 To PLD_Calc_A.Pins.Count - 1
            PinName = UCase(PLD_Calc_A.Pins.item(i))
            PLD_R_Path.AddPin (PinName)
            PLD_R_Path.Pins(PinName) = CurrentJob_Card_RAK.Pins(PinName).Subtract(R_Term).Negate
        Next i
        
    
        PLD_Calc_B1 = PLD_R_Path.Math.Multiply(SD_I2_I1)
        PLD_Calc_B = pld_V1.Math.Subtract(pld_V2).Add(PLD_Calc_B1)
        PLD_V_DP_K = PLD_Calc_A.Math.divide(PLD_Calc_B)
       
        Dim OutputTname_format() As String
        Dim TestNameInput As String
        Dim TESTLINT_MAX As Long: TESTLINT_MAX = TheExec.Flow.TestLimitIndex
        For i = 0 To PLD_Calc_A.Pins.Count - 1
            PinName = UCase(PLD_Calc_A.Pins.item(i))
            TheExec.Flow.TestLimitIndex = TESTLINT_MAX
            TestNameInput = Report_TName_From_Instance(CalcV, PinName, vbNullString, CInt(j))
            TheExec.Flow.TestLimit resultVal:=PLD_V_DP_K.Pins(i), PinName:=PinName, ForceResults:=tlForceFlow, Tname:=TestNameInput
        Next i
        If k = 0 Then
            PLD_V_DP_K_P = PLD_V_DP_K
        Else
            PLD_V_DP_K_N = PLD_V_DP_K
        End If
    Next k
        
        TESTLINT_MAX = TheExec.Flow.TestLimitIndex

        V_Diff = PLD_V_DP_K_N.Math.Subtract(PLD_V_DP_K_P).Abs
        For i = 0 To PLD_Calc_A.Pins.Count - 1
            PinName = UCase(PLD_Calc_A.Pins.item(i))
            TheExec.Flow.TestLimitIndex = TESTLINT_MAX
            TestNameInput = Report_TName_From_Instance(CalcV, PinName, vbNullString, CInt(j))
            TheExec.Flow.TestLimit resultVal:=V_Diff.Pins(i), PinName:=PinName, ForceResults:=tlForceFlow, Tname:=TestNameInput
        Next i



Exit Function
errHandler:
    Call Print_Error_Message(Error_Info, "LIB_HardIP_Calc", "TX_Level_Pingroup")
    If AbortTest Then Exit Function Else Resume Next
End Function

'[20230915][All][Neil] DRAM MRR New feature implement
Public Function MRR_DRAM_TYPE_Calc(argc As Integer, argv() As String) As Long
    'CUS_Str_DigCapData  : StoreDigAll&DSSC_OUT,16:GRP2:GRP2_DIC,16:GRP3:GRP3_DIC,....
    'DRAM_Check_Infor : [MRR Keyword1]_[Dictioanry Name]_[DRAM_Index]|[MRR Keyword2]_[Dictioanry Name]_[DRAM_Index]
    'Ex :MRR8_GRP1-MRR8-ch1-byte0-DQ2-DDR03_512,544,592|MRR5_GRP1-MRR5-ch1-byte0-DQ2-DDR03_0,32,80,112,128,160|MRR6_GRP1-MRR6-ch1-byte0-DQ2-DDR03_1024,1056,1104,1136
    '===== Note======
    'Format : MRR_DRAM_TYPE_Calc(DRAM_Check_Infor,[Register_Size=8])
    'Must Store All Digital capture data (Ex. CUS_Str_DigCapData:StoreDigAll&DSSC_OUT,16:GRP2:GRP2_DIC,16:GRP3:GRP3_DIC,....)
On Error GoTo errHandler
    Dim sDRAM_Check_Infor As String
    Dim sDRAM_Check_Infor_Ary() As String   ' Use for multi MRR data assignment -- 20221227
    Dim SourceBitStrmWf As New DSPWave      ' Get all of DigCap data from DSP
    Dim Check_Bit_width As Long             ' Use for continue parsing bit size per register index -- 20221227
    Dim DigCap_Sample_Size As Long          ' Store DigCap date sample size
    Dim Data_width As Long                  ' Use for Convert MRR Register data(Binary) to Decimal format
    Dim TestNameInput As String
    Dim temp_i As Integer
    Dim i, j As Long
    Dim vsite As Variant
    
    Dim BitForce0_Num As New SiteLong       'Update optione BitForce0_Num for multi-floating bits force 0 for MRR8 --20230626
    BitForce0_Num = 2                       '0:No ignore bit(Default) -- 20230626  [T-BraC: 2 of floating bits]

    'Get DigCap data
    SourceBitStrmWf = GetStoreDataAllType("StoreDigAll")   ' Get All of Digital capture data from Dictionary
    For Each vsite In TheExec.sites
        DigCap_Sample_Size = SourceBitStrmWf.sampleSize
        Exit For
    Next vsite
    'Get Input information from argument
    sDRAM_Check_Infor = argv(0)         ' Use for Store DRAM Register search information -- 20221227
    Check_Bit_width = CLng(argv(1))     ' Use for continue parsing bit size per register index -- 20221227

    sDRAM_Check_Infor_Ary = Split(sDRAM_Check_Infor, "|")
    For temp_i = 0 To UBound(sDRAM_Check_Infor_Ary)
        sDRAM_Check_Infor = sDRAM_Check_Infor_Ary(temp_i)
    
        Dim s_CUSMainAry() As String
        Dim s_DDRIndexAry() As String
        Dim dsp_DDRIndexAry As New DSPWave
        Dim MRR_width As New DSPWave
        Dim MRR_width1 As New DSPWave
        Dim MRR_OutWf As New DSPWave
        Dim MRR_OutWf1 As New DSPWave
        Dim MRR_Store As New SiteDouble
        Dim DicName As String
        
        'New judge rule for new format -- 20230621
        Dim sb_RamType_result_flag As New SiteBoolean   'MRR5
        Dim sb_RamSize_result_flag As New SiteBoolean   'MRR8
        Dim sb_RamNode_result_flag As New SiteBoolean   'MRR6
        
        '===== Use for Single/Dual Rank check ===== ' -- 20230109
        Dim DSPWF_RANK1_MRR5 As New DSPWave
        Dim DSPWF_RANK1_MRR6 As New DSPWave
        Dim DSPWF_RANK1_MRR8 As New DSPWave
        Dim SD_RANK0_MRR5 As New SiteDouble
        Dim SD_RANK0_MRR6 As New SiteDouble
        Dim SD_RANK0_MRR8 As New SiteDouble
        
        Set DSPWF_RANK1_MRR5 = Nothing
        Set DSPWF_RANK1_MRR6 = Nothing
        Set DSPWF_RANK1_MRR8 = Nothing
        Set SD_RANK0_MRR5 = Nothing
        Set SD_RANK0_MRR6 = Nothing
        Set SD_RANK0_MRR8 = Nothing
        
        Dim sMRR5_DictName As String
        Dim sMRR6_DictName As String
        Dim sMRR8_DictName As String
        
        Set dsp_DDRIndexAry = Nothing
        Set MRR_width = Nothing
        Set MRR_width1 = Nothing
        Set MRR_OutWf = Nothing
        Set MRR_OutWf1 = Nothing
        Set MRR_Store = Nothing
        
        Dim Split_bit_index As Long
        
        s_CUSMainAry = Split(sDRAM_Check_Infor, "_")
        
        If InStr(UCase(Split(sDRAM_Check_Infor, "_")(0)), "RANK1") = 0 Then
            If UBound(s_CUSMainAry) > 0 Then
                'Ex :ALG::MRR_DRAM_TYPE_Calc(MRR8_GRP1-MRR8-ch1-byte0-DQ2-DDR03_0,8)
                ReDim s_DDRIndexAry(UBound(s_CUSMainAry) - 1)
                For j = 1 To UBound(s_CUSMainAry)
                    s_DDRIndexAry(j - 1) = s_CUSMainAry(j)
                Next j
                dsp_DDRIndexAry.CreateConstant 1, UBound(s_DDRIndexAry) + 1, DspLong
                
                For i = 0 To UBound(s_DDRIndexAry)
                    dsp_DDRIndexAry.Element(i) = CLng(s_DDRIndexAry(i))
                Next i
            Else
                'Ex :ALG::MRR_DRAM_TYPE_Calc(MRR8_GRP1-MRR8-ch1-byte0-DQ2-DDR03,8)
                Split_bit_index = 0  ' Apply for no define select data index
                dsp_DDRIndexAry.CreateConstant 1, (DigCap_Sample_Size / Check_Bit_width), DspLong
                For i = 0 To CLng(DigCap_Sample_Size / Check_Bit_width) - 1
                    dsp_DDRIndexAry.Element(i) = Split_bit_index
                    Split_bit_index = Split_bit_index + Check_Bit_width
                Next i
            End If

            'MRR_width.CreateConstant 8, DigCap_Sample_Size / 16 / (UBound(sDRAM_Check_Infor_Ary) + 1)
            'MRR_width1.CreateConstant 16, DigCap_Sample_Size / 32 / (UBound(sDRAM_Check_Infor_Ary) + 1)
            
            'Continue search data lengh(Check_Bit_width) * 2 = max search length 1 times
            'Data_width : Use for convert MRR register value to Decimal
            'Ex. Check_Bit_width = 16bits, Check_Bit_width*2 = 32bits per search lenght 1 times
            
            'If Check_Bit_width <> 8 Then
            If UBound(s_CUSMainAry) > 0 Then
                '==== Setup MRR check width for HDC rule =====
                If UCase(argv(0)) Like "*RANK*" Then
                    'Bypass redim array size for RANK parameter
                    MRR_width.CreateConstant (Check_Bit_width / 2), DigCap_Sample_Size / Check_Bit_width / UBound(sDRAM_Check_Infor_Ary)
                    MRR_width1.CreateConstant Check_Bit_width, (DigCap_Sample_Size / (Check_Bit_width * 2)) / UBound(sDRAM_Check_Infor_Ary)
                Else
                    MRR_width.CreateConstant 8, DigCap_Sample_Size / 8 / (UBound(sDRAM_Check_Infor_Ary) + 1)
                    MRR_width1.CreateConstant Check_Bit_width, UBound(s_DDRIndexAry) + 1 ' DigCap_Sample_Size / (Check_Bit_width) / (UBound(sDRAM_Check_Infor_Ary) + 1)
                End If
                
                If Split(sDRAM_Check_Infor, "_")(0) = "MRR5" Or Split(sDRAM_Check_Infor, "_")(0) = "MRR6" Then
                    rundsp.Split_Dspwave_Partial_Index_New SourceBitStrmWf, MRR_width1, MRR_OutWf1, dsp_DDRIndexAry, 0  'Update optione BitForce0_Num for multi-floating bits force 0; 0:No ignore(Default) -- 20230626
                End If
                    
                If Split(sDRAM_Check_Infor, "_")(0) = "MRR8" Then
                    rundsp.Split_Dspwave_Partial_Index_New SourceBitStrmWf, MRR_width1, MRR_OutWf1, dsp_DDRIndexAry, BitForce0_Num  'Update optione BitForce0_Num for multi-floating bits force 0; 0:No ignore(Default) -- 20230626
                End If
            Else
                '==== Setup MRR check width for CUP rule =====
                If UCase(argv(0)) Like "*RANK*" Then
                    'Bypass redim array size for RANK parameter
                    MRR_width.CreateConstant Check_Bit_width, DigCap_Sample_Size / Check_Bit_width / UBound(sDRAM_Check_Infor_Ary)
                    MRR_width1.CreateConstant Check_Bit_width, DigCap_Sample_Size / Check_Bit_width / UBound(sDRAM_Check_Infor_Ary)
                Else
                    If Split_bit_index = CLng(DigCap_Sample_Size) Then
                        '=== Processing All of Segments data ===
                        MRR_width.CreateConstant Check_Bit_width, DigCap_Sample_Size / Check_Bit_width / (UBound(sDRAM_Check_Infor_Ary) + 1)
                        MRR_width1.CreateConstant Check_Bit_width, DigCap_Sample_Size / Check_Bit_width / (UBound(sDRAM_Check_Infor_Ary) + 1)
                    Else
                        '=== Processing Partial of Segments data ===
                        MRR_width.CreateConstant Check_Bit_width, dsp_DDRIndexAry.sampleSize
                        MRR_width1.CreateConstant Check_Bit_width, dsp_DDRIndexAry.sampleSize
                    End If
                End If
                
                'True: Force specified bit as "0" to ignore judge --20221227
                If Split(sDRAM_Check_Infor, "_")(0) = "MRR5" Or Split(sDRAM_Check_Infor, "_")(0) = "MRR6" Then
                    rundsp.Split_Dspwave_Partial_Index_New SourceBitStrmWf, MRR_width1, MRR_OutWf1, dsp_DDRIndexAry, 0  'Update optione BitForce0_Num for multi-floating bits force 0; 0:No ignore(Default) -- 20230626
                End If
                
                If Split(sDRAM_Check_Infor, "_")(0) = "MRR8" Then
                    rundsp.Split_Dspwave_Partial_Index_New SourceBitStrmWf, MRR_width1, MRR_OutWf1, dsp_DDRIndexAry, BitForce0_Num  'Update optione BitForce0_Num for multi-floating bits force 0; 0:No ignore(Default) -- 20230626
                End If
            End If
    
            rundsp.Split_Dspwave_MSB1st MRR_OutWf1, MRR_width, MRR_OutWf    'Convert data via MSB first for DigCap data
            
            If Split(sDRAM_Check_Infor, "_")(0) = "MRR5" Then
                For Each vsite In TheExec.sites
                    TheExec.Datalog.WriteComment "Site:" & vsite & ", Value =" & MRR_OutWf.Element(0)
                    sb_RamType_result_flag = MRR_OutWf.CalcMaximumValue - MRR_OutWf.CalcMinimumValue
                    MRR_Store = MRR_OutWf.Element(0)
                Next vsite
                
                DicName = s_CUSMainAry(0)         'New format from C651 -- 20230621
                sMRR5_DictName = s_CUSMainAry(0)  'Store Keyword for RANK -- 20230201
                Call StoreDataAllType(DicName & "_para", MRR_Store)
                
                MRR5_Get_Result_Flag = True  'Get MRR information flag -- 20230626
                
                '===== Use for Single/Dual Rank check ===== ' -- 20230109
                If glb_TestInstance Like "*MRRRANK0*" Then Call StoreDataAllType(DicName & "-RANK0", MRR_Store)
                If glb_TestInstance Like "*MRRRANK1*" Then Call StoreDataAllType(DicName & "-RANK1", MRR_OutWf)
                
                If sb_RamType_result_flag.All(False) Then
                    TheExec.Datalog.WriteComment "All sites, All RAM type check are the same."
                Else
                    For Each vsite In TheExec.sites
                        If sb_RamType_result_flag(vsite) = False Then
                            TheExec.Datalog.WriteComment "Site: " & vsite & " All RAM type check are the same."
                        Else
                            TheExec.Datalog.WriteComment "Site: " & vsite & " All RAM type check are not the same."
                        End If
                    Next vsite
                End If
                TestNameInput = Report_TName_From_Instance(CalcC, "Type", ForceResult:=tlForceFlow)
                TheExec.Flow.TestLimit sb_RamType_result_flag.LogicalNot, -1, -1, Tname:=TestNameInput, ForceResults:=tlForceFlow
            '===============================================added for MRR6===============================================
            ElseIf Split(sDRAM_Check_Infor, "_")(0) = "MRR6" Then
                For Each vsite In TheExec.sites
                    TheExec.Datalog.WriteComment "Site:" & vsite & ", Value =" & MRR_OutWf.Element(0)
                    sb_RamNode_result_flag = MRR_OutWf.CalcMaximumValue - MRR_OutWf.CalcMinimumValue
                    MRR_Store = MRR_OutWf.Element(0)
                Next vsite
                
                DicName = s_CUSMainAry(0)           'New format from C651 -- 20230621
                sMRR6_DictName = s_CUSMainAry(0)    'Store Keyword for RANK -- 20230201
                Call StoreDataAllType(DicName & "_para", MRR_Store)
                
                MRR6_Get_Result_Flag = True  'Get MRR information flag -- 20230626
                
                '===== Use for Single/Dual Rank check ===== ' -- 20230109
                If glb_TestInstance Like "*MRRRANK0*" Then Call StoreDataAllType(DicName & "-RANK0", MRR_Store)
                If glb_TestInstance Like "*MRRRANK1*" Then Call StoreDataAllType(DicName & "-RANK1", MRR_OutWf)

                If sb_RamNode_result_flag.All(False) Then
                    TheExec.Datalog.WriteComment "All sites, All RAM Node check are the same."
                Else
                    For Each vsite In TheExec.sites
                        If sb_RamNode_result_flag(vsite) = False Then
                            TheExec.Datalog.WriteComment "Site: " & vsite & " All RAM node check are the same."
                        Else
                            TheExec.Datalog.WriteComment "Site: " & vsite & " All RAM node check are not the same."
                        End If
                    Next vsite
                End If
                
                'TestNameInput = Report_TName_From_Instance(CalcC, "", "MRR6", 0, ForceResult:=tlForceFlow)
                'theexec.Flow.TestLimit MRR_Store, , , Tname:=TestNameInput, ForceResults:=tlForceFlow
                
                TestNameInput = Report_TName_From_Instance(CalcC, "Node", ForceResult:=tlForceFlow)
                TheExec.Flow.TestLimit sb_RamNode_result_flag.LogicalNot, -1, -1, Tname:=TestNameInput, ForceResults:=tlForceFlow
                '===============================================added for MRR6===============================================
            ElseIf Split(sDRAM_Check_Infor, "_")(0) = "MRR8" Then
                For Each vsite In TheExec.sites
                    TheExec.Datalog.WriteComment "Site:" & vsite & ", Value =" & MRR_OutWf.Element(0)
                    sb_RamSize_result_flag = MRR_OutWf.CalcMaximumValue - MRR_OutWf.CalcMinimumValue
                    MRR_Store = MRR_OutWf.Element(0)
                Next vsite
                    
                DicName = s_CUSMainAry(0)           'New format from C651 -- 20230621
                sMRR8_DictName = s_CUSMainAry(0)    'Store Keyword for RANK -- 20230201
                Call StoreDataAllType(DicName & "_para", MRR_Store)
                
                MRR8_Get_Result_Flag = True  'Get MRR information flag -- 20230626
                
                '===== Use for Single/Dual Rank check ===== ' -- 20230109
                If glb_TestInstance Like "*MRRRANK0*" Then Call StoreDataAllType(DicName & "-RANK0", MRR_Store)
                If glb_TestInstance Like "*MRRRANK1*" Then Call StoreDataAllType(DicName & "-RANK1", MRR_OutWf)
                
                If sb_RamSize_result_flag.All(False) Then
                    TheExec.Datalog.WriteComment "All sites, All RAM Size check are the same."
                Else
                    For Each vsite In TheExec.sites
                        If sb_RamSize_result_flag(vsite) = False Then
                            TheExec.Datalog.WriteComment "Site: " & vsite & " All RAM size check are the same."
                        Else
                            TheExec.Datalog.WriteComment "Site: " & vsite & " All RAM size check are not the same."
                        End If
                    Next vsite
                End If
                
                TestNameInput = Report_TName_From_Instance(CalcC, "Size", ForceResult:=tlForceFlow)
                TheExec.Flow.TestLimit sb_RamSize_result_flag.LogicalNot, -1, -1, Tname:=TestNameInput, ForceResults:=tlForceFlow
            Else
                TheExec.Datalog.WriteComment "<Error> No DRAM Case match, Please check Argument setup !!"
            End If
        Else
            '===== Use for Single/Dual Rank check ===== ' -- 20230109
            If Split(sDRAM_Check_Infor, "_")(0) = "RANK1" Then
                '===== Judge Single/Dual Rank Flag =====
                Dim Match_Count As Long: Match_Count = 0
                Dim CHK_RANK_Flag As Long: CHK_RANK_Flag = 999          '999:Initial/1:Single Rank(Not match)/2:Dual Rank(Parital match)/3:Dual Rank(All-match)
                Dim RANK_MRR5_Flag As New SiteLong                      '999:Initial /1:no match /2:parital match /3:all match
                Dim RANK_MRR6_Flag As New SiteLong                      '999:Initial /1:no match /2:parital match /3:all match
                Dim RANK_MRR8_Flag As New SiteLong                      '999:Initial /1:no match /2:parital match /3:all match
                Dim SL_RANK_Flag As New SiteLong: SL_RANK_Flag = 999    '999:Initial/1:Single Rank(Not match)/2:Dual Rank(Parital match)/3:Dual Rank(All-match)
                Set RANK_MRR5_Flag = Nothing
                Set RANK_MRR6_Flag = Nothing
                Set RANK_MRR8_Flag = Nothing
                
                SD_RANK0_MRR5 = GetStoreDataAllType(sMRR5_DictName & "-RANK0")
                SD_RANK0_MRR6 = GetStoreDataAllType(sMRR6_DictName & "-RANK0")
                SD_RANK0_MRR8 = GetStoreDataAllType(sMRR8_DictName & "-RANK0")
                DSPWF_RANK1_MRR5 = GetStoreDataAllType(sMRR5_DictName & "-RANK1")
                DSPWF_RANK1_MRR6 = GetStoreDataAllType(sMRR6_DictName & "-RANK1")
                DSPWF_RANK1_MRR8 = GetStoreDataAllType(sMRR8_DictName & "-RANK1")
                
                For Each vsite In TheExec.sites
                    Match_Count = 0 'Counter reset
                    '===== RANK CHECK for MRR5 =====
                    For i = 0 To DSPWF_RANK1_MRR5.sampleSize - 1
                        If SD_RANK0_MRR5 = DSPWF_RANK1_MRR5.Element(i) Then
                            Match_Count = Match_Count + 1
                        End If
                    Next i
                    If Match_Count = 0 Then
                        RANK_MRR5_Flag(vsite) = 1
                    ElseIf Match_Count = DSPWF_RANK1_MRR5.sampleSize Then
                        RANK_MRR5_Flag(vsite) = 3
                    Else
                        RANK_MRR5_Flag(vsite) = 2
                    End If
                    
                    Match_Count = 0 'Counter reset
                    '===== RANK CHECK for MRR6 =====
                    For i = 0 To DSPWF_RANK1_MRR6.sampleSize - 1
                        If SD_RANK0_MRR6 = DSPWF_RANK1_MRR6.Element(i) Then
                            Match_Count = Match_Count + 1
                        End If
                    Next i
                    If Match_Count = 0 Then
                        RANK_MRR6_Flag(vsite) = 1
                    ElseIf Match_Count = DSPWF_RANK1_MRR6.sampleSize Then
                        RANK_MRR6_Flag(vsite) = 3
                    Else
                        RANK_MRR6_Flag(vsite) = 2
                    End If
                    
                    Match_Count = 0 'Counter reset
                    '===== RANK CHECK for MRR8 =====
                    For i = 0 To DSPWF_RANK1_MRR8.sampleSize - 1
                        If SD_RANK0_MRR8 = DSPWF_RANK1_MRR8.Element(i) Then
                            Match_Count = Match_Count + 1
                        End If
                    Next i
                    If Match_Count = 0 Then
                        RANK_MRR8_Flag(vsite) = 1
                    ElseIf Match_Count = DSPWF_RANK1_MRR8.sampleSize Then
                        RANK_MRR8_Flag(vsite) = 3
                    Else
                        RANK_MRR8_Flag(vsite) = 2
                    End If
                    
                    '===== Judge Rank type =====
                    'RANK0(MRR5/MRR6/MRR8) Data compare with RANK1(MRR5/MRR6/MRR8)
                    '-- MRR5 any channel match & MRR6 any channel match MRR6 & MRR8 any channel match MRR8       => DUAL RANK
                    '-- MRR5 all_channel match MRR5 & MRR6 all_channel match MRR6 & MRR8 all_channel match MRR8  => DUAL RANK
                    '-- MRR5 no_channel match MRR5 or MRR6 no_channel match MRR6 or MRR8 no_channel match MRR8   => SINGLE RANK
                    
                    'Flag description: 999:Initial /1:no match /2:parital match /3:all match
                    If RANK_MRR5_Flag = 1 Or RANK_MRR6_Flag = 1 Or RANK_MRR8_Flag = 1 Then
                        TheExec.Datalog.WriteComment " ------- Enable Single_RANK for DRAM tests-------"
                        'theexec.Flow.EnableWord("DUAL_RANK") = False
                        TheExec.Datalog.WriteComment "Site: " & vsite & " DRAM Size = SINGLE RANK"
                        CHK_RANK_Flag = 1       '1:Single Rank(Not match)/2:Dual Rank(Parital match)/3:Dual Rank(All match)
                    Else
                        If RANK_MRR5_Flag = 3 And RANK_MRR6_Flag = 3 And RANK_MRR8_Flag = 3 Then
                            TheExec.Datalog.WriteComment "Site: " & vsite & " DRAM Size = DUAL RANK-All-Match"
                            'theexec.Flow.EnableWord("DUAL_RANK") = True
                            'theexec.Datalog.WriteComment " ------- Enable DUAL_RANK for DRAM tests-------"
                            CHK_RANK_Flag = 3       '1:Single Rank(Not match)/2:Dual Rank(Parital match)/3:Dual Rank(All match)
                        Else
                            TheExec.Datalog.WriteComment "Site: " & vsite & " DRAM Size = DUAL RANK-Parital-Match"
                            'theexec.Flow.EnableWord("DUAL_RANK") = True
                            'theexec.Datalog.WriteComment " ------- Enable DUAL_RANK for DRAM tests-------"
                            CHK_RANK_Flag = 2       '1:Single Rank(Not match)/2:Dual Rank(Parital match)/3:Dual Rank(All match)
                        End If
                    End If
                    If CHK_RANK_Flag = 1 Then
                        MRR_Store(vsite) = 0    'Single Rank
                    Else
                        MRR_Store(vsite) = 1    'Dual Rank
                    End If
                Next vsite
                'Return Rank status to dictionary
                DicName = s_CUSMainAry(0)                           'New format from C651 -- 20230621
                Call StoreDataAllType(DicName & "_para", MRR_Store)    '0:Single Rank/1:Dual Rank
            End If
        End If
    Next temp_i
    
    '=== Printing DRAM Information ==  -- 20230621
    Dim MRR5_Data As New SiteDouble
    Dim MRR6_Data As New SiteDouble
    Dim MRR8_Data As New SiteDouble
    'Check DRAM MRR Result is match with DRAM EnableWord Setup or not -- 20230711
    Dim DRAM_Result_Check As New SiteBoolean   'True:Check Pass False:Check Fail
    Dim Print_indx As Variant
    Dim TYPE_COUNT As Long: TYPE_COUNT = 0
    If InStr(UCase(argv(0)), "MRR5") Then MRR5_Data = GetStoreDataAllType("MRR5_para"): TYPE_COUNT = TYPE_COUNT + 1
    If InStr(UCase(argv(0)), "MRR8") Then MRR8_Data = GetStoreDataAllType("MRR8_para"): TYPE_COUNT = TYPE_COUNT + 1
    If InStr(UCase(argv(0)), "MRR6") Then MRR6_Data = GetStoreDataAllType("MRR6_para"): TYPE_COUNT = TYPE_COUNT + 1
    
    For Each vsite In TheExec.sites
     '   TheExec.Datalog.WriteComment "*** Site:" & vsite & ", Prodduction Set [" & BdfDataBase.DicDramHipMap(vsite)("FlowEnableWord") & "] Enable Word !!! ***"
        If MRR5_Get_Result_Flag = True And MRR8_Get_Result_Flag = True Then
            If sb_RamType_result_flag = True Or sb_RamSize_result_flag = True Or sb_RamNode_result_flag = True Then
                Call Print_Error_Message(Warning_Info, "LIB_HardIP", "Site: " & vsite & " Can't not judge DRAM Information!!")
                DRAM_Result_Check = False
            Else
                If TYPE_COUNT = 2 Then
                    'If BdfDataBase.DicDramMap.Exists(MRR5_Data & "_" & MRR8_Data) = True Then
                    If CLng(BdfDataBase.DicDramHipMap(vsite)("mrr5")) = MRR5_Data And CLng(BdfDataBase.DicDramHipMap(vsite)("mrr8")) = MRR8_Data Then
                        For Each Print_indx In BdfDataBase.DicDramHipMap(vsite)("Descriptor")
                            TheExec.Datalog.WriteComment "Site: " & vsite & " " & Print_indx & " = " & BdfDataBase.DicDramHipMap(vsite)("Descriptor")(Print_indx)
                        Next
'                        TheExec.Datalog.WriteComment "Site: " & VSite & " DRAM Vendor = " & BdfDataBase.DicDramHipMap(VSite)("vendor")
'                        TheExec.Datalog.WriteComment "Site: " & VSite & " DRAM TotalSize = " & BdfDataBase.DicDramHipMap(VSite)("totalsize")
'                        TheExec.Datalog.WriteComment "Site: " & VSite & " DRAM DieSize = " & BdfDataBase.DicDramHipMap(VSite)("diesize")
                        DRAM_Result_Check = True    'True:Check Pass False:Check Fail
                    Else
                        Call Print_Error_Message(Error_Info, "LIB_HardIP", "MRR_DRAM_TYPE_Calc", "Site[" & vsite & "] No Match DRAM information from Table!!")
                        If gl_Disable_HIP_debug_log = False Then
                            TheExec.Datalog.WriteComment "Site: " & vsite & " MRR5 [Table Define]= " & BdfDataBase.DicDramHipMap(vsite)("mrr5") & " ,[HIP Read]=" & MRR5_Data
                            TheExec.Datalog.WriteComment "Site: " & vsite & " MRR8 [Table Define]= " & BdfDataBase.DicDramHipMap(vsite)("mrr8") & " ,[HIP Read]=" & MRR8_Data
                        End If
                        DRAM_Result_Check = False    'True:Check Pass False:Check Fail
                    End If
                    MRR5_Get_Result_Flag = False 'Reset flag
                    MRR8_Get_Result_Flag = False 'Reset flag
                ElseIf TYPE_COUNT = 3 And MRR6_Get_Result_Flag = True Then
                    'If BdfDataBase.DicDramMap.Exists(MRR5_Data & "_" & MRR8_Data & "_" & MRR6_Data) = True Then
                    If CLng(BdfDataBase.DicDramHipMap(vsite)("mrr5")) = MRR5_Data And CLng(BdfDataBase.DicDramHipMap(vsite)("mrr8")) = MRR8_Data And CLng(BdfDataBase.DicDramHipMap(vsite)("mrr6")) = MRR6_Data Then
                        For Each Print_indx In BdfDataBase.DicDramHipMap(vsite)("Descriptor")
                            TheExec.Datalog.WriteComment "Site: " & vsite & " " & Print_indx & " = " & BdfDataBase.DicDramHipMap(vsite)("Descriptor")(Print_indx)
                        Next
'                        TheExec.Datalog.WriteComment "Site: " & VSite & " DRAM Vendor = " & BdfDataBase.DicDramHipMap(VSite)("vendor")
'                        TheExec.Datalog.WriteComment "Site: " & VSite & " DRAM TotalSize = " & BdfDataBase.DicDramHipMap(VSite)("totalsize")
'                        TheExec.Datalog.WriteComment "Site: " & VSite & " DRAM DieSize = " & BdfDataBase.DicDramHipMap(VSite)("diesize")
'                        TheExec.Datalog.WriteComment "Site: " & VSite & " DRAM Rev = " & BdfDataBase.DicDramHipMap(VSite)("rev")
                        DRAM_Result_Check = True    'True:Check Pass False:Check Fail
                    Else
                        Call Print_Error_Message(Error_Info, "LIB_HardIP", "MRR_DRAM_TYPE_Calc", "Site[" & vsite & "] No Match DRAM information from Table!!")
                        If gl_Disable_HIP_debug_log = False Then
                            TheExec.Datalog.WriteComment "Site: " & vsite & " MRR5 [Table Define]= " & BdfDataBase.DicDramHipMap(vsite)("mrr5") & " ,[HIP Read]=" & MRR5_Data
                            TheExec.Datalog.WriteComment "Site: " & vsite & " MRR8 [Table Define]= " & BdfDataBase.DicDramHipMap(vsite)("mrr8") & " ,[HIP Read]=" & MRR8_Data
                            TheExec.Datalog.WriteComment "Site: " & vsite & " MRR6 [Table Define]= " & BdfDataBase.DicDramHipMap(vsite)("mrr6") & " ,[HIP Read]=" & MRR6_Data
                        End If
                        DRAM_Result_Check = False    'True:Check Pass False:Check Fail
                    End If
                    MRR5_Get_Result_Flag = False 'Reset flag
                    MRR8_Get_Result_Flag = False 'Reset flag
                    MRR6_Get_Result_Flag = False 'Reset flag
                End If
            End If
        End If
    Next vsite
    
    'Check DRAM MRR Result is match with DRAM EnableWord Setup or not -- 20230711
    TestNameInput = Report_TName_From_Instance(CalcC, "MRR-Check", ForceResult:=tlForceFlow)
    TheExec.Flow.TestLimit DRAM_Result_Check, -1, -1, Tname:=TestNameInput, ForceResults:=tlForceFlow
    
    Exit Function
errHandler:
    Call Print_Error_Message(Error_Info, "LIB_HardIP", "MRR_DRAM_TYPE_Calc")
    If AbortTest Then Exit Function Else Resume Next
End Function