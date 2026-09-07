Attribute VB_Name = "VBT_LIB_HardIP_Customize"
#Const isUFP = True
Option Explicit
Enum VIR_EntryPoint ''VBT_CUSTOMIZE_HardIP
    VIR_MI_AFTER_MEASUREMENT = 0
    VIR_MI_AFTER_TESTLIMIT = 1
End Enum

Public AMP_EYE_VT_CZ_Flag As Boolean

Public Function CUS_VIR_MainProgram_MeasV_CalR(TestPinArrayIV() As String, TestSeqNum As Integer, _
                            CUS_CalR_Seq() As String, ForceI() As String, MeasV As PinListData, CUS_VDD As Double) As Long ''VBT_CUSTOMIZE_HardIP
On Error GoTo errHandler 'Add ErrHandler 2023/05/29
    
    Dim ForceI_Ary() As String
    Dim ForceISeqIndexPerSeq As Long ''Split by "+"
                                                                                                                                                                                                                   
    Dim R As New PinListData
    Dim R_AK() As Double
                                                                                                                                                                                                                   
    Dim v As Double ''voltage
    Dim i As Double ''current
    Dim p As Long
    Dim PinAry() As String
    Dim FirstSite As Integer
    Dim site As Variant 'Carter, 20240304
    If (UBound(ForceI) <> 0) Then
        ForceISeqIndexPerSeq = TestSeqNum
    Else
        ForceISeqIndexPerSeq = 0
    End If
                                                                                                                                                                                                                   
    ForceI_Ary = Split(ForceI(ForceISeqIndexPerSeq), ",")
                                                                                                                                                                                                                   
    PinAry = Split(TestPinArrayIV(TestSeqNum), ",")
                                                                                                                                                                                                                   
    FirstSite = 0
    
    If (UBound(ForceI_Ary) = 0) Then
                                                                                                                                                                                                                   
        For Each site In theexec.sites
                                                                                                                                                                                                                   
            For p = 0 To MeasV.Pins.Count - 1
                                                                                                                                                                                                                   
                If (FirstSite = 0) Then
                    R.AddPin (MeasV.Pins(p))
                End If
                                                                                                                                                                                                                   
                If (UCase(CUS_CalR_Seq(TestSeqNum)) Like "*RVOH*") Then
                                                                                                                                                                                                                   
                    R.Pins(p).value(site) = CUS_VDD - MeasV.Pins(p).value(site)
                    R.Pins(p).value(site) = R.Pins(p).divide(CDbl(ForceI_Ary(0)))
                                                                                                                                                                                                                   
                ElseIf (UCase(CUS_CalR_Seq(TestSeqNum)) Like "*RVOL*") Then
                                                                                                                                                                                                                   
                    R.Pins(p).value(site) = MeasV.Pins(p).value(site)
                    R.Pins(p).value(site) = R.Pins(p).divide(CDbl(ForceI_Ary(0)))
                Else 'Do nothing '20230601
                End If
                                                                                                                                                                                                                   
                'R_AK = TheHdw.PPMU.ReadRakValuesByPinnames(MeasV.Pins(p), site)  ''Get instrament impedance
                                                                                                                                                                                                                   
''                If InStr(UCase(TheExec.CurrentChanMap), "FT") <> 0 Then                          ''Get DIB impedance
''                    R_AK(0) = R_AK(0) + FT_Card_RAK.Pins(MeasV.Pins(p)).Value(Site)
''                Else
''                    R_AK(0) = R_AK(0) + CP_Card_RAK.Pins(MeasV.Pins(p)).Value(Site)
''                End If
                R_AK(0) = CurrentJob_Card_RAK.Pins(MeasV.Pins(p)).value(site)
                                                                                                                                                                                                                   
                R.Pins(p).value(site) = R.Pins(p).value(site) - R_AK(0)
                                                                                                                                                                                                                   
            Next p
            FirstSite = FirstSite + 1
        Next site
                                                                                                                                                                                                                   
        theexec.Flow.TestLimit resultVal:=R, unit:=unitCustom, Tname:="Calculate_" + CUS_CalR_Seq(TestSeqNum), customUnit:="ohm", ForceResults:=tlForceNone 'Un-Used
                                                                                                                                                                                                                   
    End If

Exit Function 'Add ErrHandler 2023/05/29
errHandler: 'Add ErrHandler 2023/05/29
    Call Print_Error_Message(Error_Info, "VBT_LIB_HardIP_Customize", "CUS_VIR_MainProgram_MeasV_CalR") 'Add ErrHandler 2023/05/29
    If AbortTest Then Exit Function Else Resume Next 'Add ErrHandler 2023/05/29
End Function

Public Function AnalyzeCusStrToCalcR(CUS_Str_MainProgram As String, TestSeqNum As Integer, ForceSequenceArray() As String, MeasCurr As PinListData, RTN_Imped As PinListData) As Long
On Error GoTo errHandler 'Add ErrHandler 2023/05/29
    
    ''ex.  CUS_Str_MainProgram = "CalR;_VDDQL_DDR0_VAR_H;RVOH,RVOL,RVOH,RVOL"
    ''========================================================================================
    Dim CUS_CalR_VDD As Double
    Dim CUS_CalR_Seq_Ary() As String
    Dim CUS_CalR_Arg_Ary() As String
    Dim Ary_str(0) As String
    
    CUS_CalR_Arg_Ary = Split(CUS_Str_MainProgram, ";")
    Ary_str(0) = CUS_CalR_Arg_Ary(1)
    Call HIP_Evaluate_ForceVal(Ary_str)
    CUS_CalR_VDD = CDbl(Ary_str(0))
    CUS_CalR_Seq_Ary = Split(CUS_CalR_Arg_Ary(2), ",")
    ''========================================================================================
    Dim p As Long
    Dim b_FirstTime As Boolean
    b_FirstTime = True
    Dim ForceVolt As Double
    Dim site As Variant 'Carter, 20240304
    If UBound(ForceSequenceArray) = 0 Then
        ForceVolt = ForceSequenceArray(0)
    Else
        ForceVolt = ForceSequenceArray(TestSeqNum)
    End If
    
    For Each site In theexec.sites
        For p = 0 To MeasCurr.Pins.Count - 1
                                                                                                                                                                                                               
            If b_FirstTime = True Then
                RTN_Imped.AddPin (MeasCurr.Pins(p))
            End If
                                                                                                                                                                                                               
            If (UCase(CUS_CalR_Seq_Ary(TestSeqNum)) Like "*RVOH*") Then
                                                                                                                                                                                                               
                RTN_Imped.Pins(p).value(site) = CUS_CalR_VDD - ForceVolt
                RTN_Imped.Pins(p).value(site) = RTN_Imped.Pins(p).divide(MeasCurr.Pins(p)).Multiply(-1)
                                                                                                                                                                                                               
            ElseIf (UCase(CUS_CalR_Seq_Ary(TestSeqNum)) Like "*RVOL*") Then
                                                                                                                                                                                                               
                RTN_Imped.Pins(p).value(site) = ForceVolt
                RTN_Imped.Pins(p).value(site) = RTN_Imped.Pins(p).divide(MeasCurr.Pins(p))
            Else 'Do nothing '20230601
            End If
        Next p
        b_FirstTime = False
    Next site
Exit Function 'Add ErrHandler 2023/05/29
errHandler: 'Add ErrHandler 2023/05/29
    Call Print_Error_Message(Error_Info, "VBT_LIB_HardIP_Customize", "AnalyzeCusStrToCalcR") 'Add ErrHandler 2023/05/29
    If AbortTest Then Exit Function Else Resume Next 'Add ErrHandler 2023/05/29
End Function

Public Function Cust_Sweep_V()
On Error GoTo errHandler 'Add ErrHandler 2023/05/29
    Dim StepIndex_Val As Long
    Dim StartVolt_Val_1 As Double, StepVolt_Val_1 As Double, ForceVolt_Val_1 As Double
    Dim StartVolt_Val_2 As Double, StepVolt_Val_2 As Double, ForceVolt_Val_2 As Double
    Dim Force_Pins_1 As String, Force_Pins_2 As String
    StartVolt_Val_1 = 0.99
    StartVolt_Val_2 = 0.82
    StepVolt_Val_1 = 0.02
    StepVolt_Val_2 = 0.02
    Force_Pins_1 = "VDDIO11_RET_DDR0,VDDIO11_RET_DDR1"
    Force_Pins_2 = "VDD_DCS_DDR0,VDD_DCS_DDR1"
    
    StepIndex_Val = CDbl(val(theexec.Flow.var("SrcCodeIndx").value))
    ForceVolt_Val_1 = StartVolt_Val_1 + StepIndex_Val * StepVolt_Val_1
    ForceVolt_Val_2 = StartVolt_Val_2 + StepIndex_Val * StepVolt_Val_2

    TheHdw.DCVS.Pins(Force_Pins_1).Voltage.value = ForceVolt_Val_1
    TheHdw.DCVS.Pins(Force_Pins_2).Voltage.value = ForceVolt_Val_2
    
    theexec.Datalog.WriteComment ("Force Pin " & Force_Pins_1 & " Value = " & ForceVolt_Val_1)
    theexec.Datalog.WriteComment ("Force Pin " & Force_Pins_2 & " Value = " & ForceVolt_Val_2)
Exit Function 'Add ErrHandler 2023/05/29
errHandler: 'Add ErrHandler 2023/05/29
    Call Print_Error_Message(Error_Info, "VBT_LIB_HardIP_Customize", "Cust_Sweep_V") 'Add ErrHandler 2023/05/29
    If AbortTest Then Exit Function Else Resume Next 'Add ErrHandler 2023/05/29
End Function
Public Function VOLH_Sweep(CUS_Str_DigSrcData As String, digsrc_assignment As String) As Long
On Error GoTo errHandler 'Add ErrHandler 2023/05/29
    Dim SrcCode_Initial_Dec As Long
    Dim SrcCode_Target_Dec As Long
    Dim SrcCode_Target_Bin() As Long
    ReDim SrcCode_Target_Bin(9) As Long
    Dim SrcCode_Target_Bin_One As String: SrcCode_Target_Bin_One = vbNullString
    Dim SrcCode_Target_Bin_Two As String: SrcCode_Target_Bin_Two = vbNullString
    Dim i As Long
    Dim SplitArray() As String
    Dim ReplaceTarget_1 As String
    Dim ReplaceTarget_2 As String
    
    '' Split with comma
    SplitArray = Split(CUS_Str_DigSrcData, ",")
    ReplaceTarget_1 = SplitArray(1)
    ReplaceTarget_2 = SplitArray(2)
    
    SrcCode_Target_Dec = theexec.Flow.var("SrcCodeIndx").value
    Call Dec2Bin(Abs(SrcCode_Target_Dec), SrcCode_Target_Bin)
    
    For i = 0 To 9
        If i < 3 Then
            SrcCode_Target_Bin_One = SrcCode_Target_Bin(i) & SrcCode_Target_Bin_One
        Else
            SrcCode_Target_Bin_Two = SrcCode_Target_Bin(i) & SrcCode_Target_Bin_Two
        End If
    Next i
    SrcCode_Target_Bin_One = SrcCode_Target_Bin_One & "0"
    SrcCode_Target_Bin_Two = SrcCode_Target_Bin_Two & "0"
    digsrc_assignment = Replace(digsrc_assignment, ReplaceTarget_1, SrcCode_Target_Bin_One)
    digsrc_assignment = Replace(digsrc_assignment, ReplaceTarget_2, SrcCode_Target_Bin_Two)
Exit Function 'Add ErrHandler 2023/05/29
errHandler: 'Add ErrHandler 2023/05/29
    Call Print_Error_Message(Error_Info, "VBT_LIB_HardIP_Customize", "VOLH_Sweep") 'Add ErrHandler 2023/05/29
    If AbortTest Then Exit Function Else Resume Next 'Add ErrHandler 2023/05/29
End Function

Public Function MTR_UVI80_Setup()
On Error GoTo errHandler 'Add ErrHandler 2023/05/29
    ' 20170227 - Set current range only for osprey Metrology 20170227
            TheHdw.DCVI.Pins("mtr_analog_test_p").CurrentRange = 0.002
            TheHdw.DCVI.Pins("mtr_analog_test_p").Current = 0.002

Exit Function 'Add ErrHandler 2023/05/29
errHandler: 'Add ErrHandler 2023/05/29
    Call Print_Error_Message(Error_Info, "VBT_LIB_HardIP_Customize", "MTR_UVI80_Setup") 'Add ErrHandler 2023/05/29
    If AbortTest Then Exit Function Else Resume Next 'Add ErrHandler 2023/05/29
End Function


Public Function CUS_DDR_Emulate_Const_Res_Loading(MeasureValue As PinListData, ForceValByPin() As String, CUS_Str_MainProgram As String, TestSeqNum As Integer, _
    Optional RAK_Flag As Enum_RAK = 0) As Long
On Error GoTo errHandler 'Add ErrHandler 2023/05/29

    Dim R0_Value As New PinListData
    Dim Final_Volt_Value As New PinListData
    Dim Final_Force_Curr_Value As New PinListData
    Dim R1_Value As New PinListData
    Dim Adjust_I_Value As New PinListData
    Dim site As Variant
    Dim pin  As Variant
    Dim Temp_Input() As String

    glb_TestInstance = vbNullString
    glb_TestInstance = UCase(theexec.DataManager.instancename)
    
''''''''''    Dim Pwr_Voltage As Double: Pwr_Voltage = TheHdw.DCVS.Pins("VDDQL_DDR").Voltage.Value
    '   Hardcode for debug
    Dim Pwr_Voltage As Double: Pwr_Voltage = TheHdw.DCVS.Pins("VDDQ0_S1").Voltage.value
    
    Dim Target_Resistance As Double
    Dim Flag_1 As Boolean: Flag_1 = False
    Dim Flag_2 As Boolean: Flag_2 = False
    Dim Initial_Setting_Flag As Boolean: Initial_Setting_Flag = True
    Dim Counter_Meas As Integer: Counter_Meas = 1
    Dim Counter_End As Integer: Counter_End = 50
    Dim ForceValue As Double: ForceValue = Abs(ForceValByPin(0))
    Dim p As Integer
    Dim RakV() As Double
    Dim Ary_str(0) As String
    Dim hiLimit As Double
    Dim loLimit As Double
    Dim Pin_Diff As String
    Dim error_flag As Boolean: error_flag = False
    Dim OutputTname_format() As String
    Dim TestNameInput As String
    
    
''''''''''    Ary_str(0) = "_VDDQL_DDR_VAR_H"
    ' Edied by Dyaln 20190909 , Hardcode for debug
    Ary_str(0) = "VDDQ0_S1_VAR"


    Call HIP_Evaluate_ForceVal(Ary_str)
    
    Temp_Input() = Split(CUS_Str_MainProgram, ":")     ' CUS_Str_MainProgram . Ex: Tname:VOL,VOH,VOL,VOH,48
    Temp_Input() = Split(Temp_Input(1), ",")
    Target_Resistance = Temp_Input(UBound(Temp_Input))
    
    If UCase(Temp_Input(TestSeqNum)) = "VOH" Then
        hiLimit = CDbl(Ary_str(0))
        loLimit = 0.65 * CDbl(Ary_str(0))
    Else
        hiLimit = 0.489 * CDbl(Ary_str(0))
        loLimit = 0
    End If
    
    
    For p = 0 To MeasureValue.Pins.Count - 1
        Flag_1 = False
        pin = MeasureValue.Pins(p).name
        
        For Each site In theexec.sites
            Initial_Setting_Flag = True
            error_flag = False
            
                For Counter_Meas = 1 To Counter_End
                    If Flag_1 = False Then
                        R0_Value.AddPin (pin)
                        Final_Volt_Value.AddPin (pin)
                        Final_Force_Curr_Value.AddPin (pin)
                        R1_Value.AddPin (pin)
                        Adjust_I_Value.AddPin (pin)
                        Flag_1 = True
                    End If
                    
                    If UCase(Temp_Input(TestSeqNum)) = "VOH" Then
                        If Initial_Setting_Flag = True Then
                            R0_Value.Pins(pin) = MeasureValue.Pins(pin).value(site) / ForceValue
                        Else
                            R0_Value.Pins(pin) = MeasureValue.Pins(pin).value(site) / Abs(Adjust_I_Value.Pins(pin).value(site))
                        End If
                    ElseIf UCase(Temp_Input(TestSeqNum)) = "VOL" Then
                        If Initial_Setting_Flag = True Then
                            R0_Value.Pins(pin) = (Pwr_Voltage - MeasureValue.Pins(pin).value(site)) / ForceValue
                        Else
                            R0_Value.Pins(pin) = (Pwr_Voltage - MeasureValue.Pins(pin).value(site)) / Adjust_I_Value.Pins(pin).value(site)
                        End If
                    Else 'Do nothing '20230601
                    End If
                    
                    If ((Abs(R0_Value.Pins(pin).value(site) - Target_Resistance)) / Target_Resistance) < 0.1 Then
                        Final_Volt_Value.Pins(pin).value(site) = MeasureValue.Pins(pin).value(site)
                        
                        If Counter_Meas = 1 Then
                            If UCase(Temp_Input(TestSeqNum)) = "VOH" Then
                                Final_Force_Curr_Value.Pins(pin).value(site) = -1 * ForceValue
                            ElseIf UCase(Temp_Input(TestSeqNum)) = "VOL" Then
                                Final_Force_Curr_Value.Pins(pin).value(site) = ForceValue
                            End If
                        Else
                            Final_Force_Curr_Value.Pins(pin).value(site) = Adjust_I_Value.Pins(pin).value(site)
                        End If
                        Counter_Meas = Counter_End + 1 ' Iteration search done , Exit from Loop
                        
                        ' 20171204 Update latest R1_Value
                        If UCase(Temp_Input(TestSeqNum)) = "VOH" Then
                            R1_Value.Pins(pin).value(site) = (Pwr_Voltage - Final_Volt_Value.Pins(pin).value(site)) / Abs(Final_Force_Curr_Value.Pins(pin).value(site))
                        ElseIf UCase(Temp_Input(TestSeqNum)) = "VOL" Then
                            R1_Value.Pins(pin).value(site) = Final_Volt_Value.Pins(pin).value(site) / Final_Force_Curr_Value.Pins(pin).value(site)
                        Else 'Do nothing '20230601
                        End If
                        
                        If R1_Value.Pins(pin).value(site) = 0 Then
                            theexec.Datalog.WriteComment ("Final " & "Site_" & site & " pin = " & pin & " R0 value = " & Format(R0_Value.Pins(pin).value(site), "0.000") & " R1 value = " & "NA" & _
                                                                            " Meas Volt = " & Format(Final_Volt_Value.Pins(pin).value(site), "0.0000"))
                        Else
                            theexec.Datalog.WriteComment ("Final " & "Site_" & site & " pin = " & pin & " R0 value = " & Format(R0_Value.Pins(pin).value(site), "0.000") & " R1 value = " & Format(R1_Value.Pins(pin).value(site), "0.000") & _
                                                                            " Meas Volt = " & Format(Final_Volt_Value.Pins(pin).value(site), "0.0000"))
                        End If
                    Else
                                                            
                        If UCase(Temp_Input(TestSeqNum)) = "VOH" Then
                            If Initial_Setting_Flag = True Then
                                R1_Value.Pins(pin).value(site) = (Pwr_Voltage - MeasureValue.Pins(pin).value(site)) / ForceValue
                            Else
                                R1_Value.Pins(pin).value(site) = (Pwr_Voltage - MeasureValue.Pins(pin).value(site)) / Abs(Adjust_I_Value.Pins(pin).value(site))
                            End If
                       ElseIf UCase(Temp_Input(TestSeqNum)) = "VOL" Then
                            If Initial_Setting_Flag = True Then
                                R1_Value.Pins(pin).value(site) = MeasureValue.Pins(pin).value(site) / ForceValue
                            Else
                                R1_Value.Pins(pin).value(site) = MeasureValue.Pins(pin).value(site) / Adjust_I_Value.Pins(pin).value(site)
                            End If
                        Else 'Do nothing '20230601
                        End If
                        
                        If (R1_Value.Pins(pin).value(site) + Target_Resistance) = 0 Then
                            Adjust_I_Value.Pins(pin).value(site) = 0
                            theexec.Datalog.WriteComment (" Error : Denominator = 0 !  ")
                            error_flag = True
                        Else
                            Adjust_I_Value.Pins(pin).value(site) = Pwr_Voltage / (R1_Value.Pins(pin).value(site) + Target_Resistance)
                        End If
                        
                        If UCase(Temp_Input(TestSeqNum)) = "VOH" Then
                            Adjust_I_Value.Pins(pin).value(site) = (-1) * Adjust_I_Value.Pins(pin).value(site)
                        End If
                        
                        Initial_Setting_Flag = False    ' False means start to use Adjust I Value for next Iteration search
                        
                        If error_flag = False Then
                            ' Update Force Condition and Measure Voltage
                            TheHdw.Digital.Pins(pin).Disconnect
                            TheHdw.PPMU.Pins(pin).ForceI 0, 0.002
                            TheHdw.PPMU.Pins(pin).Connect
                            TheHdw.PPMU.Pins(pin).Gate = tlOn
                         
                             'if PPMU > 50 mA set Warning and set PPMU = 50 mA
                             If Abs(Adjust_I_Value.Pins(pin).value(site)) <= 50 * mA Then
                                TheHdw.PPMU.Pins(pin).ForceI Adjust_I_Value.Pins(pin).value(site), Abs(Adjust_I_Value.Pins(pin).value(site))
                               
                                'NEW 20170728
                                If UCase(pin) Like UCase("DDR*_P*") Or UCase(pin) Like UCase("DDR*_N*") Then  ' Differential pair needs to force opposite current
                                    If UCase(pin) Like ("DDR*_DQS_P*") Then             'DDR0_DQS_P0
                                        Pin_Diff = Replace(UCase(pin), "DQS_P", "DQS_N")
                                    ElseIf UCase(pin) Like ("DDR*_DQS_N*") Then
                                        Pin_Diff = Replace(UCase(pin), "DQS_N", "DQS_P")
                                    ElseIf UCase(pin) Like ("DDR*_CK*_P*") Then          'DDR0_CK_P
                                        Pin_Diff = Replace(UCase(pin), "_P", "_N")
                                    ElseIf UCase(pin) Like ("DDR*_CK*_N*") Then
                                        Pin_Diff = Replace(UCase(pin), "_N", "_P")
                                    Else 'Do nothing '20230601
                                    End If
                                    TheHdw.PPMU.Pins(Pin_Diff).ForceI (-1) * Adjust_I_Value.Pins(pin).value(site), Abs(Adjust_I_Value.Pins(pin).value(site))
                                    theexec.Datalog.WriteComment ("Pin Diff: " & Pin_Diff & " Site (" & site & ")" & " , Force Value : " & Format((-1) * Adjust_I_Value.Pins(pin).value(site) * 1000, "0.000") & "mA")
                                End If
                                
                            Else
                                theexec.Datalog.WriteComment (" Error : Irange >= 50mA , Bypass Pin " & pin & " Measurement ")
                                MeasureValue.Pins(pin).value(site) = 0
                                Final_Volt_Value.Pins(pin).value(site) = MeasureValue.Pins(pin).value(site)
                                Final_Force_Curr_Value.Pins(pin).value(site) = Adjust_I_Value.Pins(pin).value(site)
                                TheHdw.PPMU.Pins(pin).Gate = tlOff
                                TheHdw.PPMU.Pins(pin).Disconnect
                                TheHdw.Digital.Pins(pin).Connect
                                Exit For
                            End If
                        
                            TheHdw.Wait 0.002
                             MeasureValue.Pins(pin).value(site) = TheHdw.PPMU.Pins(pin).Read(tlPPMUReadMeasurements, 10)
                            
                            TheHdw.PPMU.Pins(pin).ForceI 0, 0
                            TheHdw.PPMU.Pins(pin).Gate = tlOff
                            TheHdw.PPMU.Pins(pin).Disconnect
                            TheHdw.Digital.Pins(pin).Connect
                         
                            '' Calculate RAK
                            If RAK_Flag = R_TraceOnly Then
                                'RakV = TheHdw.PPMU.ReadRakValuesByPinnames(pin, Site)
                                
''                                If InStr(TheExec.CurrentChanMap, "CP") <> 0 Then
''                                    MeasureValue.Pins(Pin).Value(Site) = MeasureValue.Pins(Pin).Value(Site) - Adjust_I_Value.Pins(Pin).Value(Site) * (CP_Card_RAK.Pins(Pin).Value(Site) + RakV(0))
''                                Else
''                                    MeasureValue.Pins(Pin).Value(Site) = MeasureValue.Pins(Pin).Value(Site) - Adjust_I_Value.Pins(Pin).Value(Site) * (FT_Card_RAK.Pins(Pin).Value(Site) + RakV(0))  ' + TheHdw.PPMU.ReadRakValuesByPinnames(FT_Card_RAK.Pins(pin).Name, Site))
''                                End If
                                MeasureValue.Pins(pin).value(site) = MeasureValue.Pins(pin).value(site) - Adjust_I_Value.Pins(pin).value(site) * (CurrentJob_Card_RAK.Pins(pin).value(site))
                            
                            ElseIf RAK_Flag = R_PathWithContact Then
                                MeasureValue.Pins(pin).value(site) = MeasureValue.Pins(pin).value(site) - Adjust_I_Value.Pins(pin).value(site) * R_Path_PLD.Pins(pin).value(site)
                            Else 'Do nothing '20230601
                            End If
                       Else
                            MeasureValue.Pins(pin).value(site) = 0
                            Final_Volt_Value.Pins(pin).value(site) = MeasureValue.Pins(pin).value(site)
                            Final_Force_Curr_Value.Pins(pin).value(site) = Adjust_I_Value.Pins(pin).value(site)
                            Counter_Meas = Counter_End + 1 'Exit from loop
                       End If
                       
                        If R1_Value.Pins(pin).value(site) = 0 Then
                            theexec.Datalog.WriteComment ("Adjust " & "Site_" & site & " pin = " & pin & " R0 value = " & Format(R0_Value.Pins(pin).value(site), "0.000") & " R1 value = " & "NA" & _
                                                                            " Meas Volt = " & Format(MeasureValue.Pins(pin).value(site), "0.0000"))
                        Else
                            theexec.Datalog.WriteComment ("Adjust " & "Site_" & site & " pin = " & pin & " R0 value = " & Format(R0_Value.Pins(pin).value(site), "0.000") & " R1 value = " & Format(R1_Value.Pins(pin).value(site), "0.000") & _
                                                                            " Meas Volt = " & Format(MeasureValue.Pins(pin).value(site), "0.0000"))
                        End If
                        
                       If Counter_Meas = Counter_End Then
                            Final_Volt_Value.Pins(pin).value(site) = MeasureValue.Pins(pin).value(site)
                            Final_Force_Curr_Value.Pins(pin).value(site) = Adjust_I_Value.Pins(pin).value(site)
                       End If
                        
                    End If
                    
               Next Counter_Meas
        Next site
    Next p


    Dim TestName As String
    Dim Temp_index
    
    

    Temp_index = theexec.Flow.TestLimitIndex

    For Each pin In Final_Volt_Value.Pins
        
        theexec.Flow.TestLimitIndex = Temp_index
        TestName = Report_TName_From_Instance(CalcC, CStr(pin))
        'TestName = Report_TName_From_Instance("V", CStr(pin))
        For Each site In theexec.sites
             If UCase(glb_TestInstance) Like "*VOLH_SWEEP*LOOP*" Then
                 theexec.Flow.TestLimit Final_Volt_Value.Pins(pin).value(site), PinName:=Final_Volt_Value.Pins(pin).name, scaletype:=scaleNone, unit:=unitVolt, formatStr:="%.3f", Tname:=CStr(Temp_Input(TestSeqNum)) & "_" & CStr(TestSeqNum), ForceVal:=Final_Force_Curr_Value.Pins(pin).value(site), ForceUnit:=unitAmp, ForceResults:=tlForceNone_CZ
             ElseIf UCase(glb_TestInstance) Like "*VOLH_SWEEP*" Then       ' 20170912 Used for VOLH_SWEEP Average ZCAL test
                Dim ZCAL_Testname As String: ZCAL_Testname = "Average_ZCAL"
                theexec.Flow.TestLimit Final_Volt_Value.Pins(pin).value(site), lowVal:=loLimit, hiVal:=hiLimit, PinName:=Final_Volt_Value.Pins(pin).name, scaletype:=scaleNone, unit:=unitVolt, formatStr:="%.3f", Tname:=ZCAL_Testname & "_" & CStr(Temp_Input(TestSeqNum)) & "_" & CStr(TestSeqNum), ForceVal:=Final_Force_Curr_Value.Pins(pin).value(site), ForceUnit:=unitAmp, ForceResults:=tlForceNone_CZ
                Else
                'TheExec.Flow.TestLimit Final_Volt_Value.Pins(Pin).Value(Site), lowVal:=LoLimit, hiVal:=HiLimit, PinName:=Final_Volt_Value.Pins(Pin).Name, scaletype:=scaleNone, Unit:=unitVolt, formatStr:="%.3f", Tname:=TestName & "_" & CStr(TestSeqNum), forceVal:=Final_Force_Curr_Value.Pins(Pin).Value(Site), ForceUnit:=unitAmp, ForceResults:=tlForceFlow
                theexec.Flow.TestLimit Final_Volt_Value.Pins(pin).value(site), lowVal:=loLimit, hiVal:=hiLimit, PinName:=Final_Volt_Value.Pins(pin).name, scaletype:=scaleNone, unit:=unitVolt, formatStr:="%.3f", Tname:=TestName, ForceVal:=Final_Force_Curr_Value.Pins(pin).value(site), ForceUnit:=unitAmp, ForceResults:=tlForceFlow
                'TheExec.Flow.TestLimit Final_Volt_Value.Pins(pin).Value(site), lowval:=LoLimit, hival:=HiLimit, PinName:=Final_Volt_Value.Pins(pin).name, ScaleType:=scaleNone, Unit:=unitVolt, FormatStr:="%.3f", TName:=CStr(Temp_Input(TestSeqNum)) & "_" & CStr(TestSeqNum), ForceVal:=Final_Force_Curr_Value.Pins(pin).Value(site), ForceUnit:=unitAmp, ForceResults:=tlForceFlow
                End If
             theexec.Flow.TestLimitIndex = theexec.Flow.TestLimitIndex - 1
        Next site
                    
        theexec.Flow.TestLimitIndex = theexec.Flow.TestLimitIndex + 1
    Next pin
                    
    '20171204 Print Out Final Pull Up / Pull Down resistance
    Dim Res_Tname As String: Res_Tname = vbNullString
    Dim Target_Resistance_Hi_Lim As Double: Target_Resistance_Hi_Lim = Target_Resistance * 1.1
    Dim Target_Resistance_Lo_Lim As Double: Target_Resistance_Lo_Lim = Target_Resistance * 0.9
                
    'Output Datalog PU/PD Info
    theexec.Datalog.WriteComment (vbNullString)
             
    For Each pin In Final_Volt_Value.Pins
        For Each site In theexec.sites
            If UCase(Temp_Input(TestSeqNum)) = "VOH" Then
                Res_Tname = "R_Pull_Up"
            ElseIf UCase(Temp_Input(TestSeqNum)) = "VOL" Then
                Res_Tname = "R_Pull_Down"
            Else 'Do nothing '20230601
                End If
                
            If UCase(glb_TestInstance) Like "*VOLH_SWEEP*LOOP*" Then
                theexec.Flow.TestLimit R1_Value.Pins(pin).value(site), PinName:=Final_Volt_Value.Pins(pin).name, scaletype:=scaleNone, unit:=unitCustom, customUnit:="Ohm", formatStr:="%.3f", Tname:=Res_Tname & "_" & CStr(TestSeqNum), ForceVal:=Final_Force_Curr_Value.Pins(pin).value(site), ForceUnit:=unitAmp, ForceResults:=tlForceNone_CZ
            ElseIf UCase(glb_TestInstance) Like "*VOLH_SWEEP*" Then
               theexec.Flow.TestLimit R1_Value.Pins(pin).value(site), lowVal:=Target_Resistance_Lo_Lim, hiVal:=Target_Resistance_Hi_Lim, PinName:=Final_Volt_Value.Pins(pin).name, scaletype:=scaleNone, unit:=unitCustom, customUnit:="Ohm", formatStr:="%.3f", Tname:=ZCAL_Testname & "_" & Res_Tname & "_" & CStr(TestSeqNum), ForceVal:=Final_Force_Curr_Value.Pins(pin).value(site), ForceUnit:=unitAmp, ForceResults:=tlForceNone_CZ
                Else
               theexec.Flow.TestLimit R1_Value.Pins(pin).value(site), lowVal:=Target_Resistance_Lo_Lim, hiVal:=Target_Resistance_Hi_Lim, PinName:=Final_Volt_Value.Pins(pin).name, scaletype:=scaleNone, unit:=unitCustom, customUnit:="Ohm", formatStr:="%.3f", Tname:=Res_Tname & "_" & CStr(TestSeqNum), ForceVal:=Final_Force_Curr_Value.Pins(pin).value(site), ForceUnit:=unitAmp, ForceResults:=tlForceNone_CZ
                End If
        Next site
    Next pin
    
Exit Function 'Add ErrHandler 2023/05/29
errHandler: 'Add ErrHandler 2023/05/29
    Call Print_Error_Message(Error_Info, "VBT_LIB_HardIP_Customize", "CUS_DDR_Emulate_Const_Res_Loading") 'Add ErrHandler 2023/05/29
    If AbortTest Then Exit Function Else Resume Next 'Add ErrHandler 2023/05/29
End Function

Public Function CUS_DDR_DCS_PrintOut() As Long
On Error GoTo errHandler 'Add ErrHandler 2023/05/29

    theexec.Datalog.WriteComment (vbNullString)
    theexec.Datalog.WriteComment (" VDD_DCS_DDR Level = " & FormatNumber(TheHdw.DCVS.Pins("VDD_DCS_DDR").Voltage.value, 3) & " v")
    theexec.Datalog.WriteComment (vbNullString)

Exit Function 'Add ErrHandler 2023/05/29
errHandler: 'Add ErrHandler 2023/05/29
    Call Print_Error_Message(Error_Info, "VBT_LIB_HardIP_Customize", "CUS_DDR_DCS_PrintOut") 'Add ErrHandler 2023/05/29
    If AbortTest Then Exit Function Else Resume Next 'Add ErrHandler 2023/05/29
End Function
Public Function MEAS_I_ABS(MeasureValue As PinListData) As Long
On Error GoTo errHandler 'Add ErrHandler 2023/05/29

Dim p As Double
For p = 0 To MeasureValue.Pins.Count - 1
    MeasureValue.Pins(p) = MeasureValue.Pins(p).Abs
Next p

Exit Function 'Add ErrHandler 2023/05/29
errHandler: 'Add ErrHandler 2023/05/29
    Call Print_Error_Message(Error_Info, "VBT_LIB_HardIP_Customize", "MEAS_I_ABS") 'Add ErrHandler 2023/05/29
    If AbortTest Then Exit Function Else Resume Next 'Add ErrHandler 2023/05/29
End Function

Public Function CUS_RREF_Rak_Calc(ByRef MeasureVolt As PinListData) As Long
On Error GoTo errHandler 'Add ErrHandler 2023/05/29

    Dim Ary_str(0) As String
    Dim VDDQL_Val As Double
    Dim MeasV As Double
    Dim pin_name As String
    Dim GetRakVal As Double
    Dim p As Long
    'Dim RakV() As Double
    Dim site As Variant 'Carter, 20240304
    Ary_str(0) = "_VDDQL_DDR_VAR"
    Call HIP_Evaluate_ForceVal(Ary_str)
    VDDQL_Val = CDbl(Ary_str(0))

    For p = 0 To MeasureVolt.Pins.Count - 1 Step 1
          pin_name = MeasureVolt.Pins(p)
          
          For Each site In theexec.sites.Active
                MeasV = MeasureVolt.Pins(pin_name).value(site)
                'RakV = TheHdw.PPMU.ReadRakValuesByPinnames(pin_name, site)
                GetRakVal = CurrentJob_Card_RAK.Pins(pin_name).value(site)
                MeasureVolt.Pins(pin_name).value(site) = MeasV - GetRakVal * (VDDQL_Val - MeasV) / 240
          Next site
          
    Next p
        
Exit Function 'Add ErrHandler 2023/05/29
errHandler: 'Add ErrHandler 2023/05/29
    Call Print_Error_Message(Error_Info, "VBT_LIB_HardIP_Customize", "CUS_RREF_Rak_Calc") 'Add ErrHandler 2023/05/29
    If AbortTest Then Exit Function Else Resume Next 'Add ErrHandler 2023/05/29
End Function
Public Function CUS_AMP_SDLL_SWP(MeasFreq As PinListData, Extra_TName As String) As Long
On Error GoTo errHandler 'Add ErrHandler 2023/05/29
   
    Dim Dict_Freq_PLD As New PinListData
    Dim Step_Freq_PLD As New PinListData
    Dim pin_name As Variant
    Dim Extra_TName_StrArr() As String
    Dim Dict_idx As Long
    Dim Step_idx As Long
    Dim OutputTname_format() As String
    Dim TestNameInput As String
    Dim site As Variant 'Carter, 20240304
    Dict_Freq_PLD = GetStoreDataAllType("Freq_SDLL_SWP")
    
    Dim p As Long
     
     For p = 0 To MeasFreq.Pins.Count - 1
         If InStr(UCase(MeasFreq.Pins(p)), "_P") Then
            pin_name = MeasFreq.Pins(p).name
         
             For Dict_idx = 0 To Dict_Freq_PLD.Pins.Count - 1
                 If (pin_name = Dict_Freq_PLD.Pins(Dict_idx).name) Then
                     Step_Freq_PLD.AddPin (pin_name)
                     
                     For Each site In theexec.sites.Active
                         If MeasFreq.Pins(p).value(site) = 0 Or Dict_Freq_PLD.Pins(Dict_idx).value(site) = 0 Then
                            Step_Freq_PLD.Pins(pin_name).value(site) = 999
                            theexec.Datalog.WriteComment (" Freq measurement = 0 , Denominator = 0  ! ")
                         Else
                            Step_Freq_PLD.Pins(pin_name).value(site) = 0.5 * ((1 / MeasFreq.Pins(p).value(site)) - (1 / Dict_Freq_PLD.Pins(Dict_idx).value(site)))
                         End If
                     Next site
                     
                     Exit For
                 End If
             Next Dict_idx
         End If
     Next p

     Extra_TName_StrArr = Split(Extra_TName, ":")
              
    For Step_idx = 0 To Step_Freq_PLD.Pins.Count - 1
            pin_name = Step_Freq_PLD.Pins(Step_idx).name
            pin_name = Replace(LCase(pin_name), "ddr", "ch")
            pin_name = Replace(LCase(pin_name), "dqs_p", "core")
            
            '' Extra_TName_StrArr (0) : Frequency
            '' Extra_TName_StrArr (1) : Octant
            '' Extra_TName_StrArr (2) : Loop_Idx
            '' Extra_TName_StrArr (3) : Sweep_Name (LSW_0X or MSW_0X)
           TestNameInput = Report_TName_From_Instance("F", Step_Freq_PLD.Pins(Step_idx), "Step_SDLL_SWP", 0, Step_idx)
           '''         OutputTname_format(8) = CStr(gl_Tname_Alg_Index)           TheExec.Flow.TestLimit resultVal:=Step_Freq_PLD.Pins(Step_idx), Unit:=unitTime, Tname:=TestNameInput, ForceResults:=tlForceFlow

    Next Step_idx
  
Exit Function 'Add ErrHandler 2023/05/29
errHandler: 'Add ErrHandler 2023/05/29
    Call Print_Error_Message(Error_Info, "VBT_LIB_HardIP_Customize", "CUS_AMP_SDLL_SWP") 'Add ErrHandler 2023/05/29
    If AbortTest Then Exit Function Else Resume Next 'Add ErrHandler 2023/05/29
End Function

Public Function D2D_Flag_Sync_Master_Slave(list As String, sync_site As String)
'[Jim][Pald2d BraC] For setting Master and Slave site by instance argument for fail flag
'sync_site ==> FT 0:2,1:3 WLFT 0:1
'Example :
'For FT instance : [list] fail flag name ; [sync_site] 0:2,1:3
'For WLFT instance : [list] fail flag name ; [sync_site] 0:1
On Error GoTo errHandler

Dim funcName As String:: funcName = "D2D_Flag_Sync_Master_Slave"

Dim Flag() As String
Dim sync_site_arr() As String
Dim sync_site_idx As Integer
Dim m_site As Integer
Dim S_site As Integer

Flag = Split(list, ",")
sync_site_arr = Split(sync_site, ",")
For sync_site_idx = 0 To UBound(sync_site_arr)
    m_site = Split(sync_site_arr(sync_site_idx), ":")(0)
    S_site = Split(sync_site_arr(sync_site_idx), ":")(1)
Dim flag_idx As Integer
    For flag_idx = 0 To UBound(Flag)
        If theexec.sites(m_site).Active = True And theexec.sites(S_site).Active = True Then
            If theexec.sites.item(m_site).FlagState(Flag(flag_idx)) = logicTrue Then
                theexec.sites.item(S_site).FlagState(Flag(flag_idx)) = logicTrue
            End If
            If theexec.sites.item(S_site).FlagState(Flag(flag_idx)) = logicTrue Then
                theexec.sites.item(m_site).FlagState(Flag(flag_idx)) = logicTrue
            End If
        End If
    Next
Next
    Exit Function
errHandler:
    Call Print_Error_Message(Error_Info, "VBT_LIB_HardIP_Customize", "D2D_Flag_Sync_Master_Slave") 'Add ErrHandler 2023/05/29
    If AbortTest Then Exit Function Else Resume Next 'Add ErrHandler 2023/05/29
End Function
