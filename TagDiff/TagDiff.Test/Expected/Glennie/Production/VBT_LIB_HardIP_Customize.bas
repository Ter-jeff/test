Attribute VB_Name = "VBT_LIB_HardIP_Customize"
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
    Dim pinAry() As String
    Dim FirstSite As Integer
                                                                                                                                                                                                                   
    If (UBound(ForceI) <> 0) Then
        ForceISeqIndexPerSeq = TestSeqNum
    Else
        ForceISeqIndexPerSeq = 0
    End If
                                                                                                                                                                                                                   
    ForceI_Ary = Split(ForceI(ForceISeqIndexPerSeq), ",")
                                                                                                                                                                                                                   
    pinAry = Split(TestPinArrayIV(TestSeqNum), ",")
                                                                                                                                                                                                                   
    FirstSite = 0
    
    If (UBound(ForceI_Ary) = 0) Then
                                                                                                                                                                                                                   
        For Each site In TheExec.sites
                                                                                                                                                                                                                   
            For p = 0 To MeasV.pins.Count - 1
                                                                                                                                                                                                                   
                If (FirstSite = 0) Then
                    R.AddPin (MeasV.pins(p))
                End If
                                                                                                                                                                                                                   
                If (UCase(CUS_CalR_Seq(TestSeqNum)) Like "*RVOH*") Then
                                                                                                                                                                                                                   
                    R.pins(p).value(site) = CUS_VDD - MeasV.pins(p).value(site)
                    R.pins(p).value(site) = R.pins(p).divide(CDbl(ForceI_Ary(0)))
                                                                                                                                                                                                                   
                ElseIf (UCase(CUS_CalR_Seq(TestSeqNum)) Like "*RVOL*") Then
                                                                                                                                                                                                                   
                    R.pins(p).value(site) = MeasV.pins(p).value(site)
                    R.pins(p).value(site) = R.pins(p).divide(CDbl(ForceI_Ary(0)))
                Else 'Do nothing '20230601
                End If
                                                                                                                                                                                                                   
                'R_AK = TheHdw.PPMU.ReadRakValuesByPinnames(MeasV.Pins(p), site)  ''Get instrament impedance
                                                                                                                                                                                                                   
''                If InStr(UCase(TheExec.CurrentChanMap), "FT") <> 0 Then                          ''Get DIB impedance
''                    R_AK(0) = R_AK(0) + FT_Card_RAK.Pins(MeasV.Pins(p)).Value(Site)
''                Else
''                    R_AK(0) = R_AK(0) + CP_Card_RAK.Pins(MeasV.Pins(p)).Value(Site)
''                End If
                R_AK(0) = CurrentJob_Card_RAK.pins(MeasV.pins(p)).value(site)
                                                                                                                                                                                                                   
                R.pins(p).value(site) = R.pins(p).value(site) - R_AK(0)
                                                                                                                                                                                                                   
            Next p
            FirstSite = FirstSite + 1
        Next site
                                                                                                                                                                                                                   
        TheExec.flow.TestLimit resultVal:=R, unit:=unitCustom, Tname:="Calculate_" + CUS_CalR_Seq(TestSeqNum), customUnit:="ohm", ForceResults:=tlForceNone 'Un-Used
                                                                                                                                                                                                                   
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
    
    If UBound(ForceSequenceArray) = 0 Then
        ForceVolt = ForceSequenceArray(0)
    Else
        ForceVolt = ForceSequenceArray(TestSeqNum)
    End If
    
    For Each site In TheExec.sites
        For p = 0 To MeasCurr.pins.Count - 1
                                                                                                                                                                                                               
            If b_FirstTime = True Then
                RTN_Imped.AddPin (MeasCurr.pins(p))
            End If
                                                                                                                                                                                                               
            If (UCase(CUS_CalR_Seq_Ary(TestSeqNum)) Like "*RVOH*") Then
                                                                                                                                                                                                               
                RTN_Imped.pins(p).value(site) = CUS_CalR_VDD - ForceVolt
                RTN_Imped.pins(p).value(site) = RTN_Imped.pins(p).divide(MeasCurr.pins(p)).Multiply(-1)
                                                                                                                                                                                                               
            ElseIf (UCase(CUS_CalR_Seq_Ary(TestSeqNum)) Like "*RVOL*") Then
                                                                                                                                                                                                               
                RTN_Imped.pins(p).value(site) = ForceVolt
                RTN_Imped.pins(p).value(site) = RTN_Imped.pins(p).divide(MeasCurr.pins(p))
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
    
    StepIndex_Val = CDbl(val(TheExec.flow.var("SrcCodeIndx").value))
    ForceVolt_Val_1 = StartVolt_Val_1 + StepIndex_Val * StepVolt_Val_1
    ForceVolt_Val_2 = StartVolt_Val_2 + StepIndex_Val * StepVolt_Val_2

    TheHdw.DCVS.pins(Force_Pins_1).Voltage.value = ForceVolt_Val_1
    TheHdw.DCVS.pins(Force_Pins_2).Voltage.value = ForceVolt_Val_2
    
    TheExec.Datalog.WriteComment ("Force Pin " & Force_Pins_1 & " Value = " & ForceVolt_Val_1)
    TheExec.Datalog.WriteComment ("Force Pin " & Force_Pins_2 & " Value = " & ForceVolt_Val_2)
Exit Function 'Add ErrHandler 2023/05/29
errHandler: 'Add ErrHandler 2023/05/29
    Call Print_Error_Message(Error_Info, "VBT_LIB_HardIP_Customize", "Cust_Sweep_V") 'Add ErrHandler 2023/05/29
    If AbortTest Then Exit Function Else Resume Next 'Add ErrHandler 2023/05/29
End Function
Public Function VOLH_Sweep(CUS_Str_DigSrcData As String, DigSrc_Assignment As String) As Long
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
    
    SrcCode_Target_Dec = TheExec.flow.var("SrcCodeIndx").value
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
    DigSrc_Assignment = Replace(DigSrc_Assignment, ReplaceTarget_1, SrcCode_Target_Bin_One)
    DigSrc_Assignment = Replace(DigSrc_Assignment, ReplaceTarget_2, SrcCode_Target_Bin_Two)
Exit Function 'Add ErrHandler 2023/05/29
errHandler: 'Add ErrHandler 2023/05/29
    Call Print_Error_Message(Error_Info, "VBT_LIB_HardIP_Customize", "VOLH_Sweep") 'Add ErrHandler 2023/05/29
    If AbortTest Then Exit Function Else Resume Next 'Add ErrHandler 2023/05/29
End Function

Public Function MTR_UVI80_Setup()
On Error GoTo errHandler 'Add ErrHandler 2023/05/29
    ' 20170227 - Set current range only for osprey Metrology 20170227
            TheHdw.DCVI.pins("mtr_analog_test_p").CurrentRange = 0.002
            TheHdw.DCVI.pins("mtr_analog_test_p").Current = 0.002

Exit Function 'Add ErrHandler 2023/05/29
errHandler: 'Add ErrHandler 2023/05/29
    Call Print_Error_Message(Error_Info, "VBT_LIB_HardIP_Customize", "MTR_UVI80_Setup") 'Add ErrHandler 2023/05/29
    If AbortTest Then Exit Function Else Resume Next 'Add ErrHandler 2023/05/29
End Function


Public Function CUS_DDR_Emulate_Const_Res_Loading(MeasureValue As PinListData, ForceValByPin() As String, CUS_Str_MainProgram As String, TestSeqNum As Integer, _
    Optional RAK_Flag As Enum_RAK = 0, Optional MeasVPinGrp As String, Optional pinGroup_sequence As Long = 0) As Long
    
    On Error GoTo err:
    Dim R0_Value As New PinListData
    Dim Final_Volt_Value As New PinListData
    Dim Final_Force_Curr_Value As New PinListData
    Dim R1_Value As New PinListData
    Dim Adjust_I_Value As New PinListData
    Dim site As Variant
    Dim pin  As Variant
    Dim Temp_Input() As String
    Dim Pwr_Voltage As Double
    Dim Target_Resistance As Double
    Dim Flag_1 As Boolean: Flag_1 = False
    Dim Flag_2 As Boolean: Flag_2 = False
    Dim Initial_Setting_Flag As Boolean: Initial_Setting_Flag = True
    Dim Counter_Meas As Integer: Counter_Meas = 1
    Dim Counter_End As Integer: Counter_End = 10
    Dim ForceValue As Double: ForceValue = Abs(ForceValByPin(0))
    Dim p As Integer
    Dim RakV() As Double
    Dim Ary_str(0) As String
    Dim hiLimit As Double
    Dim loLimit As Double
    Dim Pin_Diff As String
    Dim error_flag As Boolean: error_flag = False
    Dim Isitevaluetmp As New SiteDouble
    Dim Vsitevaluetmp As New SiteDouble
    
    Dim OutputTname_format() As String
    Dim TestNameInput As String
    '20211129 Cater @YM update for Prinia
    Dim lPinCnt As Long
    Dim sPinAry() As String
    Dim mode As String
    TheExec.DataManager.DecomposePinList MeasVPinGrp, sPinAry, lPinCnt
    
    glb_TestInstance = vbNullString
    glb_TestInstance = UCase(TheExec.DataManager.instancename)

    Pwr_Voltage = TheHdw.DCVS.pins("VDDQL_DDR").Voltage.value
    Ary_str(0) = "_VDDQL_DDR_VAR"

    Call HIP_Evaluate_ForceVal(Ary_str)
    hiLimit = 0.55 * CDbl(Ary_str(0))
    loLimit = 0.45 * CDbl(Ary_str(0))
    
    Temp_Input() = Split(CUS_Str_MainProgram, ":")     ' CUS_Str_MainProgram . Ex: Tname:VOL&VOH,VOL&VOH,48
    Temp_Input() = Split(Temp_Input(1), ",")
    Target_Resistance = Temp_Input(UBound(Temp_Input))
    mode = UCase(Split(Temp_Input(TestSeqNum), "&")(pinGroup_sequence))
    
    For p = 0 To UBound(sPinAry) ''MeasureValue.Pins.Count - 1
        Flag_1 = False
        pin = sPinAry(p) ''pin = MeasureValue.Pins(p).Name
        
        For Each site In TheExec.sites
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
                    
                    If mode = "VOH" Then
                        If Initial_Setting_Flag = True Then
                            R0_Value.pins(pin) = MeasureValue.pins(pin).value(site) / ForceValue
                        Else
                            R0_Value.pins(pin) = MeasureValue.pins(pin).value(site) / Abs(Adjust_I_Value.pins(pin).value(site))
                        End If
                    ElseIf mode = "VOL" Then
                        If Initial_Setting_Flag = True Then
                            R0_Value.pins(pin) = (Pwr_Voltage - MeasureValue.pins(pin).value(site)) / ForceValue
                        Else
                            R0_Value.pins(pin) = (Pwr_Voltage - MeasureValue.pins(pin).value(site)) / Adjust_I_Value.pins(pin).value(site)
                        End If
                    End If
                    
                    If ((Abs(R0_Value.pins(pin).value(site) - Target_Resistance)) / Target_Resistance) < 0.001 Then
                        Final_Volt_Value.pins(pin).value(site) = MeasureValue.pins(pin).value(site)
                        
                        If Counter_Meas = 1 Then
                            If mode = "VOH" Then
                                Final_Force_Curr_Value.pins(pin).value(site) = -1 * ForceValue
                            ElseIf mode = "VOL" Then
                                Final_Force_Curr_Value.pins(pin).value(site) = ForceValue
                            End If
                        Else
                            Final_Force_Curr_Value.pins(pin).value(site) = Adjust_I_Value.pins(pin).value(site)
                        End If
                        Counter_Meas = Counter_End + 1 ' Iteration search done , Exit from Loop
                        
                        ' 20171204 Update latest R1_Value
                        If mode = "VOH" Then
                            R1_Value.pins(pin).value(site) = (Pwr_Voltage - Final_Volt_Value.pins(pin).value(site)) / Abs(Final_Force_Curr_Value.pins(pin).value(site))
                        ElseIf mode = "VOL" Then
                            R1_Value.pins(pin).value(site) = Final_Volt_Value.pins(pin).value(site) / Final_Force_Curr_Value.pins(pin).value(site)
                        End If
                        
                        If R1_Value.pins(pin).value(site) = 0 Then
                            TheExec.Datalog.WriteComment ("Final " & "Site_" & site & " pin = " & pin & " R0 value = " & Format(R0_Value.pins(pin).value(site), "0.000") & " R1 value = " & "NA" & _
                                                                            " Meas Volt = " & Format(Final_Volt_Value.pins(pin).value(site), "0.0000"))
                        Else
                            TheExec.Datalog.WriteComment ("Final " & "Site_" & site & " pin = " & pin & " R0 value = " & Format(R0_Value.pins(pin).value(site), "0.000") & " R1 value = " & Format(R1_Value.pins(pin).value(site), "0.000") & _
                                                                            " Meas Volt = " & Format(Final_Volt_Value.pins(pin).value(site), "0.0000"))
                        End If
                    Else
                                                            
                        If mode = "VOH" Then
                            If Initial_Setting_Flag = True Then
                                R1_Value.pins(pin).value(site) = (Pwr_Voltage - MeasureValue.pins(pin).value(site)) / ForceValue
                            Else
                                R1_Value.pins(pin).value(site) = (Pwr_Voltage - MeasureValue.pins(pin).value(site)) / Abs(Adjust_I_Value.pins(pin).value(site))
                            End If
                       ElseIf mode = "VOL" Then
                            If Initial_Setting_Flag = True Then
                                R1_Value.pins(pin).value(site) = MeasureValue.pins(pin).value(site) / ForceValue
                            Else
                                R1_Value.pins(pin).value(site) = MeasureValue.pins(pin).value(site) / Adjust_I_Value.pins(pin).value(site)
                            End If
                        End If
                        
                        If (R1_Value.pins(pin).value(site) + Target_Resistance) = 0 Then
                            Adjust_I_Value.pins(pin).value(site) = 0
                            TheExec.Datalog.WriteComment (" Error : Denominator = 0 !  ")
                            error_flag = True
                        Else
                            Adjust_I_Value.pins(pin).value(site) = Pwr_Voltage / (R1_Value.pins(pin).value(site) + Target_Resistance)
                        End If
                        
                        If TheExec.TesterMode = testModeOffline Then
                            Adjust_I_Value.pins(pin).value(site) = 0.005 / Counter_Meas
                        End If
                        
                        If mode = "VOH" Then
                            Adjust_I_Value.pins(pin).value(site) = (-1) * Adjust_I_Value.pins(pin).value(site)
                        End If
                        
                        Initial_Setting_Flag = False    ' False means start to use Adjust I Value for next Iteration search
                        
                        If error_flag = False Then
                            ' Update Force Condition and Measure Voltage
                            Isitevaluetmp = 0
                            TheHdw.Digital.pins(pin).Disconnect
                            If glb_TesterType = "Jaguar" Then
                                TheHdw.PPMU.pins(pin).ForceI 0, 0.002
                            ElseIf glb_TesterType = "UltraFLEXplus" Then
                                TheHdw.PPMU.pins(pin).ForceIPerSite Isitevaluetmp, 0.002
                            End If
                            TheHdw.PPMU.pins(pin).Connect
                            TheHdw.PPMU.pins(pin).Gate = tlOn
                         
                             'if PPMU > 50 mA set Warning and set PPMU = 50 mA
                             If Abs(Adjust_I_Value.pins(pin).value(site)) <= 50 * mA Then
                                TheHdw.PPMU.pins(pin).ForceI Adjust_I_Value.pins(pin).value(site), Abs(Adjust_I_Value.pins(pin).value(site))
                               
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
                                    TheHdw.PPMU.pins(Pin_Diff).ForceI (-1) * Adjust_I_Value.pins(pin).value(site), Abs(Adjust_I_Value.pins(pin).value(site))
                                    TheExec.Datalog.WriteComment ("Pin Diff: " & Pin_Diff & " Site (" & site & ")" & " , Force Value : " & Format((-1) * Adjust_I_Value.pins(pin).value(site) * 1000, "0.000") & "mA")
                                End If
                                
                            Else
                                TheExec.Datalog.WriteComment (" Error : Irange >= 50mA , Bypass Pin " & pin & " Measurement ")
                                MeasureValue.pins(pin).value(site) = 0
                                Final_Volt_Value.pins(pin).value(site) = MeasureValue.pins(pin).value(site)
                                Final_Force_Curr_Value.pins(pin).value(site) = Adjust_I_Value.pins(pin).value(site)
                                TheHdw.PPMU.pins(pin).Gate = tlOff
                                TheHdw.PPMU.pins(pin).Disconnect
                                TheHdw.Digital.pins(pin).Connect
                                Exit For
                            End If
                        
                            TheHdw.Wait 0.002
                             MeasureValue.pins(pin).value(site) = TheHdw.PPMU.pins(pin).Read(tlPPMUReadMeasurements, 10)
                            Isitevaluetmp = 0
                            If glb_TesterType = "Jaguar" Then
                                TheHdw.PPMU.pins(pin).ForceI 0, 0
                            ElseIf glb_TesterType = "UltraFLEXplus" Then
                                TheHdw.PPMU.pins(pin).ForceIPerSite Isitevaluetmp, 0
                            End If
                            TheHdw.PPMU.pins(pin).Gate = tlOff
                            TheHdw.PPMU.pins(pin).Disconnect
                            TheHdw.Digital.pins(pin).Connect

                            '' Calculate RAK
                            If RAK_Flag = R_TraceOnly Then
                                'RakV = TheHdw.PPMU.ReadRakValuesByPinnames(pin, Site)
                                
''                                If InStr(TheExec.CurrentChanMap, "CP") <> 0 Then
''                                    MeasureValue.Pins(Pin).Value(Site) = MeasureValue.Pins(Pin).Value(Site) - Adjust_I_Value.Pins(Pin).Value(Site) * (CP_Card_RAK.Pins(Pin).Value(Site) + RakV(0))
''                                Else
''                                    MeasureValue.Pins(Pin).Value(Site) = MeasureValue.Pins(Pin).Value(Site) - Adjust_I_Value.Pins(Pin).Value(Site) * (FT_Card_RAK.Pins(Pin).Value(Site) + RakV(0))  ' + TheHdw.PPMU.ReadRakValuesByPinnames(FT_Card_RAK.Pins(pin).Name, Site))
''                                End If
                                MeasureValue.pins(pin).value(site) = MeasureValue.pins(pin).value(site) - Adjust_I_Value.pins(pin).value(site) * (CurrentJob_Card_RAK.pins(pin).value(site))
                            
                            ElseIf RAK_Flag = R_PathWithContact Then
                                MeasureValue.pins(pin).value(site) = MeasureValue.pins(pin).value(site) - Adjust_I_Value.pins(pin).value(site) * R_Path_PLD.pins(pin).value(site)
                            Else 'Do nothing '20230601
                            End If
                            
                       Else
                            MeasureValue.pins(pin).value(site) = 0
                            Final_Volt_Value.pins(pin).value(site) = MeasureValue.pins(pin).value(site)
                            Final_Force_Curr_Value.pins(pin).value(site) = Adjust_I_Value.pins(pin).value(site)
                            Counter_Meas = Counter_End + 1 'Exit from loop
                       End If
                       
                        If R1_Value.pins(pin).value(site) = 0 Then
                            TheExec.Datalog.WriteComment ("Adjust " & "Site_" & site & " pin = " & pin & " R0 value = " & Format(R0_Value.pins(pin).value(site), "0.000") & " R1 value = " & "NA" & _
                                                                            " Meas Volt = " & Format(MeasureValue.pins(pin).value(site), "0.0000"))
                        Else
                            TheExec.Datalog.WriteComment ("Adjust " & "Site_" & site & " pin = " & pin & " R0 value = " & Format(R0_Value.pins(pin).value(site), "0.000") & " R1 value = " & Format(R1_Value.pins(pin).value(site), "0.000") & _
                                                                            " Meas Volt = " & Format(MeasureValue.pins(pin).value(site), "0.0000"))
                        End If
                        
                       If Counter_Meas = Counter_End Then
                            Final_Volt_Value.pins(pin).value(site) = MeasureValue.pins(pin).value(site)
                            Final_Force_Curr_Value.pins(pin).value(site) = Adjust_I_Value.pins(pin).value(site)
                       End If
                        
                    End If
                    
               Next Counter_Meas
        Next site
    Next p
    
    
    Dim TestName As String
    Dim Temp_index
    
    
    
    Temp_index = TheExec.flow.TestLimitIndex
    
    For Each pin In Final_Volt_Value.pins
        
        TheExec.flow.TestLimitIndex = Temp_index
        TestName = Report_TName_From_Instance("V", CStr(pin))
        
        For Each site In TheExec.sites
            TheExec.flow.TestLimitIndex = Temp_index
             If glb_TestInstance Like "*VOLH_SWEEP*LOOP*" Then
                 TheExec.flow.TestLimit Final_Volt_Value.pins(pin).value(site), PinName:=Final_Volt_Value.pins(pin).name, scaletype:=scaleNone, unit:=unitVolt, formatStr:="%.3f", Tname:=CStr(Temp_Input(TestSeqNum)) & "_" & CStr(TestSeqNum), ForceVal:=Final_Force_Curr_Value.pins(pin).value(site), ForceUnit:=unitAmp, ForceResults:=tlForceNone
             ElseIf glb_TestInstance Like "*VOLH_SWEEP*" Then       ' 20170912 Used for VOLH_SWEEP Average ZCAL test
                Dim ZCAL_Testname As String
ZCAL_Testname = "Average_ZCAL"
                TheExec.flow.TestLimit Final_Volt_Value.pins(pin).value(site), lowVal:=loLimit, hiVal:=hiLimit, PinName:=Final_Volt_Value.pins(pin).name, scaletype:=scaleNone, unit:=unitVolt, formatStr:="%.3f", Tname:=ZCAL_Testname & "_" & CStr(Temp_Input(TestSeqNum)) & "_" & CStr(TestSeqNum), ForceVal:=Final_Force_Curr_Value.pins(pin).value(site), ForceUnit:=unitAmp, ForceResults:=tlForceNone
             Else
               ' Mason 20210730 Remove & "_" & CStr(TestSeqNum) for Ivo request
               ' TheExec.Flow.TestLimit Final_Volt_Value.Pins(pin).value(Site), lowVal:=LoLimit, hiVal:=HiLimit, PinName:=Final_Volt_Value.Pins(pin).Name, scaletype:=scaleNone, unit:=unitVolt, formatStr:="%.3f", Tname:=TestName & "_" & CStr(TestSeqNum), forceVal:=Final_Force_Curr_Value.Pins(pin).value(Site), ForceUnit:=unitAmp, ForceResults:=tlForceFlow
               'TheExec.Flow.TestLimit Final_Volt_Value.Pins(pin).value(site), lowVal:=LoLimit, hiVal:=HiLimit, PinName:=Final_Volt_Value.Pins(pin).Name, scaletype:=scaleNone, unit:=unitVolt, formatStr:="%.3f", Tname:=TestName, ForceVal:=Final_Force_Curr_Value.Pins(pin).value(site), ForceUnit:=unitAmp, ForceResults:=tlForceFlow
               TheExec.flow.TestLimit Final_Volt_Value.pins(pin).value(site), PinName:=Final_Volt_Value.pins(pin).name, scaletype:=scaleNone, unit:=unitVolt, formatStr:="%.3f", Tname:=TestName, ForceVal:=Final_Force_Curr_Value.pins(pin).value(site), ForceUnit:=unitAmp, ForceResults:=tlForceFlow
               'TheExec.Flow.TestLimit Final_Volt_Value.Pins(pin).Value(site), lowval:=LoLimit, hival:=HiLimit, PinName:=Final_Volt_Value.Pins(pin).name, ScaleType:=scaleNone, Unit:=unitVolt, FormatStr:="%.3f", TName:=CStr(Temp_Input(TestSeqNum)) & "_" & CStr(TestSeqNum), ForceVal:=Final_Force_Curr_Value.Pins(pin).Value(site), ForceUnit:=unitAmp, ForceResults:=tlForceNone
             End If
        Next site
    Next pin

    '20171204 Print Out Final Pull Up / Pull Down resistance
    Dim Res_Tname As String: Res_Tname = vbNullString
    Dim Target_Resistance_Hi_Lim As Double: Target_Resistance_Hi_Lim = Target_Resistance * 1.1
    Dim Target_Resistance_Lo_Lim As Double: Target_Resistance_Lo_Lim = Target_Resistance * 0.9
    
    'Output Datalog PU/PD Info
    TheExec.Datalog.WriteComment (vbNullString)
    
    
    'TheExec.Flow.TestLimitIndex = TheExec.Flow.TestLimitIndex - 1
    Temp_index = TheExec.flow.TestLimitIndex
    For Each pin In Final_Volt_Value.pins
        
        TheExec.flow.TestLimitIndex = Temp_index
        
        For Each site In TheExec.sites
            TheExec.flow.TestLimitIndex = Temp_index
            If mode = "VOH" Then
                Res_Tname = Report_TName_From_Instance(CalcR, CStr(pin), "RPullUp")
                'Res_Tname = "R_Pull_Up"
            ElseIf mode = "VOL" Then
                Res_Tname = Report_TName_From_Instance(CalcR, CStr(pin), "RPullDown")
                'Res_Tname = "R_Pull_Down"
            End If
             
            If UCase(glb_TestInstance) Like "*VOLH_SWEEP*LOOP*" Then
                TheExec.flow.TestLimit R1_Value.pins(pin).value(site), PinName:=Final_Volt_Value.pins(pin).name, scaletype:=scaleNone, unit:=unitCustom, customUnit:="Ohm", formatStr:="%.3f", Tname:=Res_Tname & "_" & CStr(TestSeqNum), ForceVal:=Final_Force_Curr_Value.pins(pin).value(site), ForceUnit:=unitAmp, ForceResults:=tlForceNone
            ElseIf UCase(glb_TestInstance) Like "*VOLH_SWEEP*" Then
               TheExec.flow.TestLimit R1_Value.pins(pin).value(site), lowVal:=Target_Resistance_Lo_Lim, hiVal:=Target_Resistance_Hi_Lim, PinName:=Final_Volt_Value.pins(pin).name, scaletype:=scaleNone, unit:=unitCustom, customUnit:="Ohm", formatStr:="%.3f", Tname:=ZCAL_Testname & "_" & Res_Tname & "_" & CStr(TestSeqNum), ForceVal:=Final_Force_Curr_Value.pins(pin).value(site), ForceUnit:=unitAmp, ForceResults:=tlForceNone
            Else
               'TheExec.Flow.TestLimit R1_Value.Pins(pin).value(site), lowVal:=Target_Resistance_Lo_Lim, hiVal:=Target_Resistance_Hi_Lim, PinName:=Final_Volt_Value.Pins(pin).Name, scaletype:=scaleNone, unit:=unitCustom, customUnit:="Ohm", formatStr:="%.3f", Tname:=Res_Tname, ForceVal:=Final_Force_Curr_Value.Pins(pin).value(site), ForceUnit:=unitAmp, ForceResults:=tlForceNone
               'TheExec.Flow.TestLimit R1_Value.Pins(pin).Value(site), lowval:=Target_Resistance_Lo_Lim, hival:=Target_Resistance_Hi_Lim, PinName:=Final_Volt_Value.Pins(pin).name, ScaleType:=scaleNone, Unit:=unitCustom, customUnit:="Ohm", FormatStr:="%.3f", TName:=Res_Tname & "_" & CStr(TestSeqNum), ForceVal:=Final_Force_Curr_Value.Pins(pin).Value(site), ForceUnit:=unitAmp, ForceResults:=tlForceNone
                TheExec.flow.TestLimit R1_Value.pins(pin).value(site), PinName:=Final_Volt_Value.pins(pin).name, scaletype:=scaleNone, unit:=unitCustom, customUnit:="Ohm", formatStr:="%.3f", Tname:=Res_Tname, ForceVal:=Final_Force_Curr_Value.pins(pin).value(site), ForceUnit:=unitAmp, ForceResults:=tlForceFlow
            End If
        Next site
    Next pin
    
    'TheExec.Flow.TestLimitIndex = TheExec.Flow.TestLimitIndex + 1

    Exit Function
err:
    TheExec.Datalog.WriteComment "<Error> " + "CUS_DDR_Emulate_Const_Res_Loading" + ":: please check it out."
    If AbortTest Then Exit Function Else Resume Next
End Function
Public Function CUS_DDR_DCS_PrintOut() As Long
On Error GoTo errHandler 'Add ErrHandler 2023/05/29

    TheExec.Datalog.WriteComment (vbNullString)
    TheExec.Datalog.WriteComment (" VDD_DCS_DDR Level = " & FormatNumber(TheHdw.DCVS.pins("VDD_DCS_DDR").Voltage.value, 3) & " v")
    TheExec.Datalog.WriteComment (vbNullString)

Exit Function 'Add ErrHandler 2023/05/29
errHandler: 'Add ErrHandler 2023/05/29
    Call Print_Error_Message(Error_Info, "VBT_LIB_HardIP_Customize", "CUS_DDR_DCS_PrintOut") 'Add ErrHandler 2023/05/29
    If AbortTest Then Exit Function Else Resume Next 'Add ErrHandler 2023/05/29
End Function
Public Function MEAS_I_ABS(MeasureValue As PinListData) As Long
On Error GoTo errHandler 'Add ErrHandler 2023/05/29

Dim p As Double
For p = 0 To MeasureValue.pins.Count - 1
    MeasureValue.pins(p) = MeasureValue.pins(p).Abs
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
    
    Ary_str(0) = "_VDDQL_DDR_VAR"
    Call HIP_Evaluate_ForceVal(Ary_str)
    VDDQL_Val = CDbl(Ary_str(0))

    For p = 0 To MeasureVolt.pins.Count - 1 Step 1
          pin_name = MeasureVolt.pins(p)
          
          For Each site In TheExec.sites.Active
                MeasV = MeasureVolt.pins(pin_name).value(site)
                'RakV = TheHdw.PPMU.ReadRakValuesByPinnames(pin_name, site)
                GetRakVal = CurrentJob_Card_RAK.pins(pin_name).value(site)
                MeasureVolt.pins(pin_name).value(site) = MeasV - GetRakVal * (VDDQL_Val - MeasV) / 240
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
    
    Dict_Freq_PLD = GetStoreDataAllType("Freq_SDLL_SWP")
    
    Dim p As Long
     
     For p = 0 To MeasFreq.pins.Count - 1
         If InStr(UCase(MeasFreq.pins(p)), "_P") Then
            pin_name = MeasFreq.pins(p).name
         
             For Dict_idx = 0 To Dict_Freq_PLD.pins.Count - 1
                 If (pin_name = Dict_Freq_PLD.pins(Dict_idx).name) Then
                     Step_Freq_PLD.AddPin (pin_name)
                     
                     For Each site In TheExec.sites.Active
                         If MeasFreq.pins(p).value(site) = 0 Or Dict_Freq_PLD.pins(Dict_idx).value(site) = 0 Then
                            Step_Freq_PLD.pins(pin_name).value(site) = 999
                            TheExec.Datalog.WriteComment (" Freq measurement = 0 , Denominator = 0  ! ")
                         Else
                            Step_Freq_PLD.pins(pin_name).value(site) = 0.5 * ((1 / MeasFreq.pins(p).value(site)) - (1 / Dict_Freq_PLD.pins(Dict_idx).value(site)))
                         End If
                     Next site
                     
                     Exit For
                 End If
             Next Dict_idx
         End If
     Next p

     Extra_TName_StrArr = Split(Extra_TName, ":")
              
    For Step_idx = 0 To Step_Freq_PLD.pins.Count - 1
            pin_name = Step_Freq_PLD.pins(Step_idx).name
            pin_name = Replace(LCase(pin_name), "ddr", "ch")
            pin_name = Replace(LCase(pin_name), "dqs_p", "core")
            
            '' Extra_TName_StrArr (0) : Frequency
            '' Extra_TName_StrArr (1) : Octant
            '' Extra_TName_StrArr (2) : Loop_Idx
            '' Extra_TName_StrArr (3) : Sweep_Name (LSW_0X or MSW_0X)
           TestNameInput = Report_TName_From_Instance("F", Step_Freq_PLD.pins(Step_idx), "Step_SDLL_SWP", 0, Step_idx)
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
        If TheExec.sites(m_site).Active = True And TheExec.sites(S_site).Active = True Then
            If TheExec.sites.item(m_site).FlagState(Flag(flag_idx)) = logicTrue Then
                TheExec.sites.item(S_site).FlagState(Flag(flag_idx)) = logicTrue
            End If
            If TheExec.sites.item(S_site).FlagState(Flag(flag_idx)) = logicTrue Then
                TheExec.sites.item(m_site).FlagState(Flag(flag_idx)) = logicTrue
            End If
        End If
    Next
Next
    Exit Function
errHandler:
    Call Print_Error_Message(Error_Info, "VBT_LIB_HardIP_Customize", "D2D_Flag_Sync_Master_Slave") 'Add ErrHandler 2023/05/29
    If AbortTest Then Exit Function Else Resume Next 'Add ErrHandler 2023/05/29
End Function

Public Function Aciphy_Switch(SW_POWER_Pin As PinList, POWER_Value As Double, SW_SEL_Pin As PinList, SEL_Value As Double, Restore As Boolean)
'Hidra
'SW_3_3V = 1.8V & SEL = L (0V)    =>>  A connect B
'SW_3_3V = 1.8V & SEL = H (1.8V)  =>>  A connect C
On Error GoTo errHandler
    If Restore = False Then
        If UCase(gl_GetInstrument_Dic(LCase(SW_POWER_Pin))) = "HEXVS" Or UCase(gl_GetInstrument_Dic(LCase(SW_POWER_Pin))) = "VHDVS" Or UCase(gl_GetInstrument_Dic(LCase(SW_POWER_Pin))) = "VS-800MA" Or UCase(gl_GetInstrument_Dic(LCase(SW_POWER_Pin))) = "VS-5A" Then
            If TheHdw.DCVS.pins(SW_POWER_Pin).Gate = False Then
                With TheHdw.DCVS.pins(SW_POWER_Pin)
                    .Alarm = tlAlarmOff
                    .Disconnect tlDCVSConnectDefault
                    .Meter.mode = tlDCVSMeterCurrent
                    .mode = tlDCVSModeVoltage
                    '.SetCurrentRanges pc_Def_UVS256HP_Init_MeasCurrRange, pc_Def_UVS256HP_Init_MeasCurrRange
                    .SetCurrentRanges .CurrentRange.Max, .CurrentRange.Max
                    .Voltage.value = 0#
                    .Connect tlDCVSConnectDefault
                    .Gate = True
                End With
            Else
                With TheHdw.DCVS.pins(SW_POWER_Pin)
                    .SetCurrentRanges .CurrentRange.Max, .CurrentRange.Max
                End With
            End If
            '20240131 Hidra by YM
            TheHdw.Wait 0.001
            TheHdw.DCVS.pins(SW_POWER_Pin).Voltage.Main.value = CDbl(POWER_Value)
            TheHdw.DCVS.pins(SW_POWER_Pin).Alarm = tlAlarmDefault
            TheExec.Datalog.WriteComment SW_POWER_Pin & ":V:" & POWER_Value
        End If
        
        If UCase(gl_GetInstrument_Dic(LCase(SW_SEL_Pin))) = "HSD-U" Or UCase(gl_GetInstrument_Dic(LCase(SW_SEL_Pin))) = "HSDP" Then
            If TheHdw.PPMU.pins(SW_SEL_Pin).Gate = tlOff Then
                '=============== UP1600/UP2200 ==================
                TheHdw.Digital.pins(SW_SEL_Pin).Disconnect
                TheHdw.Wait 0.001
                TheHdw.PPMU.pins(SW_SEL_Pin).ForceV (CDbl(SEL_Value)), 0.05
                TheHdw.PPMU.pins(SW_SEL_Pin).Connect
                TheHdw.PPMU.pins(SW_SEL_Pin).Gate = tlOn
                TheExec.Datalog.WriteComment SW_SEL_Pin & ":V:" & SEL_Value
            Else
                TheHdw.PPMU.pins(SW_SEL_Pin).ForceV (CDbl(SEL_Value)), 0.05
                TheExec.Datalog.WriteComment SW_SEL_Pin & ":V:" & SEL_Value
            End If
        End If
    Else
        If UCase(gl_GetInstrument_Dic(LCase(SW_POWER_Pin))) = "HEXVS" Or UCase(gl_GetInstrument_Dic(LCase(SW_POWER_Pin))) = "VHDVS" Or UCase(gl_GetInstrument_Dic(LCase(SW_POWER_Pin))) = "VS-800MA" Or UCase(gl_GetInstrument_Dic(LCase(SW_POWER_Pin))) = "VS-5A" Then
            With TheHdw.DCVS.pins(SW_POWER_Pin)
                    .mode = tlDCVSModeHighImpedance   'Fix Mode alarm issue by cs 20201020
                    .Voltage.value = 0
                    .Gate = False
                    .Meter.mode = tlDCVSMeterVoltage '20201006 CT add to fix MTRGR CZ error  "ERROR DCVS:0074 : DIB connect at force current mode is not allowed. Please set to high impedance or force voltage mode prior DIB connect. "
                    .Disconnect
            End With
        End If
        
        If UCase(gl_GetInstrument_Dic(LCase(SW_SEL_Pin))) = "HSD-U" Or UCase(gl_GetInstrument_Dic(LCase(SW_SEL_Pin))) = "HSDP" Then
            TheHdw.PPMU.pins(SW_SEL_Pin).ForceV 0, 0.05
            TheHdw.PPMU.pins(SW_SEL_Pin).Gate = tlOff
            TheHdw.PPMU.pins(SW_SEL_Pin).Disconnect
            TheHdw.Digital.pins(SW_SEL_Pin).Connect
        End If
        TheExec.Datalog.WriteComment SW_POWER_Pin & ":DisConnectDCVS" & ";" & SW_SEL_Pin & ":DisConnectPPMU;" & SW_SEL_Pin & ":ConnectDigital"
    End If
    
    Exit Function
errHandler:
    Call Print_Error_Message(Error_Info, "VBT_LIB_HardIP_Customize", "Aciphy_Switch") 'Add ErrHandler 2023/05/29
    If AbortTest Then Exit Function Else Resume Next 'Add ErrHandler 2023/05/29
End Function

Public Function Create_Pseudo_Dictionary(Dictionary As String, Bit_Size As String, Code_Value As String, Optional Flag As String)
On Error GoTo errHandler
Dim i As Long
Dim Dic_ary() As String: Dic_ary = Split(Dictionary, ";")
Dim Bit_ary() As String: Bit_ary = Split(Bit_Size, ";")
Dim Code_Value_ary() As String: Code_Value_ary = Split(Code_Value, ";")
Dim Flag_ary() As String

'Set Flag false for eFuse Prewrite
If Flag <> "" Then
    Flag_ary = Split(Flag, ";")
    For i = 0 To UBound(Flag_ary)
        For Each site In TheExec.sites
            TheExec.sites.item(site).FlagState(Flag_ary(i)) = logicFalse
        Next site
    Next i
End If

If UBound(Dic_ary) <> UBound(Bit_ary) Or UBound(Dic_ary) <> UBound(Code_Value_ary) Then
    TheExec.Datalog.WriteComment "Dictionary Count not equal to Bit_Size Count or Dictionary Count not equal to Code Value Count"
    GoTo errHandler
End If
    
For i = 0 To UBound(Dic_ary)
    Dim DSP_Dec As New DSPWave: DSP_Dec.CreateConstant CInt(Code_Value_ary(i)), 1, DspLong
    Dim DSP_Bin As New DSPWave: DSP_Bin.CreateConstant 0, CInt(Bit_ary(i)), DspLong
    Call rundsp.DSPWf_Dec2Binary(DSP_Dec, CInt(Bit_ary(i)), DSP_Bin)
    Call StoreDataAllType(Dic_ary(i), DSP_Bin)
    Set DSP_Dec = Nothing
    Set DSP_Bin = Nothing
Next i

Exit Function
errHandler:
    Call Print_Error_Message(Error_Info, "VBT_LIB_HardIP_Customize", "Create_Pesudo_Dictionary") 'Add ErrHandler 2023/05/29
    If AbortTest Then Exit Function Else Resume Next 'Add ErrHandler 2023/05/29
End Function
