Attribute VB_Name = "LIB_HardIP_ADC_AP"
Public Function ADC_BandGap_DigCapDataProcessByDSP(pat As String, srcPin As PinList, code() As SiteLong, ByRef Res() As SiteDouble, TrimCodeSize As Long, TestSequenceNumber As Long, ByRef Rtn_DigCapData() As PinListData, DigSrc_Sample_Size As String, DigSrc_Equation As String, DigSrc_Assignment As String, TrimStoreName() As String, DigCap_Pin As PinList, DigCap_DataWidth As Long, DigCap_Sample_Size As Long, CUS_Str_MainProgram As String, CUS_Str_DigCapData As String, width_Wf As DSPWave)
    Dim srcWave() As New DSPWave: ReDim srcWave(UBound(code))
    Dim site As Variant
    Dim InDSPwave As New DSPWave
    Dim i As Long, j As Long
    Dim FlowTestNme() As String
    Dim HighLimitVal() As Double, LowLimitVal() As Double
    Call GetFlowSingleUseLimit(HighLimitVal, LowLimitVal)
    Dim srcwave_array() As Long: ReDim srcwave_array(TrimCodeSize - 1)
    Dim OutDspWave As New DSPWave
    Dim OutBinWf As New DSPWave

    ByPassTestLimit = True
    glb_Disable_CurrRangeSetting_Print = True

    For i = 0 To UBound(code)
    srcWave(i).CreateConstant 0, TrimCodeSize, DspLong
        For Each site In TheExec.sites
            For j = 0 To TrimCodeSize - 1
                If j = 0 Then
                    srcwave_array(j) = code(i) And 1
                Else
                    srcwave_array(j) = (code(i) And (2 ^ j)) \ (2 ^ j)
                End If
            Next j
        srcWave(i).data = srcwave_array
        Next site
        Call StoreDataAllType(TrimStoreName(i), srcWave(i))
    Next i
    
    Call GeneralDigSrcSettingWithBurst(LCase(pat), srcPin, InDSPwave)
  
    TheHdw.patterns(pat).Load
    
    Set OutDspWave = Nothing
    Call GeneralDigCapSetting(pat, DigCap_Pin, DigCap_Sample_Size, OutDspWave)
    Call PrintDigCapSetting(DigCap_Pin, DigCap_Sample_Size, CUS_Str_DigCapData)
    Call TheHdw.patterns(pat).test(pfNever, 0)
    If glb_Power_down_Flag = True Then              ''Add for fix Tunm shift when shut down site  @ Brian 20231221
        For Each site In TheExec.sites
            TheExec.sites.item(site).TestNumber = TheExec.sites.item(site).TestNumber
        Next site
    glb_Power_down_Flag = False                 ''Add for fix Tunm shift when shut down site  @ Brian 20231221
    End If
        'Fixed for testnumber shift issue for pattern execute --20221115
        For Each site In TheExec.sites
                TheExec.sites.item(site).TestNumber = TheExec.sites.item(site).TestNumber - 1
        Next site
    Call CreateSimulateDataDSPWave(OutDspWave, Instance_Data.DigCap_Sample_Size, Instance_Data.DigCap_DataWidth)
    rundsp.Split_Dspwave OutDspWave, width_Wf, OutBinWf
    
    
        For i = 0 To TestSequenceNumber - 1
        Rtn_DigCapData(i).pins(0).value = OutBinWf.Element(i)
        Call DebugPrintFunc_PPMU(vbNullString)
        For Each site In TheExec.sites
            Res(i) = Rtn_DigCapData(i).pins(0).value
            If gl_Disable_HIP_debug_log = False Then TheExec.Datalog.WriteComment "Site " & site & ",Code " & code(i) & ", DSSC Captured Register : " & Rtn_DigCapData(i).pins(0) & ", DSSC Captured Data = " & Res(i)
        Next site
   Next i


    TheHdw.Digital.Patgen.HaltWait
    ByPassTestLimit = False
    glb_Disable_CurrRangeSetting_Print = False

    Exit Function
    
errHandler:
    TheExec.Datalog.WriteComment "error in ADC_BandGap_DigCapDataProcessByDSP"
    If AbortTest Then Exit Function Else Resume Next
End Function
'markchen

'ADC Calculate final efuse trim code after 85C trimming
'CDNS   => REFERENCE_CTRL_DIG = round(0.25*REFERENCE_CTRL_DIG_25 + 0.75*REFERENCE_CTRL_DIG_85)
'Sicily => ADC0_VREF_85C = round(0.25*ADC0_VREF_25C + 0.75*ADC0_VREF_85C_IM)

Public Function Calc_ADC_Dict_Store(argc As Integer, argv() As String) As Long

Dim DSPWave_Dict_DIG_25C As New DSPWave: DSPWave_Dict_DIG_25C = GetStoreDataAllType(argv(0))
Dim DSPWave_Dict_DIG_85C As New DSPWave: DSPWave_Dict_DIG_85C = GetStoreDataAllType(argv(1))
Dim ADC_Trim_Code_DIG_25C As New DSPWave: ADC_Trim_Code_DIG_25C.CreateConstant 0, 1, DspLong
Dim ADC_Trim_Code_DIG_85C As New DSPWave: ADC_Trim_Code_DIG_85C.CreateConstant 0, 1, DspLong
Dim ADC_Trim_Code_DIG_sum As New DSPWave: ADC_Trim_Code_DIG_sum.CreateConstant 0, 1, DspLong
Dim ADC_Trim_Code_DIG_final As New DSPWave
Dim ADC_Trim_Code_DIG_final_Array() As Double
Dim ADC_Trim_Code_DIG_25C_Array() As Long
Dim ADC_Trim_Code_DIG_85C_Array() As Long
Dim ADC_Trim_Code_DIG_sum_Array() As Long
Dim eFuse_CTRL_DIG As New DSPWave
Dim eFuse_CTRL_DIG_Array() As Double
Dim bit_num As Long: bit_num = argv(3) '20230531 add bit_num
Dim Fuse_REFERENCE_CTRL_DIG_Name As String: Fuse_REFERENCE_CTRL_DIG_Name = argv(2)

Dim site As Variant

Call HardIP_Bin2Dec(ADC_Trim_Code_DIG_25C, DSPWave_Dict_DIG_25C)
Call HardIP_Bin2Dec(ADC_Trim_Code_DIG_85C, DSPWave_Dict_DIG_85C)


For Each site In TheExec.sites.Active
    ADC_Trim_Code_DIG_25C_Array = ADC_Trim_Code_DIG_25C.data
    ADC_Trim_Code_DIG_85C_Array = ADC_Trim_Code_DIG_85C.data
    ADC_Trim_Code_DIG_sum_Array = ADC_Trim_Code_DIG_sum.data
    ADC_Trim_Code_DIG_sum_Array(0) = FormatNumber(ADC_Trim_Code_DIG_25C_Array(0) * 0.25 + ADC_Trim_Code_DIG_85C_Array(0) * 0.75, 0)
'    ADC_Trim_Code_DIG_sum_Array(site).Element(0) = FormatNumber(ADC_Trim_Code_DIG_25C(site).Element(0) * 0.25 + ADC_Trim_Code_DIG_85C(site).Element(0) * 0.75, 0)
'        Call HardIP_Dec2Bin(ADC_Trim_Code_DIG_final, ADC_Trim_Code_DIG_sum, 8)
        
        If InStr(UCase(argv(0)), UCase("ADC0")) <> 0 Then
            TheExec.Datalog.WriteComment "site " & site & " ADC0_Trim_Code_25C :" & ADC_Trim_Code_DIG_25C_Array(0)
            TheExec.Datalog.WriteComment "site " & site & " ADC0_Trim_Code_85C :" & ADC_Trim_Code_DIG_85C_Array(0)
            TheExec.Datalog.WriteComment "site " & site & " ADC0_Trim_Code_sum :" & ADC_Trim_Code_DIG_sum_Array(0)
            
         ElseIf InStr(UCase(argv(0)), UCase("ADC1")) <> 0 Then
            TheExec.Datalog.WriteComment "site " & site & " ADC1_Trim_Code_25C :" & ADC_Trim_Code_DIG_25C_Array(0)
            TheExec.Datalog.WriteComment "site " & site & " ADC1_Trim_Code_85C :" & ADC_Trim_Code_DIG_85C_Array(0)
            TheExec.Datalog.WriteComment "site " & site & " ADC1_Trim_Code_sum :" & ADC_Trim_Code_DIG_sum_Array(0)
        
         ElseIf InStr(UCase(argv(0)), UCase("ADC2")) <> 0 Then
            TheExec.Datalog.WriteComment "site " & site & " ADC2_Trim_Code_25C :" & ADC_Trim_Code_DIG_25C_Array(0)
            TheExec.Datalog.WriteComment "site " & site & " ADC2_Trim_Code_85C :" & ADC_Trim_Code_DIG_85C_Array(0)
            TheExec.Datalog.WriteComment "site " & site & " ADC2_Trim_Code_sum :" & ADC_Trim_Code_DIG_sum_Array(0)
        Else
                'From T-Col
                TheExec.Datalog.WriteComment "site " & site & " ," & argv(0) & ":" & ADC_Trim_Code_DIG_25C_Array(0)
                TheExec.Datalog.WriteComment "site " & site & " ," & argv(1) & ":" & ADC_Trim_Code_DIG_85C_Array(0)
                TheExec.Datalog.WriteComment "site " & site & " ," & argv(2) & ":" & ADC_Trim_Code_DIG_sum_Array(0)
    
        End If
    ADC_Trim_Code_DIG_sum.data = ADC_Trim_Code_DIG_sum_Array
Next site
Call HardIP_Dec2Bin(ADC_Trim_Code_DIG_final, ADC_Trim_Code_DIG_sum, CLng(argv(3)))

' Dim Data_Temp As String
Dim final_Bin2_Str1(7) As String
Dim final_Bin2_Str As String
Dim efuse_REFERENCE_CTRL_DIG_Str1(7) As String
Dim efuse_REFERENCE_CTRL_DIG_Str As String
Dim i As Integer
For Each site In TheExec.sites.Active
        For i = 0 To CLng(argv(3)) - 1
           ' Data_Temp = Data_Temp & (ADC_Trim_Code_DIG_final(site).Element(i))
             ADC_Trim_Code_DIG_final_Array = ADC_Trim_Code_DIG_final.data
             final_Bin2_Str1(i) = CStr(ADC_Trim_Code_DIG_final_Array(i))
                                             
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
        For i = 0 To CLng(argv(3)) - 1
          eFuse_CTRL_DIG_Array = eFuse_CTRL_DIG.data
          efuse_REFERENCE_CTRL_DIG_Str1(i) = CStr(eFuse_CTRL_DIG_Array(i))
                                             
        Next i
        efuse_REFERENCE_CTRL_DIG_Str = Join(efuse_REFERENCE_CTRL_DIG_Str1, vbNullString)
        
        If InStr(UCase(argv(0)), UCase("ADC0")) <> 0 Then
          TheExec.Datalog.WriteComment "site " & site & " Fuse ADC0_VREF_85C :" & efuse_REFERENCE_CTRL_DIG_Str
        ElseIf InStr(UCase(argv(0)), UCase("ADC1")) <> 0 Then
          TheExec.Datalog.WriteComment "site " & site & " Fuse ADC1_VREF_85C :" & efuse_REFERENCE_CTRL_DIG_Str
        ElseIf InStr(UCase(argv(0)), UCase("ADC2")) <> 0 Then
          TheExec.Datalog.WriteComment "site " & site & " Fuse ADC2_VREF_85C :" & efuse_REFERENCE_CTRL_DIG_Str
        Else
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

Public Function Calc_ADC_REF_Dict_Store(argc As Integer, argv() As String) As Long

        '20211224 for Ibiza ADC @CWCIOU
        Dim DSPWave_Dict_DIG_25C As New DSPWave: DSPWave_Dict_DIG_25C = GetStoreDataAllType(argv(0))

        Dim ADC_Trim_Code_DIG_25C As New DSPWave: ADC_Trim_Code_DIG_25C.CreateConstant 0, 1, DspLong

        Dim ADC_Trim_Code_DIG_sum As New DSPWave: ADC_Trim_Code_DIG_sum.CreateConstant 0, 1, DspLong
        Dim ADC_Trim_Code_DIG_final As New DSPWave
        Dim ADC_Trim_Code_DIG_final_Array() As Double
        Dim ADC_Trim_Code_DIG_25C_Array() As Long
        Dim bit_num As Long: bit_num = argv(2)

        Dim ADC_Trim_Code_DIG_sum_Array() As Long
        Dim eFuse_CTRL_DIG As New DSPWave
        Dim eFuse_CTRL_DIG_Array() As Double

        Dim Fuse_REFERENCE_CTRL_DIG_Name As String: Fuse_REFERENCE_CTRL_DIG_Name = argv(1)

        Dim site As Variant
        Dim s_ErrorStatement As String
        
        Call HardIP_Bin2Dec(ADC_Trim_Code_DIG_25C, DSPWave_Dict_DIG_25C)
        'Call HardIP_Bin2Dec(ADC_Trim_Code_DIG_85C, DSPWave_Dict_DIG_85C)


        For Each site In TheExec.sites.Active
                ADC_Trim_Code_DIG_25C_Array = ADC_Trim_Code_DIG_25C.data
                ADC_Trim_Code_DIG_sum_Array = ADC_Trim_Code_DIG_sum.data
                If ADC_Trim_Code_DIG_25C_Array(0) = 0 Then
                        s_ErrorStatement = "warning : division by zero, ADC_Trim_Code_DIG_sum_array=0, Please check with customer!!!"
                        Call Print_Error_Message(Error_Info, "LIB_HardIP_ADC_AP", "Calc_ADC_REF_Dict_Store", s_ErrorStatement)
                Else
                        ADC_Trim_Code_DIG_sum_Array(0) = FormatNumber(2.7962 * (10) ^ 6 / ADC_Trim_Code_DIG_25C_Array(0) - 1333.33, 0)
                End If  '   ADC_Trim_Code_DIG_sum_Array(site).Element(0) = FormatNumber(ADC_Trim_Code_DIG_25C(site).Element(0) * 0.25 + ADC_Trim_Code_DIG_85C(site).Element(0) * 0.75, 0)
        '   Call HardIP_Dec2Bin(ADC_Trim_Code_DIG_final, ADC_Trim_Code_DIG_sum, 8)
         
                TheExec.Datalog.WriteComment "site " & site & ", " & argv(0) & "_dec:" & ADC_Trim_Code_DIG_25C_Array(0)
                
                'Add for prevent run time error issue when execute Dec2Bin -- 20220615
                If ADC_Trim_Code_DIG_sum_Array(0) < 0 Then
                        ADC_Trim_Code_DIG_sum_Array(0) = 0
                        TheExec.Datalog.WriteComment "site " & site & ", " & argv(1) & " Decimal value is Negative!! Please check the calculate result !! "
                Else
                        TheExec.Datalog.WriteComment "site " & site & ", " & argv(1) & "_dec:" & ADC_Trim_Code_DIG_sum_Array(0)
                End If
        
                TheExec.Datalog.WriteComment "site " & site & ", " & argv(1) & "_dec:" & ADC_Trim_Code_DIG_sum_Array(0)
         
                ADC_Trim_Code_DIG_sum.data = ADC_Trim_Code_DIG_sum_Array
        Next site
        Call HardIP_Dec2Bin(ADC_Trim_Code_DIG_final, ADC_Trim_Code_DIG_sum, bit_num)

        ' Dim Data_Temp As String
        Dim final_Bin2_Str1() As String
        ReDim final_Bin2_Str1(bit_num - 1)
        Dim final_Bin2_Str As String
        Dim efuse_REFERENCE_CTRL_DIG_Str1() As String
        ReDim efuse_REFERENCE_CTRL_DIG_Str1(bit_num - 1)
        Dim efuse_REFERENCE_CTRL_DIG_Str As String
        Dim i As Integer
        For Each site In TheExec.sites.Active
                        For i = 0 To bit_num - 1
                           ' Data_Temp = Data_Temp & (ADC_Trim_Code_DIG_final(site).Element(i))
                                 ADC_Trim_Code_DIG_final_Array = ADC_Trim_Code_DIG_final.data
                                 final_Bin2_Str1(i) = CStr(ADC_Trim_Code_DIG_final_Array(i))
                                                                                                 
                        Next i
                        final_Bin2_Str = Join(final_Bin2_Str1, vbNullString)
                        
                        TheExec.Datalog.WriteComment "site " & site & ", " & argv(1) & " _binary:" & final_Bin2_Str

                          
                        final_Bin2_Str = vbNullString
                   ' Data_Temp = ""
        Next site

        Call StoreDataAllType(Fuse_REFERENCE_CTRL_DIG_Name, ADC_Trim_Code_DIG_final)
        TheExec.Datalog.WriteComment ("DigCap data store in dictionary " & "<<" & Fuse_REFERENCE_CTRL_DIG_Name & ">>")
           
        TestNameInput = Report_TName_From_Instance(CalcC, vbNullString)
        TheExec.flow.TestLimit resultVal:=ADC_Trim_Code_DIG_sum.Element(0), Tname:=TestNameInput, ForceResults:=tlForceFlow

End Function
Public Function Reg_Assign_Processing_MTRICTS_BPBN(ByRef DigSrc_Equation As String, ByRef DigSrc_Assignment As String, ByRef CUS_Str_DigCapData As String, ByRef Calc_Eqn As String) As Long
    Dim i As Long
    Dim SplitInstance() As String
    Dim ReCombineAssign() As String
    Dim CombineAssignTemp As String
    On Error GoTo errHandler:
    
    If UCase(DigSrc_Equation) Like "*REG_ASSIGN*" Then
        DigSrc_Equation = RegAssign_PatBurstReConstruction(DigSrc_Equation, "Equation_")

    End If
        
    If UCase(DigSrc_Assignment) Like "*REG_ASSIGN*" Then
        DigSrc_Assignment = RegAssign_PatBurstReConstruction(DigSrc_Assignment, "Assignment_")

    End If
    
    If UCase(CUS_Str_DigCapData) Like "*REG_ASSIGN*" Then
        SplitInstance = Split(CUS_Str_DigCapData, ":")
        CUS_Str_DigCapData = Replace(CUS_Str_DigCapData, SplitInstance(0) & ":" & SplitInstance(1), RegDict(LCase("DigCapData_" & SplitInstance(1))))
    End If
    
    If UCase(Calc_Eqn) Like "*REG_ASSIGN*" Then
        SplitInstance = Split(Calc_Eqn, ":")
        Calc_Eqn = RegDict(LCase("Calc_Eqn_" & SplitInstance(1)))
    End If
        
'    If mid(CUS_Str_DigSrcData, 1, 1) = "+" Then
'        DigSrc_Equation = DigSrc_Equation + CUS_Str_DigSrcData
'        CUS_Str_DigSrcData = ""
'    ElseIf mid(CUS_Str_DigSrcData, 1, 1) = "-" Then
'        Calc_Eqn = Calc_Eqn + mid(CUS_Str_DigSrcData, 2)
'        CUS_Str_DigSrcData = ""
'    ElseIf mid(CUS_Str_DigSrcData, 1, 1) = "," Then
'        CUS_Str_DigCapData = CUS_Str_DigCapData + CUS_Str_DigSrcData
'        CUS_Str_DigSrcData = ""
'    End If
    
Exit Function

errHandler:
    Call Print_Error_Message(Error_Info, "LIB_HardIP", "Reg_Assign_Processing_MTRICTS_BPBN")
    If AbortTest Then Exit Function Else Resume Next
End Function
Public Function ICTS_DAP_PN_TRIM_DigCapDataProcessByDSP(pat As String, srcPin As PinList, DAC_bp() As SiteLong, DAC_bn() As SiteLong, ByRef Res() As SiteDouble, TrimCodeSize As Long, TestSequenceNumber As Long, ByRef Rtn_DigCapData() As PinListData, DigSrc_Sample_Size As String, DigSrc_Equation As String, DigSrc_Assignment As String, TrimStoreName_Array_BP() As String, TrimStoreName_Array_BN() As String, DigCap_Pin As PinList, Cap_Trimwidth As Long, DigCap_Sample_Size As Long, CUS_Str_MainProgram As String, CUS_Str_DigCapData As String, width_Wf As DSPWave, width_Wf_2S As DSPWave, ByRef OutWf_2s As DSPWave)
    Dim srcWave_bp() As New DSPWave: ReDim srcWave_bp(UBound(DAC_bp))
    Dim srcWave_bn() As New DSPWave: ReDim srcWave_bn(UBound(DAC_bn))
    Dim capWave_bp() As New DSPWave: ReDim capWave_bp(UBound(DAC_bp))
    Dim capWave_bn() As New DSPWave: ReDim capWave_bn(UBound(DAC_bn))
    Dim site As Variant
    Dim InDSPwave As New DSPWave
    Dim i As Integer, j As Integer
    Dim FlowTestNme() As String
    Dim HighLimitVal() As Double, LowLimitVal() As Double
    Call GetFlowSingleUseLimit(HighLimitVal, LowLimitVal)
    Dim srcwave_array_bp() As Long: ReDim srcwave_array_bp(TrimCodeSize - 1)
    Dim srcwave_array_bn() As Long: ReDim srcwave_array_bn(TrimCodeSize - 1)
    Dim OutDspWave As New DSPWave
    Dim OutBinWf As New DSPWave
    Dim TestOutDspWave As New DSPWave
    Dim OutWf As New DSPWave
    Dim DecWave_bp() As New DSPWave: ReDim DecWave_bp(UBound(DAC_bp))
    Dim DecWave_bn() As New DSPWave: ReDim DecWave_bn(UBound(DAC_bn))
'    Dim trim_Wf As New DSPWave: trim_Wf.CreateConstant 0, (TestSequenceNumber / 2 - 1)
    On Error GoTo errHandler
    
    ByPassTestLimit = True
    glb_Disable_CurrRangeSetting_Print = True

    For i = 0 To UBound(DAC_bp)
    srcWave_bp(i).CreateConstant 0, TrimCodeSize, DspLong
    srcWave_bn(i).CreateConstant 0, TrimCodeSize, DspLong
    DecWave_bp(i).CreateConstant 0, 1, DspLong
    DecWave_bn(i).CreateConstant 0, 1, DspLong
    


        For Each site In TheExec.sites
        DecWave_bp(i).Element(0) = DAC_bp(i)
        DecWave_bn(i).Element(0) = DAC_bn(i)
            For j = 0 To TrimCodeSize - 1
                If j = 0 Then
                    srcwave_array_bp(j) = DAC_bp(i) And 1
                    srcwave_array_bn(j) = DAC_bn(i) And 1
                Else
                    srcwave_array_bp(j) = (DAC_bp(i) And (2 ^ j)) \ (2 ^ j)
                    srcwave_array_bn(j) = (DAC_bn(i) And (2 ^ j)) \ (2 ^ j)
                End If
            Next j
         
        srcWave_bp(i).data = srcwave_array_bp
        srcWave_bn(i).data = srcwave_array_bn
        Next site
        Call StoreDataAllType(TrimStoreName_Array_BP(i), srcWave_bp(i))
        Call StoreDataAllType(TrimStoreName_Array_BP(i) & "_para", DecWave_bp(i))
        Call StoreDataAllType(TrimStoreName_Array_BN(i), srcWave_bn(i))
        Call StoreDataAllType(TrimStoreName_Array_BN(i) & "_para", DecWave_bn(i))
        Next i
        
    Call GeneralDigSrcSettingWithBurst(LCase(pat), srcPin, InDSPwave)
  


    TheHdw.patterns(pat).Load

    Set OutDspWave = Nothing
    Set TestOutDspWave = Nothing
    Call GeneralDigCapSetting(pat, DigCap_Pin, DigCap_Sample_Size, OutDspWave)
    Call PrintDigCapSetting(DigCap_Pin, DigCap_Sample_Size, CUS_Str_DigCapData)
    Call TheHdw.patterns(pat).test(pfNever, 0)
    Call CreateSimulateDataDSPWave(OutDspWave, Instance_Data.DigCap_Sample_Size, Cap_Trimwidth)
'    TestOutDspWave.CreateConstant 0, DigCap_Sample_Size, DspLong
    For Each site In TheExec.sites
    TestOutDspWave = OutDspWave.COPY
    Next site
    rundsp.Split_Dspwave OutDspWave, width_Wf, OutBinWf
'    theexec.Flow.TestLimit OutDspWave.Element(0), , ForceResults:=tlForceNone
'    rundsp.Split_2SComplementDSPWave_To_SignDec OutDspWave, width_Wf_2S, OutWf_2s
    rundsp.Split_2SComplementDSPWave_To_SignDec TestOutDspWave, width_Wf_2S, OutWf_2s
    

    For i = 0 To TestSequenceNumber - 1
            Rtn_DigCapData(i).pins(0).value = OutBinWf.Element(i)
            Call DebugPrintFunc_PPMU("")
            For Each site In TheExec.sites
                Res(i) = Rtn_DigCapData(i).pins(0).value
                If gl_Disable_HIP_debug_log = False Then
                    If i Mod 2 = 0 Then
                        TheExec.Datalog.WriteComment "Site " & site & ", Source Code DAC_BP:" & DAC_bp(Floor(i / 2)) & ", DSSC Captured Register : " & Rtn_DigCapData(i).pins(0) & ", DSSC Captured Data = " & Res(i)
                    Else
                        TheExec.Datalog.WriteComment "Site " & site & ", Source Code DAC_BP:" & DAC_bp(Floor(i / 2)) & ", DSSC Captured Register : " & Rtn_DigCapData(i).pins(0) & ", DSSC Captured Data = " & Res(i)
                    End If
                End If
            Next site
    Next i


    TheHdw.Digital.Patgen.HaltWait
    ByPassTestLimit = False
    glb_Disable_CurrRangeSetting_Print = False

    Exit Function
    
errHandler:
    TheExec.Datalog.WriteComment "error in ICTS_DAP_PN_TRIM_DigCapDataProcessByDSP"
    If AbortTest Then Exit Function Else Resume Next
End Function

Public Function ICTS_DAP_PN_AutoTRIM_DigCapDataProcessByDSP(pat As String, srcPin As PinList, DAC_bp() As SiteLong, DAC_bn() As SiteLong, ByRef Res() As SiteDouble, TrimCodeSize As Long, TestSequenceNumber As Long, ByRef Rtn_DigCapData() As PinListData, DigSrc_Sample_Size As String, DigSrc_Equation As String, DigSrc_Assignment As String, TrimStoreName_Array_BP() As String, TrimStoreName_Array_BN() As String, DigCap_Pin As PinList, Cap_Trimwidth As Long, DigCap_Sample_Size As Long, CUS_Str_MainProgram As String, CUS_Str_DigCapData As String, width_Wf As DSPWave, width_Wf_2S As DSPWave, ByRef OutWf_2s As DSPWave, ByRef OutBinWf As DSPWave, DecomposeParseDigCapBit() As String)
    Dim srcWave_bp() As New DSPWave: ReDim srcWave_bp(UBound(DAC_bp))
    Dim srcWave_bn() As New DSPWave: ReDim srcWave_bn(UBound(DAC_bn))
    Dim capWave_bp() As New DSPWave: ReDim capWave_bp(UBound(DAC_bp))
    Dim capWave_bn() As New DSPWave: ReDim capWave_bn(UBound(DAC_bn))
    Dim site As Variant
    Dim InDSPwave As New DSPWave
    Dim i As Integer, j As Integer
    Dim FlowTestNme() As String
    Dim HighLimitVal() As Double, LowLimitVal() As Double
    Call GetFlowSingleUseLimit(HighLimitVal, LowLimitVal)
    Dim srcwave_array_bp() As Long: ReDim srcwave_array_bp(TrimCodeSize - 1)
    Dim srcwave_array_bn() As Long: ReDim srcwave_array_bn(TrimCodeSize - 1)
    Dim OutDspWave As New DSPWave
    Dim TestOutDspWave As New DSPWave
    Dim DecWave_bp() As New DSPWave: ReDim DecWave_bp(UBound(DAC_bp))
    Dim DecWave_bn() As New DSPWave: ReDim DecWave_bn(UBound(DAC_bn))
'    Dim trim_Wf As New DSPWave: trim_Wf.CreateConstant 0, (TestSequenceNumber / 2 - 1)
    On Error GoTo errHandler
    
    ByPassTestLimit = True
    glb_Disable_CurrRangeSetting_Print = True

    For i = 0 To UBound(DAC_bp)
    srcWave_bp(i).CreateConstant 0, TrimCodeSize, DspLong
    srcWave_bn(i).CreateConstant 0, TrimCodeSize, DspLong
    DecWave_bp(i).CreateConstant 0, 1, DspLong
    DecWave_bn(i).CreateConstant 0, 1, DspLong
    


        For Each site In TheExec.sites
        DecWave_bp(i).Element(0) = DAC_bp(i)
        DecWave_bn(i).Element(0) = DAC_bn(i)
            For j = 0 To TrimCodeSize - 1
                If j = 0 Then
                    srcwave_array_bp(j) = DAC_bp(i) And 1
                    srcwave_array_bn(j) = DAC_bn(i) And 1
                Else
                    srcwave_array_bp(j) = (DAC_bp(i) And (2 ^ j)) \ (2 ^ j)
                    srcwave_array_bn(j) = (DAC_bn(i) And (2 ^ j)) \ (2 ^ j)
                End If
            Next j
         
        srcWave_bp(i).data = srcwave_array_bp
        srcWave_bn(i).data = srcwave_array_bn
        Next site
        Call StoreDataAllType(TrimStoreName_Array_BP(i), srcWave_bp(i))
        Call StoreDataAllType(TrimStoreName_Array_BP(i) & "_para", DecWave_bp(i))
        Call StoreDataAllType(TrimStoreName_Array_BN(i), srcWave_bn(i))
        Call StoreDataAllType(TrimStoreName_Array_BN(i) & "_para", DecWave_bn(i))
        Next i
        
    Call GeneralDigSrcSettingWithBurst(LCase(pat), srcPin, InDSPwave)
  


    TheHdw.patterns(pat).Load

    Set OutDspWave = Nothing
    Set TestOutDspWave = Nothing
    Call GeneralDigCapSetting(pat, DigCap_Pin, DigCap_Sample_Size, OutDspWave)
    Call PrintDigCapSetting(DigCap_Pin, DigCap_Sample_Size, CUS_Str_DigCapData)
    Call TheHdw.patterns(pat).test(pfNever, 0)
    Call CreateSimulateDataDSPWave(OutDspWave, Instance_Data.DigCap_Sample_Size, Cap_Trimwidth)
'    TestOutDspWave.CreateConstant 0, DigCap_Sample_Size, DspLong
    
    For Each site In TheExec.sites
        TestOutDspWave = OutDspWave.COPY
    Next site
    
    rundsp.Split_Dspwave OutDspWave, width_Wf, OutBinWf
    
  
'    theexec.Flow.TestLimit OutDspWave.Element(0), , ForceResults:=tlForceNone
'    rundsp.Split_2SComplementDSPWave_To_SignDec OutDspWave, width_Wf_2S, OutWf_2s
    rundsp.Split_2SComplementDSPWave_To_SignDec TestOutDspWave, width_Wf_2S, OutWf_2s
    
    Dim TestNameInput As String
    
    For i = 0 To UBound(DecomposeParseDigCapBit)
        If EnableDigitalTestLimitTTR = True Then
            TestNameInput = Report_TName_From_Instance("C", "X", , CInt(i), 0, , , , tlForceFlow)  ' Update for Turks 20190523
        Else
            TestNameInput = Report_TName_From_Instance("C", DigCap_PinName, , CInt(i), 0, , , , tlForceFlow)  ' Update for Turks 20190523
        End If
        TheExec.flow.TestLimit OutBinWf.Element(i), 0, 2 ^ DecomposeParseDigCapBit(i) - 1, PinName:=DigCap_Pin, Tname:=TestNameInput, ForceResults:=tlForceFlow, scaletype:=scaleNoScaling, formatStr:="%.0f"
    Next i

    For i = 0 To TestSequenceNumber - 1
        Rtn_DigCapData(i).pins(0).value = OutBinWf.Element(i)
        Call DebugPrintFunc_PPMU("")
        For Each site In TheExec.sites
            Res(i) = Rtn_DigCapData(i).pins(0).value
            If gl_Disable_HIP_debug_log = False Then
                TheExec.Datalog.WriteComment "Site " & site & ", Source Code DAC_BP:" & DAC_bp(Floor(i / 2)) & ",DAC_BN:" & DAC_bn(Floor(i / 2)) & ", DSSC Captured Register : " & Rtn_DigCapData(i).pins(0) & ", DSSC Captured Data = " & Res(i)
            End If
        Next site
    Next i
  
    TheHdw.Digital.Patgen.HaltWait
    ByPassTestLimit = False
    glb_Disable_CurrRangeSetting_Print = False

    Exit Function
    
errHandler:
    TheExec.Datalog.WriteComment "error in ICTS_DAP_PN_TRIM_DigCapDataProcessByDSP"
    If AbortTest Then Exit Function Else Resume Next
End Function
