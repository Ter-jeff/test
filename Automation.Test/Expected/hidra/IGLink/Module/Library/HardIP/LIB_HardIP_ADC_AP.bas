Attribute VB_Name = "LIB_HardIP_ADC_AP"
Public Function ADC_BandGap_DigCapDataProcessByDSP(Pat As String, srcPin As PinList, code() As SiteLong, ByRef Res() As SiteDouble, TrimCodeSize As Long, TestSequenceNumber As Long, ByRef Rtn_DigCapData() As PinListData, DigSrc_Sample_Size As String, DigSrc_Equation As String, digsrc_assignment As String, TrimStoreName() As String, DigCap_Pin As PinList, DigCap_DataWidth As Long, DigCap_Sample_Size As Long, CUS_Str_MainProgram As String, CUS_Str_DigCapData As String, width_Wf As DSPWave)
    Dim srcWave() As New DSPWave: ReDim srcWave(UBound(code))
    Dim site As Variant
    Dim InDSPWave As New DSPWave
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
    
    Call GeneralDigSrcSettingWithBurst(LCase(Pat), srcPin, InDSPWave)
  
    TheHdw.Patterns(Pat).Load
    
    Set OutDspWave = Nothing
    Call GeneralDigCapSetting(Pat, DigCap_Pin, DigCap_Sample_Size, OutDspWave)
    Call PrintDigCapSetting(DigCap_Pin, DigCap_Sample_Size, CUS_Str_DigCapData)
    Call TheHdw.Patterns(Pat).test(pfNever, 0)
        
        'Fixed for testnumber shift issue for pattern execute --20221115
        For Each site In TheExec.sites
                TheExec.sites.item(site).TestNumber = TheExec.sites.item(site).TestNumber - 1
        Next site
    Call CreateSimulateDataDSPWave(OutDspWave, Instance_Data.DigCap_Sample_Size, Instance_Data.DigCap_DataWidth)
    rundsp.Split_Dspwave OutDspWave, width_Wf, OutBinWf
    
    
        For i = 0 To TestSequenceNumber - 1
        Rtn_DigCapData(i).Pins(0).value = OutBinWf.Element(i)
        Call DebugPrintFunc_PPMU(vbNullString)
        For Each site In TheExec.sites
            Res(i) = Rtn_DigCapData(i).Pins(0).value
            If gl_Disable_HIP_debug_log = False Then TheExec.Datalog.WriteComment "Site " & site & ",Code " & code(i) & ", DSSC Captured Register : " & Rtn_DigCapData(i).Pins(0) & ", DSSC Captured Data = " & Res(i)
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

End Function
