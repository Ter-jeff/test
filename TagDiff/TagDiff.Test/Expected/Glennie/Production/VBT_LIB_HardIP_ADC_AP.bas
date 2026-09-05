Attribute VB_Name = "VBT_LIB_HardIP_ADC_AP"
Public Function ADC_BandGap_Calibration(Optional pat As Pattern, Optional DigCap_Pin As PinList, Optional DigCap_DataWidth As Long, Optional DigCap_Sample_Size As Long, Optional DigSrc_pin As PinList, Optional DigSrc_Sample_Size As String, Optional DigSrc_Equation As String, Optional DigSrc_Assignment As String, Optional CUS_Str_MainProgram As String = vbNullString, Optional CUS_Str_DigCapData As String = vbNullString, Optional TrimStoreName As String, Optional TrimTarget As Double, Optional TrimStart As Long, Optional TrimCodeSize As Long, Optional Calc_Eqn As String, Optional Interpose_PrePat As String, Optional TrimMethod As String, Optional Trimming_Direction_Increase As Boolean, Optional Validating_ As Boolean)
On Error GoTo errHandler 'Add ErrHandler 2023/05/29
    
    glb_TestInstance = vbNullString
    glb_TestInstance = UCase(TheExec.DataManager.instancename)
    
    If Validating_ Then
        Call PrLoadPattern(pat.value)
        Exit Function    ' Exit after validation
    End If

    
    Dim site As Variant
    Dim i As Integer
    Dim j As Integer
    Dim k As Integer
    Dim pats() As String
    Dim DSSC_Out_DecompseByComma() As String: DSSC_Out_DecompseByComma = Split(CUS_Str_DigCapData, ",")
    Dim DSSC_Out_DecompseByColon() As String
    Dim ParseStringByBits As String: ParseStringByBits = vbNullString
    Dim ParseStringForTestName As String: ParseStringForTestName = vbNullString
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
    Dim DecomposeTestName() As String: DecomposeTestName = Split(ParseStringForTestName, ",")
    Dim DecomposeParseDigCapBit() As String: DecomposeParseDigCapBit = Split(ParseStringByBits, ",")
    
    Dim TestSequenceNumber As Long: TestSequenceNumber = UBound(DecomposeTestName) + 1
    Dim code() As New SiteLong: ReDim code(TestSequenceNumber - 1)
    Dim BinarySearch_TrimCodeValue_Min() As New SiteLong: ReDim BinarySearch_TrimCodeValue_Min(TestSequenceNumber - 1)
    Dim BinarySearch_TrimCodeValue_Max() As New SiteLong: ReDim BinarySearch_TrimCodeValue_Max(TestSequenceNumber - 1)
    Dim BestCode() As New SiteLong: ReDim BestCode(TestSequenceNumber - 1)
    Dim vout() As New SiteDouble: ReDim vout(TestSequenceNumber - 1)
        For i = 0 To TestSequenceNumber - 1
            For Each site In TheExec.sites.Active
                code(i) = TrimStart
                vout(i) = 0
            Next site
        Next i
    Dim PreviousNegative() As New SiteBoolean: ReDim PreviousNegative(TestSequenceNumber - 1)
    Dim PreviousPositive() As New SiteBoolean: ReDim PreviousPositive(TestSequenceNumber - 1)
    Dim DecideTrim() As New SiteBoolean: ReDim DecideTrim(TestSequenceNumber - 1)
    Dim Trim_Flag As Boolean
        For i = 0 To TestSequenceNumber - 1
            For Each site In TheExec.sites.Active
                PreviousNegative(i) = False
                PreviousPositive(i) = False
                DecideTrim(i) = False
            Next site
        Next i
    Dim blockName() As String: blockName = Split(glb_TestInstance, "_")
    Dim MeasValue() As New PinListData: ReDim MeasValue(TestSequenceNumber - 1)
    Dim BestVal() As New PinListData: ReDim BestVal(TestSequenceNumber - 1)
    Dim stepcount As Long: stepcount = 0
    Dim TestNameInput As String
    Dim PatCount As Long, PattArray() As String
    Dim BestTargetCompare() As New SiteDouble: ReDim BestTargetCompare(TestSequenceNumber - 1)
    Dim TrimStoreName_Array() As String: TrimStoreName_Array = Split(TrimStoreName, ",")
    Dim PinName As String
    Dim TempVal As Integer
    Dim FinalTrimCode() As New DSPWave: ReDim FinalTrimCode(UBound(Split(TrimStoreName, ",")))
    Dim FinalTrimCode_Array() As Long: ReDim FinalTrimCode_Array(TrimCodeSize - 1) As Long
    Dim width_Wf  As New DSPWave: width_Wf.CreateConstant 0, UBound(DecomposeParseDigCapBit) + 1
    Dim DecomposeParseDigCapBit_long() As Long: ReDim DecomposeParseDigCapBit_long(UBound(DecomposeParseDigCapBit))
    For i = 0 To UBound(DecomposeParseDigCapBit)
        DecomposeParseDigCapBit_long(i) = CLng(DecomposeParseDigCapBit(i))  'deliver data to dsp array
    Next i
    For Each site In TheExec.sites
        width_Wf.data = DecomposeParseDigCapBit_long  'deliver data to dsp array
    Next site
    Call ProcessInputToGLB(pat, , , , , , , , , , , , , , , , , DigCap_Pin, DigCap_DataWidth, DigCap_Sample_Size, DigSrc_pin, CStr(DigSrc_Sample_Size), CStr(DigSrc_Sample_Size), DigSrc_Equation, DigSrc_Assignment, , , , CUS_Str_MainProgram)
    
    For i = 0 To UBound(DecomposeTestName)
        MeasValue(i).AddPin (DecomposeTestName(i) & "_" & i)
        MeasValue(i).pins(DecomposeTestName(i) & "_" & i).value = 0
        BestVal(i).AddPin (DecomposeTestName(i) & "_" & i)
        BestVal(i).pins(DecomposeTestName(i) & "_" & i).value = 0
    Next i

    Call GetFlowTName

    
    
    TheHdw.Digital.ApplyLevelsTiming True, True, True, tlPowered
    TheHdw.Digital.Patgen.Continue 0, cpuA + cpuB + cpuC + cpuD
    
    PATT_GetPatListFromPatternSet pat.value, pats, PatCount
        
    Dim TrimCodeValue_Min As Long, TrimCodeValue_Max As Long
    If UCase(CUS_Str_MainProgram) = "TWOSCOMPLEMENT_DSSC_SOURCE" Then
        TrimCodeValue_Min = (-1) * 2 ^ (TrimCodeSize - 1)
        TrimCodeValue_Max = 2 ^ (TrimCodeSize - 1) - 1
    Else
    TrimCodeValue_Min = 0
    TrimCodeValue_Max = 2 ^ TrimCodeSize - 1
    End If
    
     If Interpose_PrePat <> "" Then Call SetForceCondition(Interpose_PrePat & ";STOREPREPAT")
     
    If gl_Disable_HIP_debug_log = False Then TheExec.Datalog.WriteComment ("**************** The Current Trimming Method is: " & TrimMethod & "****************")
    If gl_Disable_HIP_debug_log = False Then TheExec.Datalog.WriteComment ("**************** The Current Trimming Direction is: " & Trimming_Direction_Increase & "****************")
    If gl_Disable_HIP_debug_log = False Then TheExec.Datalog.WriteComment ("**************** The DSSC Captured Data at Trim Start Point ****************")
    Call ADC_BandGap_DigCapDataProcessByDSP(pats(0), DigSrc_pin, code(), vout(), TrimCodeSize, TestSequenceNumber, MeasValue(), DigSrc_Sample_Size, DigSrc_Equation, DigSrc_Assignment, TrimStoreName_Array(), DigCap_Pin, DigCap_DataWidth, DigCap_Sample_Size, CUS_Str_MainProgram, CUS_Str_DigCapData, width_Wf)

    For Each site In TheExec.sites.Active
        For i = 0 To TestSequenceNumber - 1
            BestCode(i) = code(i)
            BestVal(i) = MeasValue(i)
            BestTargetCompare(i) = vout(i).Subtract(TrimTarget).Abs
        Next i
    Next site
    For i = 0 To TestSequenceNumber - 1
        If LCase(TrimMethod) = "linearsearch" Then
            For Each site In TheExec.sites.Active
                If Trimming_Direction_Increase = True Then
                    If vout(i).compare(GreaterThan, TrimTarget) Then
                        code(i) = code(i) - 1
                    ElseIf vout(i).compare(LessThan, TrimTarget) Then
                        code(i) = code(i) + 1
                    End If
                ElseIf Trimming_Direction_Increase = False Then
                    If vout(i).compare(LessThan, TrimTarget) Then
                        code(i) = code(i) - 1
                    ElseIf vout(i).compare(GreaterThan, TrimTarget) Then
                        code(i) = code(i) + 1
                    Else 'Do nothing '20230601
                    End If
                End If
                DecideTrim(i) = True
            Next site
        ElseIf LCase(TrimMethod) = "binarysearch" Then
            For Each site In TheExec.sites.Active
                If Trimming_Direction_Increase = True Then
                    If vout(i).compare(GreaterThan, TrimTarget) Then
                        BinarySearch_TrimCodeValue_Min(i) = TrimCodeValue_Min
                        BinarySearch_TrimCodeValue_Max(i) = code(i)
                        code(i) = FormatNumber(0.5 * (code(i) + TrimCodeValue_Min), 0)
                    ElseIf vout(i).compare(LessThan, TrimTarget) Then
                        BinarySearch_TrimCodeValue_Min(i) = code(i)
                        BinarySearch_TrimCodeValue_Max(i) = TrimCodeValue_Max
                        code(i) = FormatNumber(0.5 * (code(i) + TrimCodeValue_Max), 0)
                    End If
                ElseIf Trimming_Direction_Increase = False Then
                    If vout(i).compare(LessThan, TrimTarget) Then
                        BinarySearch_TrimCodeValue_Min(i) = TrimCodeValue_Min
                        BinarySearch_TrimCodeValue_Max(i) = code(i)
                        code(i) = FormatNumber(0.5 * (code(i) + TrimCodeValue_Min), 0)
                    ElseIf vout(i).compare(GreaterThan, TrimTarget) Then
                        BinarySearch_TrimCodeValue_Min(i) = code(i)
                        BinarySearch_TrimCodeValue_Max(i) = TrimCodeValue_Max
                        code(i) = FormatNumber(0.5 * (code(i) + TrimCodeValue_Max), 0)
                    End If
                End If
                DecideTrim(i) = True
            Next site
        Else
            Call Print_Error_Message(Warning_Info, "VBT_LIB_HardIP_ADC_AP", "ADC_BandGap_Calibration", "TrimMethod type misalignment.")
            Exit Function
        End If
    Next i
StartTrim:
    For i = 0 To TestSequenceNumber - 1
        If i = 0 Then
            Trim_Flag = DecideTrim(i).Any(True)
        Else
            Trim_Flag = Trim_Flag Or DecideTrim(i).Any(True)
        End If
    Next i
        If Trim_Flag Then
            stepcount = stepcount + 1
            
            If gl_Disable_HIP_debug_log = False Then
                If right(CStr(stepcount), 1) = "1" Then
                    TheExec.Datalog.WriteComment ("**************** The " & stepcount & "st Trim Process ****************")
                ElseIf right(CStr(stepcount), 1) = "2" Then
                    TheExec.Datalog.WriteComment ("**************** The " & stepcount & "nd Trim Process ****************")
                ElseIf right(CStr(stepcount), 1) = "3" Then
                    TheExec.Datalog.WriteComment ("**************** The " & stepcount & "rd Trim Process ****************")
                Else
                    TheExec.Datalog.WriteComment ("**************** The " & stepcount & "th Trim Process ****************")
                End If
            End If
                Call ADC_BandGap_DigCapDataProcessByDSP(pats(0), DigSrc_pin, code(), vout(), TrimCodeSize, TestSequenceNumber, MeasValue(), DigSrc_Sample_Size, DigSrc_Equation, DigSrc_Assignment, TrimStoreName_Array(), DigCap_Pin, DigCap_DataWidth, DigCap_Sample_Size, CUS_Str_MainProgram, CUS_Str_DigCapData, width_Wf)
        End If
    For k = 0 To TestSequenceNumber - 1
        For Each site In TheExec.sites.Active
            If DecideTrim(k) And vout(k).Subtract(TrimTarget).Abs < BestTargetCompare(k) Then
                BestTargetCompare(k) = vout(k).Subtract(TrimTarget).Abs
                BestCode(k) = code(k)
                BestVal(k).pins(0).value = MeasValue(k).pins(0).value
            End If
            If DecideTrim(k) And LCase(TrimMethod) = "linearsearch" Then
                If Trimming_Direction_Increase = True Then
                    If stepcount > (TrimCodeValue_Max - TrimCodeValue_Min) Then
                        DecideTrim(k) = False
                    ElseIf code(k).compare(GreaterThan, TrimCodeValue_Max) Then
                        code(k) = TrimCodeValue_Max
                        DecideTrim(k) = True
                    ElseIf code(k).compare(LessThan, TrimCodeValue_Min) Then
                        code(k) = TrimCodeValue_Min
                        DecideTrim(k) = True
                    ElseIf code(k).compare(EqualTo, TrimCodeValue_Max) Or code(k).compare(EqualTo, TrimCodeValue_Min) Then
                        DecideTrim(k) = False
                    ElseIf vout(k).compare(GreaterThan, TrimTarget) And PreviousPositive(k) Then
                        DecideTrim(k) = False
                    ElseIf vout(k).compare(GreaterThan, TrimTarget) And Not (PreviousPositive(k)) Then
                        code(k) = code(k) - 1
                        PreviousNegative(k) = True
                        DecideTrim(k) = True
                    ElseIf vout(k).compare(LessThan, TrimTarget) And PreviousNegative(k) Then
                        DecideTrim(k) = False
                    ElseIf vout(k).compare(LessThan, TrimTarget) And Not (PreviousNegative(k)) Then
                        code(k) = code(k) + 1
                        PreviousPositive(k) = True
                        DecideTrim(k) = True
                    ElseIf vout(k).compare(EqualTo, TrimTarget) Then
                        DecideTrim(k) = False
                    End If
                ElseIf Trimming_Direction_Increase = False Then
                    If stepcount > (TrimCodeValue_Max - TrimCodeValue_Min) Then
                        DecideTrim(k) = False
                    ElseIf code(k).compare(GreaterThan, TrimCodeValue_Max) Then
                        code(k) = TrimCodeValue_Max
                        DecideTrim(k) = True
                    ElseIf code(k).compare(LessThan, TrimCodeValue_Min) Then
                        code(k) = TrimCodeValue_Min
                        DecideTrim(k) = True
                    ElseIf code(k).compare(EqualTo, TrimCodeValue_Max) Or code(k).compare(EqualTo, TrimCodeValue_Min) Then
                        DecideTrim(k) = False
                    ElseIf vout(k).compare(LessThan, TrimTarget) And PreviousPositive(k) Then
                        DecideTrim(k) = False
                    ElseIf vout(k).compare(LessThan, TrimTarget) And Not (PreviousPositive(k)) Then
                        code(k) = code(k) - 1
                        PreviousNegative(k) = True
                        DecideTrim(k) = True
                    ElseIf vout(k).compare(GreaterThan, TrimTarget) And PreviousNegative(k) Then
                        DecideTrim(k) = False
                    ElseIf vout(k).compare(GreaterThan, TrimTarget) And Not (PreviousNegative(k)) Then
                        code(k) = code(k) + 1
                        PreviousPositive(k) = True
                        DecideTrim(k) = True
                    ElseIf vout(k).compare(EqualTo, TrimTarget) Then
                        DecideTrim(k) = False
                    End If
                End If
            ElseIf DecideTrim(k) And LCase(TrimMethod) = "binarysearch" Then
                If Trimming_Direction_Increase = True Then
                    If stepcount > (TrimCodeValue_Max - TrimCodeValue_Min) Then
                        DecideTrim(k) = False
                    ElseIf code(k).compare(GreaterThan, TrimCodeValue_Max) Then
                        code(k) = TrimCodeValue_Max
                        DecideTrim(k) = True
                    ElseIf code(k).compare(LessThan, TrimCodeValue_Min) Then
                        code(k) = TrimCodeValue_Min
                        DecideTrim(k) = True
                    ElseIf code(k).compare(EqualTo, BinarySearch_TrimCodeValue_Max(k)) Or code(k).compare(EqualTo, BinarySearch_TrimCodeValue_Min(k)) Then
                        DecideTrim(k) = False
                    ElseIf vout(k).compare(GreaterThan, TrimTarget) Then
                        If code(k) = FormatNumber(0.5 * (code(k) + BinarySearch_TrimCodeValue_Min(k)), 0) Then
                            DecideTrim(k) = False
                        Else
                            BinarySearch_TrimCodeValue_Max(k) = code(k)
                            code(k) = FormatNumber(0.5 * (code(k) + BinarySearch_TrimCodeValue_Min(k)), 0)
                            DecideTrim(k) = True
                        End If
                    ElseIf vout(k).compare(LessThan, TrimTarget) Then
                        If code(k) = FormatNumber(0.5 * (code(k) + BinarySearch_TrimCodeValue_Max(k)), 0) Then
                            DecideTrim(k) = False
                        Else
                            BinarySearch_TrimCodeValue_Min(k) = code(k)
                            code(k) = FormatNumber(0.5 * (code(k) + BinarySearch_TrimCodeValue_Max(k)), 0)
                            DecideTrim(k) = True
                        End If
                    ElseIf vout(k).compare(EqualTo, TrimTarget) Then
                        DecideTrim(k) = False
                    End If
                ElseIf Trimming_Direction_Increase = False Then
                    If stepcount > (TrimCodeValue_Max - TrimCodeValue_Min) Then
                        DecideTrim(k) = False
                    ElseIf code(k).compare(GreaterThan, TrimCodeValue_Max) Then
                        code(k) = TrimCodeValue_Max
                        DecideTrim(k) = True
                    ElseIf code(k).compare(LessThan, TrimCodeValue_Min) Then
                        code(k) = TrimCodeValue_Min
                        DecideTrim(k) = True
                    ElseIf code(k).compare(EqualTo, BinarySearch_TrimCodeValue_Max(k)) Or code(k).compare(EqualTo, BinarySearch_TrimCodeValue_Min(k)) Then
                        DecideTrim(k) = False
                    ElseIf vout(k).compare(LessThan, TrimTarget) Then
                        If code(k) = FormatNumber(0.5 * (code(k) + BinarySearch_TrimCodeValue_Min(k)), 0) Then
                            DecideTrim(k) = False
                        Else
                            BinarySearch_TrimCodeValue_Max(k) = code(k)
                            code(k) = FormatNumber(0.5 * (code(k) + BinarySearch_TrimCodeValue_Min(k)), 0)
                            DecideTrim(k) = True
                        End If
                    ElseIf vout(k).compare(GreaterThan, TrimTarget) Then
                        If code(k) = FormatNumber(0.5 * (code(k) + BinarySearch_TrimCodeValue_Max(k)), 0) Then
                            DecideTrim(k) = False
                        Else
                            BinarySearch_TrimCodeValue_Min(k) = code(k)
                            code(k) = FormatNumber(0.5 * (code(k) + BinarySearch_TrimCodeValue_Max(k)), 0)
                            DecideTrim(k) = True
                        End If
                    ElseIf vout(k).compare(EqualTo, TrimTarget) Then
                        DecideTrim(k) = False
                    Else 'Do nothing '20230601
                    End If
                End If
            End If
        Next site
    Next k
    
    For i = 0 To TestSequenceNumber - 1
        If i = 0 Then
            Trim_Flag = DecideTrim(i).Any(True)
        Else
            Trim_Flag = Trim_Flag Or DecideTrim(i).Any(True)
        End If
    Next i
    
    If Trim_Flag Then GoTo StartTrim

     If Interpose_PrePat <> "" Then Call SetForceCondition("RESTOREPREPAT")

    Dim TempBestVal As New SiteDouble
    
    For i = 0 To TestSequenceNumber - 1
    
        TempBestVal = BestVal(i).pins(DecomposeTestName(i) & "_" & i)
        TestNameInput = Report_TName_From_Instance("C", DigCap_Pin.value, vbNullString, i, 0)
        TheExec.flow.TestLimit TempBestVal, 0, 2 ^ DecomposeParseDigCapBit(i) - 1, PinName:=DigCap_Pin.value, Tname:=TestNameInput, ForceResults:=tlForceFlow, scaletype:=scaleNoScaling, formatStr:="%.0f"
        
    Next i
    For i = 0 To TestSequenceNumber - 1
        TestNameInput = Report_TName_From_Instance("C", vbNullString, blockName(0) & "Trim", i, 0)
        TheExec.flow.TestLimit resultVal:=BestCode(i), Tname:=TestNameInput, ForceResults:=tlForceFlow
    Next i
    
        ' Modify for Ibiza ADC_TRIMT2 item -- 20220116
    Dim ADC_TRIM_ENG() As New DSPWave: ReDim ADC_TRIM_ENG(UBound(Split(TrimStoreName, ",")))
    Dim TempBestVal_buf As New SiteDouble
    
     For i = 0 To TestSequenceNumber - 1
         For Each site In TheExec.sites
             If UCase(CUS_Str_MainProgram) = "TWOSCOMPLEMENT_DSSC_SOURCE" And BestCode(i) < 0 Then
                 TempVal = 2 ^ TrimCodeSize + BestCode(i)
             Else
                 TempVal = BestCode(i)
             End If
             For j = 0 To TrimCodeSize - 1
                 FinalTrimCode_Array(j) = TempVal Mod 2
                 TempVal = TempVal \ 2
             Next j
             FinalTrimCode(i).data = FinalTrimCode_Array
            If gl_Disable_HIP_debug_log = False Then                    'Printing for store info
                TheExec.Datalog.WriteComment "Site : " & site & ", Store Value : " & BestCode(i) & ", Binary Bits : " & TrimCodeSize & ", Store Name : " & TrimStoreName_Array(i)
            End If
         Next site
         Call StoreDataAllType(TrimStoreName_Array(i), FinalTrimCode(i))
         
         Call StoreDataAllType(TrimStoreName_Array(i) & "_para", BestCode(i))
         ' Add for Ibiza ADC_TRIMT2 to store Best Value for Calculate -- 20220116
        TempBestVal_buf = BestVal(i).pins(DecomposeTestName(i) & "_" & i)
         For Each site In TheExec.sites
            ADC_TRIM_ENG(i).CreateConstant 0, Instance_Data.DigCap_DataWidth, DspDouble
            For j = 0 To Instance_Data.DigCap_DataWidth - 1
                ADC_TRIM_ENG(i).Element(j) = TempBestVal_buf(site) Mod 2
                TempBestVal_buf(site) = TempBestVal_buf(site) \ 2
            Next j
        Next site
        Call StoreDataAllType(TrimStoreName_Array(i) & "_target", ADC_TRIM_ENG(i))
     Next i
         
    DebugPrintFunc pat.value
    
    Call HardIP_WriteFuncResult(, , glb_TestInstance)
    
    If Calc_Eqn <> "" Then: Call ProcessCalcEquation(Calc_Eqn)
    
    ' Check implicit alarms
    TheHdw.Alarms.Check

    Exit Function
    
errHandler:
    Call Print_Error_Message(Error_Info, "VBT_LIB_HardIP_ADC_AP", "ADC_BandGap_Calibration") 'Add ErrHandler 2023/05/29
    If AbortTest Then Exit Function Else Resume Next
End Function
Public Function ADC_BandGap_Calibration_Reverse(Optional pat As Pattern, Optional DigCap_Pin As PinList, Optional DigCap_DataWidth As Long, Optional DigCap_Sample_Size As Long, Optional DigSrc_pin As PinList, Optional DigSrc_Sample_Size As String, Optional DigSrc_Equation As String, Optional DigSrc_Assignment As String, Optional CUS_Str_MainProgram As String = vbNullString, Optional CUS_Str_DigCapData As String = vbNullString, Optional TrimStoreName As String, Optional TrimTarget As Double, Optional TrimStart As Long, Optional TrimCodeSize As Long, Optional Calc_Eqn As String, Optional Interpose_PrePat As String, Optional TrimMethod As String, Optional Validating_ As Boolean)
    'From T-Col
    
    glb_TestInstance = vbNullString
    glb_TestInstance = UCase(TheExec.DataManager.instancename)
    
    If Validating_ Then
        Call PrLoadPattern(pat.value)
        Exit Function    ' Exit after validation
    End If
    
    On Error GoTo errHandler
    
    Dim b_TnameConstructionTTR As Boolean
    b_TnameConstructionTTR = True
    If EnableHardIPTnameConstructionTTR = True Then
        EnableHardIPTnameConstructionTTR = False
        b_TnameConstructionTTR = False
    End If
    
    Dim site As Variant
    Dim i As Integer
    Dim k As Integer
    Dim pats() As String
    Dim DSSC_Out_DecompseByComma() As String: DSSC_Out_DecompseByComma = Split(CUS_Str_DigCapData, ",")
    Dim DSSC_Out_DecompseByColon() As String
    Dim ParseStringByBits As String: ParseStringByBits = vbNullString
    Dim ParseStringForTestName As String: ParseStringForTestName = vbNullString
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
    Dim DecomposeTestName() As String: DecomposeTestName = Split(ParseStringForTestName, ",")
    Dim DecomposeParseDigCapBit() As String: DecomposeParseDigCapBit = Split(ParseStringByBits, ",")
    
    Dim TestSequenceNumber As Long: TestSequenceNumber = UBound(DecomposeTestName) + 1
    Dim code() As New SiteLong: ReDim code(TestSequenceNumber - 1)
    Dim BinarySearch_TrimCodeValue_Min() As New SiteLong: ReDim BinarySearch_TrimCodeValue_Min(TestSequenceNumber - 1)
    Dim BinarySearch_TrimCodeValue_Max() As New SiteLong: ReDim BinarySearch_TrimCodeValue_Max(TestSequenceNumber - 1)
    Dim BestCode() As New SiteLong: ReDim BestCode(TestSequenceNumber - 1)
    Dim vout() As New SiteDouble: ReDim vout(TestSequenceNumber - 1)
        For i = 0 To TestSequenceNumber - 1
            For Each site In TheExec.sites.Active
                code(i) = TrimStart
                vout(i) = 0
            Next site
        Next i
    Dim PreviousNegative() As New SiteBoolean: ReDim PreviousNegative(TestSequenceNumber - 1)
    Dim PreviousPositive() As New SiteBoolean: ReDim PreviousPositive(TestSequenceNumber - 1)
    Dim DecideTrim() As New SiteBoolean: ReDim DecideTrim(TestSequenceNumber - 1)
    Dim Trim_Flag As Boolean
        For i = 0 To TestSequenceNumber - 1
            For Each site In TheExec.sites.Active
                PreviousNegative(i) = False
                PreviousPositive(i) = False
                DecideTrim(i) = False
            Next site
        Next i
    Dim blockName() As String: blockName = Split(glb_TestInstance, "_")
    Dim MeasValue() As New PinListData: ReDim MeasValue(TestSequenceNumber - 1)
    Dim BestVal() As New PinListData: ReDim BestVal(TestSequenceNumber - 1)
    Dim stepcount As Long: stepcount = 0
    Dim TestNameInput As String
    Dim PatCount As Long, PattArray() As String
    Dim BestTargetCompare() As New SiteDouble: ReDim BestTargetCompare(TestSequenceNumber - 1)
    Dim TrimStoreName_Array() As String: TrimStoreName_Array = Split(TrimStoreName, ",")
    Dim PinName As String
    Dim TempVal As Integer
    Dim FinalTrimCode() As New DSPWave: ReDim FinalTrimCode(UBound(Split(TrimStoreName, ",")))
    Dim FinalTrimCode_Array() As Long: ReDim FinalTrimCode_Array(TrimCodeSize - 1) As Long
    Dim width_Wf  As New DSPWave: width_Wf.CreateConstant 0, UBound(DecomposeParseDigCapBit) + 1
    Dim DecomposeParseDigCapBit_long() As Long: ReDim DecomposeParseDigCapBit_long(UBound(DecomposeParseDigCapBit))
    For i = 0 To UBound(DecomposeParseDigCapBit)
        DecomposeParseDigCapBit_long(i) = CLng(DecomposeParseDigCapBit(i))  'deliver data to dsp array
    Next i
    For Each site In TheExec.sites
        width_Wf.data = DecomposeParseDigCapBit_long  'deliver data to dsp array
    Next site
    Call ProcessInputToGLB(pat, , , , , , , , , , , , , , , , , DigCap_Pin, DigCap_DataWidth, DigCap_Sample_Size, DigSrc_pin, CStr(DigSrc_Sample_Size), CStr(DigSrc_Sample_Size), DigSrc_Equation, DigSrc_Assignment, , , , CUS_Str_MainProgram)
    
    For i = 0 To UBound(DecomposeTestName)
        MeasValue(i).AddPin (DecomposeTestName(i) & "_" & i)
        MeasValue(i).pins(DecomposeTestName(i) & "_" & i).value = 0
        BestVal(i).AddPin (DecomposeTestName(i) & "_" & i)
        BestVal(i).pins(DecomposeTestName(i) & "_" & i).value = 0
    Next i

    Call GetFlowTName
    
    TheHdw.Digital.ApplyLevelsTiming True, True, True, tlPowered
    TheHdw.Digital.Patgen.Continue 0, cpuA + cpuB + cpuC + cpuD
    
    PATT_GetPatListFromPatternSet pat.value, pats, PatCount
        
    Dim TrimCodeValue_Min As Long, TrimCodeValue_Max As Long
    If UCase(CUS_Str_MainProgram) = "TWOSCOMPLEMENT_DSSC_SOURCE" Then
        TrimCodeValue_Min = (-1) * 2 ^ (TrimCodeSize - 1)
        TrimCodeValue_Max = 2 ^ (TrimCodeSize - 1) - 1
    Else
        TrimCodeValue_Min = 0
        TrimCodeValue_Max = 2 ^ TrimCodeSize - 1
    End If
  
    
     If Interpose_PrePat <> "" Then Call SetForceCondition(Interpose_PrePat & ";STOREPREPAT")

    If gl_Disable_HIP_debug_log = False Then TheExec.Datalog.WriteComment ("**************** The DSSC Captured Data at Trim Start Point ****************")
    Call ADC_BandGap_DigCapDataProcessByDSP(pats(0), DigSrc_pin, code(), vout(), TrimCodeSize, TestSequenceNumber, MeasValue(), DigSrc_Sample_Size, DigSrc_Equation, DigSrc_Assignment, TrimStoreName_Array(), DigCap_Pin, DigCap_DataWidth, DigCap_Sample_Size, CUS_Str_MainProgram, CUS_Str_DigCapData, width_Wf)

    For Each site In TheExec.sites.Active
        For i = 0 To TestSequenceNumber - 1
            BestCode(i) = code(i)
            BestVal(i) = MeasValue(i)
            BestTargetCompare(i) = vout(i).Subtract(TrimTarget).Abs
        Next i
    Next site
    For i = 0 To TestSequenceNumber - 1
        If LCase(TrimMethod) = "linearsearch" Then
            For Each site In TheExec.sites.Active
                If vout(i).compare(GreaterThan, TrimTarget) Then
                    code(i) = code(i) - 1
                ElseIf vout(i).compare(LessThan, TrimTarget) Then
                    code(i) = code(i) + 1
                End If
                DecideTrim(i) = True
            Next site
        ElseIf LCase(TrimMethod) = "binarysearch" Then
            For Each site In TheExec.sites.Active
                If vout(i).compare(GreaterThan, TrimTarget) Then
                    BinarySearch_TrimCodeValue_Min(i) = TrimCodeValue_Min
                    BinarySearch_TrimCodeValue_Max(i) = code(i)
                    code(i) = FormatNumber(0.5 * (code(i) + TrimCodeValue_Min), 0)
                ElseIf vout(i).compare(LessThan, TrimTarget) Then
                    BinarySearch_TrimCodeValue_Min(i) = code(i)
                    BinarySearch_TrimCodeValue_Max(i) = TrimCodeValue_Max
                    code(i) = FormatNumber(0.5 * (code(i) + TrimCodeValue_Max), 0)
                End If
                DecideTrim(i) = True
            Next site
        End If
    Next i
StartTrim:
    For i = 0 To TestSequenceNumber - 1
        If i = 0 Then
            Trim_Flag = DecideTrim(i).Any(True)
        Else
            Trim_Flag = Trim_Flag Or DecideTrim(i).Any(True)
        End If
    Next i
        If Trim_Flag Then
            stepcount = stepcount + 1
            
            If gl_Disable_HIP_debug_log = False Then
                If right(CStr(stepcount), 1) = "1" Then
                    TheExec.Datalog.WriteComment ("**************** The " & stepcount & "st Trim Process ****************")
                ElseIf right(CStr(stepcount), 1) = "2" Then
                    TheExec.Datalog.WriteComment ("**************** The " & stepcount & "nd Trim Process ****************")
                ElseIf right(CStr(stepcount), 1) = "3" Then
                    TheExec.Datalog.WriteComment ("**************** The " & stepcount & "rd Trim Process ****************")
                Else
                    TheExec.Datalog.WriteComment ("**************** The " & stepcount & "th Trim Process ****************")
                End If
            End If
                Call ADC_BandGap_DigCapDataProcessByDSP(pats(0), DigSrc_pin, code(), vout(), TrimCodeSize, TestSequenceNumber, MeasValue(), DigSrc_Sample_Size, DigSrc_Equation, DigSrc_Assignment, TrimStoreName_Array(), DigCap_Pin, DigCap_DataWidth, DigCap_Sample_Size, CUS_Str_MainProgram, CUS_Str_DigCapData, width_Wf)
        End If
    For k = 0 To TestSequenceNumber - 1
        For Each site In TheExec.sites.Active
            If DecideTrim(k) And vout(k).Subtract(TrimTarget).Abs < BestTargetCompare(k) Then
                BestTargetCompare(k) = vout(k).Subtract(TrimTarget).Abs
                BestCode(k) = code(k)
                BestVal(k).pins(0).value = MeasValue(k).pins(0).value
            End If
            If DecideTrim(k) And LCase(TrimMethod) = "linearsearch" Then
                If stepcount > (TrimCodeValue_Max - TrimCodeValue_Min) Then
                    DecideTrim(k) = False
                ElseIf code(k).compare(GreaterThan, TrimCodeValue_Max) Then
                    code(k) = TrimCodeValue_Max
                    DecideTrim(k) = True
                ElseIf code(k).compare(LessThan, TrimCodeValue_Min) Then
                    code(k) = TrimCodeValue_Min
                    DecideTrim(k) = True
                ElseIf code(k).compare(EqualTo, TrimCodeValue_Max) Or code(k).compare(EqualTo, TrimCodeValue_Min) Then
                    DecideTrim(k) = False
                ElseIf vout(k).compare(GreaterThan, TrimTarget) And PreviousPositive(k) Then
                    DecideTrim(k) = False
                ElseIf vout(k).compare(GreaterThan, TrimTarget) And Not (PreviousPositive(k)) Then
                    code(k) = code(k) - 1
                    PreviousNegative(k) = True
                    DecideTrim(k) = True
                ElseIf vout(k).compare(LessThan, TrimTarget) And PreviousNegative(k) Then
                    DecideTrim(k) = False
                ElseIf vout(k).compare(LessThan, TrimTarget) And Not (PreviousNegative(k)) Then
                    code(k) = code(k) + 1
                    PreviousPositive(k) = True
                    DecideTrim(k) = True
                ElseIf vout(k).compare(EqualTo, TrimTarget) Then
                    DecideTrim(k) = False
                End If
            ElseIf DecideTrim(k) And LCase(TrimMethod) = "binarysearch" Then
                If stepcount > (TrimCodeValue_Max - TrimCodeValue_Min) Then
                    DecideTrim(k) = False
                ElseIf code(k).compare(GreaterThan, TrimCodeValue_Max) Then
                    code(k) = TrimCodeValue_Max
                    DecideTrim(k) = True
                ElseIf code(k).compare(LessThan, TrimCodeValue_Min) Then
                    code(k) = TrimCodeValue_Min
                    DecideTrim(k) = True
                ElseIf code(k).compare(EqualTo, BinarySearch_TrimCodeValue_Max(k)) Or code(k).compare(EqualTo, BinarySearch_TrimCodeValue_Min(k)) Then
                    DecideTrim(k) = False
                ElseIf vout(k).compare(GreaterThan, TrimTarget) Then
                    If code(k) = FormatNumber(0.5 * (code(k) + BinarySearch_TrimCodeValue_Min(k)), 0) Then
                        DecideTrim(k) = False
                    Else
                        BinarySearch_TrimCodeValue_Max(k) = code(k)
                        code(k) = FormatNumber(0.5 * (code(k) + BinarySearch_TrimCodeValue_Min(k)), 0)
                        DecideTrim(k) = True
                    End If
                ElseIf vout(k).compare(LessThan, TrimTarget) Then
                    If code(k) = FormatNumber(0.5 * (code(k) + BinarySearch_TrimCodeValue_Max(k)), 0) Then
                        DecideTrim(k) = False
                    Else
                        BinarySearch_TrimCodeValue_Min(k) = code(k)
                        code(k) = FormatNumber(0.5 * (code(k) + BinarySearch_TrimCodeValue_Max(k)), 0)
                        DecideTrim(k) = True
                    End If
                ElseIf vout(k).compare(EqualTo, TrimTarget) Then
                    DecideTrim(k) = False
                End If
            End If
        Next site
    Next k
    
    For i = 0 To TestSequenceNumber - 1
        If i = 0 Then
            Trim_Flag = DecideTrim(i).Any(True)
        Else
            Trim_Flag = Trim_Flag Or DecideTrim(i).Any(True)
        End If
    Next i
    
    If Trim_Flag Then GoTo StartTrim

     If Interpose_PrePat <> "" Then Call SetForceCondition("RESTOREPREPAT")

    Dim TempBestVal As New SiteDouble
    
    For i = 0 To TestSequenceNumber - 1
    
        TempBestVal = BestVal(i).pins(DecomposeTestName(i) & "_" & i)
        TestNameInput = Report_TName_From_Instance("C", DigCap_Pin.value, vbNullString, i, 0)
        TheExec.flow.TestLimit TempBestVal, 0, 2 ^ DecomposeParseDigCapBit(i) - 1, PinName:=DigCap_Pin.value, Tname:=TestNameInput, ForceResults:=tlForceFlow, scaletype:=scaleNoScaling, formatStr:="%.0f"
        
    Next i
    For i = 0 To TestSequenceNumber - 1
        TestNameInput = Report_TName_From_Instance("C", vbNullString, blockName(0) & "Trim", i, 0)
        TheExec.flow.TestLimit resultVal:=BestCode(i), Tname:=TestNameInput, ForceResults:=tlForceFlow
    Next i

   ' Modify for Ibiza ADC_TRIMT2 item -- 20220116
    Dim ADC_TRIM_ENG() As New DSPWave: ReDim ADC_TRIM_ENG(UBound(Split(TrimStoreName, ",")))
    Dim TempBestVal_buf As New SiteDouble
    
     For i = 0 To TestSequenceNumber - 1
        For Each site In TheExec.sites
             If UCase(CUS_Str_MainProgram) = "TWOSCOMPLEMENT_DSSC_SOURCE" And BestCode(i) < 0 Then
                 TempVal = 2 ^ TrimCodeSize + BestCode(i)
             Else
                 TempVal = BestCode(i)
             End If
             For j = 0 To TrimCodeSize - 1
                 FinalTrimCode_Array(j) = TempVal Mod 2
                 TempVal = TempVal \ 2
             Next j
            FinalTrimCode(i).data = FinalTrimCode_Array
        Next site
        Call AddStoredCaptureData(TrimStoreName_Array(i), FinalTrimCode(i))
        
        ' Add Store Decimal value for CalcC use for Coll -- 20220613
        Call AddStoredData(TrimStoreName_Array(i) & "_para", BestCode(i))
        
        ' Add for Ibiza ADC_TRIMT2 to store Best Value for Calculate -- 20220116
        TempBestVal_buf = BestVal(i).pins(DecomposeTestName(i) & "_" & i)
         For Each site In TheExec.sites
            ADC_TRIM_ENG(i).CreateConstant 0, Instance_Data.DigCap_DataWidth, DspDouble
            For j = 0 To Instance_Data.DigCap_DataWidth - 1
                ADC_TRIM_ENG(i).Element(j) = TempBestVal_buf(site) Mod 2
                TempBestVal_buf(site) = TempBestVal_buf(site) \ 2
            Next j
        Next site
        Call AddStoredCaptureData(TrimStoreName_Array(i) & "_target", ADC_TRIM_ENG(i))
     Next i
     
    DebugPrintFunc pat.value
    
    Call HardIP_WriteFuncResult(, , glb_TestInstance)
    
    If Calc_Eqn <> "" And InStr(LCase(TestSequence), "p") = 0 Then: Call ProcessCalcEquation(Calc_Eqn)
    
    ' Check implicit alarms
    TheHdw.Alarms.Check
    
    If b_TnameConstructionTTR = False Then
        EnableHardIPTnameConstructionTTR = True
    End If
    
    Exit Function
    
errHandler:
    TheExec.Datalog.WriteComment "error in ADC_BandGap_Calibration"
    If AbortTest Then Exit Function Else Resume Next
End Function

Public Function ADC_LDO_Calibration(Optional pat As Pattern, Optional TestSequence As String, Optional MeasV_PinS As String, Optional MeaV_WaitTime As String, Optional DigSrc_pin As PinList, _
                        Optional DigSrc_Sample_Size As String, Optional DigSrc_Equation As String, Optional DigSrc_Assignment As String, Optional CUS_Str_MainProgram As String = vbNullString, Optional TrimStoreName As String, Optional TrimTarget As Double, Optional TrimStart As Long, Optional TrimCodeSize As Long, Optional TrimMethod As String, Optional TrimStepSize As Double, Optional Validating_ As Boolean)
    
    glb_TestInstance = vbNullString
    glb_TestInstance = UCase(TheExec.DataManager.instancename)
    
    If Validating_ Then
        Call PrLoadPattern(pat.value)
        Exit Function    ' Exit after validation
    End If
    
    On Error GoTo errHandler
    
    Dim site As Variant
    Dim i As Integer
    Dim j As Integer
    Dim k As Integer
    Dim pats() As String
    Dim code() As New SiteLong: ReDim code(UBound(Split(TestSequence, ",")))
    Dim BinarySearch_TrimCodeValue_Min() As New SiteLong: ReDim BinarySearch_TrimCodeValue_Min(UBound(Split(TestSequence, ",")))
    Dim BinarySearch_TrimCodeValue_Max() As New SiteLong: ReDim BinarySearch_TrimCodeValue_Max(UBound(Split(TestSequence, ",")))
    Dim BestCode() As New SiteLong: ReDim BestCode(UBound(Split(TestSequence, ",")))
    Dim vout() As New SiteDouble: ReDim vout(UBound(Split(TestSequence, ",")))
        For i = 0 To UBound(Split(TestSequence, ","))
            For Each site In TheExec.sites.Active
                code(i) = TrimStart
                vout(i) = 0
            Next site
        Next i
    Dim NumberOfMeasV As Integer: NumberOfMeasV = UBound(Split(TestSequence, ",")) + 1
    Dim PreviousNegative() As New SiteBoolean: ReDim PreviousNegative(UBound(Split(TestSequence, ",")))
    Dim PreviousPositive() As New SiteBoolean: ReDim PreviousPositive(UBound(Split(TestSequence, ",")))
    Dim DecideTrim() As New SiteBoolean: ReDim DecideTrim(UBound(Split(TestSequence, ",")))
    Dim Trim_Flag As Boolean
        For i = 0 To UBound(Split(TestSequence, ","))
            For Each site In TheExec.sites.Active
                PreviousNegative(i) = False
                PreviousPositive(i) = False
                DecideTrim(i) = False
            Next site
        Next i
    Dim blockName() As String: blockName = Split(glb_TestInstance, "_")
    Dim MeasValue() As New PinListData: ReDim MeasValue(NumberOfMeasV - 1)
    Dim BestVal() As New PinListData: ReDim BestVal(NumberOfMeasV - 1)
    Dim stepcount As Long: stepcount = 0
    Dim TestNameInput As String
    Dim PatCount As Long, PattArray() As String
    Dim BestTargetCompare() As New SiteDouble: ReDim BestTargetCompare(UBound(Split(TestSequence, ",")))
    Dim TrimStoreName_Array() As String: TrimStoreName_Array = Split(TrimStoreName, ",")
    Dim PinName() As String
    Dim TempVal As Integer
    Dim FinalTrimCode() As New DSPWave: ReDim FinalTrimCode(UBound(TrimStoreName_Array))
    Dim FinalTrimCode_Array() As Long: ReDim FinalTrimCode_Array(TrimCodeSize - 1) As Long
    
    Call ProcessInputToGLB(pat, TestSequence, True, , , , , MeasV_PinS, , , , , , , , , , , , , DigSrc_pin, DigSrc_Sample_Size, DigSrc_Sample_Size, DigSrc_Equation, DigSrc_Assignment, , , , CUS_Str_MainProgram, , , , , , , , , , , , , , , , , , , , , , MeaV_WaitTime)


    PinName = Split(MeasV_PinS, "+")
   For i = 0 To NumberOfMeasV - 1
    
       If InStr(MeasV_PinS, "+") <> 0 Then
        ''For j = 0 To UBound(Split(MeasV_PinS, "+"))
            MeasValue(i).AddPin PinName(i)
            MeasValue(i).pins(PinName(i)).value = 0
            BestVal(i).AddPin PinName(i)
            BestVal(i).pins(PinName(i)).value = 0
        ''Next j
       Else
            MeasValue(i).AddPin (PinName(0))
            MeasValue(i).pins(PinName(0)).value = 0
            BestVal(i).AddPin (PinName(0))
            BestVal(i).pins(PinName(0)).value = 0
       End If
       
    Next i
    
    Call GetFlowTName
    
    If TheExec.DevChar.Setups.IsRunning Then
        If TheExec.DevChar.Setups(TheExec.DevChar.Setups.ActiveSetupName).Shmoo.axes.Contains(tlDevCharShmooAxis_Y) Then
            If gl_Flag_HardIP_Trim_Set_PrePoint And Not (gl_Flag_HardIP_Characterization_1stRun) Then
                TheHdw.Digital.ApplyLevelsTiming True, True, True, tlPowered
                Call TheExec.Overlays.ApplyUniformSpecToHW("XI0_Shmoo_Freq_VAR", TheExec.DevChar.results(TheExec.DevChar.Setups.ActiveSetupName).Shmoo.CurrentPoint.axes(tlDevCharShmooAxis_Y).value)
            ElseIf gl_Flag_HardIP_Trim_Set_PostPoint Then
                TheHdw.Digital.ApplyLevelsTiming True, True, True, tlPowered
                Call TheExec.Overlays.ApplyUniformSpecToHW("XI0_Shmoo_Freq_VAR", 24000000#)
            Else
                TheHdw.Digital.ApplyLevelsTiming True, True, True, tlPowered
            End If
        End If
    Else
        TheHdw.Digital.ApplyLevelsTiming True, True, True, tlPowered
    End If
    TheHdw.Digital.Patgen.Continue 0, cpuA + cpuB + cpuC + cpuD
    
    PATT_GetPatListFromPatternSet pat.value, pats, PatCount
        
    Dim TrimCodeValue_Min As Long, TrimCodeValue_Max As Long
    If UCase(CUS_Str_MainProgram) = "TWOSCOMPLEMENT_DSSC_SOURCE" Then
        TrimCodeValue_Min = (-1) * 2 ^ (TrimCodeSize - 1)
        TrimCodeValue_Max = 2 ^ (TrimCodeSize - 1) - 1
    Else
        TrimCodeValue_Min = 0
        TrimCodeValue_Max = 2 ^ TrimCodeSize - 1
    End If
  

    If gl_Disable_HIP_debug_log = False Then TheExec.Datalog.WriteComment ("**************** The Measurement at Trim Start Point ****************")
    Call LDO_Measurement_Process(pats(0), DigSrc_pin, code(), vout(), TrimCodeSize, NumberOfMeasV, MeasValue(), DigSrc_Sample_Size, DigSrc_Equation, DigSrc_Assignment, TrimStoreName_Array(), MeaV_WaitTime)

    For Each site In TheExec.sites.Active
        For i = 0 To NumberOfMeasV - 1
            BestCode(i) = code(i)
            BestVal(i) = MeasValue(i)
            BestTargetCompare(i) = vout(i).Subtract(TrimTarget).Abs
        Next i
    Next site
    For i = 0 To UBound(Split(TestSequence, ","))
        If LCase(TrimMethod) = "linearsearch" Then
            For Each site In TheExec.sites.Active
                If vout(i).compare(LessThan, TrimTarget) Then
                    code(i) = code(i) + 1
                ElseIf vout(i).compare(GreaterThan, TrimTarget) Then
                    code(i) = code(i) - 1
                End If
                DecideTrim(i) = True
            Next site
        ElseIf LCase(TrimMethod) = "binarysearch" Then
            For Each site In TheExec.sites.Active
                If vout(i).compare(LessThan, TrimTarget) Then
                    BinarySearch_TrimCodeValue_Min(i) = code(i)
                    BinarySearch_TrimCodeValue_Max(i) = TrimCodeValue_Max
                    code(i) = FormatNumber(0.5 * (code(i) + TrimCodeValue_Max), 0)
                ElseIf vout(i).compare(GreaterThan, TrimTarget) Then
                    BinarySearch_TrimCodeValue_Min(i) = TrimCodeValue_Min
                    BinarySearch_TrimCodeValue_Max(i) = code(i)
                    code(i) = FormatNumber(0.5 * (code(i) + TrimCodeValue_Min), 0)
                End If
                DecideTrim(i) = True
            Next site
        Else
            For Each site In TheExec.sites.Active
                code(i) = code(i) + Fix((TrimTarget - vout(i)) / TrimStepSize)
                DecideTrim(i) = True
            Next site
        End If
    Next i
StartTrim:
    For i = 0 To UBound(Split(TestSequence, ","))
        If i = 0 Then
            Trim_Flag = DecideTrim(i).Any(True)
        Else
            Trim_Flag = Trim_Flag Or DecideTrim(i).Any(True)
        End If
    Next i
        If Trim_Flag Then
            stepcount = stepcount + 1
            
            If gl_Disable_HIP_debug_log = False Then
                If right(CStr(stepcount), 1) = "1" Then
                    TheExec.Datalog.WriteComment ("**************** The " & stepcount & "st Trim Process ****************")
                ElseIf right(CStr(stepcount), 1) = "2" Then
                    TheExec.Datalog.WriteComment ("**************** The " & stepcount & "nd Trim Process ****************")
                ElseIf right(CStr(stepcount), 1) = "3" Then
                    TheExec.Datalog.WriteComment ("**************** The " & stepcount & "rd Trim Process ****************")
                Else
                    TheExec.Datalog.WriteComment ("**************** The " & stepcount & "th Trim Process ****************")
                End If
            End If

            Call LDO_Measurement_Process(pats(0), DigSrc_pin, code(), vout(), TrimCodeSize, NumberOfMeasV, MeasValue(), DigSrc_Sample_Size, DigSrc_Equation, DigSrc_Assignment, TrimStoreName_Array(), MeaV_WaitTime)
        End If
    For k = 0 To UBound(Split(TestSequence, ","))
        For Each site In TheExec.sites.Active
            If DecideTrim(k) And vout(k).Subtract(TrimTarget).Abs < BestTargetCompare(k) Then
                BestTargetCompare(k) = vout(k).Subtract(TrimTarget).Abs
                BestCode(k) = code(k)
                
                If InStr(MeasV_PinS, "+") <> 0 Then
                    BestVal(k).pins(PinName(k)).value = MeasValue(k).pins(PinName(k)).value
                Else
                    BestVal(k).pins(PinName(0)).value = MeasValue(k).pins(PinName(0)).value
                End If
                
            End If
            If DecideTrim(k) And LCase(TrimMethod) = "linearsearch" Then
                If stepcount > (TrimCodeValue_Max - TrimCodeValue_Min) Then
                    DecideTrim(k) = False
                ElseIf code(k).compare(GreaterThan, TrimCodeValue_Max) Then
                    code(k) = TrimCodeValue_Max
                    DecideTrim(k) = True
                ElseIf code(k).compare(LessThan, TrimCodeValue_Min) Then
                    code(k) = TrimCodeValue_Min
                    DecideTrim(k) = True
                ElseIf code(k).compare(EqualTo, TrimCodeValue_Max) Or code(k).compare(EqualTo, TrimCodeValue_Min) Then
                    DecideTrim(k) = False
                ElseIf vout(k).compare(LessThan, TrimTarget) And PreviousPositive(k) Then
                    DecideTrim(k) = False
                ElseIf vout(k).compare(LessThan, TrimTarget) And Not (PreviousPositive(k)) Then
                    code(k) = code(k) + 1
                    PreviousNegative(k) = True
                    DecideTrim(k) = True
                ElseIf vout(k).compare(GreaterThan, TrimTarget) And PreviousNegative(k) Then
                    DecideTrim(k) = False
                ElseIf vout(k).compare(GreaterThan, TrimTarget) And Not (PreviousNegative(k)) Then
                    code(k) = code(k) - 1
                    PreviousPositive(k) = True
                    DecideTrim(k) = True
                End If
            ElseIf DecideTrim(k) And LCase(TrimMethod) = "binarysearch" Then
                If stepcount > (TrimCodeValue_Max - TrimCodeValue_Min) Then
                    DecideTrim(k) = False
                ElseIf code(k).compare(GreaterThan, TrimCodeValue_Max) Then
                    code(k) = TrimCodeValue_Max
                    DecideTrim(k) = True
                ElseIf code(k).compare(LessThan, TrimCodeValue_Min) Then
                    code(k) = TrimCodeValue_Min
                    DecideTrim(k) = True
                ElseIf code(k).compare(EqualTo, BinarySearch_TrimCodeValue_Max(k)) Or code(k).compare(EqualTo, BinarySearch_TrimCodeValue_Min(k)) Then
                    DecideTrim(k) = False
                ElseIf vout(k).compare(LessThan, TrimTarget) Then
                    If code(k) = FormatNumber(0.5 * (code(k) + BinarySearch_TrimCodeValue_Max(k)), 0) Then
                        DecideTrim(k) = False
                    Else
                        BinarySearch_TrimCodeValue_Min(k) = code(k)
                        code(k) = FormatNumber(0.5 * (code(k) + BinarySearch_TrimCodeValue_Max(k)), 0)
                        DecideTrim(k) = True
                    End If
                ElseIf vout(k).compare(GreaterThan, TrimTarget) Then
                    If code(k) = FormatNumber(0.5 * (code(k) + BinarySearch_TrimCodeValue_Min(k)), 0) Then
                        DecideTrim(k) = False
                    Else
                        BinarySearch_TrimCodeValue_Max(k) = code(k)
                        code(k) = FormatNumber(0.5 * (code(k) + BinarySearch_TrimCodeValue_Min(k)), 0)
                        DecideTrim(k) = True
                    End If
                ElseIf vout(k).compare(EqualTo, TrimTarget) Then
                    DecideTrim(k) = False
                End If
            End If
        Next site
    Next k
    
    For i = 0 To UBound(Split(TestSequence, ","))
        If i = 0 Then
            Trim_Flag = DecideTrim(i).Any(True)
        Else
            Trim_Flag = Trim_Flag Or DecideTrim(i).Any(True)
        End If
    Next i
    
    If Trim_Flag Then GoTo StartTrim

        
    For i = 0 To NumberOfMeasV - 1
       
        If InStr(MeasV_PinS, "+") <> 0 Then
            
            TestNameInput = Report_TName_From_Instance("V", PinName(i), vbNullString, i, 0)
            TheExec.flow.TestLimit resultVal:=BestVal(i), Tname:=TestNameInput, PinName:=PinName(i), ForceResults:=tlForceFlow

        Else
      
            TestNameInput = Report_TName_From_Instance("V", PinName(0), vbNullString, i, 0)
            TheExec.flow.TestLimit resultVal:=BestVal(i), Tname:=TestNameInput, PinName:=PinName(0), ForceResults:=tlForceFlow
        End If
    Next i
    For i = 0 To UBound(Split(TestSequence, ","))
        TestNameInput = Report_TName_From_Instance("C", vbNullString, blockName(0) & "Trim", i, 0)
        TheExec.flow.TestLimit resultVal:=BestCode(i), Tname:=TestNameInput, ForceResults:=tlForceFlow
    Next i
    For i = 0 To UBound(Split(TestSequence, ","))
        For Each site In TheExec.sites
            If UCase(CUS_Str_MainProgram) = "TWOSCOMPLEMENT_DSSC_SOURCE" And BestCode(i) < 0 Then
                TempVal = 2 ^ TrimCodeSize + BestCode(i)
            Else
                TempVal = BestCode(i)
            End If
            For j = 0 To TrimCodeSize - 1
                FinalTrimCode_Array(j) = TempVal Mod 2
                TempVal = TempVal \ 2
            Next j
            FinalTrimCode(i).data = FinalTrimCode_Array
        Next site
        Call AddStoredCaptureData(TrimStoreName_Array(i), FinalTrimCode(i))
    Next i
    DebugPrintFunc pat.value
    
    Call HardIP_WriteFuncResult(, , glb_TestInstance)
    
    TheHdw.Alarms.Check
    Exit Function
    
errHandler:
    TheExec.Datalog.WriteComment "error in ADC_LDO_Calibration"
    If AbortTest Then Exit Function Else Resume Next
End Function

Public Function MetrologyICTS_DAP_PN_AutoTRIM(Optional pat As Pattern, Optional DigCap_Pin As PinList, Optional Cap_Trimwidth As Long, Optional DigCap_Sample_Size As Long, Optional DigSrc_pin As PinList, Optional DigSrc_Sample_Size As String, Optional DigSrc_Equation As String, Optional DigSrc_Assignment As String, Optional CUS_Str_MainProgram As String = "", Optional CUS_Str_DigCapData As String = "", Optional TrimStoreName As String, Optional TrimTarget As Double, Optional TrimStart As String, Optional TrimCodeSize As Long, Optional Calc_Eqn As String, Optional Interpose_PrePat As String, Optional TrimMethod As String, Optional Trimming_Direction_Increase As Boolean, Optional Validating_ As Boolean, Optional DataOut As String)
    
    glb_TestInstance = ""
    glb_TestInstance = UCase(TheExec.DataManager.instancename)
    
    If Validating_ Then
        Call PrLoadPattern(pat.value)
        Exit Function    ' Exit after validation
    End If
    
    On Error GoTo errHandler
    
    Dim b_TnameConstructionTTR As Boolean
    b_TnameConstructionTTR = True
    If EnableHardIPTnameConstructionTTR = True Then
        EnableHardIPTnameConstructionTTR = False
        b_TnameConstructionTTR = False
    End If
    
    Call Reg_Assign_Processing_MTRICTS_BPBN(DigSrc_Equation, DigSrc_Assignment, CUS_Str_DigCapData, Calc_Eqn)
        
    Dim site As Variant
    Dim i As Integer
    Dim k As Integer
    Dim j As Integer
    Dim pats() As String
    Dim DSSC_Out_DecompseByComma() As String: DSSC_Out_DecompseByComma = Split(CUS_Str_DigCapData, ",")
    Dim DSSC_Out_DecompseByColon() As String
    Dim ParseStringByBits As String: ParseStringByBits = ""
    Dim ParseStringForTestName As String: ParseStringForTestName = ""
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
    Dim DecomposeTestName() As String: DecomposeTestName = Split(DataOut, ",") 'ICTS__Te008_ARCAL_DATA_OUT
    Dim DecomposeDigCapTestName() As String: DecomposeDigCapTestName = Split(ParseStringForTestName, ",")
    Dim DecomposeParseDigCapBit() As String: DecomposeParseDigCapBit = Split(ParseStringByBits, ",")
    Dim TestSequenceNumber As Long: TestSequenceNumber = UBound(DecomposeTestName) + 1
    Dim TrimStoreName_Array() As String: TrimStoreName_Array = Split(TrimStoreName, ",")
    Dim TrimStoreName_Array_BP() As String: ReDim TrimStoreName_Array_BP(TestSequenceNumber - 1)
    Dim TrimStoreName_Array_BN() As String: ReDim TrimStoreName_Array_BN(TestSequenceNumber - 1)


    'Dim code() As New SiteLong: ReDim code(TestSequenceNumber / 2 - 1)          'adc_bandgap
    Dim DAC_bp() As New SiteLong: ReDim DAC_bp(UBound(TrimStoreName_Array) \ 2)
    Dim DAC_bn() As New SiteLong: ReDim DAC_bn(UBound(TrimStoreName_Array) \ 2)
    Dim BestTrimgap() As New SiteDouble: ReDim BestTrimgap(TestSequenceNumber - 1)
    Dim BestTrimgap_previous() As New SiteDouble: ReDim BestTrimgap_previous(TestSequenceNumber - 1)
    
    Dim BinarySearch_TrimCodeValue_Min() As New SiteLong: ReDim BinarySearch_TrimCodeValue_Min(TestSequenceNumber - 1)
    Dim BinarySearch_TrimCodeValue_Max() As New SiteLong: ReDim BinarySearch_TrimCodeValue_Max(TestSequenceNumber - 1)
    'Dim BestCode() As New SiteLong: ReDim BestCode(TestSequenceNumber / 2 - 1)     'adc_bandgap
    Dim TrimCode_bp() As New PinListData: ReDim TrimCode_bp(UBound(TrimStoreName_Array) \ 2)
    Dim TrimCode_bn() As New PinListData: ReDim TrimCode_bn(UBound(TrimStoreName_Array) \ 2)
    Dim vout() As New SiteDouble: ReDim vout(TestSequenceNumber - 1)
    Dim temp_calc() As New SiteDouble: ReDim temp_calc(TestSequenceNumber - 1)
    Dim temp_TrimStart() As String: temp_TrimStart = Split(TrimStart, ";")
    Dim temp_DAC_bp_data() As String
    Dim temp_DAC_bn_data() As String
    Dim temp_Trim_step() As String
    Dim DAC_bp_data As Long
    Dim DAC_bn_data As Long
    Dim Trim_step As Long
    
    temp_DAC_bp_data = Split(temp_TrimStart(0), ":")
    temp_DAC_bn_data = Split(temp_TrimStart(1), ":")
    temp_Trim_step = Split(temp_TrimStart(2), ":")
    DAC_bp_data = temp_DAC_bp_data(1)
    DAC_bn_data = temp_DAC_bn_data(1)
    Trim_step = temp_Trim_step(1)
        
        For i = 0 To TestSequenceNumber - 1
            For Each site In TheExec.sites.Active
                    DAC_bp(i) = DAC_bp_data
                    DAC_bn(i) = DAC_bn_data
                    temp_calc(i) = 0
            Next site
        Next i
        For i = 0 To TestSequenceNumber - 1
            vout(i) = 0
        Next i
    Dim DecideTrim() As New SiteBoolean: ReDim DecideTrim(TestSequenceNumber - 1)
    Dim Trim_Flag As Boolean
        For i = 0 To TestSequenceNumber - 1
            For Each site In TheExec.sites.Active
                DecideTrim(i) = False
            Next site
        Next i
    Dim blockName() As String: blockName = Split(glb_TestInstance, "_")
    Dim MeasValue() As New PinListData: ReDim MeasValue(TestSequenceNumber - 1)
    Dim BestVal_BP() As New PinListData: ReDim BestVal_BP(TestSequenceNumber - 1)
    Dim BestVal_BN() As New PinListData: ReDim BestVal_BN(TestSequenceNumber - 1)
    Dim TestNameInput As String
    Dim PatCount As Long, PattArray() As String
    Dim BestTargetCompare() As New SiteDouble: ReDim BestTargetCompare(TestSequenceNumber - 1)

        For i = 0 To UBound(TrimStoreName_Array)
            If i Mod 2 = 1 Then
                TrimStoreName_Array_BN(Floor(i / 2)) = TrimStoreName_Array(i)
            Else
                TrimStoreName_Array_BP(Floor(i / 2)) = TrimStoreName_Array(i)
            End If
        Next i
    Dim PinName As String
    Dim TempVal As Integer
    Dim width_Wf As New DSPWave: width_Wf.CreateConstant 0, UBound(DecomposeParseDigCapBit) + 1
    Dim DecomposeParseDigCapBit_long() As Long: ReDim DecomposeParseDigCapBit_long(UBound(DecomposeParseDigCapBit))
    Dim DecomposeParseDigCapBit_long_temp() As Long: ReDim DecomposeParseDigCapBit_long_temp(TestSequenceNumber - 1)
    Dim temp_bp() As Long: ReDim temp_bp(TestSequenceNumber - 1)
    Dim temp_bn() As Long: ReDim temp_bn(TestSequenceNumber - 1)
    Dim Strtemp_bp() As String: ReDim Strtemp_bp(TestSequenceNumber - 1)
    Dim Strtemp_bn() As String: ReDim Strtemp_bn(TestSequenceNumber - 1)
    Dim Bintemp_bp() As Long: ReDim Bintemp_bp(15)
    Dim Bintemp_bn() As Long: ReDim Bintemp_bn(11)
    Dim width_Wf_2S  As New DSPWave: width_Wf_2S.CreateConstant 0, (TestSequenceNumber - 1)
    Dim OutWf_2s As New DSPWave
    Dim OutWf_DEC As New DSPWave
    Dim Bin_bp() As Long: ReDim Bin_bp(TestSequenceNumber - 1)
    Dim Bin_bn() As Long: ReDim Bin_bn(TestSequenceNumber - 1)
    Dim trim_Wf As New DSPWave: trim_Wf.CreateConstant 0, (TestSequenceNumber - 1)
    Dim trim_data() As New PinListData: ReDim trim_data(TestSequenceNumber - 1)
    
    For Each site In TheExec.sites
        For i = 0 To TestSequenceNumber - 1
            DecomposeParseDigCapBit_long_temp(i) = Cap_Trimwidth  'deliver data to dsp array
            BestTrimgap_previous(i) = 999999
        Next i
        width_Wf_2S.data = DecomposeParseDigCapBit_long_temp  'deliver data to dsp array
    Next site
    
    
    For i = 0 To UBound(DecomposeParseDigCapBit)
        DecomposeParseDigCapBit_long(i) = CLng(DecomposeParseDigCapBit(i))  'deliver data to dsp array
    Next i
    For Each site In TheExec.sites
        width_Wf.data = DecomposeParseDigCapBit_long  'deliver data to dsp array
    Next site
    Call ProcessInputToGLB(pat, , , , , , , , , , , , , , , , , DigCap_Pin, Cap_Trimwidth, DigCap_Sample_Size, DigSrc_pin, CStr(DigSrc_Sample_Size), CStr(DigSrc_Sample_Size), DigSrc_Equation, DigSrc_Assignment, , , , CUS_Str_MainProgram)
    For i = 0 To UBound(DecomposeTestName)
        MeasValue(i).AddPin (DecomposeTestName(i))
        MeasValue(i).pins(DecomposeTestName(i)).value = 0
    Next i
    
    For i = 0 To TestSequenceNumber - 1
        BestVal_BP(i).AddPin (TrimStoreName_Array_BP(i))
        BestVal_BP(i).pins(TrimStoreName_Array_BP(i)).value = 0
        BestVal_BN(i).AddPin (TrimStoreName_Array_BN(i))
        BestVal_BN(i).pins(TrimStoreName_Array_BN(i)).value = 0
    Next i

    Call GetFlowTName
    
    TheHdw.Digital.ApplyLevelsTiming True, True, True, tlPowered
    TheHdw.Digital.Patgen.Continue 0, cpuA + cpuB + cpuC + cpuD
    
    PATT_GetPatListFromPatternSet pat.value, pats, PatCount
        
    Dim TrimCodeValue_Min As Long, TrimCodeValue_Max As Long
    If UCase(CUS_Str_MainProgram) = "TWOSCOMPLEMENT_DSSC_SOURCE" Then
        TrimCodeValue_Min = (-1) * 2 ^ (TrimCodeSize - 1)
        TrimCodeValue_Max = 2 ^ (TrimCodeSize - 1) - 1
    Else
        TrimCodeValue_Min = 0
        TrimCodeValue_Max = 2 ^ TrimCodeSize - 1
    End If

    If Interpose_PrePat <> "" Then Call SetForceCondition(Interpose_PrePat & ";STOREPREPAT")
    For j = 0 To Trim_step - 1
        TheExec.Datalog.WriteComment "Now Running Trim of " & glb_TestInstance
'        For i = 0 To TestSequenceNumber - 1
'            If i = 0 Then
'                Trim_Flag = DecideTrim(i).Any(False)
'            Else
'                Trim_Flag = Trim_Flag Or DecideTrim(i).Any(False)
'            End If
'        Next i
        Trim_Flag = True
        
        If Trim_Flag Then
            If j = 0 Then
                TheExec.Datalog.WriteComment ("**************** The DSSC Captured Data at Trim Start Point =" & j & " ****************")
            ElseIf j = Trim_step Then
                TheExec.Datalog.WriteComment ("**************** The DSSC Captured Data at Trim End Point =" & j & " ****************")
            Else
                TheExec.Datalog.WriteComment ("**************** The DSSC Captured Data at Trim Step Point =" & j & " ****************")
            End If
        
            Call ICTS_DAP_PN_AutoTRIM_DigCapDataProcessByDSP(pats(0), DigSrc_pin, DAC_bp(), DAC_bn(), vout(), TrimCodeSize, TestSequenceNumber, MeasValue(), DigSrc_Sample_Size, DigSrc_Equation, DigSrc_Assignment, TrimStoreName_Array_BP(), TrimStoreName_Array_BN(), DigCap_Pin, Cap_Trimwidth, DigCap_Sample_Size, CUS_Str_MainProgram, CUS_Str_DigCapData, width_Wf, width_Wf_2S, OutWf_2s, OutWf_DEC, DecomposeParseDigCapBit)
            For Each site In TheExec.sites.Active
                trim_Wf = OutWf_2s
                For i = 0 To TestSequenceNumber - 1
                    temp_calc(i) = trim_Wf.Element(i)
                    If gl_Disable_HIP_debug_log = False Then TheExec.Datalog.WriteComment ("**************** Site :" & site & " Reg :" & MeasValue(i).pins.item(0) & " " & Cap_Trimwidth & " bit 2scomplement data is :" & temp_calc(i) & " ****************")
'                    If DecideTrim(i) = False Then
                        If j > 0 Then
                            BestTrimgap(i) = temp_calc(i).Subtract(TrimTarget).Abs
                            If BestTrimgap(i) < BestTrimgap_previous(i) Then
                                BestTrimgap_previous(i) = BestTrimgap(i)
                                BestVal_BP(i) = DAC_bp(i)
                                BestVal_BN(i) = DAC_bn(i)
                            End If
                        End If
                        If j = 0 Then
                            DAC_bp(i) = OutWf_DEC.Element(1)
                            DAC_bn(i) = OutWf_DEC.Element(2)
                            Instance_Data.DigSrc_Assignment = Replace(Instance_Data.DigSrc_Assignment, "rcal_en=1", "rcal_en=0")
                        ElseIf j = 1 Then
                            DAC_bp(i) = DAC_bp(i) - 1
                            DAC_bn(i) = DAC_bn(i) + 1
                            If DAC_bp(i) > TrimCodeValue_Max Then DAC_bp(i) = TrimCodeValue_Max
                            If DAC_bn(i) > TrimCodeValue_Max Then DAC_bn(i) = TrimCodeValue_Max
                            If DAC_bp(i) < TrimCodeValue_Min Then DAC_bp(i) = TrimCodeValue_Min
                            If DAC_bn(i) < TrimCodeValue_Min Then DAC_bn(i) = TrimCodeValue_Min
                        ElseIf j = 2 Then
                            DAC_bp(i) = DAC_bp(i) + 1
                            DAC_bn(i) = DAC_bn(i) - 1
                            If DAC_bp(i) > TrimCodeValue_Max Then DAC_bp(i) = TrimCodeValue_Max
                            If DAC_bn(i) > TrimCodeValue_Max Then DAC_bn(i) = TrimCodeValue_Max
                            If DAC_bp(i) < TrimCodeValue_Min Then DAC_bp(i) = TrimCodeValue_Min
                            If DAC_bn(i) < TrimCodeValue_Min Then DAC_bn(i) = TrimCodeValue_Min
                        ElseIf j = 3 Then
                        
                        End If
'                        If temp_calc(i) = TrimTarget Then
'                            DecideTrim(i) = True
'                        End If
'                    End If
                Next i
            Next site
        End If
        
    Next j

    Dim Best_srcwave_array_bp() As Long: ReDim Best_srcwave_array_bp(TrimCodeSize - 1)
    Dim Best_srcwave_array_bn() As Long: ReDim Best_srcwave_array_bn(TrimCodeSize - 1)
    Dim Best_DecWave_bp() As New DSPWave: ReDim Best_DecWave_bp(UBound(DAC_bp))
    Dim Best_DecWave_bn() As New DSPWave: ReDim Best_DecWave_bn(UBound(DAC_bn))
    Dim Best_srcWave_bp() As New DSPWave: ReDim Best_srcWave_bp(UBound(DAC_bp))
    Dim Best_srcWave_bn() As New DSPWave: ReDim Best_srcWave_bn(UBound(DAC_bn))
    
    For i = 0 To UBound(DAC_bp)
     Best_srcWave_bp(i).CreateConstant 0, TrimCodeSize, DspLong
     Best_srcWave_bn(i).CreateConstant 0, TrimCodeSize, DspLong
     Best_DecWave_bp(i).CreateConstant 0, 1, DspLong
     Best_DecWave_bn(i).CreateConstant 0, 1, DspLong
    For Each site In TheExec.sites
        Best_DecWave_bp(i).Element(0) = BestVal_BP(i)
        Best_DecWave_bn(i).Element(0) = BestVal_BN(i)
            For j = 0 To TrimCodeSize - 1
                If j = 0 Then
                    Best_srcwave_array_bp(j) = BestVal_BP(i) And 1
                    Best_srcwave_array_bn(j) = BestVal_BN(i) And 1
                Else
                    Best_srcwave_array_bp(j) = (BestVal_BP(i) And (2 ^ j)) \ (2 ^ j)
                    Best_srcwave_array_bn(j) = (BestVal_BN(i) And (2 ^ j)) \ (2 ^ j)
                End If
            Next j
         
        Best_srcWave_bp(i).data = Best_srcwave_array_bp
        Best_srcWave_bn(i).data = Best_srcwave_array_bn
        
        If gl_Disable_HIP_debug_log = False Then
            TheExec.Datalog.WriteComment "Site : " & site & ", Store Value : " & Best_DecWave_bp(i).Element(0) & ", Binary Bits : " & TrimCodeSize & ", Store Name : " & TrimStoreName_Array_BP(i)
            TheExec.Datalog.WriteComment "Site : " & site & ", Store Value : " & Best_DecWave_bn(i).Element(0) & ", Binary Bits : " & TrimCodeSize & ", Store Name : " & TrimStoreName_Array_BN(i)

        End If
    Next site

        Call StoreDataAllType(TrimStoreName_Array_BP(i), Best_srcWave_bp(i))
        Call StoreDataAllType(TrimStoreName_Array_BP(i) & "_para", Best_DecWave_bp(i))
        Call StoreDataAllType(TrimStoreName_Array_BN(i), Best_srcWave_bn(i))
        Call StoreDataAllType(TrimStoreName_Array_BN(i) & "_para", Best_DecWave_bn(i))
    Next i
    
    If Interpose_PrePat <> "" Then Call SetForceCondition("RESTOREPREPAT")

    Dim TempBestVal As New SiteDouble

'    For i = 0 To TestSequenceNumber - 1
'
'        TempBestVal = MeasValue(i).Pins(DecomposeTestName(i))
'        TestNameInput = Report_TName_From_Instance("C", DigCap_Pin.value, "", i, 0)
'        TheExec.Flow.TestLimit TempBestVal, 0, 2 ^ DecomposeParseDigCapBit(i) - 1, PinName:=DecomposeTestName(i), Tname:=TestNameInput, ForceResults:=tlForceFlow, scaletype:=scaleNoScaling, formatStr:="%.0f"
'
'    Next i
    For i = 0 To TestSequenceNumber - 1

        TempBestVal = BestVal_BP(i).pins(TrimStoreName_Array_BP(i))
        TestNameInput = Report_TName_From_Instance("C", DigCap_Pin.value, "", i, 0)
        TheExec.flow.TestLimit TempBestVal, 0, 2 ^ DecomposeParseDigCapBit(i) - 1, PinName:=TrimStoreName_Array_BP(i), Tname:=TestNameInput, ForceResults:=tlForceFlow, scaletype:=scaleNoScaling, formatStr:="%.0f"

        TempBestVal = BestVal_BN(i).pins(TrimStoreName_Array_BN(i))
        TestNameInput = Report_TName_From_Instance("C", DigCap_Pin.value, "", i, 0)
        TheExec.flow.TestLimit TempBestVal, 0, 2 ^ DecomposeParseDigCapBit(i) - 1, PinName:=TrimStoreName_Array_BN(i), Tname:=TestNameInput, ForceResults:=tlForceFlow, scaletype:=scaleNoScaling, formatStr:="%.0f"

    Next i
     
    DebugPrintFunc pat.value
    
    Call HardIP_WriteFuncResult(, , glb_TestInstance)
    
    If Calc_Eqn <> "" Then: Call ProcessCalcEquation(Calc_Eqn)
    
    ' Check implicit alarms
    TheHdw.Alarms.Check
    
    If b_TnameConstructionTTR = False Then
        EnableHardIPTnameConstructionTTR = True
    End If
    
    Exit Function
    
errHandler:
    TheExec.Datalog.WriteComment "error in MetrologyICTS_DAP_PN_AutoTRIM"
    If AbortTest Then Exit Function Else Resume Next
End Function


