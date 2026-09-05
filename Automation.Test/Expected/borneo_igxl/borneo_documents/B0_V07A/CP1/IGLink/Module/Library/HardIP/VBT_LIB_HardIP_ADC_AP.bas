Attribute VB_Name = "VBT_LIB_HardIP_ADC_AP"
#Const isUFP = True
Option Explicit 'Add ErrHandler 2023/05/29
Public Function ADC_BandGap_Calibration(Optional Pat As Pattern, Optional DigCap_Pin As PinList, Optional DigCap_DataWidth As Long, Optional DigCap_Sample_Size As Long, Optional DigSrc_pin As PinList, Optional DigSrc_Sample_Size As String, Optional DigSrc_Equation As String, Optional digsrc_assignment As String, Optional CUS_Str_MainProgram As String = vbNullString, Optional CUS_Str_DigCapData As String = vbNullString, Optional TrimStoreName As String, Optional TrimTarget As Double, Optional TrimStart As Long, Optional TrimCodeSize As Long, Optional Calc_Eqn As String, Optional Interpose_PrePat As String, Optional TrimMethod As String, Optional Trimming_Direction_Increase As Boolean, Optional Validating_ As Boolean)
On Error GoTo errHandler 'Add ErrHandler 2023/05/29
    
    glb_TestInstance = vbNullString
    glb_TestInstance = UCase(TheExec.DataManager.instancename)
    
    If Validating_ Then
        Call PrLoadPattern(Pat.value)
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
    Dim BlockName() As String: BlockName = Split(glb_TestInstance, "_")
    Dim MeasValue() As New PinListData: ReDim MeasValue(TestSequenceNumber - 1)
    Dim BestVal() As New PinListData: ReDim BestVal(TestSequenceNumber - 1)
    Dim StepCount As Long: StepCount = 0
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
    Call ProcessInputToGLB(Pat, , , , , , , , , , , , , , , , , DigCap_Pin, DigCap_DataWidth, DigCap_Sample_Size, DigSrc_pin, CStr(DigSrc_Sample_Size), CStr(DigSrc_Sample_Size), DigSrc_Equation, digsrc_assignment, , , , CUS_Str_MainProgram)
    
    For i = 0 To UBound(DecomposeTestName)
        MeasValue(i).AddPin (DecomposeTestName(i) & "_" & i)
        MeasValue(i).Pins(DecomposeTestName(i) & "_" & i).value = 0
        BestVal(i).AddPin (DecomposeTestName(i) & "_" & i)
        BestVal(i).Pins(DecomposeTestName(i) & "_" & i).value = 0
    Next i

    Call GetFlowTName

    
    
    TheHdw.Digital.ApplyLevelsTiming True, True, True, tlPowered
    TheHdw.Digital.Patgen.Continue 0, cpuA + cpuB + cpuC + cpuD
    
    PATT_GetPatListFromPatternSet Pat.value, pats, PatCount
        
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
    Call ADC_BandGap_DigCapDataProcessByDSP(pats(0), DigSrc_pin, code(), vout(), TrimCodeSize, TestSequenceNumber, MeasValue(), DigSrc_Sample_Size, DigSrc_Equation, digsrc_assignment, TrimStoreName_Array(), DigCap_Pin, DigCap_DataWidth, DigCap_Sample_Size, CUS_Str_MainProgram, CUS_Str_DigCapData, width_Wf)

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
                Call ADC_BandGap_DigCapDataProcessByDSP(pats(0), DigSrc_pin, code(), vout(), TrimCodeSize, TestSequenceNumber, MeasValue(), DigSrc_Sample_Size, DigSrc_Equation, digsrc_assignment, TrimStoreName_Array(), DigCap_Pin, DigCap_DataWidth, DigCap_Sample_Size, CUS_Str_MainProgram, CUS_Str_DigCapData, width_Wf)
        End If
    For k = 0 To TestSequenceNumber - 1
        For Each site In TheExec.sites.Active
            If DecideTrim(k) And vout(k).Subtract(TrimTarget).Abs < BestTargetCompare(k) Then
                BestTargetCompare(k) = vout(k).Subtract(TrimTarget).Abs
                BestCode(k) = code(k)
                BestVal(k).Pins(0).value = MeasValue(k).Pins(0).value
            End If
            If DecideTrim(k) And LCase(TrimMethod) = "linearsearch" Then
                If Trimming_Direction_Increase = True Then
                    If StepCount > (TrimCodeValue_Max - TrimCodeValue_Min) Then
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
                    If StepCount > (TrimCodeValue_Max - TrimCodeValue_Min) Then
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
                    If StepCount > (TrimCodeValue_Max - TrimCodeValue_Min) Then
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
                    If StepCount > (TrimCodeValue_Max - TrimCodeValue_Min) Then
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
    
        TempBestVal = BestVal(i).Pins(DecomposeTestName(i) & "_" & i)
        TestNameInput = Report_TName_From_Instance("C", DigCap_Pin.value, vbNullString, i, 0)
        TheExec.Flow.TestLimit TempBestVal, 0, 2 ^ DecomposeParseDigCapBit(i) - 1, PinName:=DigCap_Pin.value, Tname:=TestNameInput, ForceResults:=tlForceFlow, scaletype:=scaleNoScaling, formatStr:="%.0f"
        
    Next i
    For i = 0 To TestSequenceNumber - 1
        TestNameInput = Report_TName_From_Instance("C", vbNullString, BlockName(0) & "Trim", i, 0)
        TheExec.Flow.TestLimit resultVal:=BestCode(i), Tname:=TestNameInput, ForceResults:=tlForceFlow
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
        TempBestVal_buf = BestVal(i).Pins(DecomposeTestName(i) & "_" & i)
         For Each site In TheExec.sites
            ADC_TRIM_ENG(i).CreateConstant 0, Instance_Data.DigCap_DataWidth, DspDouble
            For j = 0 To Instance_Data.DigCap_DataWidth - 1
                ADC_TRIM_ENG(i).Element(j) = TempBestVal_buf(site) Mod 2
                TempBestVal_buf(site) = TempBestVal_buf(site) \ 2
            Next j
        Next site
        Call StoreDataAllType(TrimStoreName_Array(i) & "_target", ADC_TRIM_ENG(i))
     Next i
         
    DebugPrintFunc Pat.value
    
    Call HardIP_WriteFuncResult(, , glb_TestInstance)
    
    If Calc_Eqn <> "" Then: Call ProcessCalcEquation(Calc_Eqn)
    
    ' Check implicit alarms
    TheHdw.Alarms.Check

    Exit Function
    
errHandler:
    Call Print_Error_Message(Error_Info, "VBT_LIB_HardIP_ADC_AP", "ADC_BandGap_Calibration") 'Add ErrHandler 2023/05/29
    If AbortTest Then Exit Function Else Resume Next
End Function


Public Function ADC_LDO_Calibration(Optional Pat As Pattern, Optional TestSequence As String, Optional MeasV_PinS As String, Optional MeaV_WaitTime As String, Optional DigSrc_pin As PinList, _
                        Optional DigSrc_Sample_Size As String, Optional DigSrc_Equation As String, Optional digsrc_assignment As String, Optional CUS_Str_MainProgram As String = vbNullString, Optional TrimStoreName As String, Optional TrimTarget As Double, Optional TrimStart As Long, Optional TrimCodeSize As Long, Optional TrimMethod As String, Optional TrimStepSize As Double, Optional Validating_ As Boolean)
On Error GoTo errHandler 'Add ErrHandler 2023/05/29
    
    glb_TestInstance = vbNullString
    glb_TestInstance = UCase(TheExec.DataManager.instancename)
    
    
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
    Dim BlockName() As String: BlockName = Split(glb_TestInstance, "_")
    Dim MeasValue() As New PinListData: ReDim MeasValue(NumberOfMeasV - 1)
    Dim BestVal() As New PinListData: ReDim BestVal(NumberOfMeasV - 1)
    Dim StepCount As Long: StepCount = 0
    Dim TestNameInput As String
    Dim PatCount As Long, PattArray() As String
    Dim BestTargetCompare() As New SiteDouble: ReDim BestTargetCompare(UBound(Split(TestSequence, ",")))
    Dim TrimStoreName_Array() As String: TrimStoreName_Array = Split(TrimStoreName, ",")
    Dim PinName() As String
    Dim TempVal As Integer
    Dim FinalTrimCode() As New DSPWave: ReDim FinalTrimCode(UBound(TrimStoreName_Array))
    Dim FinalTrimCode_Array() As Long: ReDim FinalTrimCode_Array(TrimCodeSize - 1) As Long
    
    Call ProcessInputToGLB(Pat, TestSequence, True, , , , , MeasV_PinS, , , , , , , , , , , , , DigSrc_pin, DigSrc_Sample_Size, DigSrc_Sample_Size, DigSrc_Equation, digsrc_assignment, , , , CUS_Str_MainProgram, , , , , , , , , , , , , , , , , , , , , , MeaV_WaitTime)


    PinName = Split(MeasV_PinS, "+")
   For i = 0 To NumberOfMeasV - 1
    
       If InStr(MeasV_PinS, "+") <> 0 Then
        ''For j = 0 To UBound(Split(MeasV_PinS, "+"))
            MeasValue(i).AddPin PinName(i)
            MeasValue(i).Pins(PinName(i)).value = 0
            BestVal(i).AddPin PinName(i)
            BestVal(i).Pins(PinName(i)).value = 0
        ''Next j
       Else
            MeasValue(i).AddPin (PinName(0))
            MeasValue(i).Pins(PinName(0)).value = 0
            BestVal(i).AddPin (PinName(0))
            BestVal(i).Pins(PinName(0)).value = 0
       End If
       
    Next i
    
    Call GetFlowTName

    If Validating_ Then
        Call PrLoadPattern(Pat.value)
        Exit Function    ' Exit after validation
    End If
    

    
    If TheExec.DevChar.Setups.IsRunning Then
        If TheExec.DevChar.Setups(TheExec.DevChar.Setups.ActiveSetupName).Shmoo.axes.Contains(tlDevCharShmooAxis_Y) Then
            If gl_Flag_HardIP_Trim_Set_PrePoint And Not (gl_Flag_HardIP_Characterization_1stRun) Then
                TheHdw.Digital.ApplyLevelsTiming True, True, True, tlPowered
                Call TheExec.Overlays.ApplyUniformSpecToHW(XI0_Shmoo & "_Freq_VAR", TheExec.DevChar.Results(TheExec.DevChar.Setups.ActiveSetupName).Shmoo.CurrentPoint.axes(tlDevCharShmooAxis_Y).value)
            ElseIf gl_Flag_HardIP_Trim_Set_PostPoint Then
                TheHdw.Digital.ApplyLevelsTiming True, True, True, tlPowered
                Call TheExec.Overlays.ApplyUniformSpecToHW(XI0_Shmoo & "_Freq_VAR", 24000000#)
            Else
                TheHdw.Digital.ApplyLevelsTiming True, True, True, tlPowered
            End If
        End If
    Else
        TheHdw.Digital.ApplyLevelsTiming True, True, True, tlPowered
    End If
    TheHdw.Digital.Patgen.Continue 0, cpuA + cpuB + cpuC + cpuD
    
    PATT_GetPatListFromPatternSet Pat.value, pats, PatCount
        
    Dim TrimCodeValue_Min As Long, TrimCodeValue_Max As Long
    If UCase(CUS_Str_MainProgram) = "TWOSCOMPLEMENT_DSSC_SOURCE" Then
        TrimCodeValue_Min = (-1) * 2 ^ (TrimCodeSize - 1)
        TrimCodeValue_Max = 2 ^ (TrimCodeSize - 1) - 1
    Else
        TrimCodeValue_Min = 0
        TrimCodeValue_Max = 2 ^ TrimCodeSize - 1
    End If
  

    If gl_Disable_HIP_debug_log = False Then TheExec.Datalog.WriteComment ("**************** The Measurement at Trim Start Point ****************")
    Call LDO_Measurement_Process(pats(0), DigSrc_pin, code(), vout(), TrimCodeSize, NumberOfMeasV, MeasValue(), DigSrc_Sample_Size, DigSrc_Equation, digsrc_assignment, TrimStoreName_Array(), MeaV_WaitTime)

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
                Else 'Do nothing '20230601
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
                Else 'Do nothing '20230601
                End If
                DecideTrim(i) = True
            Next site
        Else
            Call Print_Error_Message(Warning_Info, "VBT_LIB_HardIP_ADC_AP", "ADC_LDO_Calibration", "TrimMethod type misalignment.")
            Exit Function
'            For Each site In TheExec.sites.Active
'                code(i) = code(i) + Fix((TrimTarget - vout(i)) / TrimStepSize)
'                DecideTrim(i) = True
'            Next site
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

            Call LDO_Measurement_Process(pats(0), DigSrc_pin, code(), vout(), TrimCodeSize, NumberOfMeasV, MeasValue(), DigSrc_Sample_Size, DigSrc_Equation, digsrc_assignment, TrimStoreName_Array(), MeaV_WaitTime)
        End If
    For k = 0 To UBound(Split(TestSequence, ","))
        For Each site In TheExec.sites.Active
            If DecideTrim(k) And vout(k).Subtract(TrimTarget).Abs < BestTargetCompare(k) Then
                BestTargetCompare(k) = vout(k).Subtract(TrimTarget).Abs
                BestCode(k) = code(k)
                
                If InStr(MeasV_PinS, "+") <> 0 Then
                    BestVal(k).Pins(PinName(k)).value = MeasValue(k).Pins(PinName(k)).value
                Else
                    BestVal(k).Pins(PinName(0)).value = MeasValue(k).Pins(PinName(0)).value
                End If
                
            End If
            If DecideTrim(k) And LCase(TrimMethod) = "linearsearch" Then
                If StepCount > (TrimCodeValue_Max - TrimCodeValue_Min) Then
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
                Else 'Do nothing '20230601
                End If
            ElseIf DecideTrim(k) And LCase(TrimMethod) = "binarysearch" Then
                If StepCount > (TrimCodeValue_Max - TrimCodeValue_Min) Then
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
                Else 'Do nothing '20230601
                End If
            Else 'Do nothing '20230601
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
            TheExec.Flow.TestLimit resultVal:=BestVal(i), Tname:=TestNameInput, PinName:=PinName(i), ForceResults:=tlForceFlow

        Else
      
            TestNameInput = Report_TName_From_Instance("V", PinName(0), vbNullString, i, 0)
            TheExec.Flow.TestLimit resultVal:=BestVal(i), Tname:=TestNameInput, PinName:=PinName(0), ForceResults:=tlForceFlow
        End If
    Next i
    For i = 0 To UBound(Split(TestSequence, ","))
        TestNameInput = Report_TName_From_Instance("C", vbNullString, BlockName(0) & "Trim", i, 0)
        TheExec.Flow.TestLimit resultVal:=BestCode(i), Tname:=TestNameInput, ForceResults:=tlForceFlow
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
        Call StoreDataAllType(TrimStoreName_Array(i), FinalTrimCode(i))
    Next i
    DebugPrintFunc Pat.value
    
    Call HardIP_WriteFuncResult(, , glb_TestInstance)
    
    TheHdw.Alarms.Check
    Exit Function
    
errHandler:
    Call Print_Error_Message(Error_Info, "VBT_LIB_HardIP_ADC_AP", "ADC_LDO_Calibration") 'Add ErrHandler 2023/05/29
    If AbortTest Then Exit Function Else Resume Next
End Function



