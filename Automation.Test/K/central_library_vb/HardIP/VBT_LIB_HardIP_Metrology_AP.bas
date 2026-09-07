Attribute VB_Name = "VBT_LIB_HardIP_Metrology_AP"
Option Explicit 'Add ErrHandler 2023/05/29

Public Function MetrologyTMPS_Calibration(Optional Pat As Pattern, Optional TestSequence As String, Optional MeasV_PinS As String, Optional MeaV_WaitTime As String, Optional DigSrc_pin As PinList, _
                    Optional DigSrc_Sample_Size As String, Optional DigSrc_Equation As String, Optional digsrc_assignment As String, Optional TrimStoreName As String, Optional TrimTarget As Double, Optional TrimStart As Long, Optional TrimCodeSize As Long, Optional MTRTMPS_Error_Fuse_BitCount As Long, Optional MTRTMPS_Error_Dict As String, Optional TrimMethod As String, Optional Validating_ As Boolean)
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
    Dim BestTarget() As New SiteDouble: ReDim BestTarget(UBound(Split(TestSequence, ",")))
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
    Dim PreviousMeasValue() As New PinListData: ReDim PreviousMeasValue(NumberOfMeasV - 1)
    Dim BestVal() As New PinListData: ReDim BestVal(NumberOfMeasV - 1)
    Dim StepCount As Long: StepCount = 0
    Dim TestNameInput As String
    Dim PatCount As Long, PattArray() As String
    Dim PreviousTargetCompare() As New SiteDouble: ReDim PreviousTargetCompare(UBound(Split(TestSequence, ",")))
    Dim BestTargetCompare() As New SiteDouble: ReDim BestTargetCompare(UBound(Split(TestSequence, ",")))
    Dim TrimStoreName_Array() As String: TrimStoreName_Array = Split(TrimStoreName, ",")
    Dim PinName As String
    Dim TempVal As Integer
    Dim FinalTrimCode() As New DSPWave: ReDim FinalTrimCode(UBound(Split(TrimStoreName, ",")))
    Dim FinalTrimCode_Array() As Long: ReDim FinalTrimCode_Array(TrimCodeSize - 1) As Long
    
    If Validating_ Then
        Call PrLoadPattern(Pat.value)
        Exit Function    ' Exit after validation
    End If
    Call ProcessInputToGLB(Pat, TestSequence, True, , , , , MeasV_PinS, , , , , , , , , , , , , DigSrc_pin, CStr(DigSrc_Sample_Size), CStr(DigSrc_Sample_Size), DigSrc_Equation, digsrc_assignment, , , , , , , , , , , , , , , , , , , , , , , , , , MeaV_WaitTime)
    
    For i = 0 To NumberOfMeasV - 1
        For j = 0 To UBound(Split(MeasV_PinS, ","))
            PreviousMeasValue(i).AddPin (Split(MeasV_PinS, ",")(j))
            PreviousMeasValue(i).Pins(Split(MeasV_PinS, ",")(j)).value = 0
            MeasValue(i).AddPin (Split(MeasV_PinS, ",")(j))
            MeasValue(i).Pins(Split(MeasV_PinS, ",")(j)).value = 0
            BestVal(i).AddPin (Split(MeasV_PinS, ",")(j))
            BestVal(i).Pins(Split(MeasV_PinS, ",")(j)).value = 0
        Next j
    Next i
    
    Call GetFlowTName

    
    
    TheHdw.Digital.ApplyLevelsTiming True, True, True, tlPowered
    TheHdw.Digital.Patgen.Continue 0, cpuA + cpuB + cpuC + cpuD
    
    PATT_GetPatListFromPatternSet Pat.value, pats, PatCount
        
    Dim TrimCodeValue_Min As Long, TrimCodeValue_Max As Long
    TrimCodeValue_Min = 0
    TrimCodeValue_Max = 2 ^ TrimCodeSize - 1
  

    If gl_Disable_HIP_debug_log = False Then TheExec.Datalog.WriteComment ("**************** The Measurement at Trim Start Point ****************")
    Call MetrologyTMPS_Measurement_Process(pats(0), DigSrc_pin, code(), vout(), TrimCodeSize, NumberOfMeasV, MeasValue(), DigSrc_Sample_Size, DigSrc_Equation, digsrc_assignment, TrimStoreName_Array(), MeaV_WaitTime)
    
    For Each site In TheExec.sites.Active
        For i = 0 To NumberOfMeasV - 1
            BestCode(i) = code(i)
            BestVal(i) = MeasValue(i)
            BestTarget(i) = vout(i)
            BestTargetCompare(i) = vout(i).Subtract(TrimTarget).Abs
        Next i
    Next site
    For i = 0 To UBound(Split(TestSequence, ","))
        If LCase(TrimMethod) = "linearsearch" Then
            For Each site In TheExec.sites.Active
                If vout(i).compare(LessThan, TrimTarget) Then
                    code(i) = code(i) - 1
                ElseIf vout(i).compare(GreaterThan, TrimTarget) Then
                    code(i) = code(i) + 1
                Else 'Do nothing '20230601
                End If
                DecideTrim(i) = True
            Next site
        ElseIf LCase(TrimMethod) = "binarysearch" Then
            For Each site In TheExec.sites.Active
                If vout(i).compare(LessThan, TrimTarget) Then
                    BinarySearch_TrimCodeValue_Min(i) = TrimCodeValue_Min
                    BinarySearch_TrimCodeValue_Max(i) = code(i)
                    code(i) = FormatNumber(0.5 * (code(i) + TrimCodeValue_Min), 0)
                ElseIf vout(i).compare(GreaterThan, TrimTarget) Then
                    BinarySearch_TrimCodeValue_Min(i) = code(i)
                    BinarySearch_TrimCodeValue_Max(i) = TrimCodeValue_Max
                    code(i) = FormatNumber(0.5 * (code(i) + TrimCodeValue_Max), 0)
                Else 'Do nothing '20230601
                End If
                DecideTrim(i) = True
            Next site
        Else
            Call Print_Error_Message(Warning_Info, "VBT_LIB_HardIP_Metrology_AP", "MetrologyTMPS_Calibration", "TrimMethod type misalignment.")
            Exit Function
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

            Call MetrologyTMPS_Measurement_Process(pats(0), DigSrc_pin, code(), vout(), TrimCodeSize, NumberOfMeasV, MeasValue(), DigSrc_Sample_Size, DigSrc_Equation, digsrc_assignment, TrimStoreName_Array(), MeaV_WaitTime)
        End If
    For k = 0 To UBound(Split(TestSequence, ","))
        For Each site In TheExec.sites.Active
            If DecideTrim(k) And vout(k).Subtract(TrimTarget).Abs < BestTargetCompare(k) Then
                BestTargetCompare(k) = vout(k).Subtract(TrimTarget).Abs
                BestCode(k) = code(k)
                BestTarget(k) = vout(k)
                For j = 0 To UBound(Split(MeasV_PinS, ","))
                    BestVal(k).Pins(j).value = MeasValue(k).Pins(j).value
                Next j
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
                    code(k) = code(k) - 1
                    PreviousNegative(k) = True
                    DecideTrim(k) = True
                ElseIf vout(k).compare(GreaterThan, TrimTarget) And PreviousNegative(k) Then
                    DecideTrim(k) = False
                ElseIf vout(k).compare(GreaterThan, TrimTarget) And Not (PreviousNegative(k)) Then
                    code(k) = code(k) + 1
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
    
    Dim sPinName As String
    Dim sPinAry() As String
    Dim lPinCnt As Long
    
    
    TheExec.DataManager.DecomposePinList MeasV_PinS, sPinAry, lPinCnt
    For i = 0 To lPinCnt - 1
        If LCase(sPinAry(i)) Like "*_p*" Then
            sPinName = sPinAry(i)
        End If
    Next i
    
    For i = 0 To NumberOfMeasV - 1
        'Oscar, DiffMeter

        TestNameInput = Report_TName_From_Instance("Vdiff", vbNullString, vbNullString, i, 0)
        TheExec.Flow.TestLimit resultVal:=BestTarget(i), Tname:=TestNameInput, ForceResults:=tlForceFlow, PinName:=sPinName
    Next i
    For i = 0 To UBound(Split(TestSequence, ","))
        'TestNameInput = Report_TName_From_Instance("C", "", blockName(0) & "Trim", i, 0)
        TestNameInput = Report_TName_From_Instance("C", vbNullString, vbNullString, i, 0)
        TheExec.Flow.TestLimit resultVal:=BestCode(i), Tname:=TestNameInput, ForceResults:=tlForceFlow, PinName:=DigSrc_pin
    Next i
    For i = 0 To UBound(Split(TestSequence, ","))
        For Each site In TheExec.sites
            TempVal = BestCode(i)
            For j = 0 To TrimCodeSize - 1
                FinalTrimCode_Array(j) = TempVal Mod 2
                TempVal = TempVal \ 2
            Next j
            FinalTrimCode(i).data = FinalTrimCode_Array
        Next site
        Call StoreDataAllType(TrimStoreName_Array(i), FinalTrimCode(i))
    Next i
    DebugPrintFunc Pat.value

'=====================================Metrology TMPS Error Calculation=====================================
    
'    Dim SiteDbl_Vref As New SiteDouble
'    For Each Site In TheExec.sites
'        SiteDbl_Vref = Abs(BestVal(0).Pins(0).Value - BestVal(0).Pins(1).Value)
'    Next Site
'    Dim SiteDbl_Vref_Error As New SiteDouble: SiteDbl_Vref_Error = SiteDbl_Vref.Divide(0.8).Subtract(1).Multiply(100).Divide(0.125)
'    Dim High_limit As Double: High_limit = Bin2Dec_rev(String(MTRTMPS_Error_Fuse_BitCount - 1, "1"))
'    Dim Low_limit As Double: Low_limit = -2 ^ (MTRTMPS_Error_Fuse_BitCount - 1)
'
'    For Each Site In TheExec.sites
'        SiteDbl_Vref_Error = Floor(SiteDbl_Vref_Error)
'    Next Site
'
'    TestNameInput = Report_TName_From_Instance("C", "X", "mtr_gr_vref_err", 0, 0)
'
'    TheExec.Flow.TestLimit resultVal:=SiteDbl_Vref_Error, LowVal:=Low_limit, HiVal:=High_limit, TName:=TestNameInput, ForceResults:=tlForceFlow
'
'    Dim SiteDbl_Vref_Error_Fuse As New DSPWave
'    Dim SiteDbl_Vref_Error_Fuse_Array(0) As Long
'
'    For Each Site In TheExec.sites
'        SiteDbl_Vref_Error = FormatNumber(SiteDbl_Vref_Error, 0)
'        If SiteDbl_Vref_Error < Low_limit Then
'            SiteDbl_Vref_Error_Fuse_Array(0) = 2 ^ (MTRTMPS_Error_Fuse_BitCount) + FormatNumber(Low_limit)
'        ElseIf SiteDbl_Vref_Error >= Low_limit And SiteDbl_Vref_Error < 0 Then
'            SiteDbl_Vref_Error_Fuse_Array(0) = 2 ^ (MTRTMPS_Error_Fuse_BitCount) + FormatNumber(SiteDbl_Vref_Error)
'        ElseIf SiteDbl_Vref_Error < High_limit And SiteDbl_Vref_Error >= 0 Then
'            SiteDbl_Vref_Error_Fuse_Array(0) = FormatNumber(SiteDbl_Vref_Error)
'        Else
'            SiteDbl_Vref_Error_Fuse_Array(0) = FormatNumber(High_limit)
'        End If
'        SiteDbl_Vref_Error_Fuse.Data = SiteDbl_Vref_Error_Fuse_Array
'    Next Site
'
'    Call StoreDataAllType(MTRTMPS_Error_Dict, SiteDbl_Vref_Error_Fuse)
    
'===============================================================================================================
    
    Call HardIP_WriteFuncResult(, , glb_TestInstance)
    
    ' Check implicit alarms
    TheHdw.Alarms.Check
    
    Exit Function
    
errHandler:
    Call Print_Error_Message(Error_Info, "VBT_LIB_HardIP_Metrology_AP", "MetrologyTMPS_Calibration") 'Add ErrHandler 2023/05/29
    If AbortTest Then Exit Function Else Resume Next
End Function

Public Function MetrologyGR_Calibration(Optional Pat As Pattern, Optional TestSequence As String, Optional MeasV_PinS As String, Optional MeaV_WaitTime As String, Optional DigSrc_pin As PinList, _
                        Optional DigSrc_Sample_Size As String, Optional DigSrc_Equation As String, Optional digsrc_assignment As String, Optional TrimStoreName As String, Optional TrimTarget As Double, Optional TrimStart As Long, Optional TrimCodeSize As Long, Optional Error_Fuse_BitCount As Long, Optional Error_Dict As String, Optional TrimStoreName_MSB_Inverse As String, Optional TrimMethod As String, Optional Validating_ As Boolean)
On Error GoTo errHandler 'Add ErrHandler 2023/05/29

    Dim site As Variant
    Dim i As Integer
    Dim j As Integer
    Dim pats() As String
    Dim code As New SiteLong: code = TrimStart
    Dim BinarySearch_TrimCodeValue_Min As New SiteLong
    Dim BinarySearch_TrimCodeValue_Max As New SiteLong
    Dim vout As New SiteDouble
    Dim BestTarget As New SiteDouble
    Dim NumberOfMeasV As Integer: NumberOfMeasV = UBound(Split(TestSequence, ",")) + 1
    Dim BestCode As New SiteLong, temp As New SiteLong
    Dim Done As New SiteBoolean
    Dim PreviousNegative As New SiteBoolean
    Dim PreviousPositive As New SiteBoolean
    Dim DecideTrim As New SiteBoolean
        For Each site In TheExec.sites.Active
            PreviousNegative = False
            PreviousPositive = False
            DecideTrim = False
        Next site
                
    glb_TestInstance = vbNullString
    glb_TestInstance = UCase(TheExec.DataManager.instancename)
        
    Dim BlockName() As String: BlockName = Split(glb_TestInstance, "_")
    Dim MeasValue() As New PinListData: ReDim MeasValue(NumberOfMeasV - 1)
    Dim PreviousMeasValue() As New PinListData: ReDim PreviousMeasValue(NumberOfMeasV - 1)
    Dim BestVal() As New PinListData: ReDim BestVal(NumberOfMeasV - 1)
    Dim StepCount As Long: StepCount = 0
    Dim TestNameInput As String
    Dim PatCount As Long, PattArray() As String
    Dim PinName_Temp As String
    Dim BestTargetCompare As New SiteDouble
    
    'Call ProcessInputToGLB(Pat, TestSequence, True, , , , , MeasV_PinS, , , , , , , , , , , , , , , , , , , , , , , , , , , , , , , , , , , , , , , , , , , MeaV_WaitTime)
    Call ProcessInputToGLB(Pat, TestSequence, True, , , , , MeasV_PinS, , , , , , , , , , , , , DigSrc_pin, CStr(DigSrc_Sample_Size), CStr(DigSrc_Sample_Size), DigSrc_Equation, digsrc_assignment, , , , MeaV_WaitTime)
    
    For i = 0 To NumberOfMeasV - 1
        For j = 0 To UBound(Split(Split(MeasV_PinS, "+")(i), ","))
            PreviousMeasValue(i).AddPin (Split(Split(MeasV_PinS, "+")(i), ",")(j))
            PreviousMeasValue(i).Pins(Split(Split(MeasV_PinS, "+")(i), ",")(j)).value = 0
            MeasValue(i).AddPin (Split(Split(MeasV_PinS, "+")(i), ",")(j))
            MeasValue(i).Pins(Split(Split(MeasV_PinS, "+")(i), ",")(j)).value = 0
            BestVal(i).AddPin (Split(Split(MeasV_PinS, "+")(i), ",")(j))
            BestVal(i).Pins(Split(Split(MeasV_PinS, "+")(i), ",")(j)).value = 0
        Next j
    Next i
    
    Call GetFlowTName

    If Validating_ Then
        Call PrLoadPattern(Pat.value)
        Exit Function    ' Exit after validation
    End If
   
    
    TheHdw.Digital.ApplyLevelsTiming True, True, True, tlPowered
    TheHdw.Digital.Patgen.Continue 0, cpuA + cpuB + cpuC + cpuD
    
    PATT_GetPatListFromPatternSet Pat.value, pats, PatCount
        
    Dim TrimCodeValue_Min As Long, TrimCodeValue_Mid As Long, TrimCodeValue_Max As Long
    TrimCodeValue_Min = 0
    TrimCodeValue_Mid = (2 ^ TrimCodeSize) / 2
    TrimCodeValue_Max = 2 ^ TrimCodeSize - 1
  

    If gl_Disable_HIP_debug_log = False Then TheExec.Datalog.WriteComment ("**************** The Measurement at Trim Start Point ****************")
    Call MetrologyGR_Measurement_Process(pats(0), DigSrc_pin, code, vout, TrimCodeSize, NumberOfMeasV, MeasValue(), DigSrc_Sample_Size, DigSrc_Equation, digsrc_assignment, TrimStoreName, MeaV_WaitTime)

    For Each site In TheExec.sites.Active
        BestCode = code
        BestTargetCompare = vout.Subtract(TrimTarget).Abs
        BestTarget = vout
        For i = 0 To NumberOfMeasV - 1
            BestVal(i) = MeasValue(i)
        Next i
    Next site
    If LCase(TrimMethod) = "linearsearch" Then
        For Each site In TheExec.sites.Active
            If vout.compare(LessThan, TrimTarget) Then
                code = code - 1
            ElseIf vout.compare(GreaterThan, TrimTarget) Then
                code = code + 1
            Else 'Do nothing '20230601
            End If
            DecideTrim = True
        Next site
    ElseIf LCase(TrimMethod) = "binarysearch" Then
        For Each site In TheExec.sites.Active
            If vout.compare(LessThan, TrimTarget) Then
                BinarySearch_TrimCodeValue_Min = TrimCodeValue_Min
                BinarySearch_TrimCodeValue_Max = code
                code = FormatNumber(0.5 * (code + TrimCodeValue_Min), 0)
            ElseIf vout.compare(GreaterThan, TrimTarget) Then
                BinarySearch_TrimCodeValue_Min = code
                BinarySearch_TrimCodeValue_Max = TrimCodeValue_Max
                code = FormatNumber(0.5 * (code + TrimCodeValue_Max), 0)
            Else 'Do nothing '20230601
            End If
            DecideTrim = True
        Next site
    Else
        Call Print_Error_Message(Warning_Info, "VBT_LIB_HardIP_Metrology_AP", "MetrologyGR_Calibration", "TrimMethod type misalignment.")
        Exit Function
    End If
    
StartTrim:
        If DecideTrim.Any(True) Then
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
            Call MetrologyGR_Measurement_Process(pats(0), DigSrc_pin, code, vout, TrimCodeSize, NumberOfMeasV, MeasValue(), DigSrc_Sample_Size, DigSrc_Equation, digsrc_assignment, TrimStoreName, MeaV_WaitTime)
        End If

    For Each site In TheExec.sites.Active
        If DecideTrim And vout.Subtract(TrimTarget).Abs < BestTargetCompare Then
            BestTargetCompare = vout.Subtract(TrimTarget).Abs
            BestCode = code
            BestTarget = vout
            For i = 0 To NumberOfMeasV - 1
                For j = 0 To UBound(Split(MeasV_PinS, ","))
                    BestVal(i).Pins(j).value = MeasValue(i).Pins(j).value
                Next j
            Next i
        End If
        If DecideTrim And LCase(TrimMethod) = "linearsearch" Then
            If StepCount > (TrimCodeValue_Max - TrimCodeValue_Min) Then
                DecideTrim = False
            ElseIf code.compare(GreaterThan, TrimCodeValue_Max) Then
                code = TrimCodeValue_Max
                DecideTrim = True
            ElseIf code.compare(LessThan, TrimCodeValue_Min) Then
                code = TrimCodeValue_Min
                DecideTrim = True
            ElseIf code.compare(EqualTo, TrimCodeValue_Max) Or code.compare(EqualTo, TrimCodeValue_Min) Then
                DecideTrim = False
            ElseIf vout.compare(LessThan, TrimTarget) And PreviousPositive Then
                DecideTrim = False
            ElseIf vout.compare(LessThan, TrimTarget) And Not (PreviousPositive) Then
                code = code - 1
                PreviousNegative = True
                DecideTrim = True
            ElseIf vout.compare(GreaterThan, TrimTarget) And PreviousNegative Then
                DecideTrim = False
            ElseIf vout.compare(GreaterThan, TrimTarget) And Not (PreviousNegative) Then
                code = code + 1
                PreviousPositive = True
                DecideTrim = True
            Else 'Do nothing '20230601
            End If
        ElseIf DecideTrim And LCase(TrimMethod) = "binarysearch" Then
            If StepCount > (TrimCodeValue_Max - TrimCodeValue_Min) Then
                DecideTrim = False
            ElseIf code.compare(GreaterThan, TrimCodeValue_Max) Then
                code = TrimCodeValue_Max
                DecideTrim = True
            ElseIf code.compare(LessThan, TrimCodeValue_Min) Then
                code = TrimCodeValue_Min
                DecideTrim = True
            ElseIf code.compare(EqualTo, BinarySearch_TrimCodeValue_Max) Or code.compare(EqualTo, BinarySearch_TrimCodeValue_Min) Then
                DecideTrim = False
            ElseIf vout.compare(LessThan, TrimTarget) Then
                If code = FormatNumber(0.5 * (code + BinarySearch_TrimCodeValue_Min), 0) Then
                    DecideTrim = False
                Else
                    BinarySearch_TrimCodeValue_Max = code
                    code = FormatNumber(0.5 * (code + BinarySearch_TrimCodeValue_Min), 0)
                    DecideTrim = True
                End If
            ElseIf vout.compare(GreaterThan, TrimTarget) Then
                If code = FormatNumber(0.5 * (code + BinarySearch_TrimCodeValue_Max), 0) Then
                    DecideTrim = False
                Else
                    BinarySearch_TrimCodeValue_Min = code
                    code = FormatNumber(0.5 * (code + BinarySearch_TrimCodeValue_Max), 0)
                    DecideTrim = True
                End If
            ElseIf vout.compare(EqualTo, TrimTarget) Then
                DecideTrim = False
            Else 'Do nothing '20230601
            End If
        Else 'Do nothing '20230601
        End If
    Next site
        
    If DecideTrim.Any(True) Then GoTo StartTrim
        
    For i = 0 To NumberOfMeasV - 1
        For j = 0 To UBound(Split(Split(MeasV_PinS, "+")(i), ","))
            PinName_Temp = Split(Split(MeasV_PinS, "+")(i), ",")(j)
            TestNameInput = Report_TName_From_Instance("V", PinName_Temp, vbNullString, i, 0)
            TheExec.Flow.TestLimit resultVal:=BestVal(i), Tname:=TestNameInput, PinName:=Split(Split(MeasV_PinS, "+")(i), ",")(j), ForceResults:=tlForceFlow
        Next j
    Next i
    TestNameInput = Report_TName_From_Instance("Vdiff", vbNullString, vbNullString, i, 0)
    TheExec.Flow.TestLimit resultVal:=BestTarget, Tname:=TestNameInput, ForceResults:=tlForceFlow
    
    TestNameInput = Report_TName_From_Instance("C", vbNullString, BlockName(0) & "Trim", 0, 0)
    TheExec.Flow.TestLimit resultVal:=BestCode, Tname:=TestNameInput, ForceResults:=tlForceFlow
    
    
    
    Dim TempVal As Integer
    Dim FinalTrimCode As New DSPWave
    Dim FinalTrimCode_MSB_Inverse As New DSPWave
    Dim FinalTrimCode_Array() As Long: ReDim FinalTrimCode_Array(TrimCodeSize - 1) As Long
    Dim FinalTrimCode_MSB_Inverse_Array() As Long: ReDim FinalTrimCode_MSB_Inverse_Array(TrimCodeSize - 1) As Long
        
    For Each site In TheExec.sites
        TempVal = BestCode(site)
        For i = 0 To TrimCodeSize - 1
            FinalTrimCode_Array(i) = TempVal Mod 2
            If i = TrimCodeSize - 1 Then
                If FinalTrimCode_Array(i) = 0 Then FinalTrimCode_MSB_Inverse_Array(i) = 1 Else FinalTrimCode_MSB_Inverse_Array(i) = 0
            Else
                FinalTrimCode_MSB_Inverse_Array(i) = FinalTrimCode_Array(i)
            End If
            TempVal = TempVal \ 2
        Next i
        FinalTrimCode.data = FinalTrimCode_Array
        FinalTrimCode_MSB_Inverse.data = FinalTrimCode_MSB_Inverse_Array
    Next site

    Call StoreDataAllType(TrimStoreName, FinalTrimCode)
    Call StoreDataAllType(TrimStoreName_MSB_Inverse, FinalTrimCode_MSB_Inverse)
    
'=====================================Metrology GR Error Calculation=====================================
    
    Dim SiteDbl_Vref As New SiteDouble
    For Each site In TheExec.sites
        SiteDbl_Vref = Abs(BestVal(0).Pins(0).value - BestVal(1).Pins(0).value)
    Next site
    Dim SiteDbl_Vref_Error As New SiteDouble: SiteDbl_Vref_Error = SiteDbl_Vref.divide(0.8).Subtract(1).Multiply(100).divide(0.125)
    Dim High_limit As Double: High_limit = Bin2Dec_rev(String(Error_Fuse_BitCount - 1, "1"))
    Dim Low_limit As Double: Low_limit = -2 ^ (Error_Fuse_BitCount - 1)

    For Each site In TheExec.sites
        SiteDbl_Vref_Error = Floor(SiteDbl_Vref_Error)
    Next site

    TestNameInput = Report_TName_From_Instance("C", "X", "MTRVBGERR0", 0, 0)

    TheExec.Flow.TestLimit resultVal:=SiteDbl_Vref_Error, lowVal:=Low_limit, hiVal:=High_limit, Tname:=TestNameInput, ForceResults:=tlForceFlow

    Dim SiteDbl_Vref_Error_Fuse As New DSPWave
    Dim SiteDbl_Vref_Error_Fuse_Array(0) As Long

    For Each site In TheExec.sites
        SiteDbl_Vref_Error = FormatNumber(SiteDbl_Vref_Error, 0)
        If SiteDbl_Vref_Error < Low_limit Then
            SiteDbl_Vref_Error_Fuse_Array(0) = 2 ^ (Error_Fuse_BitCount) + FormatNumber(Low_limit)
        ElseIf SiteDbl_Vref_Error >= Low_limit And SiteDbl_Vref_Error < 0 Then
            SiteDbl_Vref_Error_Fuse_Array(0) = 2 ^ (Error_Fuse_BitCount) + FormatNumber(SiteDbl_Vref_Error)
        ElseIf SiteDbl_Vref_Error < High_limit And SiteDbl_Vref_Error >= 0 Then
            SiteDbl_Vref_Error_Fuse_Array(0) = FormatNumber(SiteDbl_Vref_Error)
        Else
            SiteDbl_Vref_Error_Fuse_Array(0) = FormatNumber(High_limit)
        End If
        SiteDbl_Vref_Error_Fuse.data = SiteDbl_Vref_Error_Fuse_Array
    Next site

    Call StoreDataAllType(Error_Dict, SiteDbl_Vref_Error_Fuse)
    
'===============================================================================================================
    DebugPrintFunc Pat.value
    
    Call HardIP_WriteFuncResult(, , glb_TestInstance)
    
    ' Check implicit alarms
    TheHdw.Alarms.Check
    
    Exit Function
    
errHandler:
    Call Print_Error_Message(Error_Info, "VBT_LIB_HardIP_Metrology_AP", "MetrologyGR_Calibration") 'Add ErrHandler 2023/05/29
    If AbortTest Then Exit Function Else Resume Next
End Function

Public Function MetrologyICTS_DAP_PN_TRIM(Optional Pat As Pattern, Optional DigCap_Pin As PinList, Optional Cap_Trimwidth As Long, Optional DigCap_Sample_Size As Long, Optional DigSrc_pin As PinList, Optional DigSrc_Sample_Size As String, Optional DigSrc_Equation As String, Optional DigSrc_Assignment As String, Optional CUS_Str_MainProgram As String = "", Optional CUS_Str_DigCapData As String = "", Optional TrimStoreName As String, Optional TrimTarget As Double, Optional TrimStart As String, Optional TrimCodeSize As Long, Optional Calc_Eqn As String, Optional Interpose_PrePat As String, Optional TrimMethod As String, Optional Trimming_Direction_Increase As Boolean,Optional Resolution As String, Optional Validating_ As Boolean)
    
    glb_TestInstance = ""
    glb_TestInstance = UCase(TheExec.DataManager.instanceName)
    
    If Validating_ Then
        Call PrLoadPattern(Pat.value)
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
    Dim DecomposeTestName() As String: DecomposeTestName = Split(ParseStringForTestName, ",")
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
    Dim BlockName() As String: BlockName = Split(glb_TestInstance, "_")
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
    Dim width_Wf_2S  As New DSPWave: width_Wf_2S.CreateConstant 0, (TestSequenceNumber / 2)
    Dim OutWf_2s As New DSPWave
    Dim Bin_bp() As Long: ReDim Bin_bp(TestSequenceNumber - 1)
    Dim Bin_bn() As Long: ReDim Bin_bn(TestSequenceNumber - 1)
    Dim trim_Wf As New DSPWave: trim_Wf.CreateConstant 0, (TestSequenceNumber / 2 - 1)
    Dim trim_data() As New PinListData: ReDim trim_data(TestSequenceNumber / 2 - 1)
    
    For Each site In TheExec.sites
        For i = 0 To TestSequenceNumber - 1
            DecomposeParseDigCapBit_long_temp(i) = Cap_Trimwidth  'deliver data to dsp array
            BestTrimgap_previous(i) = 9999
        Next i
        width_Wf_2S.data = DecomposeParseDigCapBit_long_temp  'deliver data to dsp array
    Next site
    
    
    For i = 0 To UBound(DecomposeParseDigCapBit)
        DecomposeParseDigCapBit_long(i) = CLng(DecomposeParseDigCapBit(i))  'deliver data to dsp array
    Next i
    For Each site In TheExec.sites
        width_Wf.data = DecomposeParseDigCapBit_long  'deliver data to dsp array
    Next site
    Call ProcessInputToGLB(Pat, , , , , , , , , , , , , , , , , DigCap_Pin, Cap_Trimwidth, DigCap_Sample_Size, DigSrc_pin, CStr(DigSrc_Sample_Size), CStr(DigSrc_Sample_Size), DigSrc_Equation, DigSrc_Assignment, , , , CUS_Str_MainProgram)
    For i = 0 To UBound(DecomposeTestName)
        MeasValue(i).AddPin (DecomposeTestName(i))
        MeasValue(i).Pins(DecomposeTestName(i)).value = 0
    Next i
    
    For i = 0 To TestSequenceNumber - 1
        BestVal_BP(i).AddPin (TrimStoreName_Array_BP(i))
        BestVal_BP(i).Pins(TrimStoreName_Array_BP(i)).value = 0
        BestVal_BN(i).AddPin (TrimStoreName_Array_BN(i))
        BestVal_BN(i).Pins(TrimStoreName_Array_BN(i)).value = 0
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
'===================Check Resolution=========================
    Dim Resolution_CHK As Double
    Resolution_CHK = CDbl(Resolution)
    Resolution_CHK = Round(Resolution_CHK)
    If Resolution_CHK < CDbl(Resolution) Then TheExec.Datalog.WriteComment " Please check the argument : Resolution !!! "
'===================Check Resolution=========================
    If Interpose_PrePat <> "" Then Call SetForceCondition(Interpose_PrePat & ";STOREPREPAT")
    For j = Trim_step To -1 Step -1
        TheExec.Datalog.WriteComment "Now Rr unning Trim of " & glb_TestInstance
        For i = 0 To TestSequenceNumber - 1
            If i = 0 Then
                Trim_Flag = DecideTrim(i).Any(False)
            Else
                Trim_Flag = Trim_Flag Or DecideTrim(i).Any(False)
            End If
        Next i
        If Trim_Flag Then
            If j = 6 Then
                TheExec.Datalog.WriteComment ("**************** The DSSC Captured Data at Trim Start Point J=" & j & " ****************")
            ElseIf j = -1 Then
                TheExec.Datalog.WriteComment ("**************** The DSSC Captured Data at Trim End Point J=" & j & " ****************")
            Else
                TheExec.Datalog.WriteComment ("**************** The DSSC Captured Data at Trim Step Point J=" & j & " ****************")
            End If
        
            Call ICTS_DAP_PN_TRIM_DigCapDataProcessByDSP(pats(0), DigSrc_pin, DAC_bp(), DAC_bn(), vout(), TrimCodeSize, TestSequenceNumber, MeasValue(), DigSrc_Sample_Size, DigSrc_Equation, DigSrc_Assignment, TrimStoreName_Array_BP(), TrimStoreName_Array_BN(), DigCap_Pin, Cap_Trimwidth, DigCap_Sample_Size, CUS_Str_MainProgram, CUS_Str_DigCapData, width_Wf, width_Wf_2S, OutWf_2s)
            For Each site In TheExec.sites.Active
                trim_Wf = OutWf_2s.divide(2 ^ Resolution_CHK)           '
                For i = 0 To TestSequenceNumber - 1
                    temp_calc(i) = trim_Wf.Element(i)
                    If gl_Disable_HIP_debug_log = False Then TheExec.Datalog.WriteComment ("**************** Site :" & site & " Reg :" & MeasValue(i).Pins.item(0) & " " & Cap_Trimwidth & " bit 2scomplement data is :" & temp_calc(i) & " ****************")
                    If DecideTrim(i) = False Then
                        BestTrimgap(i) = temp_calc(i).Subtract(TrimTarget).Abs
                        If BestTrimgap(i) < BestTrimgap_previous(i) Then
                            BestTrimgap_previous(i) = BestTrimgap(i)
                            BestVal_BP(i) = DAC_bp(i)
                            BestVal_BN(i) = DAC_bn(i)
                        End If
                        If temp_calc(i) > TrimTarget Then
                            If Trimming_Direction_Increase = True Then
                                DAC_bp(i) = DAC_bp(i) - 2 ^ j
                                DAC_bn(i) = DAC_bn(i) + 2 ^ j
                            Else
                                DAC_bp(i) = DAC_bp(i) + 2 ^ j
                                DAC_bn(i) = DAC_bn(i) - 2 ^ j
                            End If
                        Else
                            If Trimming_Direction_Increase = True Then
                                DAC_bp(i) = DAC_bp(i) + 2 ^ j
                                DAC_bn(i) = DAC_bn(i) - 2 ^ j
                            Else
                                DAC_bp(i) = DAC_bp(i) - 2 ^ j
                                DAC_bn(i) = DAC_bn(i) + 2 ^ j
                            End If
                        End If
                        If temp_calc(i) = TrimTarget Then
                            DecideTrim(i) = True
                        End If
                    End If
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

    For i = 0 To TestSequenceNumber - 1

        TempBestVal = MeasValue(i).Pins(DecomposeTestName(i))
        TestNameInput = Report_TName_From_Instance("C", DigCap_Pin.value, "", i, 0)
        TheExec.Flow.TestLimit TempBestVal, 0, 2 ^ DecomposeParseDigCapBit(i) - 1, PinName:=DecomposeTestName(i), Tname:=TestNameInput, ForceResults:=tlForceFlow, scaletype:=scaleNoScaling, formatStr:="%.0f"

    Next i
    For i = 0 To TestSequenceNumber - 1

        TempBestVal = BestVal_BP(i).Pins(TrimStoreName_Array_BP(i))
        TestNameInput = Report_TName_From_Instance("C", DigCap_Pin.value, "", i, 0)
        TheExec.Flow.TestLimit TempBestVal, 0, 2 ^ DecomposeParseDigCapBit(i) - 1, PinName:=TrimStoreName_Array_BP(i), Tname:=TestNameInput, ForceResults:=tlForceFlow, scaletype:=scaleNoScaling, formatStr:="%.0f"

        TempBestVal = BestVal_BN(i).Pins(TrimStoreName_Array_BN(i))
        TestNameInput = Report_TName_From_Instance("C", DigCap_Pin.value, "", i, 0)
        TheExec.Flow.TestLimit TempBestVal, 0, 2 ^ DecomposeParseDigCapBit(i) - 1, PinName:=TrimStoreName_Array_BN(i), Tname:=TestNameInput, ForceResults:=tlForceFlow, scaletype:=scaleNoScaling, formatStr:="%.0f"

    Next i
     
    DebugPrintFunc Pat.value
    
    Call HardIP_WriteFuncResult(, , glb_TestInstance)
    
    If Calc_Eqn <> "" Then: Call ProcessCalcEquation(Calc_Eqn)
    
    ' Check implicit alarms
    TheHdw.Alarms.Check
    
    If b_TnameConstructionTTR = False Then
        EnableHardIPTnameConstructionTTR = True
    End If
    
    Exit Function
    
errHandler:
    Call Print_Error_Message(Error_Info, "VBT_LIB_HardIP_Metrology_AP", "MetrologyICTS_DAP_PN_TRIM") 'Add ErrHandler 2023/05/29
    If AbortTest Then Exit Function Else Resume Next
End Function



