Attribute VB_Name = "VBT_LIB_Shmoo_ND"
Option Explicit

Public Function Print_Shmoo_ND_Summary(Setup_name As String)
On Error GoTo errHandler
    Dim inst As Variant
    Dim i As Long, j As Long, k As Long
    Dim o_Str As String
    Dim p As Variant, p_val As Long
    Dim val_str As String, Tname As String
    Dim opt_cnt As Long
    Dim site As Variant
    Dim HLvccMarginStr As String
    Dim HLvccForceStr As String
    Dim HLvcc_SpecNameAry() As String
    Dim HLvcc_SpecName As String
    Dim HLvcc_Pin As String
    Dim HLvcc_unit As String
    Dim PinVal As Double
    Dim VolType As String
    Dim pass_list() As Variant
    
    
    Shmoo_PreCond_Info = Shmoo_ND_Setups(Shmoo_ND_Setup_Dict.item(Setup_name)).Spec
    Shmoo_PreCond_Info_Fast = Shmoo_ND_Setups(Shmoo_ND_Setup_Dict(Setup_name)).char
    Shmoo_PreCond_Spec_N = Shmoo_ND_Setups(Shmoo_ND_Setup_Dict.item(Setup_name)).N_spec
    
    
    For Each inst In InstanceInUse.Keys ' list all instance
        If inst = "" Then Exit For
        theexec.Datalog.WriteComment vbNullString
        theexec.Datalog.WriteComment inst & "_OptimizeTiming Summary:"
        theexec.Datalog.WriteComment "Block_Type: " & Setup_name
        theexec.Datalog.WriteComment "=================================================="
        For Each site In theexec.sites
            If Shmoo_PreCond_LVCC_N <> 0 Then
                For i = 0 To Shmoo_PreCond_LVCC_N - 1
                    HLvcc_SpecName = Shmoo_PreCond_Info_LVCC(i).SpecVoltageName
                    HLvcc_SpecNameAry = Split(HLvcc_SpecName, "_")
                    HLvcc_unit = Shmoo_PreCond_Info_LVCC(i).unit
                    If Shmoo_PreCond_Info_LVCC(i).SearchType = "LVCC" Then
                        If Shmoo_PreCond_Info_LVCC(i).LVCC_Value(site) = Empty Then
                            HLvccForceStr = "N/A"
                            HLvccMarginStr = "N/A"
                            theexec.Datalog.WriteComment "Please check the setting: " & HLvcc_SpecName & " on the test flow"
                            theexec.Datalog.WriteComment HLvcc_SpecName & " LVCC Value: " & HLvccMarginStr
                        Else
                            HLvccForceStr = Unit_format((Shmoo_PreCond_Info_LVCC(i).LVCC_Value(site)), HLvcc_unit)
                            HLvccMarginStr = Unit_format(Shmoo_PreCond_Info_LVCC(i).LVCC_Value(site) - Unit_Calc(Abs(CDbl(Shmoo_PreCond_Info_LVCC(i).constraint)), HLvcc_unit), HLvcc_unit)
                            theexec.Datalog.WriteComment HLvcc_SpecName & " Force Value: " & HLvccForceStr & " with LVCC Value: " & HLvccMarginStr & " and Constraint: " & CDbl(Shmoo_PreCond_Info_LVCC(i).constraint) & HLvcc_unit
                            If Shmoo_PreCond_Info_LVCC(i).TrackingCnt > 0 Then
                                For k = 0 To UBound(Shmoo_PreCond_Info_LVCC(i).TrackingSpecN)
                                    theexec.Datalog.WriteComment HLvcc_SpecName & "'s Tracking Spec: " & Shmoo_PreCond_Info_LVCC(i).TrackingSpecN(k)
                                Next k
                            End If
                        End If
                    ElseIf Shmoo_PreCond_Info_LVCC(i).SearchType = "HVCC" Then
                        If Shmoo_PreCond_Info_LVCC(i).HVCC_Value(site) = Empty Then
                            HLvccForceStr = "N/A"
                            HLvccMarginStr = "N/A"
                            theexec.Datalog.WriteComment "Please check the setting: " & HLvcc_SpecName & " on the test flow"
                            theexec.Datalog.WriteComment HLvcc_SpecName & " HVCC Value: " & HLvccMarginStr
                        Else
                            HLvccForceStr = Unit_format((Shmoo_PreCond_Info_LVCC(i).HVCC_Value(site)), HLvcc_unit)
                            HLvccMarginStr = Unit_format(Shmoo_PreCond_Info_LVCC(i).HVCC_Value(site) + Unit_Calc(Abs(CDbl(Shmoo_PreCond_Info_LVCC(i).constraint)), HLvcc_unit), HLvcc_unit)
                            theexec.Datalog.WriteComment HLvcc_SpecName & " Force Value: " & HLvccForceStr & " with HVCC Value: " & HLvccMarginStr & " and Constraint: " & CDbl(Shmoo_PreCond_Info_LVCC(i).constraint) & HLvcc_unit
                            If Shmoo_PreCond_Info_LVCC(i).TrackingCnt > 0 Then
                                For k = 0 To UBound(Shmoo_PreCond_Info_LVCC(i).TrackingSpecN)
                                    theexec.Datalog.WriteComment HLvcc_SpecName & "'s Tracking Spec: " & Shmoo_PreCond_Info_LVCC(i).TrackingSpecN(k)
                                Next k
                            End If
                        End If
                    End If
                Next
            End If
            theexec.Datalog.WriteComment " "
    
            pass_list = Split(Result_ND(InstanceInUse(inst)).FlowStep_Constraint, ",")
            opt_cnt = 0
            For Each p In pass_list  ' list all pass condition
                p_val = CLng(p)
                For j = Shmoo_PreCond_Spec_N - 1 To 0 Step -1 'list all pass spec
                    val_str = Unit_format(Shmoo_PreCond_Val(p_val, j), Shmoo_PreCond_Info(j).unit)
                    theexec.Datalog.WriteComment site & " " & inst & " " & Shmoo_PreCond_Info(j).Spec & "opt" & opt_cnt & "_" & Shmoo_PreCond_Info(j).SpecVoltageName & " = " & val_str
                Next j
                val_str = Unit_format(Result_ND(InstanceInUse(inst)).val(site), Shmoo_PreCond_Info_Fast(0).unit)
                    theexec.Datalog.WriteComment site & " " & inst & " " & Shmoo_PreCond_Info_Fast(0).Spec & Shmoo_PreCond_Info_Fast(0).constraint & opt_cnt & "_" & Shmoo_PreCond_Info_Fast(0).SpecVoltageName & " = " & val_str
                opt_cnt = opt_cnt + 1
                theexec.Datalog.WriteComment vbNullString
            Next p
         Next site
    Next inst
    
    
    ' log in PTR for best value
    For Each inst In InstanceInUse.Keys ' list all instance
        If inst = "" Then Exit For
        With Shmoo_PreCond_Info_Fast(0)
            Tname = inst & "_" & .Spec & .constraint & "_" & .SpecVoltageName
        End With
        Dim current_context As String
        current_context = theexec.Contexts.ActiveSelection
        theexec.Datalog.WriteComment " Current Levels Sheet: " & theexec.Contexts(current_context).Sheets.PinLevels
        
        theexec.Datalog.SetDynamicTestName Tname, False
        For Each site In theexec.sites
            If Result_ND(InstanceInUse(inst)).val <> Empty Then
                theexec.Flow.TestLimit Result_ND(InstanceInUse(inst)).val
            Else
                theexec.Flow.TestLimit 0
            End If
        Next site
    Next inst
    Exit Function
errHandler:
    Call Print_Error_Message(Error_Info, "VBT_LIB_Shmoo_ND", "Print_Shmoo_ND_Summary")
    If AbortTest Then Exit Function Else Resume Next
End Function
Public Function Setup_FlowLoop(Setup_name As String)
On Error GoTo errHandler
    Dim i As Long
    Dim j As Long
    Dim k As Long
    Dim li() As PreCond_Info
    Dim Point_Count As Double
    Dim reftime As Double
    Dim o_Str As String
    Dim inst_name As String
    Dim shmoo_nd_setup1 As Shmoo_ND_Setup
    Dim ND_CharSetName As String
    Dim ShiftIn_From_Val As Double
    Dim ShiftIn_To_Val As Double
    Dim ShiftIn_Step_Val As Double
    Dim ShiftInStepNameAry() As String
    Dim tmpLVCCSetupName As String
    Dim tmpLVCCSetupNameAry() As String
    Dim tmpLVCCFrom As String
    Dim tmpLVCCTo As String
    Dim tmpLVCCStep As String
    Dim SpecType As String
    Dim m_step As Variant
    Dim tracking_Idx As Long
    Dim TrackingStepAry() As String
    Dim Shmoo_Var_tmp_string As String
    Dim Shmoo_Var_split() As String
    Dim tmp As Double
    Dim index_value As Integer

    Shmoo_ND_Info_now.Setup_name = Setup_name
    Shmoo_ND_Info_now.Setup_Idx = Shmoo_ND_Setup_Dict.item(Setup_name)
    theexec.Datalog.WriteComment "<Shmoo_ND_" & Setup_name & ">"
    Shmoo_PreCond_Info = Shmoo_ND_Setups(Shmoo_ND_Info_now.Setup_Idx).Spec
    Shmoo_PreCond_Info_Fast = Shmoo_ND_Setups(Shmoo_ND_Info_now.Setup_Idx).char
    Shmoo_PreCond_Info_LVCC = Shmoo_ND_Setups(Shmoo_ND_Info_now.Setup_Idx).Lvcc
    Shmoo_PreCond_Info_StaticV = Shmoo_ND_Setups(Shmoo_ND_Info_now.Setup_Idx).StaticV
    Shmoo_PreCond_Spec_N = Shmoo_ND_Setups(Shmoo_ND_Info_now.Setup_Idx).N_spec
    Shmoo_PreCond_LVCC_N = Shmoo_ND_Setups(Shmoo_ND_Info_now.Setup_Idx).N_lvcc
    Shmoo_PreCond_Static_N = Shmoo_ND_Setups(Shmoo_ND_Info_now.Setup_Idx).N_static
    
    li = Shmoo_PreCond_Info
    Point_Count = 1
    
    'Setup ShiftIn Freq char setting
    With Shmoo_PreCond_Info_Fast(0)
        ShiftIn_From_Val = Unit_Calc(CDbl(.from), .unit)
        ShiftIn_To_Val = Unit_Calc(CDbl(.to), .unit)
        ShiftIn_Step_Val = Unit_Calc(CDbl(.StepSize), .unit)
    End With
    
'    'Setup with Margin Shmoo
'    For Each m_step In TheExec.DevChar.Setups.Item(ShiftIn_SetupName).Margins.List
'        ND_DummyShift = TheExec.DevChar.Setups(ShiftIn_SetupName).Margins(m_step).Parameter.name
'        With TheExec.DevChar.Setups(ShiftIn_SetupName).Margins(m_step).Parameter.range
'            .from = ShiftIn_From_Val
'            .to = ShiftIn_To_Val
'            .stepSize = ShiftIn_Step_Val
'        End With
'    Next m_step
    'Setup with X Shmoo
    If Shmoo_PreCond_Info_Fast(0).constraint = "max" Then
        ND_CharSetName = "ND_ShiftIn_Freq"
    ElseIf Shmoo_PreCond_Info_Fast(0).constraint = "min" Then
        ND_CharSetName = "ND_PatGenRunTime"
    End If
    With theexec.DevChar.Setups(ND_CharSetName).Shmoo.Axes(tlDevCharShmooAxis_X).Parameter.range
        .from = ShiftIn_From_Val
        .to = ShiftIn_To_Val
        .StepSize = ShiftIn_Step_Val
    End With
    
    'Setup LVCC char setting
    If Shmoo_PreCond_LVCC_N >= 1 Then
    ReDim Lvcc_SpecName(Shmoo_PreCond_LVCC_N - 1) As String
        For i = 0 To Shmoo_PreCond_LVCC_N - 1
            tmpLVCCSetupNameAry = Split(Shmoo_PreCond_Info_LVCC(i).SpecVoltageName, "_")
            Select Case UCase(tmpLVCCSetupNameAry(UBound(tmpLVCCSetupNameAry)))
            Case "VMAIN"
                SpecType = "_VAR"
            Case "VALT"
                SpecType = "_VOP_VAR"
            End Select
            ReDim Preserve tmpLVCCSetupNameAry(UBound(tmpLVCCSetupNameAry) - 1)
            tmpLVCCSetupName = Join(tmpLVCCSetupNameAry, "_")
            Lvcc_SpecName(i) = tmpLVCCSetupName & SpecType
            With Shmoo_PreCond_Info_LVCC(i)
                tmpLVCCFrom = Unit_Calc(CDbl(.from), .unit)
                tmpLVCCTo = Unit_Calc(CDbl(.to), .unit)
                tmpLVCCStep = Unit_Calc(CDbl(.StepSize), .unit)
            End With
'            'Setup with Margin Shmoo
'            With TheExec.DevChar.Setups("ND_" & tmpLVCCSetupName & "_LVCC").Margins(tmpLVCCSetupName).Parameter.range
'                .from = tmpLVCCFrom
'                .to = tmpLVCCTo
'                .stepSize = tmpLVCCStep
'            End With
            'Setup with X Shmoo
            
'            With TheExec.DevChar.Setups("ND_" & tmpLVCCSetupName & "_LVCC").Shmoo.axes(tlDevCharShmooAxis_X)
            With theexec.DevChar.Setups("ND_" & Shmoo_PreCond_Info_LVCC(i).SpecVoltageName).Shmoo.Axes(tlDevCharShmooAxis_X)
                .Parameter.range.from = tmpLVCCFrom
                .Parameter.range.to = tmpLVCCTo
                .Parameter.range.StepSize = tmpLVCCStep
                'Setup Tracking case
                If .TrackingParameters.Count > 0 Then
                    TrackingStepAry = .TrackingParameters.list
                    For tracking_Idx = 0 To .TrackingParameters.Count - 1
                        .TrackingParameters(TrackingStepAry(tracking_Idx)).range.from = tmpLVCCFrom
                        .TrackingParameters(TrackingStepAry(tracking_Idx)).range.to = tmpLVCCTo
                    Next tracking_Idx
                End If
            End With
        Next i
    End If
    
    'for i=0 to
    
    reftime = theexec.Timer
        
    o_Str = vbNullString
    For i = Shmoo_PreCond_Spec_N - 1 To 0 Step -1
        If i = Shmoo_PreCond_Spec_N - 1 Then
            o_Str = Shmoo_PreCond_Info(i).SpecVoltageName & "(" & Shmoo_PreCond_Info(i).Spec & ")"
        Else
            o_Str = o_Str & "," & Shmoo_PreCond_Info(i).SpecVoltageName & "(" & Shmoo_PreCond_Info(i).Spec & ")"
        End If
    Next i
       
    ' Store Value to Shmoo Variable
    Shmoo_Var_split = Split(Shmoo_Var, ",")
    For i = 0 To Shmoo_PreCond_Count - 1
        For j = 0 To Shmoo_PreCond_Spec_N - 1
            If (li(j).SpecVoltageName = Shmoo_Var_split(i)) Then
                Shmoo_Var_tmp_string = "Shmoo_PreCond_LoopVal" & i
                ' If li(j).from has Shmoo Variable, the followed process will be error
                StoreVal_df Shmoo_Var_tmp_string, CDbl(li(j).from)
            End If
        Next j
    Next i
    
    ' Decide Number of flow shmoo point
    For i = Shmoo_PreCond_Spec_N - 1 To 0 Step -1
        tmp = Decide_loop_Count(i, li)
        Point_Count = Point_Count * tmp
    Next i
    
    Call SetVarValue(li)
    
    theexec.Flow.var("FlowLoopMax").value = Int(Point_Count) - 1

    'Setup PreCondition Value for each shmoo point
    ReDim Shmoo_PreCond_Val(Point_Count - 1, Shmoo_PreCond_Spec_N - 1)
    
    For i = 0 To Point_Count - 1
        For j = Shmoo_PreCond_Spec_N - 1 To 0 Step -1
            index_value = Decide_index_value(i, j, li)
            Shmoo_PreCond_Val(i, j) = Get_Shmoo_Value(li(j), index_value)
            For k = 0 To UBound(Shmoo_Var_split)
                If li(j).SpecVoltageName = Shmoo_Var_split(k) Then
                    StoreVal_df "Shmoo_PreCond_LoopVal" & k, Shmoo_PreCond_Val(i, j)
                End If
            Next k
        Next j
    Next i
    
    Debug.Print "Test Time(" & Point_Count & "pt):" & theexec.Timer(reftime)

    Init_Result
    Exit Function
errHandler:
    Call Print_Error_Message(Error_Info, "VBT_LIB_Shmoo_ND", "Setup_FlowLoop")
    If AbortTest Then Exit Function Else Resume Next
End Function


Public Function Read_Shmoo_ND_Table()
On Error GoTo errHandler
    Dim ws_ND As Worksheet
    Dim wb As Workbook
    Dim Setup_sheet_name As String
    Dim Row As Long, Col As Long, lRow As Long, lCol As Long
    Dim ND_setup_info(100) As Shmoo_ND_setup_info
    Dim Setup_name As String
    Dim i As Long, j As Long
    Dim N_lvcc As Long
    Dim N_row As Long
    Dim N_static As Long
    
    Setup_sheet_name = "Shmoo_ND"
    Shmoo_PreCond_Count = 0
        
    Application.ScreenUpdating = False
    
    Set wb = Application.ActiveWorkbook
    Set ws_ND = wb.Sheets(Setup_sheet_name)
    ws_ND.Activate
    
    ' Get Info for Shmoo_ND
    Row = 1
    Shmoo_PreCond_Setup_N = 0
    While ws_ND.Application.WorksheetFunction.CountA(ws_ND.Rows(Row)) <> 0
        ' start Setup
        If LCase(ws_ND.Cells(Row, 1)) Like "spec" And LCase(ws_ND.Cells(Row, ShmooNDSetupCol.Default)) Like "default" Then
            N_lvcc = 0
            N_static = 0
            ND_setup_info(Shmoo_PreCond_Setup_N).name = Setup_name
            ND_setup_info(Shmoo_PreCond_Setup_N).Nth_setup = Shmoo_PreCond_Setup_N
            ND_setup_info(Shmoo_PreCond_Setup_N).start_row = Row + 1
        End If
        ' end Setup
        If Row <> 1 And LCase(ws_ND.Cells(Row, ShmooNDSetupCol.SpecVoltageName)) Like "" _
        And LCase(ws_ND.Cells(Row, ShmooNDSetupCol.Default)) Like "" Then
            ND_setup_info(Shmoo_PreCond_Setup_N).end_row = Row - 1
            ND_setup_info(Shmoo_PreCond_Setup_N).N_lvcc = N_lvcc
            ND_setup_info(Shmoo_PreCond_Setup_N).N_static = N_static
            ND_setup_info(Shmoo_PreCond_Setup_N).N_row = ND_setup_info(Shmoo_PreCond_Setup_N).end_row - ND_setup_info(Shmoo_PreCond_Setup_N).start_row + 1
            ND_setup_info(Shmoo_PreCond_Setup_N).N_spec = ND_setup_info(Shmoo_PreCond_Setup_N).N_row - N_lvcc - N_static - 1
            Shmoo_PreCond_Setup_N = Shmoo_PreCond_Setup_N + 1
        End If
        'HLVCC
        If LCase(ws_ND.Cells(Row, ShmooNDSetupCol.Default)) Like "*lvcc" Then
            N_lvcc = N_lvcc + 1
        End If
        If LCase(ws_ND.Cells(Row, ShmooNDSetupCol.from)) = "" And LCase(ws_ND.Cells(Row, ShmooNDSetupCol.to_1)) = "" Then
            N_static = N_static + 1
        End If
        Setup_name = ws_ND.Cells(Row, 1)
        Row = Row + 1
    Wend
    ND_setup_info(Shmoo_PreCond_Setup_N).end_row = Row - 1
    ND_setup_info(Shmoo_PreCond_Setup_N).N_lvcc = N_lvcc
    ND_setup_info(Shmoo_PreCond_Setup_N).N_static = N_static
    ND_setup_info(Shmoo_PreCond_Setup_N).N_row = ND_setup_info(Shmoo_PreCond_Setup_N).end_row - ND_setup_info(Shmoo_PreCond_Setup_N).start_row + 1
    ND_setup_info(Shmoo_PreCond_Setup_N).N_spec = ND_setup_info(Shmoo_PreCond_Setup_N).N_row - N_lvcc - N_static - 1
    Shmoo_PreCond_Setup_N = Shmoo_PreCond_Setup_N + 1
    
    
    ' fill in Shmoo_PreCond_Info for each setup
    ReDim Shmoo_ND_Setups(Shmoo_PreCond_Setup_N - 1)
    ReDim LOOP_DEPTH(Shmoo_PreCond_Setup_N - 1)
    Shmoo_ND_Setup_Dict.RemoveAll
    InstanceInUsePerSite.RemoveAll
    Row = 1
    For i = 0 To Shmoo_PreCond_Setup_N - 1
        Setup_name = ND_setup_info(i).name
        LOOP_DEPTH(i) = ND_setup_info(i).N_spec
'        ReDim Shmoo_ND_Setups(i).spec(ND_setup_info(i).N_spec - 1)
        ReDim Shmoo_ND_Setups(i).Spec(LOOP_DEPTH(i) - 1)
        ReDim Shmoo_ND_Setups(i).char(0)
        If ND_setup_info(i).N_lvcc > 0 Then ReDim Shmoo_ND_Setups(i).Lvcc(ND_setup_info(i).N_lvcc - 1)
        If ND_setup_info(i).N_static > 0 Then ReDim Shmoo_ND_Setups(i).StaticV(ND_setup_info(i).N_static - 1)
        Row = Row + 2 ' skip the row of the setup name and header
        ' spec
        Shmoo_ND_Spec_Dict.RemoveAll
        For j = 0 To ND_setup_info(i).N_spec - 1
            Call Shmoo_ND_TableParseRow("spec", Row, Shmoo_ND_Setups(i).Spec(ND_setup_info(i).N_spec - 1 - j), ws_ND, ND_setup_info(i), j)
            Row = Row + 1
        Next j
        Shmoo_ND_Setups(i).N_spec = ND_setup_info(i).N_spec
        ' 1D char
        Call Shmoo_ND_TableParseRow("char", Row, Shmoo_ND_Setups(i).char(0), ws_ND, ND_setup_info(i), 0)
        Row = Row + 1
        ' lvcc and static power
        If ND_setup_info(i).N_static > 0 Then
            Dim idx_lvcc As Long: idx_lvcc = 0
            Dim idx_static As Long: idx_static = 0
            For j = 0 To ND_setup_info(i).N_lvcc + ND_setup_info(i).N_static - 1

                If LCase(ws_ND.Cells(Row, ShmooNDSetupCol.from)) = "" And LCase(ws_ND.Cells(Row, ShmooNDSetupCol.to_1)) = "" Then
                    Call Shmoo_ND_TableParseRow("static", Row, Shmoo_ND_Setups(i).StaticV(idx_static), ws_ND, ND_setup_info(i), ND_setup_info(i).N_spec - 1 - idx_static)
                    idx_static = idx_static + 1
                Else
                    Call Shmoo_ND_TableParseRow("lvcc", Row, Shmoo_ND_Setups(i).Lvcc(idx_lvcc), ws_ND, ND_setup_info(i), ND_setup_info(i).N_spec - 1 - idx_lvcc)
                    idx_lvcc = idx_lvcc + 1
                End If
                Row = Row + 1
            Next j
        Else
            For j = 0 To ND_setup_info(i).N_lvcc - 1
                Call Shmoo_ND_TableParseRow("lvcc", Row, Shmoo_ND_Setups(i).Lvcc(j), ws_ND, ND_setup_info(i), ND_setup_info(i).N_spec - 1 - j)
                Row = Row + 1
            Next j
        End If
        Shmoo_ND_Setups(i).N_lvcc = ND_setup_info(i).N_lvcc
        Shmoo_ND_Setups(i).N_static = ND_setup_info(i).N_static
        If Shmoo_ND_Setup_Dict.Exists(Setup_name) Then
            Call Print_Error_Message(Error_Info, "VBT_LIB_Shmoo_ND", "Read_Shmoo_ND_Table", "Duplicated Shmoo ND setup " & Setup_name)
'            Call TheExec.ErrorLogMessage("Duplicated Shmoo ND setup " & Setup_name)
        Else
            Shmoo_ND_Setup_Dict.Add Setup_name, i
        End If
    Next i
        
    Application.ScreenUpdating = True
        
    Exit Function
errHandler:
    Call Print_Error_Message(Error_Info, "VBT_LIB_Shmoo_ND", "Read_Shmoo_ND_Table")
    If AbortTest Then Exit Function Else Resume Next
End Function

Public Function Calc_Pass_List()
On Error GoTo errHandler
    Dim inst As Variant
    Dim i As Long, j As Long, o_Str As String
    Dim p As Variant, p_val As Long
    Dim val_str As String, Tname As String
    Dim opt_cnt As Long
    Dim Setup_name As String
    Dim pass_list() As Variant
    Dim site As Variant 'Carter, 20240304
    Result_Pass_Dict.RemoveAll
    
    Setup_name = Shmoo_ND_Info_now.Setup_name
    Shmoo_ND_Info_now.Setup_name_CalcPass = Setup_name
    Shmoo_PreCond_Info = Shmoo_ND_Setups(Shmoo_ND_Setup_Dict.item(Setup_name)).Spec
    Shmoo_PreCond_Info_Fast = Shmoo_ND_Setups(Shmoo_ND_Setup_Dict(Setup_name)).char
    Shmoo_PreCond_Spec_N = UBound(Shmoo_ND_Setups(Shmoo_ND_Setup_Dict.item(Setup_name)).Spec) + 1
    For Each inst In InstanceInUse.Keys ' list all instance
        If inst = "" Then Exit For
        For Each site In theexec.sites
            pass_list = Split(Result_ND(InstanceInUse(inst)).FlowStep_Pass, ",")
            For Each p In pass_list  ' list all pass condition
                If p <> "" Then
                    If Result_Pass_Dict.Exists(p) = False Then
                        Result_Pass_Dict.Add p, CLng(p)
                    End If
                End If
            Next p
        Next site
    Next inst
    Exit Function
errHandler:
    Call Print_Error_Message(Error_Info, "VBT_LIB_Shmoo_ND", "Calc_Pass_List")
    If AbortTest Then Exit Function Else Resume Next
End Function

Public Function Skip_All_Fail_Start()
On Error GoTo errHandler
    ' skip until it find a pass flow step
    ' Need to put Skip_All_Fail instance right after the "for" of the flow for loop
    While Result_Pass_Dict.Exists(str(theexec.Flow.var("FlowLoop").value)) = False _
            And theexec.Flow.var("FlowLoop").value < theexec.Flow.var("FlowLoopMax").value
        theexec.Datalog.WriteComment "Skip Condition " & Replace(Shmoo_Precond_Str(theexec.Flow.var("FlowLoop").value, vbNullString), ".", "p")
        theexec.Flow.var("FlowLoop").value = theexec.Flow.var("FlowLoop").value + 1
    Wend

    Exit Function
errHandler:
    Call Print_Error_Message(Error_Info, "VBT_LIB_Shmoo_ND", "Skip_All_Fail_Start")
    If AbortTest Then Exit Function Else Resume Next
End Function

Public Function Skip_All_Fail_End()
On Error GoTo errHandler

    ' Prevent last flow step's execution when it is in all fail list
    ' Need to put Skip_All_Fail instance right before the "next" of the flow for loop
    While Result_Pass_Dict.Exists(str(theexec.Flow.var("FlowLoop").value + 1)) = False _
            And theexec.Flow.var("FlowLoop").value + 1 <= theexec.Flow.var("FlowLoopMax").value
        theexec.Datalog.WriteComment "Skip Condition " & Replace(Shmoo_Precond_Str(theexec.Flow.var("FlowLoop").value + 1, vbNullString), ".", "p")
        theexec.Flow.var("FlowLoop").value = theexec.Flow.var("FlowLoop").value + 1
    Wend

    Exit Function
errHandler:
    Call Print_Error_Message(Error_Info, "VBT_LIB_Shmoo_ND", "Skip_All_Fail_End")
    If AbortTest Then Exit Function Else Resume Next
End Function
