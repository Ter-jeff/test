Attribute VB_Name = "LIB_Shmoo_ND"
Option Explicit
Public LOOP_DEPTH() As Integer
Public Const Max_Instance_N = 1000
Public Shmoo_PreCond_Count As Integer
Public Shmoo_Var As String
Type Shmoo_ND_setup_info
    name As String
    Nth_setup As Long
    start_row As Long
    end_row As Long
    N_row As Long
    N_lvcc As Long
    N_spec As Long
    N_static As Long
End Type
Enum flag_par
    flagFalse = 0
    flagFrom = 1
    flagTo = 2
    flagFromTo = 3
    flagPar = 4
End Enum
Enum ShmooNDSetupCol
    Spec = 1
    SpecVoltageName = 2
    Default = 3
    from = 4
    to_1 = 5
    stepSizenum = 6
    constraint = 7
    COMMENT = 8
End Enum
Enum PreCond_Type
    Spec_AC = 0
    Spec_DC = 1
    VDD_MAIN = 2
    VDD_ALT = 2
End Enum
Type PreCond_Info
    Spec As String
    SpecVoltageName As String
    Var_name As String
    type As PreCond_Type
    Default As String
    from As String
    to As String
    StepSize As String
    constraint As String
    COMMENT As String
    start As Double ' for creating flow loop sequence
    stop As Double
    step As Double
    step_count As Double
    step_count_ave As Double
    from_eval As String
    par_flag As Integer
    to_eval As String
    stepSize_eval As String
    unit As String
    LVCC_Value As New SiteDouble
    HVCC_Value As New SiteDouble
    Static_Value As Double
    SearchType As String
    TrackingSpecN() As String
    TrackingCnt As Long
End Type
Type Shmoo_ND_Setup
    Spec() As PreCond_Info
    char() As PreCond_Info
    Lvcc() As PreCond_Info
    StaticV() As PreCond_Info
    N_lvcc As Long
    N_spec As Long
    N_static As Long
End Type
Type Result_ND_Info
     FlowStep_Constraint As New SiteVariant 'Save list of precondition point with max or min transition val
     FlowStep_Pass As New SiteVariant 'Save list of passing precondition point
     val As New SiteDouble  'Transition value
End Type
Type Shmoo_ND_Info
    Setup_name As String
    ''
    Setup_Idx As Long
    N_InstanceInUse As Long
    Setup_name_CalcPass As String
End Type
Public InstanceInUse As New Dictionary
Public InstanceInUsePerSite As New Dictionary
Public Shmoo_ND_Info_now As Shmoo_ND_Info
Public Result_ND() As Result_ND_Info
''''
Public Shmoo_PreCond_Spec_N As Long
Public Shmoo_PreCond_Setup_N As Long
Public Shmoo_PreCond_LVCC_N As Long
Public Shmoo_PreCond_Static_N As Long
'''''
Public Shmoo_PreCond_Val() As Double
Public Shmoo_PreCond_Info() As PreCond_Info
Public Shmoo_PreCond_Info_Fast() As PreCond_Info
Public Shmoo_PreCond_Info_LVCC() As PreCond_Info
Public Shmoo_PreCond_Info_StaticV() As PreCond_Info
Public Shmoo_ND_Setups() As Shmoo_ND_Setup
Public Shmoo_ND_Setup_Dict As New Dictionary
Public Shmoo_ND_Spec_Dict As New Dictionary
Public Result_Pass_Dict As New Dictionary


Public Shmoo_Result() As New SiteLong
Public Accumulated_Shmoo_Result() As New SiteLong
Public Flag_Accumulated_Shmoo_First As Boolean
Enum Result_ND_Type
    NotTest = 0
    fail = 1
    Pass = 2
End Enum
Public result() As Result_ND_Type
Public Lvcc_SpecName() As String
Public ND_DummyShift As String
Public RefTimer As Double

Public Function Unit_format(val As Double, unit As String) As String
On Error GoTo errHandler
Dim val_str As String
    Select Case unit
        Case "%": val_str = Format(val * 100, "0.0") & " %"
        Case "Mhz": val_str = Format(val / 1000000#, "0.0") & " Mhz"
        Case "mV": val_str = Format(val * 1000, "0.0") & " mV"
        Case "Sec": val_str = Format(val, "0.000000") & " Sec"
        Case Else:
            Call Print_Error_Message(Error_Info, "LIB_Shmoo_ND", "Unit_format", "unit " & unit & "in Unit_format is not supported")
'            TheExec.ErrorLogMessage "unit " & unit & "in Unit_format is not supported"
    End Select
    Unit_format = val_str
    Exit Function
errHandler:
    Call Print_Error_Message(Error_Info, "LIB_Shmoo_ND", "Unit_format")
    If AbortTest Then Exit Function Else Resume Next
End Function
Public Function Unit_Calc(val As Double, unit As String) As Double
On Error GoTo errHandler
    Select Case unit
        Case "%": Unit_Calc = val
        Case "Mhz": Unit_Calc = val * 1000000
        Case "mV": Unit_Calc = val / 1000
        Case "Sec": Unit_Calc = Format(val, "0.000000")
        Case Else:
            Call Print_Error_Message(Error_Info, "LIB_Shmoo_ND", "Unit_Calc", "Unit " & unit & " is not supported in Unit_Calc")
'            TheExec.ErrorLogMessage "Unit " & unit & " is not supported in Unit_Calc"
    End Select
    Exit Function
errHandler:
    Call Print_Error_Message(Error_Info, "LIB_Shmoo_ND", "Unit_Calc")
    If AbortTest Then Exit Function Else Resume Next
End Function
Function Tokenizer(ByVal TokenString As String, TokenSeparators As String)
On Error GoTo errHandler
    Dim i As Long
    Dim seperator_ary() As String
    ' Get boundaries of the separators array
    seperator_ary = Split(TokenSeparators, " ")
    
    ' Replace every separator by a NULL-character
    For i = 0 To UBound(seperator_ary)
        TokenString = Replace(TokenString, seperator_ary(i), vbNullChar)
    Next i
    ' Prevent empty tokens by multiple NULL-characters
    While InStr(TokenString, vbNullChar & vbNullChar) > 0
        TokenString = Replace(TokenString, vbNullChar & vbNullChar, vbNullChar)
    Wend
    ' Prevent from first empty token
    If left(TokenString, 1) = vbNullChar Then TokenString = mid(TokenString, 2)
    ' Prevent from last empty token
    If right(TokenString, 1) = vbNullChar Then TokenString = left(TokenString, Len(TokenString) - 1)
    ' Split string and return value
    Tokenizer = Split(TokenString, vbNullChar)
    Exit Function
errHandler:
    Call Print_Error_Message(Error_Info, "LIB_Shmoo_ND", "Tokenizer")
    If AbortTest Then Exit Function Else Resume Next
End Function
Public Function Decide_index_value(ByVal Point_number As Integer, ByVal Spec_index As Integer, li() As PreCond_Info)
On Error GoTo errHandler
    Dim result As Double
    Dim i As Integer
    Dim loop_in As Double
    Dim loop_out As Double
    Dim tmp As Double
    Dim li_par_index As Integer
    
    Decide_index_value = Point_number
    loop_in = 1
    loop_out = 1
    If li(Spec_index).par_flag = flag_par.flagFrom Then
        li(Spec_index).start = Unit_Calc(Evaluate(li(Spec_index).from_eval), li(Spec_index).unit)
        li(Spec_index).step_count = Celling((li(Spec_index).stop - li(Spec_index).start) / li(Spec_index).step + 1)
        For i = UBound(li) To 0 Step -1
            If (li(i).SpecVoltageName = li(Spec_index).Var_name) Then
                li_par_index = i
                Exit For
            End If
        Next i
        For i = Spec_index To 0 Step -1
            loop_out = loop_out * li(i).step_count_ave
        Next i
        loop_out = loop_out * li(li_par_index).step_count_ave
        
        While CInt(Decide_index_value) >= CInt(loop_out)
            Decide_index_value = Decide_index_value - loop_out
        Wend
        loop_out = (loop_out / li(li_par_index).step_count_ave) / li(Spec_index).step_count_ave
        
        For i = 0 To li(li_par_index).step_count_ave - 1
            tmp = Get_Shmoo_Value(li(li_par_index), i)
            loop_in = Celling((li(Spec_index).stop - CDbl(li(Spec_index).from) - tmp) / li(Spec_index).step + 1)
            loop_in = loop_in * loop_out
            If (CInt(loop_in - 1) < CInt(Decide_index_value + 1)) Then
                li(li_par_index).par_flag = li(li_par_index).par_flag + 1
            End If
            If (CInt(Decide_index_value) > CInt(loop_in - 1)) Then
                Decide_index_value = Decide_index_value - loop_in
            Else
                Exit For
            End If
        Next i
    ElseIf li(Spec_index).par_flag = flag_par.flagTo Then
        li(Spec_index).stop = Unit_Calc(Evaluate(li(Spec_index).to_eval), li(Spec_index).unit)
        li(Spec_index).step_count = Celling((li(Spec_index).stop - li(Spec_index).start) / li(Spec_index).step + 1)
        For i = UBound(li) To 0 Step -1
            If (li(i).SpecVoltageName = li(Spec_index).Var_name) Then
                li_par_index = i
                Exit For
            End If
        Next i
        For i = Spec_index To 0 Step -1
            loop_out = loop_out * li(i).step_count_ave
        Next i
        loop_out = loop_out * li(li_par_index).step_count_ave
        
        While CInt(Decide_index_value) >= CInt(loop_out)
            Decide_index_value = Decide_index_value - loop_out
        Wend
        loop_out = (loop_out / li(li_par_index).step_count_ave) / li(Spec_index).step_count_ave
                
        For i = 0 To li(li_par_index).step_count_ave - 1
            tmp = Get_Shmoo_Value(li(li_par_index), i)
            loop_in = Celling((CDbl(li(Spec_index).to) - li(Spec_index).start + tmp) / li(Spec_index).step + 1)
            loop_in = loop_in * loop_out
            If (CInt(loop_in - 1) < CInt(Decide_index_value + 1)) Then
                li(li_par_index).par_flag = li(li_par_index).par_flag + 1
            End If
            If (CInt(Decide_index_value) > CInt(loop_in - 1)) Then
                Decide_index_value = Decide_index_value - loop_in
            Else
                Exit For
            End If
        Next i
    ElseIf li(Spec_index).par_flag = flag_par.flagFromTo Then
        li(Spec_index).start = Unit_Calc(Evaluate(li(Spec_index).from_eval), li(Spec_index).unit)
        li(Spec_index).stop = Unit_Calc(Evaluate(li(Spec_index).to_eval), li(Spec_index).unit)
        
    ElseIf li(Spec_index).par_flag >= flag_par.flagPar Then
        If (li(Spec_index).par_flag >= flag_par.flagPar + li(Spec_index).step_count_ave) Then
            li(Spec_index).par_flag = li(Spec_index).par_flag - li(Spec_index).step_count_ave
        End If
        Decide_index_value = li(Spec_index).par_flag - flag_par.flagPar
        li(Spec_index).par_flag = flag_par.flagPar
        Exit Function
    Else
        Call Print_Error_Message(Warning_Info, "LIB_Shmoo_ND", "Check the ND shmoo type!")
'        TheExec.Datalog.WriteComment "<Warning> Check the ND shmoo type!"
    End If
    
    tmp = 1
    If Spec_index = 0 Then
        Decide_index_value = Decide_index_value Mod li(Spec_index).step_count
    Else
        For i = Spec_index To 1 Step -1
            tmp = tmp * li(i - 1).step_count_ave
            'Decide_index_value = Fix(CInt(Decide_index_value) / li(i - 1).step_count_ave)
        Next i
        Decide_index_value = Fix(Decide_index_value / tmp)
        Decide_index_value = Decide_index_value Mod li(Spec_index).step_count
    End If
    Exit Function
errHandler:
    Call Print_Error_Message(Error_Info, "LIB_Shmoo_ND", "Decide_index_value")
    If AbortTest Then Exit Function Else Resume Next
End Function
Public Function Celling(input_value As Double) As Integer
On Error GoTo errHandler
    ' For Double resoultion issue
    Celling = IIf(Int(input_value / 0.0001) > Int(input_value) / 0.0001, Int(input_value) + 1, Int(input_value))
    Exit Function
errHandler:
    Call Print_Error_Message(Error_Info, "LIB_Shmoo_ND", "Celling")
    If AbortTest Then Exit Function Else Resume Next
End Function
Public Function Get_Shmoo_Value(li As PreCond_Info, li_index As Integer) As Double
On Error GoTo errHandler
    Dim tmp As Double
    
    tmp = li.start + li.step * li_index
    Get_Shmoo_Value = IIf(tmp > li.stop, li.stop, tmp)
Exit Function
errHandler:
    Call Print_Error_Message(Error_Info, "LIB_Shmoo_ND", "Get_Shmoo_Value")
    If AbortTest Then Exit Function Else Resume Next
End Function
Public Function SetVarValue(li() As PreCond_Info)
On Error GoTo errHandler
    Dim i As Integer
    Dim var_value As Double
    Dim li_array() As String
    
    For i = 0 To UBound(li)
        If li(i).par_flag = 1 Then
            li_array = Tokenizer(li(i).from, "+ - * / ( )")
            li(i).from = CStr(li(i).start)
            li(i).Var_name = li_array(0)
        ElseIf li(i).par_flag = 2 Then
            li_array = Tokenizer(li(i).to, "+ - * / ( )")
            li(i).to = CStr(li(i).stop)
            li(i).Var_name = li_array(0)
        ElseIf li(i).par_flag = 3 Then
            li_array = Tokenizer(li(i).from, "+ - * / ( )")
            li(i).from = CStr(li(i).start)
            li(i).to = CStr(li(i).stop)
            li(i).Var_name = li_array(0)
        End If
    Next i
    Exit Function
errHandler:
    Call Print_Error_Message(Error_Info, "LIB_Shmoo_ND", "SetVarValue")
    If AbortTest Then Exit Function Else Resume Next
End Function
Public Function Decide_loop_Count(Nth_D As Long, li() As PreCond_Info)
On Error GoTo errHandler
    Dim li_from_array() As String
    Dim dic_from_array() As String
    Dim li_to_array() As String
    Dim dic_to_array() As String
    Dim i, j As Integer
    li(Nth_D).start = Unit_Calc(Evaluate(li(Nth_D).from_eval), li(Nth_D).unit)
    li(Nth_D).stop = Unit_Calc(Evaluate(li(Nth_D).to_eval), li(Nth_D).unit)
    li(Nth_D).step = Abs(Unit_Calc(Evaluate(li(Nth_D).stepSize_eval), li(Nth_D).unit))
    li(Nth_D).par_flag = flag_par.flagFalse
    
    li_from_array = Tokenizer(li(Nth_D).from, "+ - * / ( )")
    li_to_array = Tokenizer(li(Nth_D).to, "+ - * / ( )")
    
    
    If (UBound(li_from_array) = 0 And UBound(li_to_array) = 0) Then ' No variable
        li(Nth_D).step_count = Celling((li(Nth_D).stop - li(Nth_D).start) / li(Nth_D).step + 1)
        li(Nth_D).step_count_ave = li(Nth_D).step_count
    ' Shmoo Var at li.From
    ElseIf (UBound(li_from_array) > 0 And UBound(li_to_array) = 0) Then
        For i = 0 To UBound(li)
            If (li(i).SpecVoltageName = li_from_array(0)) Then
                li(i).par_flag = flag_par.flagPar
                Exit For
            End If
        Next i
        For j = 0 To li(i).step_count_ave - 1
            li(Nth_D).step_count_ave = Celling((li(Nth_D).stop - li(Nth_D).start - Get_Shmoo_Value(li(i), j)) / li(Nth_D).step + 1) + li(Nth_D).step_count_ave
        Next j
        li(Nth_D).step_count_ave = li(Nth_D).step_count_ave / li(i).step_count_ave
        li(Nth_D).step_count = Celling((li(Nth_D).stop - li(Nth_D).start) / li(Nth_D).step + 1)
        li(Nth_D).par_flag = flag_par.flagFrom
    ' Shmoo Var at li.To
    ElseIf (UBound(li_from_array) = 0 And UBound(li_to_array) > 0) Then
        For i = 0 To UBound(li)
            If (li(i).SpecVoltageName = li_to_array(0)) Then
                li(i).par_flag = flag_par.flagPar
                Exit For
            End If
        Next i
        For j = 0 To li(i).step_count_ave - 1
            li(Nth_D).step_count_ave = Celling((li(Nth_D).stop - li(Nth_D).start - Get_Shmoo_Value(li(i), j)) / li(Nth_D).step + 1) + li(Nth_D).step_count_ave
        Next j
        li(Nth_D).step_count_ave = li(Nth_D).step_count_ave / li(i).step_count_ave
        li(Nth_D).step_count = Celling((li(Nth_D).stop - li(Nth_D).start) / li(Nth_D).step + 1)
        li(Nth_D).par_flag = flag_par.flagTo
    Else
        If Not (li_from_array(0) = li_to_array(0)) Then
            Debug.Print "Note!! The parameter must be same!!"
            MsgBox "Note! The Shmoo Variable in Spec must be same."
            Resume Next
        End If
        li(Nth_D).step_count = (li(Nth_D).stop - li(Nth_D).start) / li(Nth_D).step + 1
        li(Nth_D).step_count_ave = li(Nth_D).step_count
        li(Nth_D).par_flag = flag_par.flagFromTo
    End If
    Decide_loop_Count = li(Nth_D).step_count_ave
    Exit Function
errHandler:
    Call Print_Error_Message(Error_Info, "LIB_Shmoo_ND", "Decide_loop_Count")
    If AbortTest Then Exit Function Else Resume Next
End Function
Public Function Shmoo_ND_TableParseRow(RowType As String, Row As Long, PreCond_Info_org As PreCond_Info, ws_ND As Worksheet, setup_info As Shmoo_ND_setup_info, nth_spec As Long)
On Error GoTo errHandler

    With PreCond_Info_org
        .Spec = Replace(ws_ND.Cells(Row, ShmooNDSetupCol.Spec), " ", vbNullString)
        .SpecVoltageName = Replace(ws_ND.Cells(Row, ShmooNDSetupCol.SpecVoltageName), " ", vbNullString)
        Shmoo_ND_Spec_Dict.Add .SpecVoltageName, setup_info.N_spec - nth_spec - 1
        If UCase(.SpecVoltageName) Like "VDD*" Then
            .type = PreCond_Type.Spec_DC
        Else
            .type = PreCond_Type.Spec_AC
        End If
        If LCase(RowType) = "static" Then
        .Default = Replace(ws_ND.Cells(Row, ShmooNDSetupCol.Default), " ", vbNullString)
        End If
        .from = Replace(ws_ND.Cells(Row, ShmooNDSetupCol.from), " ", vbNullString)
        .to = Replace(ws_ND.Cells(Row, ShmooNDSetupCol.to_1), " ", vbNullString)
        .StepSize = Replace(ws_ND.Cells(Row, ShmooNDSetupCol.stepSizenum), " ", vbNullString)
        .constraint = Replace(ws_ND.Cells(Row, ShmooNDSetupCol.constraint), " ", vbNullString)
        If LCase(.constraint) = "maximize" Then
            .constraint = Replace(LCase(.constraint), "maximize", "max")
        ElseIf LCase(.constraint) = "minimize" Then
            .constraint = Replace(LCase(.constraint), "minimize", "min")
        End If
        'unit
        .unit = Replace(LCase(ws_ND.Cells(Row, ShmooNDSetupCol.COMMENT)), " ", vbNullString)
        If InStr(.unit, "%") Then
            .unit = "%"
        ElseIf InStr(.unit, "mhz") Then
            .unit = "Mhz"
        ElseIf InStr(.unit, "mv") Then
            .unit = "mV"
        ElseIf InStr(.unit, "sec") Then
            .unit = "Sec"
        Else
            Call Print_Error_Message(Error_Info, "LIB_Shmoo_ND", "Shmoo_ND_TableParseRow", "Unit " & ws_ND.Cells(Row, ShmooNDSetupCol.COMMENT) & " in Shmoo_ND_TableParseRow is not recognized !")
'            TheExec.ErrorLogMessage ("Unit " & ws_ND.Cells(Row, ShmooNDSetupCol.comment) & " in Shmoo_ND_TableParseRow is not recognized !")
        End If
        .from_eval = Shmoo_ND_TableParseCell(.from, .unit)
        .to_eval = Shmoo_ND_TableParseCell(.to, .unit)
        .stepSize_eval = Shmoo_ND_TableParseCell(.StepSize, .unit)
    End With
    Exit Function
errHandler:
    Call Print_Error_Message(Error_Info, "LIB_Shmoo_ND", "Shmoo_ND_TableParseRow")
    If AbortTest Then Exit Function Else Resume Next
End Function
Public Function Shmoo_ND_TableParseCell(org As String, unit As String) As String
On Error GoTo errHandler
    Dim modified As String, i, j As Long
    Dim token_ary() As String
    Dim Shmoo_PreCond_arr() As String
    Dim tmp As String
    Dim var_flag As Boolean
    Dim var_index As Integer
    tmp = org
    Shmoo_PreCond_arr = Split(Shmoo_Var, ",")
    token_ary = Tokenizer(org, "+ - * / ( )")
    var_flag = False
    var_index = 0
    For i = 0 To UBound(token_ary)
        If Shmoo_ND_Spec_Dict.Exists(token_ary(i)) Then
            For j = 0 To UBound(Shmoo_PreCond_arr)
                If Shmoo_PreCond_arr(j) = token_ary(i) Then
                    var_flag = True
                    var_index = j
                End If
            Next j
            If Not var_flag Then
                Create_df "Shmoo_PreCond_LoopVal" & Shmoo_PreCond_Count
                Shmoo_PreCond_Count = Shmoo_PreCond_Count + 1
                Shmoo_Var = IIf(Shmoo_Var = "", token_ary(i), Shmoo_Var & "," & token_ary(i))
                var_index = Shmoo_PreCond_Count - 1
            End If
            tmp = Replace(org, token_ary(i), "Shmoo_PreCond_LoopVal" & var_index)
        End If
    Next i
    Shmoo_ND_TableParseCell = "=" & tmp
    Exit Function
errHandler:
    Call Print_Error_Message(Error_Info, "LIB_Shmoo_ND", "Shmoo_ND_TableParseCell")
    If AbortTest Then Exit Function Else Resume Next
End Function
Public Function Show_Margin_Result(argc As Integer, argv() As String) As Long
On Error GoTo errHandler
    Dim DevChar_StepName As Variant
    Dim ShmooAxis As Variant
    Dim x_steps As Long
    Dim site As Variant
    Dim x1 As Long
    Dim result_str As String
    Dim result_tmp As String
    Dim Transition_val As New SiteDouble, Transition_str As New SiteVariant
    Dim head_str As String
    Dim Nth_point As Long
    Dim Total_point As Long
    Dim inst_name As String
    Dim TransCnt As Long
    Dim p_max As Double
    Dim p_min As Double
    Dim debug_var As String
    Dim ExecTime As Double
    Dim SitePF As New SiteBoolean
    Dim AllSitePassChecker As Boolean
    Dim DevChar_Setup As String
    
    AllSitePassChecker = True
    If Shmoo_PreCond_Info_Fast(0).constraint = "min" Then
        ExecTime = TheExec.Timer(RefTimer)
    End If
    
    inst_name = TheExec.DataManager.instancename
    DevChar_Setup = TheExec.DevChar.Setups.ActiveSetupName
    Nth_point = TheExec.flow.var("FlowLoop").value
    Total_point = TheExec.flow.var("FlowLoopMax").value
    
    For Each ShmooAxis In TheExec.DevChar.Setups(DevChar_Setup).Shmoo.axes.list
        With TheExec.DevChar.results(DevChar_Setup).Shmoo
            TheExec.Datalog.WriteComment inst_name
            For Each site In TheExec.sites
            
                result_str = vbNullString
                With Shmoo_PreCond_Info_Fast(0)
                    x_steps = Abs(CDbl(.to) - CDbl(.from)) \ CDbl(.StepSize)
                End With
                For x1 = 0 To x_steps
                     If x1 = 0 Then
                        result_str = MarginResultConv(.Points(x1).ExecutionResult)
                     Else
                        result_str = result_str & MarginResultConv(.Points(x1).ExecutionResult)
                     End If
                Next x1
                head_str = site & " " & " " & Shmoo_Precond_Str(Nth_point, inst_name) & "_" & Shmoo_PreCond_Info_Fast(0).Spec & Shmoo_PreCond_Info_Fast(0).constraint  'DevChar_StepName
    '            head_str = Replace(head_str, ".", "p")
                
                If TheExec.TesterMode = testModeOffline Then 'Make up psuedo val
                    Transition_val = Floor(10# * Rnd(1)) * 5000000#
                    If Not (Transition_val = 20 * 1000000# Or Transition_val = 25 * 1000000# _
                        Or Transition_val = 30 * 1000000# Or Transition_val = 35 * 1000000# _
                        Or TheExec.flow.var("FlowLoop").value = 0 _
                        Or TheExec.flow.var("FlowLoop").value = TheExec.flow.var("FlowLoopMax").value - 1 _
                        Or TheExec.flow.var("FlowLoop").value = TheExec.flow.var("FlowLoopMax").value) Then ' psuedo fail
                        Transition_str = Unit_format(Transition_val(site), Shmoo_PreCond_Info_Fast(0).unit)
                        Update_Result_ND Transition_val, inst_name
                    Else
                        Transition_str = "N/A"
                    End If
                Else
                    ' Need to modify if shmoo hole
                    With Shmoo_PreCond_Info_Fast(0)
                        AnalysisTrans result_str, TransCnt, CDbl(.from), CDbl(.to), CDbl(.StepSize), p_min, p_max
                        If .constraint = "max" Then
                            Transition_val(site) = Unit_Calc(p_max, Shmoo_PreCond_Info_Fast(0).unit)
                        ElseIf .constraint = "min" Then
                            If p_min <> 0 Then
                                Transition_val(site) = Unit_Calc(ExecTime, Shmoo_PreCond_Info_Fast(0).unit)
                            End If
                        End If
                    End With
    '                Transition_val(site) = Unit_Calc(p_max, Shmoo_PreCond_Info_Fast(0).unit)
                    
                    If TransCnt > 2 Or Transition_val(site) = 0 Then
                        Transition_str = "N/A"
                        SitePF(site) = False
                    ElseIf TransCnt = 2 And UCase(mid(result_str, 1, 1)) = "P" Then
                        Transition_str = "N/A"
                    Else
                        Transition_str = Unit_format(Transition_val(site), Shmoo_PreCond_Info_Fast(0).unit)
    '                    If Shmoo_PreCond_Info_Fast(0).constraint = "min" Then
    '                        SitePF(site) = True
    '                    Else
                            Update_Result_ND Transition_val, inst_name
    '                    End If
                    End If
                End If
                TheExec.Datalog.WriteComment head_str & " = " & Transition_str
                TheExec.Datalog.WriteComment vbNullString
                TheExec.Datalog.WriteComment "********** Loop Index " & Nth_point + 1 & " of " & Total_point + 1 & " Spec Combinations **********"
                TheExec.Datalog.WriteComment "Result String: " & result_str
                TheExec.Datalog.WriteComment "Transition Count: " & TransCnt
                
    '            AllSitePassChecker = AllSitePassChecker And SitePF(site)
            Next site
        
'        If AllSitePassChecker = True Then
'            For Each site In TheExec.sites
'                Update_Result_ND Transition_val, inst_name
'            Next site
'        End If
        
        End With
'               Debug.Print Site & ": " & result_str
    Next ShmooAxis
    If TheExec.Overlays.Count > 10000 Then
        TheExec.Overlays.RemoveAll
    End If
   
    Exit Function
errHandler:
    Call Print_Error_Message(Error_Info, "LIB_Shmoo_ND", "Show_Margin_Result")
    If AbortTest Then Exit Function Else Resume Next
End Function

Public Function Store_LVCC_Value(argc As Integer, argv() As String) As Long
On Error GoTo errHandler
    Dim DevChar_StepName As Variant
    Dim power_pin As String
    Dim site As Variant
    Dim x_steps As Long
    Dim x1 As Long
    Dim LVCC_val As New SiteDouble, HLvcc_Str As New SiteVariant
    Dim HVCC_val As New SiteDouble
    Dim Nth_point As Long
    Dim inst_name As String
    Dim LVCC_Idx As Long
    Dim back_off As Double
    Dim i As Long
    Dim result_str As String
    Dim TransCnt As Long
    Dim p_min As Double
    Dim p_max As Double
    Dim TrackingStep() As String
    Dim DevChar_Setup As String
    Dim trans_cnt As Long
    
    inst_name = TheExec.DataManager.instancename
    DevChar_Setup = TheExec.DevChar.Setups.ActiveSetupName
    'Nth_point = TheExec.Flow.var("FlowLoop").Value
    power_pin = TheExec.DevChar.Setups(DevChar_Setup).Shmoo.axes(tlDevCharShmooAxis_X).ApplyTo.pins
    For i = 0 To UBound(Shmoo_PreCond_Info_LVCC)
        If UCase(Shmoo_PreCond_Info_LVCC(i).SpecVoltageName) Like power_pin & "*" Then
            LVCC_Idx = i
        End If
        'Stop
    Next i
    'Get Tracking StepName
    With TheExec.DevChar.Setups(DevChar_Setup).Shmoo.axes(tlDevCharShmooAxis_X)
        If .TrackingParameters.Count > 0 Then
            ReDim Shmoo_PreCond_Info_LVCC(LVCC_Idx).TrackingSpecN(.TrackingParameters.Count - 1)
            TrackingStep = .TrackingParameters.list
            Shmoo_PreCond_Info_LVCC(LVCC_Idx).TrackingCnt = .TrackingParameters.Count
            For i = 1 To .TrackingParameters.Count
                Select Case UCase(.TrackingParameters(TrackingStep(i - 1)))
                Case "VMAIN"
                    Shmoo_PreCond_Info_LVCC(LVCC_Idx).TrackingSpecN(i - 1) = TrackingStep(i - 1) & "_VAR"
                Case "VALT"
                    Shmoo_PreCond_Info_LVCC(LVCC_Idx).TrackingSpecN(i - 1) = TrackingStep(i - 1) & "_VOP_VAR"
                End Select
            Next i
        End If
    End With
        'Judge Search LVCC or HVCC
'        If inst_name Like "*_HBV" Then
'            Shmoo_PreCond_Info_LVCC(LVCC_Idx).SearchType = "HVCC"
'        ElseIf inst_name Like "*_BV" Then
'            Shmoo_PreCond_Info_LVCC(LVCC_Idx).SearchType = "LVCC"
'        Else
'            Shmoo_PreCond_Info_LVCC(LVCC_Idx).SearchType = "LVCC"
'        End If
        
    With Shmoo_PreCond_Info_LVCC(LVCC_Idx)
        If .constraint Like "-*" Then
            .SearchType = "HVCC"
        Else
            .SearchType = "LVCC"
        End If
        back_off = Abs(Unit_Calc(CDbl(.constraint), .unit))
    End With
    TheExec.Datalog.WriteComment inst_name
    For Each site In TheExec.sites
        With TheExec.DevChar.results(DevChar_Setup).Shmoo
            With Shmoo_PreCond_Info_LVCC(LVCC_Idx)
                x_steps = Abs(CDbl(.to) - CDbl(.from)) / CDbl(.StepSize)
            End With
            'Check Shmoo Result
            result_str = vbNullString
            For x1 = 0 To x_steps 'need mod
                If x1 = 0 Then
                    result_str = MarginResultConv(.Points(x1).ExecutionResult)
                Else
                    result_str = result_str & MarginResultConv(.Points(x1).ExecutionResult)
                End If
            Next x1
''                TheExec.Datalog.WriteComment result_str
            '''''''''''''''''''''''''''''''''''''''''''''''''''''''''
            With Shmoo_PreCond_Info_LVCC(LVCC_Idx)
                AnalysisTrans result_str, TransCnt, CDbl(.from), CDbl(.to), CDbl(.StepSize), p_min, p_max
                
                LVCC_val(site) = p_min
                HVCC_val(site) = p_max
                If trans_cnt > 2 Then
                    HLvcc_Str = "N/A"
                    Shmoo_PreCond_Info_LVCC(LVCC_Idx).LVCC_Value(site) = 0
                Else
                    If .SearchType = "LVCC" Then
                        HLvcc_Str = CStr(LVCC_val(site)) & " " & .unit
                        Shmoo_PreCond_Info_LVCC(LVCC_Idx).LVCC_Value(site) = Unit_Calc(LVCC_val(site), .unit) + back_off
                    ElseIf .SearchType = "HVCC" Then
                        HLvcc_Str = CStr(HVCC_val(site)) & " " & .unit
                        Shmoo_PreCond_Info_LVCC(LVCC_Idx).HVCC_Value(site) = Unit_Calc(HVCC_val(site), .unit) - back_off
                    Else
                    End If
                End If
            
            End With
            With Shmoo_PreCond_Info_LVCC(LVCC_Idx)
                TheExec.Datalog.WriteComment "**********"
                TheExec.Datalog.WriteComment "Search Type: " & .SearchType
                If .SearchType = "LVCC" Then
                    TheExec.Datalog.WriteComment "Site: " & site & ", Value for timing search: " & .SpecVoltageName & " = " & _
                                                Replace(HLvcc_Str, .unit, vbNullString) & "+ " & .constraint & " = " & Unit_format(.LVCC_Value(site), .unit)
                ElseIf .SearchType = "HVCC" Then
                    TheExec.Datalog.WriteComment "Site: " & site & ", Value for timing search: " & .SpecVoltageName & " = " & _
                                                Replace(HLvcc_Str, .unit, vbNullString) & "- " & Replace(.constraint, "-", vbNullString) & " = " & Unit_format(.HVCC_Value(site), .unit)
                End If
                TheExec.Datalog.WriteComment "**********"
            End With
            'theexec.Datalog.WriteComment head_str & " = " & LVCC_str
       End With
    Next site
    Exit Function
errHandler:
    Call Print_Error_Message(Error_Info, "LIB_Shmoo_ND", "Store_LVCC_Value")
    If AbortTest Then Exit Function Else Resume Next
End Function

Public Function Update_Result_ND(Transition_val As SiteDouble, inst_name As String)
On Error GoTo errHandler
    Dim flow_step As String
    Dim inst_perSite As String
    
    If Not InstanceInUse.Exists(inst_name) Then
        InstanceInUse.Add inst_name, Shmoo_ND_Info_now.N_InstanceInUse
        Shmoo_ND_Info_now.N_InstanceInUse = Shmoo_ND_Info_now.N_InstanceInUse + 1
    End If
    inst_perSite = inst_name & CStr(TheExec.sites.siteNumber)
    flow_step = str(TheExec.flow.var("FlowLoop").value)
    If Not InstanceInUsePerSite.Exists(inst_perSite) Then ' check if it is the first time for each site
        InstanceInUsePerSite.Add inst_perSite, Shmoo_ND_Info_now.N_InstanceInUse
        If Shmoo_PreCond_Info_Fast(0).constraint = "max" Then
            Result_ND(InstanceInUse(inst_name)).val = -9999
        ElseIf Shmoo_PreCond_Info_Fast(0).constraint = "min" Then
            Result_ND(InstanceInUse(inst_name)).val = 9999
        End If
    End If
    
    Result_ND(InstanceInUse(inst_name)).FlowStep_Pass = Result_ND(InstanceInUse(inst_name)).FlowStep_Pass & "," & flow_step
    
    If Shmoo_PreCond_Info_Fast(0).constraint = "max" Then
        If Result_ND(InstanceInUse(inst_name)).val < Transition_val Or Result_ND(InstanceInUse(inst_name)).val = -9999 Then
            Result_ND(InstanceInUse(inst_name)).val = Transition_val
            Result_ND(InstanceInUse(inst_name)).FlowStep_Constraint = flow_step
        ElseIf Result_ND(InstanceInUse(inst_name)).val = Transition_val Then
            Result_ND(InstanceInUse(inst_name)).FlowStep_Constraint = Result_ND(InstanceInUse(inst_name)).FlowStep_Constraint & "," & flow_step
        End If
    ElseIf Shmoo_PreCond_Info_Fast(0).constraint = "min" Then
        If Result_ND(InstanceInUse(inst_name)).val > Transition_val Or Result_ND(InstanceInUse(inst_name)).val = 9999 Then
            Result_ND(InstanceInUse(inst_name)).val = Transition_val
            Result_ND(InstanceInUse(inst_name)).FlowStep_Constraint = flow_step
        ElseIf Result_ND(InstanceInUse(inst_name)).val = Transition_val Then
            Result_ND(InstanceInUse(inst_name)).FlowStep_Constraint = Result_ND(InstanceInUse(inst_name)).FlowStep_Constraint & "," & flow_step
        End If
    End If
    Exit Function
errHandler:
    Call Print_Error_Message(Error_Info, "LIB_Shmoo_ND", "Update_Result_ND")
    If AbortTest Then Exit Function Else Resume Next
End Function
Public Function MarginResultConv(result As tlDevCharResult) As String
On Error GoTo errHandler
    Select Case result
        Case tlDevCharResult_Pass: MarginResultConv = "P"                           '0
        Case tlDevCharResult_AssumedPass: MarginResultConv = "P"                    '2
        Case tlDevCharResult_Fail: MarginResultConv = "F"                           '1
        Case tlDevCharResult_AssumedFail: MarginResultConv = "F"                    '3
        Case tlDevCharResult_Alarm, tlDevCharResult_Error: MarginResultConv = "A"   '8
        Case Else:  MarginResultConv = "?" 'unknown case"
    End Select
    
    Exit Function
errHandler:
    Call Print_Error_Message(Error_Info, "LIB_Shmoo_ND", "MarginResultConv")
    If AbortTest Then Exit Function Else Resume Next
End Function

Public Function Shmoo_Precond_Str(Nth_point As Long, inst_name As String) As String
On Error GoTo errHandler
    Dim i As Long, j As Long, o_Str As String
    Dim SpecVal_Str As String
    For j = 0 To Shmoo_PreCond_Spec_N - 1
        SpecVal_Str = CStr(Shmoo_PreCond_Val(Nth_point, j))
        If (CDbl(SpecVal_Str) / 1000000) > 1 Then
            SpecVal_Str = CStr((CDbl(SpecVal_Str) / 1000000)) & "M"
        End If
        If InStr(SpecVal_Str, ".") > 0 Then
            SpecVal_Str = Replace(SpecVal_Str, ".", "p")
        End If
        
        If j = 0 Then
'            o_Str = Shmoo_PreCond_Info(j).spec & Shmoo_PreCond_Val(Nth_point, j)
            o_Str = Shmoo_PreCond_Info(j).Spec & SpecVal_Str
        Else
'            o_Str = Shmoo_PreCond_Info(j).spec & Shmoo_PreCond_Val(Nth_point, j) & "_" & o_Str
            o_Str = Shmoo_PreCond_Info(j).Spec & SpecVal_Str & "_" & o_Str
        End If
        If j = Shmoo_PreCond_Spec_N - 1 Then o_Str = inst_name & " " & o_Str
    Next j
    Shmoo_Precond_Str = o_Str
    Exit Function
errHandler:
    Call Print_Error_Message(Error_Info, "LIB_Shmoo_ND", "Shmoo_Precond_Str")
    If AbortTest Then Exit Function Else Resume Next
End Function
Public Function Init_Result()
On Error GoTo errHandler
    InstanceInUse.RemoveAll
    Shmoo_ND_Info_now.N_InstanceInUse = 0
    ReDim Result_ND(Max_Instance_N)
    Exit Function
errHandler:
    Call Print_Error_Message(Error_Info, "LIB_Shmoo_ND", "Init_Result")
    If AbortTest Then Exit Function Else Resume Next
End Function
Public Function Apply_Flow_PreCond(argc As Integer, argv() As String) As Long
On Error GoTo errHandler
    Dim ac_spec As String, val As Double
    Dim DevChar_Setup As String
    Dim site As Variant
    Dim i As Long
    Dim j As Long
    Dim curr_flow_pt As Long
    Dim OverlayName As String, specName As String
    Dim HLVCC_Name As String
    Dim char_val As Double
    Dim HLvccMarginStr As String
    Dim HLvccForceStr As String
    Dim HLvcc_SpecNameAry() As String
    Dim HLvcc_Pin As String
    Dim HLvcc_unit As String
    Dim PinVal As Double
    Dim VolType As String
    Dim tmp_SpecPeriod As String
    Dim SpecPeriodAry() As String
    Dim tmp_SPVal() As Double
    Dim Period_cnt As Long
    Dim Period_Bool As Boolean
    Dim temp_comma() As String
    Dim i_comma As Integer
    Period_cnt = 0
    Period_Bool = False
    OverlayName = "Any"
    curr_flow_pt = TheExec.flow.var("FlowLoop").value
    DevChar_Setup = TheExec.DevChar.Setups.ActiveSetupName
    For Each site In TheExec.sites
        char_val = TheExec.DevChar.results(DevChar_Setup).Shmoo.CurrentPoint.axes(tlDevCharShmooAxis_X).value
'        char_val = TheExec.DevChar.Results(DevChar_Setup).Margins.Item(Shmoo_PreCond_Info_Fast(0).SpecVoltageName).CurrentPoint.Value
'        char_val = TheExec.Specs.AC.Item(ND_DummyShift).CurrentValue(site)
        For i = 0 To Shmoo_PreCond_Spec_N - 1
            tmp_SpecPeriod = Shmoo_PreCond_Info(i).SpecVoltageName
''            If UCase(tmp_SpecPeriod) = UCase("Buffer_VAR") Then
            If UCase(Shmoo_PreCond_Info(i).unit) = "MHZ" Then
                Period_Bool = True
'                If Period_cnt = 0 Then
                If tmp_SpecPeriod Like "*,*" Then 'tracking case
                    temp_comma = Split(tmp_SpecPeriod, ",")
                    For i_comma = 0 To UBound(temp_comma)
                        ReDim Preserve SpecPeriodAry(Period_cnt) As String
                        ReDim Preserve tmp_SPVal(Period_cnt) As Double
                        SpecPeriodAry(Period_cnt) = temp_comma(i_comma) & "_1"
                        tmp_SPVal(Period_cnt) = Shmoo_PreCond_Val(curr_flow_pt, i)
                        Period_cnt = Period_cnt + 1
                    Next
                Else
                        ReDim Preserve SpecPeriodAry(Period_cnt) As String
                        ReDim Preserve tmp_SPVal(Period_cnt) As Double
                        SpecPeriodAry(Period_cnt) = tmp_SpecPeriod & "_1"
                        tmp_SPVal(Period_cnt) = Shmoo_PreCond_Val(curr_flow_pt, i)
                        Period_cnt = Period_cnt + 1
                End If
'                Else
'                    If tmp_SpecPeriod Like "*,*" Then 'tracking case
'                        temp_comma = Split(tmp_SpecPeriod, ",")
'                        For i_comma = 0 To UBound(temp_comma)
'                            ReDim Preserve SpecPeriodAry(Period_cnt) As String
'                            ReDim Preserve tmp_SPVal(Period_cnt) As Double
'                            SpecPeriodAry(Period_cnt) = temp_comma(i_comma) & "_1"
'                            tmp_SPVal(Period_cnt) = Shmoo_PreCond_Val(curr_flow_pt, i)
'                            Period_cnt = Period_cnt + 1
'                        Next
'                    Else
'                            ReDim Preserve SpecPeriodAry(Period_cnt) As String
'                            ReDim Preserve tmp_SPVal(Period_cnt) As Double
'                            SpecPeriodAry(Period_cnt) = tmp_SpecPeriod & "_1"
'                            tmp_SPVal(Period_cnt) = Shmoo_PreCond_Val(curr_flow_pt, i)
'                            Period_cnt = Period_cnt + 1
'                    End If
'                End If
                    
'                tmp_SPVal = Shmoo_PreCond_Val(curr_flow_pt, i)
'                TheExec.Overlays.ApplyUniformSpecToHW tmp_SpecPeriod & "_1", tmp_SPVal
            End If
        Next i
        Exit For
    Next site
    If Period_Bool Then
        For i = 0 To UBound(SpecPeriodAry)
            If SpecPeriodAry(i) Like "*,*" Then
                temp_comma = Split(SpecPeriodAry(i), ",")
                For i_comma = 0 To UBound(temp_comma)
                    TheExec.Overlays.ApplyUniformSpecToHW temp_comma(i_comma), tmp_SPVal(i)
                Next
            Else
                TheExec.Overlays.ApplyUniformSpecToHW SpecPeriodAry(i), tmp_SPVal(i)
            End If
        Next i
    End If
    For Each site In TheExec.sites
        With TheExec.Overlays
             If (.Contains(OverlayName) <> False) Then .Remove OverlayName
            .Add (OverlayName)
            For i = 0 To Shmoo_PreCond_Spec_N - 1
                specName = Shmoo_PreCond_Info(i).SpecVoltageName
                If specName Like "*,*" Then
                    temp_comma = Split(specName, ",")
                    For i_comma = 0 To UBound(temp_comma)
                        TheExec.Overlays(OverlayName).Specs.Add temp_comma(i_comma)
                        TheExec.Overlays(OverlayName).Specs(temp_comma(i_comma)).value = Shmoo_PreCond_Val(curr_flow_pt, i)
                    Next
                Else
                    TheExec.Overlays(OverlayName).Specs.Add specName
                    TheExec.Overlays(OverlayName).Specs(specName).value = Shmoo_PreCond_Val(curr_flow_pt, i)
                End If
'                theexec.Datalog.WriteComment "Set " & SpecName & " to " & Shmoo_PreCond_Val(curr_flow_pt, i)
                'Call TheExec.Overlays(OverlayName).Apply(True, False)
            Next i
            TheExec.Overlays(OverlayName).Specs.Add Shmoo_PreCond_Info_Fast(0).SpecVoltageName
            TheExec.Overlays(OverlayName).Specs(Shmoo_PreCond_Info_Fast(0).SpecVoltageName).value(site) = char_val
            ' Setup LVCC Result
            If Shmoo_PreCond_LVCC_N > 0 Then
                For i = 0 To Shmoo_PreCond_LVCC_N - 1
''                    LVCC_Name = Shmoo_ND_Setups(Shmoo_ND_Info_now.Setup_Idx).lvcc(i).SpecVoltageName
                    HLVCC_Name = Lvcc_SpecName(i)
                    If Shmoo_PreCond_Info_LVCC(i).SearchType = "LVCC" Then
                        If Not Shmoo_PreCond_Info_LVCC(i).LVCC_Value(site) = Empty Then
                            TheExec.Overlays(OverlayName).Specs.Add HLVCC_Name
                            TheExec.Overlays(OverlayName).Specs(HLVCC_Name).value(site) = Shmoo_PreCond_Info_LVCC(i).LVCC_Value(site)
                            'Add Tracking Overlay
                            If Shmoo_PreCond_Info_LVCC(i).TrackingCnt > 0 Then
                                For j = 0 To UBound(Shmoo_PreCond_Info_LVCC(i).TrackingSpecN)
                                    TheExec.Overlays(OverlayName).Specs.Add Shmoo_PreCond_Info_LVCC(i).TrackingSpecN(j)
                                    TheExec.Overlays(OverlayName).Specs(Shmoo_PreCond_Info_LVCC(i).TrackingSpecN(j)).value(site) = Shmoo_PreCond_Info_LVCC(i).LVCC_Value(site)
                                Next j
                            End If
                        End If
                    ElseIf Shmoo_PreCond_Info_LVCC(i).SearchType = "HVCC" Then
                        If Not Shmoo_PreCond_Info_LVCC(i).HVCC_Value(site) = Empty Then
                            TheExec.Overlays(OverlayName).Specs.Add HLVCC_Name
                            TheExec.Overlays(OverlayName).Specs(HLVCC_Name).value(site) = Shmoo_PreCond_Info_LVCC(i).HVCC_Value(site)
                            'Add Tracking Overlay
                            If Shmoo_PreCond_Info_LVCC(i).TrackingCnt > 0 Then
                                For j = 0 To UBound(Shmoo_PreCond_Info_LVCC(i).TrackingSpecN)
                                    TheExec.Overlays(OverlayName).Specs.Add Shmoo_PreCond_Info_LVCC(i).TrackingSpecN(j)
                                    TheExec.Overlays(OverlayName).Specs(Shmoo_PreCond_Info_LVCC(i).TrackingSpecN(j)).value(site) = Shmoo_PreCond_Info_LVCC(i).HVCC_Value(site)
                                Next j
                            End If
                        End If
                    End If
                Next i
            End If
            ' Setup Static Power Value
            If Shmoo_PreCond_Static_N > 0 Then
                ''''''''''''''''Add Code Here
                For i = 0 To Shmoo_PreCond_Static_N - 1
                    Dim tmpStaticAry() As String
                    Dim StaticSpecName As String
                    Dim SpecType As String
                    tmpStaticAry = Split(Shmoo_PreCond_Info_StaticV(i).SpecVoltageName, "_")
                    Select Case UCase(tmpStaticAry(UBound(tmpStaticAry)))
                    Case "VALT"
                        SpecType = "_VOP_VAR"
                    Case "VMAIN"
                        SpecType = "_VAR"
                    Case Else
                        SpecType = vbNullString
                        TheExec.Datalog.WriteComment "Please Check Static Power Naming " & Shmoo_PreCond_Info_StaticV(i).SpecVoltageName
                    End Select
                    ReDim Preserve tmpStaticAry(UBound(tmpStaticAry) - 1)
                    StaticSpecName = Join(tmpStaticAry, "_") & SpecType
                    TheExec.Overlays(OverlayName).Specs.Add StaticSpecName
                    TheExec.Overlays(OverlayName).Specs(StaticSpecName).value(site) = Unit_Calc(CDbl(Shmoo_PreCond_Info_StaticV(i).Default), Shmoo_PreCond_Info_StaticV(i).unit)
                Next i
            End If
        End With
        Call TheExec.Overlays(OverlayName).Apply(True, False)
        
        If TheExec.flow.enableWord("char_powercheck") Then
            If Shmoo_PreCond_LVCC_N <> 0 Then
                For i = 0 To Shmoo_PreCond_LVCC_N - 1
                    HLvcc_SpecNameAry = Split(Shmoo_PreCond_Info_LVCC(i).SpecVoltageName, "_")
                    HLvcc_unit = Shmoo_PreCond_Info_LVCC(i).unit
                    If Shmoo_PreCond_Info_LVCC(i).SearchType = "LVCC" Then
                        If Shmoo_PreCond_Info_LVCC(i).LVCC_Value(site) = Empty Then
                            HLvccForceStr = "N/A"
                            HLvccMarginStr = "N/A"
                        Else
                            HLvccForceStr = Unit_format((Shmoo_PreCond_Info_LVCC(i).LVCC_Value(site)), HLvcc_unit)
                            HLvccMarginStr = Unit_format(Shmoo_PreCond_Info_LVCC(i).LVCC_Value(site) - Unit_Calc(Abs(CDbl(Shmoo_PreCond_Info_LVCC(i).constraint)), HLvcc_unit), HLvcc_unit)
                        End If
                    ElseIf Shmoo_PreCond_Info_LVCC(i).SearchType = "HVCC" Then
                        If Shmoo_PreCond_Info_LVCC(i).HVCC_Value(site) = Empty Then
                            HLvccForceStr = "N/A"
                            HLvccMarginStr = "N/A"
                        Else
                            HLvccForceStr = Unit_format((Shmoo_PreCond_Info_LVCC(i).HVCC_Value(site)), HLvcc_unit)
                            HLvccMarginStr = Unit_format(Shmoo_PreCond_Info_LVCC(i).HVCC_Value(site) + Unit_Calc(Abs(CDbl(Shmoo_PreCond_Info_LVCC(i).constraint)), HLvcc_unit), HLvcc_unit)
                        End If
                    End If
                    
                    VolType = UCase(HLvcc_SpecNameAry(UBound(HLvcc_SpecNameAry)))
                    ReDim Preserve HLvcc_SpecNameAry(UBound(HLvcc_SpecNameAry) - 1)
                    HLvcc_Pin = Join(HLvcc_SpecNameAry, "_")
                    Select Case VolType
                    Case "VMAIN"
                        PinVal = Format(TheHdw.DCVS.pins(HLvcc_Pin).Voltage.Main, "0.000")
                    Case "VALT"
                        PinVal = Format(TheHdw.DCVS.pins(HLvcc_Pin).Voltage.Alt, "0.000")
                    End Select
                    TheExec.Datalog.WriteComment "Site: " & site & ", " & HLvcc_Pin & "'s " & VolType & " HW Value: " & PinVal & " V, With " & Shmoo_PreCond_Info_LVCC(i).SearchType & "_Value: " & HLvccMarginStr & " and Constraint: " & CDbl(Shmoo_PreCond_Info_LVCC(i).constraint) & HLvcc_unit
                    
                    'Check Tracking Voltage
                    If Shmoo_PreCond_Info_LVCC(i).TrackingCnt > 0 Then
                    Dim tmpTpin As String
                    Dim TmpVal As Double
                    Dim TSpecAry() As String
                    TSpecAry = Shmoo_PreCond_Info_LVCC(i).TrackingSpecN
                        For j = 0 To UBound(TSpecAry)
                            If TSpecAry(j) Like "*_VOP_VAR" Then
                                tmpTpin = Replace(TSpecAry(j), "_VOP_VAR", vbNullString)
                                TmpVal = Format(TheHdw.DCVS.pins(tmpTpin).Voltage.Alt, "0.000")
                            ElseIf TSpecAry(j) Like "*_VAR" Then
                                tmpTpin = Replace(TSpecAry(j), "_VAR", vbNullString)
                                TmpVal = Format(TheHdw.DCVS.pins(tmpTpin).Voltage.Main, "0.000")
                            End If
                            TheExec.Datalog.WriteComment "Site: " & site & ", Tracking Pin: " & tmpTpin & " HW Value: " & TmpVal & " V"
                        Next j
                    End If
                Next
                TheExec.Datalog.WriteComment " "
            End If
        End If
        
    Next site
    
    '''''''Print out timing
    If TheExec.flow.enableWord("char_timingcheck") = True Then
        CheckTiming char_val
    End If
    RefTimer = 0
    If Shmoo_PreCond_Info_Fast(0).constraint = "min" Then
        RefTimer = TheExec.Timer
    End If
    Exit Function
errHandler:
    Call Print_Error_Message(Error_Info, "LIB_Shmoo_ND", "Apply_Flow_PreCond")
    If AbortTest Then Exit Function Else Resume Next
End Function

Public Function StoreVal_df(df_name As String, val As Double)
On Error GoTo errHandler
     '   ============================================
     '   Save a value in a Defined Name
     '   ============================================
    Names(df_name).RefersTo = "=" & str(val)
    Exit Function
errHandler:
    Call Print_Error_Message(Error_Info, "LIB_Shmoo_ND", "StoreVal_df")
    If AbortTest Then Exit Function Else Resume Next
End Function
Public Function Create_df(df_name As String)
     '   ============================================
     '   Save a value in a Defined Name
     '   ============================================
    Dim df_Cell As name
         
    On Error Resume Next
    Set df_Cell = ThisWorkbook.Names(df_name)
    If err.number > 0 Then Names.Add df_name, 1 'create define name if not exist
End Function

Public Function AnalysisTrans(result_str As String, Transition_Cnt As Long, start_val As Double, stop_val As Double, step_val As Double, Optional p_min As Double, Optional p_max As Double)
On Error GoTo errHandler
    Dim current_val As Double
    Dim i As Long
    p_min = 0
    p_max = 0
    If start_val > stop_val Then
        step_val = -1 * step_val
    Else
        step_val = step_val
    End If
    Transition_Cnt = 0
    For i = 1 To Len(result_str)
        If mid(result_str, i, 1) = "P" Then
            current_val = start_val + step_val * (i - 1)
            
            If p_min = 0 And p_max = 0 Then
                p_min = current_val
                p_max = current_val
            Else
                If current_val > p_min Then
                    p_min = p_min
                    p_max = current_val
                ElseIf current_val < p_max Then
                    p_min = current_val
                    p_max = p_max
                End If
            End If
        End If
        If i <> Len(result_str) And mid(result_str, i, 1) <> mid(result_str, i + 1, 1) Then
            Transition_Cnt = Transition_Cnt + 1
        End If
    Next i
    Exit Function
errHandler:
    Call Print_Error_Message(Error_Info, "LIB_Shmoo_ND", "AnalysisTrans")
    If AbortTest Then Exit Function Else Resume Next
End Function


Public Function CheckTiming(char_val As Double)
On Error GoTo errHandler
    Dim timing_period As Double
    Dim timing_freq As Double
    Dim D1_edge As Double
    Dim D3_edge As Double
    Dim R0_edge As Double
    Dim site As Variant
    
    Dim confirm_timeset As String
    Dim confirm_pinname As String
    confirm_timeset = "tset5"
    confirm_pinname = "AOP_FUNC22"
    For Each site In TheExec.sites
        timing_freq = char_val
        timing_period = 1 / timing_freq
        D1_edge = TheHdw.Digital.pins(confirm_pinname).Timing.EdgeTime(confirm_timeset, chEdgeD1).value
        'D3_edge = TheHdw.Digital.Pins(confirm_pinname).Timing.EdgeTime(confirm_timeset, chEdgeD3).Value
        R0_edge = TheHdw.Digital.pins(confirm_pinname).Timing.EdgeTime(confirm_timeset, chEdgeR0).value
'        timing_freq = 1 / timing_period
        TheExec.Datalog.WriteComment "*****************************************************************"
'        TheExec.Datalog.WriteComment "Site :" & site & " Current Period Read From HW: " & Format(timing_period * 1000000000#, "0.000") & " nsec"
'        TheExec.Datalog.WriteComment "Site :" & site & " Current ShiftIn_Freq Read From HW: " & Unit_format(timing_freq, "Mhz")
        TheExec.Datalog.WriteComment "Site :" & site & " Current ShiftIn_Freq Read From Overlay: " & Unit_format(timing_freq, "Mhz")
        TheExec.Datalog.WriteComment "Site :" & site & " Current Period Calc From Overlay: " & Format(timing_period * 1000000000#, "0.000") & " nsec"
        TheExec.Datalog.WriteComment "Site :" & site & " Current D1 Timing Read From HW: " & Format(D1_edge * 1000000000#, "0.000") & " nsec"
        TheExec.Datalog.WriteComment "Site :" & site & " The D1 Edge is on " & Format((D1_edge / timing_period) * 100, "000.0") & " % of Current Period"
        TheExec.Datalog.WriteComment "Site :" & site & " Current R0 Timing Read From HW: " & Format(R0_edge * 1000000000#, "0.000") & " nsec"
        TheExec.Datalog.WriteComment "Site :" & site & " The R0 Edge is on " & Format((R0_edge / timing_period) * 100, "000.0") & " % of Current Period"
        TheExec.Datalog.WriteComment "*****************************************************************"
    Next site
    Exit Function
errHandler:
    Call Print_Error_Message(Error_Info, "LIB_Shmoo_ND", "CheckTiming")
    If AbortTest Then Exit Function Else Resume Next
End Function
