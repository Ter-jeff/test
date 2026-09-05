Attribute VB_Name = "LIB_Common_Pattern_Freq"
Option Explicit
'Revision History:
'V0.0 initial bring up
Public Function PatExculdePath(pat As Variant) As String
On Error GoTo errHandler 'Add ErrHandler 2023/08/18
    Dim patt_ary_temp() As String
    patt_ary_temp = Split(pat, "\")
    PatExculdePath = patt_ary_temp(UBound(patt_ary_temp))
Exit Function 'Add ErrHandler 2023/08/18
errHandler: 'Add ErrHandler 2023/08/18
    Call Print_Error_Message(Error_Info, "LIB_Common_Pattern_Freq", "PatExculdePath") 'Add ErrHandler 2023/08/18
    If AbortTest Then Exit Function Else Resume Next 'Add ErrHandler 2023/08/18
End Function

'*****************************************
'******        pattern set, patterns******
'*****************************************
' Decompose patset recursively and return a string "patt" with a list of .pat
Public Function PatSetToPat(ByVal patset As Pattern) As String
On Error GoTo errHandler 'Add ErrHandler 2023/08/18
'   pat1    pat1a,pat1b,pat1c.pat
'   pat1a   pat1a.pat
'   pat1b   pat1b1, pat1b2
'   pat1b1  pat1b1.pat
'   pat1b2  pat1b2.pat

    Dim pat_ary() As String, patcnt As Long
    Dim Pat_ary1() As String, PatCnt1 As Long
    Dim patset_ary() As String, i As Long
    Dim patset1 As New Pattern
    Dim patt_str As String
    Dim patt As String
    
    patset_ary = Split(patset.value, ",")
    patt = vbNullString
    For i = 0 To UBound(patset_ary)
        Current_Patterns = vbNullString
        Call PatsetDecompose(patset_ary(i))
        If patt <> "" Then
            patt = patt & "," & Current_Patterns
        Else
            patt = Current_Patterns
        End If
    Next i
    PatSetToPat = patt
Exit Function 'Add ErrHandler 2023/08/18
errHandler: 'Add ErrHandler 2023/08/18
    Call Print_Error_Message(Error_Info, "LIB_Common_Pattern_Freq", "PatSetToPat") 'Add ErrHandler 2023/08/18
    If AbortTest Then Exit Function Else Resume Next 'Add ErrHandler 2023/08/18
End Function

Public Function PatSetToPat_EFuse(ByVal patset As Pattern, ByRef patt As String)

On Error GoTo errHandler
    Dim funcName As String:: funcName = "PatSetToPat"

    Dim pat_ary() As String, patcnt As Long

    pat_ary = TheExec.DataManager.Raw.GetPatternsInSet(patset, patcnt)
    patset.value = pat_ary(0)
    While Not (LCase(pat_ary(0)) Like "*.pat*")
        pat_ary = TheExec.DataManager.Raw.GetPatternsInSet(patset, patcnt)
        If patcnt > 1 Then TheExec.ErrorLogMessage (patset & " is with more than one pattern in the pattern set")
        patset.value = pat_ary(0)
    Wend
    patt = pat_ary(0)
    TheHdw.patterns(patt).Load

Exit Function
errHandler:
     TheExec.Datalog.WriteComment "<Error> " + funcName + ":: please check it out."
     If AbortTest Then Exit Function Else Resume Next
End Function

Public Function GetPatListFromPatternSet(TestPat As String, _
                              rtnPatNames() As String, _
                              rtnPatCnt As Long) As Boolean
On Error GoTo errHandler 'Add ErrHandler 2023/08/18
    'used to be GetPatFromPatternSet
    Dim patt_list As String
    Dim patt As New Pattern
    
    GetPatListFromPatternSet = False
    patt.value = TestPat
    patt_list = PatSetToPat(patt)
    rtnPatNames = Split(patt_list, ",")
    rtnPatCnt = UBound(rtnPatNames) + 1
    If (UBound(rtnPatNames) >= 0) Then
        If LCase(rtnPatNames(0)) Like "*.pat*" Then
            GetPatListFromPatternSet = True
        End If
    End If

    Exit Function
    
Exit Function
errHandler:
    GetPatListFromPatternSet = False
    rtnPatCnt = -1

    If AbortTest Then Exit Function Else Resume Next
End Function
' do not use: only as the sub function recursively called in the PatSetToPat()
Public Function PatsetDecompose(PatSetName As String) As String
On Error GoTo errHandler 'Add ErrHandler 2023/08/18
    Dim patcnt As Long                          '<- Number of patterns in set
    Dim RawNameData() As String                 '<- Raw pattern name data
    Dim pIndex As Long
    Dim patt_str As String
    
    RawNameData = TheExec.DataManager.Raw.GetPatternsInSet(PatSetName, patcnt)
    If patcnt = 0 Then
        Current_Patterns = PatSetName
        Exit Function
    End If
    patcnt = UBound(RawNameData)
    For pIndex = 0 To patcnt
        If InStr(1, RawNameData(pIndex), ".pat", vbTextCompare) Then
            Current_Patterns = CombineStringList(Current_Patterns, RawNameData(pIndex))
        Else
            Call PatsetDecompose(RawNameData(pIndex))
        End If
    Next pIndex
Exit Function 'Add ErrHandler 2023/08/18
errHandler: 'Add ErrHandler 2023/08/18
    Call Print_Error_Message(Error_Info, "LIB_Common_Pattern_Freq", "PatsetDecompose") 'Add ErrHandler 2023/08/18
    If AbortTest Then Exit Function Else Resume Next 'Add ErrHandler 2023/08/18
End Function

'*****************************************
'******         frequecy measurement******
'*****************************************
Public Function Freq_MeasFreqSetup(pin As PinList, Interval As Double, Optional MeasF_EventSource As FreqCtrEventSrcSel = 1)
On Error GoTo errHandler 'Add ErrHandler 2023/08/18
    With TheHdw.Digital.pins(pin).FreqCtr
        .EventSource = MeasF_EventSource '' VOH
        .EventSlope = Positive
        .Interval = Interval
        .Enable = IntervalEnable
        .Clear
    End With
Exit Function 'Add ErrHandler 2023/08/18
errHandler: 'Add ErrHandler 2023/08/18
    Call Print_Error_Message(Error_Info, "LIB_Common_Pattern_Freq", "Freq_MeasFreqSetup") 'Add ErrHandler 2023/08/18
    If AbortTest Then Exit Function Else Resume Next 'Add ErrHandler 2023/08/18
End Function


Public Function Freq_MeasFreqStart(pin As PinList, Interval As Double, freq As PinListData)
On Error GoTo errHandler 'Add ErrHandler 2023/08/18
    Dim CounterValue As New PinListData
    Dim site As Variant
    TheHdw.Digital.pins(pin).FreqCtr.Clear
    TheHdw.Digital.pins(pin).FreqCtr.start
    
    For Each site In TheExec.sites
        CounterValue = TheHdw.Digital.pins(pin).FreqCtr.Read
        freq = CounterValue.Math.divide(Interval)
    Next site
Exit Function 'Add ErrHandler 2023/08/18
errHandler: 'Add ErrHandler 2023/08/18
    Call Print_Error_Message(Error_Info, "LIB_Common_Pattern_Freq", "Freq_MeasFreqStart") 'Add ErrHandler 2023/08/18
    If AbortTest Then Exit Function Else Resume Next 'Add ErrHandler 2023/08/18
End Function
