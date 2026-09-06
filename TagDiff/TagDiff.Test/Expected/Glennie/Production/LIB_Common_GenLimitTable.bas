Attribute VB_Name = "LIB_Common_GenLimitTable"
Option Explicit
'Revision History:
'V0.0 initial bring up

'variable declaration
Public Const Version_Lib_GenLimitTable = "0.1"  'lib version

Public Const LT_RowsShift As Integer = 2
Public Const b_DisplayPinName As Boolean = True
Public Const b_DisplayTestNumber As Boolean = False

Public Const b_EnableCP_Update As Boolean = True
Public Const b_EnableFT_Update As Boolean = False

Public Const b_Disable_CPFT_JobToEmpty As Boolean = True

Public Const b_Enable_CPFT_UseTheSameLimit As Boolean = True
Sub GenParaLimitTable()

Dim sourcesheet As Worksheet
Dim list As Worksheet
Dim i As Integer
Dim j As Integer
Dim TempValue As String
Dim TempSheet As String
Dim maxcol As Integer
Dim MaxRow As Integer
Dim Count As Integer
Dim rowIndex As Integer
Dim TempInst As String
Dim TempBin As Long
Dim TempSoft As Long
Dim TempLo As Double
Dim TempHi As Double
Dim TempTestName As String
Dim TempTestNum As Long
Dim k As Integer
Dim ADD_L As String
Dim ADD_H As String
Dim FormulaStrL As String
Dim FormulaStrH As String
Dim FlowTableName As String


Set list = Sheets.Add(Before:=Worksheets(1))
list.name = "A_Limits_Table_1"

Count = ActiveWorkbook.Sheets.Count

rowIndex = 3

list.Cells(2, 1).value = "Flow"
list.Cells(2, 2).value = "TestInstance"
list.Cells(2, 3).value = "TestName"
list.Cells(2, 4).value = "TestNum"
list.Cells(2, 5).value = "Low Limit"
list.Cells(2, 6).value = "High Limit"

For k = 2 To Count

  Set sourcesheet = Sheets(k)
  
  If UCase(sourcesheet.name) Like "*FLOW*" And Not UCase(sourcesheet.name) Like "*CHAR*" Then
  
          MaxRow = sourcesheet.Cells(Rows.Count, 7).End(xlUp).Row
          
          
          For i = 5 To MaxRow
          
          If sourcesheet.Cells(i, 7).value = "Use-Limit" Then
          
''                TempSoft = CLng(sourcesheet.Cells(i, 19).Value)
''                TempBin = CLng(sourcesheet.Cells(i, 17).Value)
                FlowTableName = sourcesheet.name
                TempInst = sourcesheet.Cells(i, 8).value
                TempTestName = sourcesheet.Cells(i, 9).value
                TempTestNum = CLng(sourcesheet.Cells(i, 10).value)
                
                If sourcesheet.Cells(i, 11).value <> "" Then
                    TempLo = CDbl(sourcesheet.Cells(i, 11).value)
                End If
                
                If sourcesheet.Cells(i, 12).value <> "" Then
                    TempHi = CDbl(sourcesheet.Cells(i, 12).value)
                End If

            
                   
'                If TempSoft <> 0 Or TempBin <> 0 Then
                   list.Cells(rowIndex, 1).value = FlowTableName
                   list.Cells(rowIndex, 2).value = TempInst
                   list.Cells(rowIndex, 3).value = TempTestName
                   list.Cells(rowIndex, 4).value = TempTestNum
                   
                   If sourcesheet.Cells(i, 11).value <> "" Then
                        list.Cells(rowIndex, 5).value = TempLo
                   End If
                   
                   If sourcesheet.Cells(i, 12).value <> "" Then
                        list.Cells(rowIndex, 6).value = TempHi
                   End If
                   
                   
                   '*********fill the formula back to the orig cell start****************"
                   ADD_L = Cells(rowIndex, 5).Address(RowAbsolute:=False, ColumnAbsolute:=False)
                   ADD_H = Cells(rowIndex, 6).Address(RowAbsolute:=False, ColumnAbsolute:=False)
                   
                   
                   FormulaStrL = "=OFFSET(A_Limits_Table!" & ADD_L & ",0,_Limit_Offset)"
                   FormulaStrH = "=OFFSET(A_Limits_Table!" & ADD_H & ",0,_Limit_Offset)"
                    
                    
                   If sourcesheet.Cells(i, 11).value <> "" Then
                     sourcesheet.Cells(i, 11).Formula = FormulaStrL  'replace limit value with formula
                   End If
                   
                   If sourcesheet.Cells(i, 12).value <> "" Then
                     sourcesheet.Cells(i, 12).Formula = FormulaStrH  'replace limit value with formula
                   End If
                   
                    '=OFFSET(A_Limits_Table!D2,0,_Limit_Offset)
                  '*********fill the formula back to the orig cell end  ****************"
                   
                
                   rowIndex = rowIndex + 1
                   
            ElseIf LCase(sourcesheet.Cells(i, 7).value) = "return" Or LCase(sourcesheet.Cells(i, 7).value) = "stop" Then
                   
                   i = MaxRow
                
          End If  'If sourcesheet.Cells(i, 7).Value = "Use-Limit"
          
          Next i
  
  
  End If
  
Next k

With list.Cells
    .Font.name = "Arial"
    .HorizontalAlignment = xlCenter
    .VerticalAlignment = xlCenter
  
End With

'''Dim TempRange As Range
'''
'''List.Activate
'''Set TempRange = List.Range(Cells(1, 1), Cells(rowindex, 3))
'''TempRange.Select
'''
'''With Selection.Borders
'''        .LineStyle = xlContinuous
'''        .ColorIndex = xlAutomatic
'''        .TintAndShade = 0
'''        .Weight = xlThin
'''End With
'''
''' Range("A1:C1").Select
'''    Selection.Font.Bold = True
'''
'''Range("A1:C1").Select
'''    With Selection.Interior
'''        .Pattern = xlSolid
'''        .PatternColorIndex = xlAutomatic
'''        .Color = 5296274
'''        .TintAndShade = 0
'''        .PatternTintAndShade = 0
'''    End With

For i = 1 To 5

 list.Cells(1, i).EntireColumn.AutoFit


Next i


End Sub
Sub LimitTableBypassCheck()

Dim i As Integer
Dim j As Integer
Dim LimitTableSheet As Worksheet
Dim LimitTableBypassSheet As Worksheet

Set LimitTableSheet = Sheets("A_Limits_Table")
Set LimitTableBypassSheet = Sheets("A_Limits_Table_Bypass")

Dim LTB_RowStart As Integer
Dim LTB_RowEnd As Integer

LimitTableSheet.Columns.range("H:H").ClearContents

LTB_RowStart = 2
LTB_RowEnd = LimitTableBypassSheet.Cells(Rows.Count, 1).End(xlUp).Row

Dim LTB_TestInstance() As String
Dim LTB_TestNamePin() As String
Dim LTB_TestLowLimit() As String
Dim LTB_TestHighLimit() As String

ReDim LTB_TestInstance(LTB_RowEnd - LTB_RowStart) As String
ReDim LTB_TestNamePin(LTB_RowEnd - LTB_RowStart) As String
ReDim LTB_TestLowLimit(LTB_RowEnd - LTB_RowStart) As String
ReDim LTB_TestHighLimit(LTB_RowEnd - LTB_RowStart) As String


For i = LTB_RowStart To LTB_RowEnd
    LTB_TestInstance(i - LTB_RowStart) = LimitTableBypassSheet.Cells(i, 1)
    LTB_TestNamePin(i - LTB_RowStart) = LimitTableBypassSheet.Cells(i, 2)
    LTB_TestLowLimit(i - LTB_RowStart) = LimitTableBypassSheet.Cells(i, 4)
    LTB_TestHighLimit(i - LTB_RowStart) = LimitTableBypassSheet.Cells(i, 5)
Next i


Dim RowStart As Integer
Dim RowEnd As Integer

RowStart = 2
RowEnd = LimitTableSheet.Cells(Rows.Count, 1).End(xlUp).Row

Dim index As Integer
Dim b_FirstTimeFlag_OnlyTestInst As Boolean
For index = 0 To UBound(LTB_TestInstance)
    b_FirstTimeFlag_OnlyTestInst = False
    
    For i = RowStart To RowEnd
    
        If UCase(LimitTableSheet.Cells(i, 1)) = UCase(LTB_TestInstance(index)) Then
            If UCase(LTB_TestNamePin(index)) = "X" Then
                LimitTableSheet.Cells(i, 8) = "F"
            Else
                If UCase(LimitTableSheet.Cells(i, 2)) = UCase(LTB_TestNamePin(index)) And UCase(LTB_TestNamePin(index)) <> "" And LTB_TestLowLimit(index) = "" And LTB_TestHighLimit(index) = "" Then
                    LimitTableSheet.Cells(i, 8) = "F"
                    
                ElseIf UCase(LimitTableSheet.Cells(i, 2)) = UCase(LTB_TestNamePin(index)) And UCase(LTB_TestNamePin(index)) <> "" And LTB_TestLowLimit(index) <> "" And LTB_TestHighLimit(index) <> "" Then
                    LimitTableSheet.Cells(i, 8) = "F"
                
                ElseIf UCase(LTB_TestNamePin(index)) = "" And LTB_TestLowLimit(index) = "" And LTB_TestHighLimit(index) = "" Then
                      If b_FirstTimeFlag_OnlyTestInst = False Then
                        b_FirstTimeFlag_OnlyTestInst = True
                      Else
                        LimitTableSheet.Cells(i, 8) = "F"
                      End If
                End If
            End If
        End If
    
    Next i

Next index


End Sub
Sub SetupExecutionTypeForLimitTable(b_CP_Update As Boolean, b_FT_Update As Boolean)

Dim LT_Sheet As Worksheet
Set LT_Sheet = Worksheets("A_Limits_Table")

If b_CP_Update = True Then
    LT_Sheet.Cells(1, 4).value = "CP1_Func"
Else
    LT_Sheet.Cells(1, 4).value = "Disable"
End If
   
If b_FT_Update = True Then
    LT_Sheet.Cells(1, 6).value = "FT1_Func"
Else
     LT_Sheet.Cells(1, 6).value = "Disable"
End If

End Sub




Sub UpdateFlowLimit()

Dim t_TimeStart As Date
Dim t_TimeStop As Date
Dim t_TimeDiff As Long

t_TimeStart = Now

Dim LimitTableSheet As Worksheet
Dim InstanceList As Worksheet

Set LimitTableSheet = Worksheets("A_Limits_Table")

Dim LT_RowsMax As Long
Dim LT_RowsIndex As Long

Dim LT_TestInstance() As String

LT_RowsMax = LimitTableSheet.Cells(Rows.Count, 1).End(xlUp).Row

ReDim LT_TestInstance(LT_RowsMax - 1 - LT_RowsShift) As String

For LT_RowsIndex = 0 To LT_RowsMax - 1 - LT_RowsShift
    LT_TestInstance(LT_RowsIndex) = LimitTableSheet.Cells(LT_RowsIndex + 1 + LT_RowsShift, 1).value
Next LT_RowsIndex


''' Check Duplicate Test Instance in limit table
Dim b_SortedTestInstDuplicateCheck() As Boolean
ReDim b_SortedTestInstDuplicateCheck(LT_RowsMax - 1 - LT_RowsShift) As Boolean

Dim LT_TestInstanceCompared As String
Dim l_Index As Long
Dim l_SortedIndex As Long

For l_Index = 0 To LT_RowsMax - 1 - LT_RowsShift
    LT_TestInstanceCompared = LT_TestInstance(l_Index)
            
    For l_SortedIndex = 0 To LT_RowsMax - 1 - LT_RowsShift
        
        If l_SortedIndex > l_Index Then
            If b_SortedTestInstDuplicateCheck(l_SortedIndex) = False Then
                If UCase(LT_TestInstanceCompared) = UCase(LT_TestInstance(l_SortedIndex)) Then
                    b_SortedTestInstDuplicateCheck(l_SortedIndex) = True
                End If
            End If
        End If
    
    Next l_SortedIndex
Next l_Index

Dim l_TestInst_Index As Long
Dim l_TestInst_Max As Long
l_TestInst_Index = 0

For l_Index = 0 To LT_RowsMax - 1 - LT_RowsShift
    If b_SortedTestInstDuplicateCheck(l_Index) = False Then
    l_TestInst_Index = l_TestInst_Index + 1
    End If
Next l_Index

l_TestInst_Max = l_TestInst_Index
l_TestInst_Index = 0

Dim LT_SortedTestInstance() As String   '' Sorted Test Instances, remove duplicate
ReDim LT_SortedTestInstance(l_TestInst_Max - 1) As String

For l_Index = 0 To LT_RowsMax - 1 - LT_RowsShift
    If b_SortedTestInstDuplicateCheck(l_Index) = False Then
        LT_SortedTestInstance(l_TestInst_Index) = LT_TestInstance(l_Index)
        l_TestInst_Index = l_TestInst_Index + 1
    End If
Next l_Index

Dim LT_CPFT_TestName() As String
Dim LT_CPFT_TestNumber() As String
Dim LT_CPFT_LowLimit() As String
Dim LT_CPFT_HighLimit() As String
Dim LT_CPFT_PinsCounterForEachTestINst() As Integer
Dim LT_CPFT_LL_Formula() As String
Dim LT_CPFT_HL_Formula() As String
Dim LT_CPFT_JobGate() As String
''' CP/FT use the same use limit
Call CPFT_SortLimitTable(LT_SortedTestInstance, LimitTableSheet, _
                                      LT_CPFT_TestName, LT_CPFT_TestNumber, LT_CPFT_LowLimit, LT_CPFT_HighLimit, LT_CPFT_PinsCounterForEachTestINst, _
                                      LT_CPFT_LL_Formula, LT_CPFT_HL_Formula, LT_CPFT_JobGate)


Dim LT_CP_TestName() As String
Dim LT_CP_TestNumber() As Long
Dim LT_CP_LowLimit() As Double
Dim LT_CP_HighLimit() As Double
Dim LT_CP_PinsCounterForEachTestINst() As Integer
Dim LT_CP_LL_Formula() As String
Dim LT_CP_HL_Formula() As String

Dim LT_FT_TestName() As String
Dim LT_FT_TestNumber() As Long
Dim LT_FT_LowLimit() As Double
Dim LT_FT_HighLimit() As Double
Dim LT_FT_PinsCounterForEachTestINst() As Integer
Dim LT_FT_LL_Formula() As String
Dim LT_FT_HL_Formula() As String

''' Use in CP/FT use-limit are independently and use in Test Instance combine _CP_NV, _CP_HV, _CP_LV, _FT_NV, _FT_HV, _FT_LV
''' Check limit table to sort correspond High/Low limit value and pins for each test instance (Remove the pin if correspond High/Low limit value is 0)
Call SortLimitTable(LT_SortedTestInstance, LimitTableSheet, _
                            LT_CP_TestName, LT_CP_TestNumber, LT_CP_LowLimit, LT_CP_HighLimit, LT_CP_PinsCounterForEachTestINst, _
                            LT_CP_LL_Formula, LT_CP_HL_Formula, _
                            LT_FT_TestName, LT_FT_TestNumber, LT_FT_LowLimit, LT_FT_HighLimit, LT_FT_PinsCounterForEachTestINst, _
                            LT_FT_LL_Formula, LT_FT_HL_Formula)


'' Use LT_SortedTestInstance to search InstanceList table to check which flow need to update test limit
Set InstanceList = Worksheets("InstanceList")

Dim IL_Index As Integer
Dim IL_RowsMax As Integer

IL_RowsMax = InstanceList.Cells(Rows.Count, 1).End(xlUp).Row

Dim IL_RowsShift As Integer
IL_RowsShift = 2
Dim IL_TestInatName() As String
ReDim IL_TestInatName(IL_RowsMax - 1 - IL_RowsShift) As String

If UBound(IL_TestInatName) <> UBound(LT_SortedTestInstance) Then
    Debug.Print "Number of Test Instance is different between Limit Table and Instance List"
End If

Dim LT_SortedIndex As Integer

Dim IL_EachFlow() As String
Dim IL_SearchFlowMax As Integer
Dim IL_SearchFlowIndex As Integer

Dim SearchedFlow As Worksheet
Dim SearchedFlowName As String
Dim SearchedFlowRowMax As Integer
Dim SearchedFlowIndex As Integer
Dim SearchedInstIndex As Integer

Dim FindoutInstMax As Integer
Dim FindoutInstIndex As Integer
Dim FindoutInstArrayIndex() As Integer
Dim IndexCounter As Integer
Dim UserLimitCountArray() As Integer

''' Check Flow whether empty: If empty, use Test Instance combine _CP_NV, _CP_HV, _CP_LV, _FT_NV, _FT_HV, _FT_LV to do seach of each Flow Table
Dim CombineInstEnd() As Variant
CombineInstEnd = Array("_CP_NV", "_CP_HV", "_CP_LV", "_FT_NV", "_FT_HV", "_FT_LV")
Dim index As Integer
Dim b_CheckflowDuplicate As Boolean
Dim CellShiftCounter As Integer

Dim SortedTestInstCombineEnd As String


For LT_SortedIndex = 0 To UBound(LT_SortedTestInstance)     ''' Test instance count after sorted of limit table
    For IL_Index = 0 To IL_RowsMax - 1 - IL_RowsShift               ''' Test instance count of instance list table (already sorted by GetFlowInfo)
        If InstanceList.Cells(IL_Index + 1 + IL_RowsShift, 3) = "" Then
        
            If UCase(LT_SortedTestInstance(LT_SortedIndex)) = UCase(InstanceList.Cells(IL_Index + 1 + IL_RowsShift, 1)) Then    ''' Compare instance name between limit table and instance list table (sorted)
                
                IL_EachFlow = Split(InstanceList.Cells(IL_Index + 1 + IL_RowsShift, 2), ",")    ''' Seprate flow table name of instance list table for current test instance
            
                For IL_SearchFlowIndex = 0 To UBound(IL_EachFlow)   ''' Remove current use limit and re-write use limit of each flow
                    Set SearchedFlow = Worksheets(IL_EachFlow(IL_SearchFlowIndex))
                    SearchedFlowName = SearchedFlow.name
                    Call SearchedFlow.OutLine.ShowLevels(2, 1) ''' Expand rows group of user limit
                    SearchedFlowRowMax = SearchedFlow.Cells(Rows.Count, 7).End(xlUp).Row
                    
                    ''' Get test instance count of flow table
                    For SearchedFlowIndex = 5 To SearchedFlowRowMax
                        If UCase(SearchedFlow.Cells(SearchedFlowIndex, 7).value) = "TEST" And UCase(SearchedFlow.Cells(SearchedFlowIndex, 8).value) = UCase(LT_SortedTestInstance(LT_SortedIndex)) Then
                            SearchedInstIndex = SearchedInstIndex + 1
                        End If
                    Next SearchedFlowIndex
                    
                    ''' Record index when opcode = Test and parameter content equal to LT_SortedTestInstance of flow table
                    FindoutInstMax = SearchedInstIndex
                    ReDim FindoutInstArrayIndex(FindoutInstMax) As Integer
                    
                    For SearchedFlowIndex = 5 To SearchedFlowRowMax
                        If UCase(SearchedFlow.Cells(SearchedFlowIndex, 7).value) = "TEST" And UCase(SearchedFlow.Cells(SearchedFlowIndex, 8).value) = UCase(LT_SortedTestInstance(LT_SortedIndex)) Then
                            FindoutInstArrayIndex(FindoutInstIndex) = SearchedFlowIndex
                            FindoutInstIndex = FindoutInstIndex + 1
                            
                            If FindoutInstIndex = UBound(FindoutInstArrayIndex) Then
                                FindoutInstArrayIndex(FindoutInstIndex) = SearchedFlowRowMax
                            End If
                        End If
                    Next SearchedFlowIndex
                    
                    ''' Caculate total user limit count of each test item in flow table
                    ReDim UserLimitCountArray(FindoutInstMax - 1) As Integer
                    For IndexCounter = 0 To FindoutInstMax - 1
                        For SearchedFlowIndex = FindoutInstArrayIndex(IndexCounter) To FindoutInstArrayIndex(IndexCounter + 1)
                            If UCase(SearchedFlow.Cells(SearchedFlowIndex, 7).value) = UCase("Use-Limit") And UCase(SearchedFlow.Cells(SearchedFlowIndex, 8).value) = UCase(LT_SortedTestInstance(LT_SortedIndex)) Then
                                UserLimitCountArray(IndexCounter) = UserLimitCountArray(IndexCounter) + 1
                            End If
                        Next SearchedFlowIndex
    
                    Next IndexCounter
                    
                    For IndexCounter = FindoutInstMax - 1 To 0 Step -1
                        If UserLimitCountArray(IndexCounter) <> 0 Then
                        
                            ''' Remove user limit
                            SearchedFlow.Rows.range(CStr((FindoutInstArrayIndex(IndexCounter) + 1)) & ":" & CStr(FindoutInstArrayIndex(IndexCounter) + UserLimitCountArray(IndexCounter))).delete
                        End If
                        
                        ''' Update user limit from sorted limit table
                        If b_Enable_CPFT_UseTheSameLimit Then
                            Call CPFT_WriteLimitTableToFlowTable(SearchedFlow, LT_SortedTestInstance(LT_SortedIndex), LT_SortedIndex, FindoutInstArrayIndex(IndexCounter), _
                                                                                      LT_CPFT_TestName, LT_CPFT_TestNumber, LT_CPFT_LowLimit, LT_CPFT_HighLimit, LT_CPFT_PinsCounterForEachTestINst, _
                                                                                      LT_CPFT_LL_Formula, LT_CPFT_HL_Formula, LT_CPFT_JobGate, b_DisplayPinName, b_DisplayTestNumber)
                        
                        
                        Else
                            Call WriteLimitTableToFlowTable(SearchedFlow, LT_SortedTestInstance(LT_SortedIndex), LT_SortedIndex, FindoutInstArrayIndex(IndexCounter), _
                                                                            LT_CP_TestName, LT_CP_TestNumber, LT_CP_LowLimit, LT_CP_HighLimit, LT_CP_PinsCounterForEachTestINst, _
                                                                            LT_CP_LL_Formula, LT_CP_HL_Formula, _
                                                                            LT_FT_TestName, LT_FT_TestNumber, LT_FT_LowLimit, LT_FT_HighLimit, LT_FT_PinsCounterForEachTestINst, _
                                                                            LT_FT_LL_Formula, LT_FT_HL_Formula, b_DisplayPinName, b_DisplayTestNumber)
                        End If
                    Next IndexCounter
                    
                    SearchedInstIndex = 0
                    FindoutInstIndex = 0
                Next IL_SearchFlowIndex
                    
            End If
        
        Else
            If UCase(LT_SortedTestInstance(LT_SortedIndex)) = UCase(InstanceList.Cells(IL_Index + 1 + IL_RowsShift, 1)) Then
                IL_EachFlow = Split(InstanceList.Cells(IL_Index + 1 + IL_RowsShift, 2), ",")
            
                For IL_SearchFlowIndex = 0 To UBound(IL_EachFlow)
                    Set SearchedFlow = Worksheets(IL_EachFlow(IL_SearchFlowIndex))
                    SearchedFlowName = SearchedFlow.name
                    Call SearchedFlow.OutLine.ShowLevels(2, 1) ''' Expand rows group of user limit
'''                    SearchedFlowRowMax = SearchedFlow.Cells(Rows.count, 7).End(xlUp).Row
                    
                    For index = 0 To UBound(CombineInstEnd)
                        ''' Get test instance count of flow table
                        SearchedFlowRowMax = SearchedFlow.Cells(Rows.Count, 7).End(xlUp).Row
                        
                        For SearchedFlowIndex = 5 To SearchedFlowRowMax
                            If UCase(SearchedFlow.Cells(SearchedFlowIndex, 7).value) = "TEST" And UCase(SearchedFlow.Cells(SearchedFlowIndex, 8).value) = UCase(LT_SortedTestInstance(LT_SortedIndex) & CombineInstEnd(index)) Then
                                SearchedInstIndex = SearchedInstIndex + 1
                            End If
                        Next SearchedFlowIndex
                        
                        ''' Record index when opcode = Test and parameter content equal to LT_SortedTestInstance of flow table
                        FindoutInstMax = SearchedInstIndex
                        If FindoutInstMax > 0 Then
                            ReDim FindoutInstArrayIndex(FindoutInstMax) As Integer
                            
                            For SearchedFlowIndex = 5 To SearchedFlowRowMax
                                If UCase(SearchedFlow.Cells(SearchedFlowIndex, 7).value) = "TEST" And UCase(SearchedFlow.Cells(SearchedFlowIndex, 8).value) = UCase(LT_SortedTestInstance(LT_SortedIndex) & CombineInstEnd(index)) Then
                                    FindoutInstArrayIndex(FindoutInstIndex) = SearchedFlowIndex
                                    FindoutInstIndex = FindoutInstIndex + 1
                                    
                                    If FindoutInstIndex = UBound(FindoutInstArrayIndex) Then
                                        FindoutInstArrayIndex(FindoutInstIndex) = SearchedFlowRowMax
                                    End If
                                End If
                            Next SearchedFlowIndex
                            
                            ''' Caculate total user limit count of each test item in flow table
                            ReDim UserLimitCountArray(FindoutInstMax - 1) As Integer
                            For IndexCounter = 0 To FindoutInstMax - 1
                                For SearchedFlowIndex = FindoutInstArrayIndex(IndexCounter) To FindoutInstArrayIndex(IndexCounter + 1)
                                    If UCase(SearchedFlow.Cells(SearchedFlowIndex, 7).value) = UCase("Use-Limit") And UCase(SearchedFlow.Cells(SearchedFlowIndex, 8).value) = UCase(LT_SortedTestInstance(LT_SortedIndex) & CombineInstEnd(index)) Then
                                        UserLimitCountArray(IndexCounter) = UserLimitCountArray(IndexCounter) + 1
                                    End If
                                Next SearchedFlowIndex
                            Next IndexCounter
                            
                            For IndexCounter = FindoutInstMax - 1 To 0 Step -1
                                ''' Remove user limit
                                If UserLimitCountArray(IndexCounter) <> 0 Then
                                    SearchedFlow.Rows.range(CStr((FindoutInstArrayIndex(IndexCounter) + 1)) & ":" & CStr(FindoutInstArrayIndex(IndexCounter) + UserLimitCountArray(IndexCounter))).delete
                                End If
                                
                                SortedTestInstCombineEnd = LT_SortedTestInstance(LT_SortedIndex) & CombineInstEnd(index)
                                ''' Update user limit from sorted limit table

                                Call WriteLimitTableToFlowTable(SearchedFlow, SortedTestInstCombineEnd, LT_SortedIndex, FindoutInstArrayIndex(IndexCounter), _
                                                                                LT_CP_TestName, LT_CP_TestNumber, LT_CP_LowLimit, LT_CP_HighLimit, LT_CP_PinsCounterForEachTestINst, _
                                                                                LT_CP_LL_Formula, LT_CP_HL_Formula, _
                                                                                LT_FT_TestName, LT_FT_TestNumber, LT_FT_LowLimit, LT_FT_HighLimit, LT_FT_PinsCounterForEachTestINst, _
                                                                                LT_FT_LL_Formula, LT_FT_HL_Formula, b_DisplayPinName, b_DisplayTestNumber)
                            Next IndexCounter
                        End If
                        
                        SearchedInstIndex = 0 ''' Total used times of current test instance of flow table
                        FindoutInstIndex = 0    ''' Index of current test instance of flow table
                    
                    Next index
                
                Next IL_SearchFlowIndex
                    
            End If
        
        End If
        
    Next IL_Index
Next LT_SortedIndex


t_TimeStop = Now
t_TimeDiff = DateDiff("s", t_TimeStart, t_TimeStop)

Debug.Print "Cost time = " & t_TimeDiff & " sec"
'''TheExec.Datalog.WriteComment ("Cost time = " & t_TimeDiff & " sec")
End Sub

Public Function UpdateLimitTable_Main() As Long
    
    '' Must setup below 4 flags in Declarations before running UpdateLimitTable_Main
    '' b_DisplayPinName      - Decide update pin name to use limit or not
    '' b_DisplayTestNumber  - Decide update test number to use limit or not
    '' b_EnableCP_Update    - Update CP use limit from limit table or not
    '' b_EnableFT_Update     - Update FT use limit from limit table or not
    
    '' 1. Transfer "Datalog" to "A_Limits_Table" format
    Call DatalogExtractDuplicate

    '' 1. According to "A_Limits_Table_Bypass" to setup User Limit Group Flag to "F" in "A_Limits_Table" to bypass update use limit.
    Call LimitTableBypassCheck

    '' 1. Sort Test Instance from "A_Limits_Table" and update to "InstanceList"
    '' 2. InstanceList: List Test Instance use for which flow table
    '' 3. Check Flow content whether empty: If empty, use Test Instance combine
    ''     _CP_NV, _CP_HV, _CP_LV, _FT_NV, _FT_HV, _FT_LV to seach each Flow Table

    Call GetFlowInfo

    '' 1. Select execution type to update limit table (CP or FT)
    If b_Enable_CPFT_UseTheSameLimit Then
        ' Force CP and FT to True
        Call SetupExecutionTypeForLimitTable(True, True)
    Else
        ' Force by global flag setting
        Call SetupExecutionTypeForLimitTable(b_EnableCP_Update, b_EnableFT_Update)
    End If
    
    ' 1. Update use-limit of to flow table by reference "A_Limits_Table" and "InstanceList"
    Call UpdateFlowLimit
    
    MsgBox ("New Test Limit Flow Done")
    
End Function



Sub DatalogExtractDuplicate() 'Remove b_RepeatCheck Instance Name

Dim RowStart As Integer
Dim RowEnd As Integer

Dim DatalogInstanceName As String
Dim AfterSortInstanceName As String

Dim DatalogTestName As String
Dim AfterSortTestName As String
Dim Count As Long
Dim DatalogPinName As String
Dim AfterSortPinName As String
Dim x As Long
Dim i As Integer
Dim j As Integer


Dim b_RepeatCheck As Boolean

Dim TargetSheet As Worksheet
Dim OutputSheet As Worksheet

Dim s_PatternName As String

Set TargetSheet = Sheets("Datalog")
Set OutputSheet = Sheets("A_Limits_Table")
OutputSheet.Cells.Clear

RowStart = 2
RowEnd = TargetSheet.Cells(Rows.Count, 4).End(xlUp).Row

Dim UnitFactor As Double
Dim b_TestInstFirstCheck As Boolean

Dim DatalogInstNameCurrent As String
Dim DatalogInstNamePrevious As String

For i = RowStart To RowEnd
    DatalogInstNameCurrent = TargetSheet.Cells(i, 4)
    If DatalogInstNameCurrent = DatalogInstNamePrevious Then
        TargetSheet.Cells(i, 16) = "True"
    Else
        TargetSheet.Cells(i, 16) = "False"
    End If
    
    DatalogInstNamePrevious = DatalogInstNameCurrent
Next i

For i = RowStart To RowEnd
        If i = RowStart Then
         
            DatalogInstanceName = UCase(TargetSheet.Cells(i, 4).value)
            AfterSortInstanceName = DatalogInstanceName
            OutputSheet.Cells(i, 1).value = UCase(AfterSortInstanceName)    '' Test Instance
            Count = 1
                     
        Else
            b_RepeatCheck = False
            DatalogInstanceName = UCase(TargetSheet.Cells(i, 4).value)   '' Test Instance

            
            For j = 0 To Count
             
                If DatalogInstanceName = OutputSheet.Cells(RowStart + j, 1).value Then '' Compare test instance between datalog and limit table to check whether duplicate
                    b_RepeatCheck = True
                    j = Count
                End If
            Next j
            
            If b_RepeatCheck <> True Then
                
                AfterSortInstanceName = DatalogInstanceName
                OutputSheet.Cells(RowStart + Count, 1).value = UCase(AfterSortInstanceName)     '' Test Instance
                
                Count = Count + 1
            End If
        End If
Next i

Dim OutputSheetRowStart As Integer
Dim OutputSheetRowEnd As Integer
OutputSheetRowStart = 2
OutputSheetRowEnd = OutputSheet.Cells(Rows.Count, 1).End(xlUp).Row

Dim TestInstMax As Integer
TestInstMax = OutputSheetRowEnd - OutputSheetRowStart + 1
Dim TestInstName() As String
ReDim TestInstName(TestInstMax - 1) As String

For i = OutputSheetRowStart To OutputSheetRowEnd
    TestInstName(i - OutputSheetRowStart) = OutputSheet.Cells(i, 1)
Next i


TargetSheet.Columns.range("Q:Q").ClearContents
Dim b_TestInstRecordFlag() As Boolean
ReDim b_TestInstRecordFlag(TestInstMax - 1) As Boolean

Dim TestInstIndex As Integer
Dim RowsIndexCurrent As Integer
Dim RowsIndexPrevious As Integer

Dim b_FirstTimeFlag As Boolean

For TestInstIndex = 0 To UBound(TestInstName)
    
    b_FirstTimeFlag = False
    b_TestInstRecordFlag(TestInstIndex) = False
    
    For i = RowStart To RowEnd
    
        If TargetSheet.Cells(i, 12).value = "N/A" Then
           
            If UCase(TargetSheet.Cells(i, 4).value) = UCase(TestInstName(TestInstIndex)) Then
            
                If b_FirstTimeFlag = False Then
                    RowsIndexCurrent = i
                    RowsIndexPrevious = RowsIndexCurrent
                    TargetSheet.Cells(i, 17).value = "V"
                    b_FirstTimeFlag = True
                Else
                    RowsIndexCurrent = i
                    If b_TestInstRecordFlag(TestInstIndex) = False Then
                        If (RowsIndexCurrent - RowsIndexPrevious) = 1 Then
                            TargetSheet.Cells(i, 17).value = "V"
                        Else
                            b_TestInstRecordFlag(TestInstIndex) = True
                        End If
                    End If
                    
                    RowsIndexPrevious = RowsIndexCurrent
                
                End If
            End If
        End If
    Next i
Next TestInstIndex

OutputSheet.Cells.Clear

Dim RPT As Integer
Dim IndexCounter As Integer
IndexCounter = 0
For i = RowStart To RowEnd
    If UCase(TargetSheet.Cells(i, 17).value) = "V" Then
     If UCase(TargetSheet.Cells(i, 15).value) = "" Then
      RPT = 1
      Else
      RPT = TargetSheet.Cells(i, 15).value
      End If
      
     For x = 1 To RPT
    
        DatalogInstanceName = UCase(TargetSheet.Cells(i, 4).value)   '' Test Instance
        DatalogTestName = UCase(TargetSheet.Cells(i, 6).value)         '' Test Name
        DatalogPinName = UCase(TargetSheet.Cells(i, 7).value)           '' Test Pin

        OutputSheet.Cells(RowStart + IndexCounter, 1).value = UCase(DatalogInstanceName)   '' Test Instance

        OutputSheet.Cells(RowStart + IndexCounter, 10).value = UCase(DatalogTestName)          '' Test Name

        OutputSheet.Cells(RowStart + IndexCounter, 2).value = UCase(DatalogPinName)             '' Test Pin

        OutputSheet.Cells(RowStart + IndexCounter, 3).value = TargetSheet.Cells(i, 5).value        '' Test Number

        If UCase(TargetSheet.Cells(i, 11).value) = UCase("GHZ") Then
            UnitFactor = 1000000000
        ElseIf UCase(TargetSheet.Cells(i, 11).value) = UCase("MHZ") Then
            UnitFactor = 1000000
        ElseIf UCase(TargetSheet.Cells(i, 11).value) = UCase("KHZ") Then
            UnitFactor = 1000
        ElseIf UCase(TargetSheet.Cells(i, 11).value) = UCase("mA") Or UCase(TargetSheet.Cells(i, 11).value) = UCase("mV") Then
            UnitFactor = 0.001
        ElseIf UCase(TargetSheet.Cells(i, 11).value) = UCase("uA") Or UCase(TargetSheet.Cells(i, 11).value) = UCase("uV") Then
            UnitFactor = 0.000001
        ElseIf UCase(TargetSheet.Cells(i, 11).value) = UCase("nA") Or UCase(TargetSheet.Cells(i, 11).value) = UCase("nV") Then
            UnitFactor = 0.000000001
        Else
            UnitFactor = 1
        End If

        OutputSheet.Cells(RowStart + IndexCounter, 4).value = TargetSheet.Cells(i, 8).value * UnitFactor    '' Low limit
        OutputSheet.Cells(RowStart + IndexCounter, 5).value = TargetSheet.Cells(i, 10).value * UnitFactor  '' High limit
        
        IndexCounter = IndexCounter + 1
      Next x
    End If
Next i

OutputSheet.Rows.range("1:1").Insert shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
OutputSheet.Cells(2, 1).value = "TestInstance"
OutputSheet.Cells(2, 2).value = "TestName (Pin)"
OutputSheet.Cells(2, 3).value = "TestNum"
OutputSheet.Cells(2, 4).value = "Low Limit"
OutputSheet.Cells(2, 5).value = "High Limit"
OutputSheet.Cells(2, 8).value = "User Limit Group Flag"

End Sub



Sub Backup_DatalogExtractDuplicate() 'Remove b_RepeatCheck Instance Name

Dim RowStart As Integer
Dim RowEnd As Integer

Dim DatalogInstanceName As String
Dim AfterSortInstanceName As String

Dim DatalogTestName As String
Dim AfterSortTestName As String

Dim DatalogPinName As String
Dim AfterSortPinName As String

Dim i As Integer
Dim j As Integer
Dim Count As Integer

Dim b_RepeatCheck As Boolean

Dim TargetSheet As Worksheet
Dim OutputSheet As Worksheet

Dim s_PatternName As String

Set TargetSheet = Sheets("Datalog")
Set OutputSheet = Sheets("A_Limits_Table")
OutputSheet.Cells.Clear

RowStart = 2
RowEnd = TargetSheet.Cells(Rows.Count, 4).End(xlUp).Row

Dim UnitFactor As Double
Dim b_TestInstFirstCheck As Boolean

For i = RowStart To RowEnd
    If TargetSheet.Cells(i, 12).value = "N/A" Then
    
        If i = RowStart Then
         
            DatalogInstanceName = UCase(TargetSheet.Cells(i, 4).value)
            AfterSortInstanceName = DatalogInstanceName
            OutputSheet.Cells(i, 1).value = UCase(AfterSortInstanceName)    '' Test Instance
            
            DatalogTestName = UCase(TargetSheet.Cells(i, 6).value) '
            AfterSortTestName = DatalogTestName
            OutputSheet.Cells(i, 10).value = UCase(AfterSortTestName)           '' Test Name
            
            DatalogPinName = UCase(TargetSheet.Cells(i, 7).value)
            AfterSortPinName = DatalogPinName
            OutputSheet.Cells(i, 2).value = UCase(AfterSortPinName)             '' Test Pin
            
            OutputSheet.Cells(i, 3).value = TargetSheet.Cells(i, 5).value         '' Test Number
            
            If UCase(TargetSheet.Cells(i, 11).value) = UCase("GHZ") Then
                UnitFactor = 1000000000
            ElseIf UCase(TargetSheet.Cells(i, 11).value) = UCase("MHZ") Then
                UnitFactor = 1000000
            ElseIf UCase(TargetSheet.Cells(i, 11).value) = UCase("KHZ") Then
                UnitFactor = 1000
            ElseIf UCase(TargetSheet.Cells(i, 11).value) = UCase("mA") Or UCase(TargetSheet.Cells(i, 11).value) = UCase("mV") Then
                UnitFactor = 0.001
            ElseIf UCase(TargetSheet.Cells(i, 11).value) = UCase("uA") Or UCase(TargetSheet.Cells(i, 11).value) = UCase("uV") Then
                UnitFactor = 0.000001
            ElseIf UCase(TargetSheet.Cells(i, 11).value) = UCase("nA") Or UCase(TargetSheet.Cells(i, 11).value) = UCase("nV") Then
                UnitFactor = 0.000000001
            Else
                UnitFactor = 1
            End If
            OutputSheet.Cells(i, 4).value = TargetSheet.Cells(i, 8).value * UnitFactor      '' Low limit
            OutputSheet.Cells(i, 5).value = TargetSheet.Cells(i, 10).value * UnitFactor    '' High limit
            
            Count = 1
         
        Else
            b_RepeatCheck = False
            DatalogInstanceName = UCase(TargetSheet.Cells(i, 4).value)   '' Test Instance
            DatalogTestName = UCase(TargetSheet.Cells(i, 6).value)         '' Test Name
            DatalogPinName = UCase(TargetSheet.Cells(i, 7).value)           '' Test Pin
            
            For j = 0 To Count
             
                If DatalogInstanceName = OutputSheet.Cells(RowStart + j, 1).value And DatalogTestName = OutputSheet.Cells(RowStart + j, 10).value And DatalogPinName = OutputSheet.Cells(RowStart + j, 2).value Then '' Compare test instance/name/pin between datalog and limit table to check whether duplicate
                    b_RepeatCheck = True
                    j = Count
                End If
                
    
            Next j
            
            If b_RepeatCheck <> True Then
                
                AfterSortInstanceName = DatalogInstanceName
                OutputSheet.Cells(RowStart + Count, 1).value = UCase(AfterSortInstanceName)     '' Test Instance
                
                AfterSortTestName = DatalogTestName
                OutputSheet.Cells(RowStart + Count, 10).value = UCase(AfterSortTestName)           '' Test Name
                
                AfterSortPinName = DatalogPinName
                OutputSheet.Cells(RowStart + Count, 2).value = UCase(AfterSortPinName)             '' Test Pin
                
                OutputSheet.Cells(RowStart + Count, 3).value = TargetSheet.Cells(i, 5).value         '' Test Number
                
                If UCase(TargetSheet.Cells(i, 11).value) = UCase("GHZ") Then
                    UnitFactor = 1000000000
                ElseIf UCase(TargetSheet.Cells(i, 11).value) = UCase("MHZ") Then
                    UnitFactor = 1000000
                ElseIf UCase(TargetSheet.Cells(i, 11).value) = UCase("KHZ") Then
                    UnitFactor = 1000
                ElseIf UCase(TargetSheet.Cells(i, 11).value) = UCase("mA") Or UCase(TargetSheet.Cells(i, 11).value) = UCase("mV") Then
                    UnitFactor = 0.001
                ElseIf UCase(TargetSheet.Cells(i, 11).value) = UCase("uA") Or UCase(TargetSheet.Cells(i, 11).value) = UCase("uV") Then
                    UnitFactor = 0.000001
                ElseIf UCase(TargetSheet.Cells(i, 11).value) = UCase("nA") Or UCase(TargetSheet.Cells(i, 11).value) = UCase("nV") Then
                    UnitFactor = 0.000000001
                Else
                    UnitFactor = 1
                End If
                
                OutputSheet.Cells(RowStart + Count, 4).value = TargetSheet.Cells(i, 8).value * UnitFactor   '' Low limit
                OutputSheet.Cells(RowStart + Count, 5).value = TargetSheet.Cells(i, 10).value * UnitFactor '' High limit
                
                
                Count = Count + 1
                
            End If
         
        End If
        
    End If
    
Next i



OutputSheet.Rows.range("1:1").Insert shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
OutputSheet.Cells(2, 1).value = "TestInstance"
OutputSheet.Cells(2, 2).value = "TestName (Pin)"
OutputSheet.Cells(2, 3).value = "TestNum"
OutputSheet.Cells(2, 4).value = "Low Limit"
OutputSheet.Cells(2, 5).value = "High Limit"
OutputSheet.Cells(2, 8).value = "User Limit Group Flag"

End Sub


Sub CPFT_WriteLimitTableToFlowTable(FlowT_Sheet As Worksheet, LT_SortedTestInstanceName As String, LT_SortedTestInstIndex As Integer, FlowT_FindoutInstRowsIndex, _
                                                ByRef LT_CPFT_TestName() As String, ByRef LT_CPFT_TestNumber() As String, ByRef LT_CPFT_LowLimit() As String, ByRef LT_CPFT_HighLimit() As String, ByRef LT_CPFT_PinsCounterForEachTestINst() As Integer, _
                                                ByRef LT_CPFT_LL_Formula() As String, ByRef LT_CPFT_HL_Formula() As String, ByRef LT_CPFT_JobGate() As String, _
                                                Optional b_DisplayPinName As Boolean = True, Optional b_DisplayTestNum As Boolean = True)

Dim LT_Sheet As Worksheet
Dim InsertRowsMax As Long

Set LT_Sheet = Worksheets("A_Limits_Table")

InsertRowsMax = LT_CPFT_PinsCounterForEachTestINst(LT_SortedTestInstIndex)
Dim InsertRowStart As Integer
Dim InsertRowEnd As Integer
InsertRowStart = FlowT_FindoutInstRowsIndex + 1
InsertRowEnd = FlowT_FindoutInstRowsIndex + InsertRowsMax

If (InsertRowEnd - InsertRowStart) < 0 Then
    Exit Sub
End If

FlowT_Sheet.Rows.range(InsertRowStart & ":" & InsertRowEnd).Insert shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove

Dim index As Integer

    FlowT_Sheet.range("G" & (FlowT_FindoutInstRowsIndex + 1) & ":G" & FlowT_FindoutInstRowsIndex + InsertRowsMax) = "Use-Limit"
    FlowT_Sheet.range("H" & (FlowT_FindoutInstRowsIndex + 1) & ":H" & FlowT_FindoutInstRowsIndex + InsertRowsMax) = LT_SortedTestInstanceName
    

For index = 0 To LT_CPFT_PinsCounterForEachTestINst(LT_SortedTestInstIndex) - 1
    
    If b_DisplayPinName Then
        FlowT_Sheet.Cells(FlowT_FindoutInstRowsIndex + 1 + index, 9).value = LT_CPFT_TestName(LT_SortedTestInstIndex, index)
    End If
    
    If b_DisplayTestNum Then
        FlowT_Sheet.Cells(FlowT_FindoutInstRowsIndex + 1 + index, 10).value = LT_CPFT_TestNumber(LT_SortedTestInstIndex, index)
    End If
    
''        FlowT_Sheet.Cells(FlowT_FindoutInstRowsIndex + 1 + Index, 11).Value = LT_CP_LowLimit(LT_SortedTestInstIndex, Index)
''        FlowT_Sheet.Cells(FlowT_FindoutInstRowsIndex + 1 + Index, 12).Value = LT_CP_HighLimit(LT_SortedTestInstIndex, Index)
    FlowT_Sheet.Cells(FlowT_FindoutInstRowsIndex + 1 + index, 11).value = LT_CPFT_LL_Formula(LT_SortedTestInstIndex, index)
    FlowT_Sheet.Cells(FlowT_FindoutInstRowsIndex + 1 + index, 12).value = LT_CPFT_HL_Formula(LT_SortedTestInstIndex, index)
    
    FlowT_Sheet.Cells(FlowT_FindoutInstRowsIndex + 1 + index, 4).value = LT_CPFT_JobGate(LT_SortedTestInstIndex, index)  '' Job gate
Next index
    
End Sub

Sub WriteLimitTableToFlowTable(FlowT_Sheet As Worksheet, LT_SortedTestInstanceName As String, LT_SortedTestInstIndex As Integer, FlowT_FindoutInstRowsIndex, _
                                                ByRef LT_CP_TestName() As String, ByRef LT_CP_TestNumber() As Long, ByRef LT_CP_LowLimit() As Double, ByRef LT_CP_HighLimit() As Double, ByRef LT_CP_PinsCounterForEachTestINst() As Integer, _
                                                ByRef LT_CP_LL_Formula() As String, ByRef LT_CP_HL_Formula() As String, _
                                                ByRef LT_FT_TestName() As String, ByRef LT_FT_TestNumber() As Long, ByRef LT_FT_LowLimit() As Double, ByRef LT_FT_HighLimit() As Double, ByRef LT_FT_PinsCounterForEachTestINst() As Integer, _
                                                ByRef LT_FT_LL_Formula() As String, ByRef LT_FT_HL_Formula() As String, _
                                                Optional b_DisplayPinName As Boolean = True, Optional b_DisplayTestNum As Boolean = True)

Dim LT_Sheet As Worksheet
Set LT_Sheet = Worksheets("A_Limits_Table")

Dim splitString() As String
Dim CheckWord As String '' Check CP or FT

splitString = Split(LT_SortedTestInstanceName, "_")
CheckWord = splitString(UBound(splitString) - 1)

Dim InsertRowsMax As Integer

If UCase(LT_Sheet.Cells(1, 4).value) = UCase("Disable") And UCase(LT_Sheet.Cells(1, 6).value) = UCase("Disable") Or _
  UCase(LT_Sheet.Cells(1, 4).value) = "" And UCase(LT_Sheet.Cells(1, 6).value) = "" Or _
  UCase(LT_Sheet.Cells(1, 4).value) = "" And UCase(LT_Sheet.Cells(1, 6).value) = UCase("Disable") Or _
  UCase(LT_Sheet.Cells(1, 4).value) = UCase("Disable") And UCase(LT_Sheet.Cells(1, 6).value) = "" Then
  
    Debug.Print "Please select update type of both orange cells"
    Exit Sub
End If

If UCase(CheckWord) <> "CP" And UCase(CheckWord) <> "FT" Then
    
    If UCase(LT_Sheet.Cells(1, 4).value) = UCase("CP1_Func") And UCase(LT_Sheet.Cells(1, 6).value) = UCase("FT1_Func") Then
        
        InsertRowsMax = LT_CP_PinsCounterForEachTestINst(LT_SortedTestInstIndex) + LT_FT_PinsCounterForEachTestINst(LT_SortedTestInstIndex)
    
    ElseIf UCase(LT_Sheet.Cells(1, 4).value) = UCase("CP1_Func") And UCase(LT_Sheet.Cells(1, 6).value) = UCase("Disable") Then
    
        InsertRowsMax = LT_CP_PinsCounterForEachTestINst(LT_SortedTestInstIndex)
    
    ElseIf UCase(LT_Sheet.Cells(1, 4).value) = UCase("Disable") And UCase(LT_Sheet.Cells(1, 6).value) = UCase("FT1_Func") Then
    
        InsertRowsMax = LT_FT_PinsCounterForEachTestINst(LT_SortedTestInstIndex)
        
    End If
    
ElseIf UCase(CheckWord) = "CP" Then
    InsertRowsMax = LT_CP_PinsCounterForEachTestINst(LT_SortedTestInstIndex)

ElseIf UCase(CheckWord) = "FT" Then
    InsertRowsMax = LT_FT_PinsCounterForEachTestINst(LT_SortedTestInstIndex)

End If

Dim InsertRowStart As Integer
Dim InsertRowEnd As Integer
InsertRowStart = FlowT_FindoutInstRowsIndex + 1
InsertRowEnd = FlowT_FindoutInstRowsIndex + InsertRowsMax

If (InsertRowEnd - InsertRowStart) < 0 Then
    Exit Sub
End If

FlowT_Sheet.Rows.range(InsertRowStart & ":" & InsertRowEnd).Insert shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove

Dim index As Integer
If UCase(CheckWord) <> "CP" And UCase(CheckWord) <> "FT" Then
    
    If UCase(LT_Sheet.Cells(1, 4).value) = UCase("CP1_Func") And UCase(LT_Sheet.Cells(1, 6).value) = UCase("FT1_Func") Then
        If b_Disable_CPFT_JobToEmpty Then
        Else
            FlowT_Sheet.range("D" & (FlowT_FindoutInstRowsIndex + 1) & ":D" & (FlowT_FindoutInstRowsIndex + 1) + LT_CP_PinsCounterForEachTestINst(LT_SortedTestInstIndex)) = "CP1"
            FlowT_Sheet.range("D" & FlowT_FindoutInstRowsIndex + LT_CP_PinsCounterForEachTestINst(LT_SortedTestInstIndex) + 1 & ":D" & FlowT_FindoutInstRowsIndex + InsertRowsMax) = "FT1"
        End If
        
    ElseIf UCase(LT_Sheet.Cells(1, 4).value) = UCase("CP1_Func") And UCase(LT_Sheet.Cells(1, 6).value) = UCase("Disable") Then
        If b_Disable_CPFT_JobToEmpty Then
        Else
            FlowT_Sheet.range("D" & (FlowT_FindoutInstRowsIndex + 1) & ":D" & FlowT_FindoutInstRowsIndex + InsertRowsMax) = "CP1"
        End If
        
    ElseIf UCase(LT_Sheet.Cells(1, 4).value) = UCase("Disable") And UCase(LT_Sheet.Cells(1, 6).value) = UCase("FT1_Func") Then
        If b_Disable_CPFT_JobToEmpty Then
        Else
            FlowT_Sheet.range("D" & (FlowT_FindoutInstRowsIndex + 1) & ":D" & FlowT_FindoutInstRowsIndex + InsertRowsMax) = "FT1"
        End If
        
    End If
    FlowT_Sheet.range("G" & (FlowT_FindoutInstRowsIndex + 1) & ":G" & FlowT_FindoutInstRowsIndex + InsertRowsMax) = "Use-Limit"
    FlowT_Sheet.range("H" & (FlowT_FindoutInstRowsIndex + 1) & ":H" & FlowT_FindoutInstRowsIndex + InsertRowsMax) = LT_SortedTestInstanceName
    
    If UCase(LT_Sheet.Cells(1, 4).value) = UCase("CP1_Func") Then
        For index = 0 To LT_CP_PinsCounterForEachTestINst(LT_SortedTestInstIndex) - 1
    '''        FlowT_Sheet.Cells(FlowT_FindoutInstRowsIndex + 1 + Index, 4).Value = "CP1"
    '''        FlowT_Sheet.Cells(FlowT_FindoutInstRowsIndex + 1 + Index, 7).Value = "Use-Limit"
    '''        FlowT_Sheet.Cells(FlowT_FindoutInstRowsIndex + 1 + Index, 8).Value = LT_SortedTestInstanceName
            
            If b_DisplayPinName Then
                FlowT_Sheet.Cells(FlowT_FindoutInstRowsIndex + 1 + index, 9).value = LT_CP_TestName(LT_SortedTestInstIndex, index)
            End If
            
            If b_DisplayTestNum Then
                FlowT_Sheet.Cells(FlowT_FindoutInstRowsIndex + 1 + index, 10).value = LT_CP_TestNumber(LT_SortedTestInstIndex, index)
            End If
            
    ''        FlowT_Sheet.Cells(FlowT_FindoutInstRowsIndex + 1 + Index, 11).Value = LT_CP_LowLimit(LT_SortedTestInstIndex, Index)
    ''        FlowT_Sheet.Cells(FlowT_FindoutInstRowsIndex + 1 + Index, 12).Value = LT_CP_HighLimit(LT_SortedTestInstIndex, Index)
            FlowT_Sheet.Cells(FlowT_FindoutInstRowsIndex + 1 + index, 11).value = LT_CP_LL_Formula(LT_SortedTestInstIndex, index)
            FlowT_Sheet.Cells(FlowT_FindoutInstRowsIndex + 1 + index, 12).value = LT_CP_HL_Formula(LT_SortedTestInstIndex, index)
        Next index
    End If

    If UCase(LT_Sheet.Cells(1, 4).value) = UCase("CP1_Func") And UCase(LT_Sheet.Cells(1, 6).value) = UCase("FT1_Func") Then
        For index = 0 To LT_FT_PinsCounterForEachTestINst(LT_SortedTestInstIndex) - 1
    '''        FlowT_Sheet.Cells(FlowT_FindoutInstRowsIndex + 1 + LT_CP_PinsCounterForEachTestINst(LT_SortedTestInstIndex) + Index, 4).Value = "FT1"
    '''        FlowT_Sheet.Cells(FlowT_FindoutInstRowsIndex + 1 + LT_CP_PinsCounterForEachTestINst(LT_SortedTestInstIndex) + Index, 7).Value = "Use-Limit"
    '''        FlowT_Sheet.Cells(FlowT_FindoutInstRowsIndex + 1 + LT_CP_PinsCounterForEachTestINst(LT_SortedTestInstIndex) + Index, 8).Value = LT_SortedTestInstanceName
            
            If b_DisplayPinName Then
                FlowT_Sheet.Cells(FlowT_FindoutInstRowsIndex + 1 + LT_CP_PinsCounterForEachTestINst(LT_SortedTestInstIndex) + index, 9).value = LT_FT_TestName(LT_SortedTestInstIndex, index)
            End If
            
            If b_DisplayTestNum Then
                FlowT_Sheet.Cells(FlowT_FindoutInstRowsIndex + 1 + LT_CP_PinsCounterForEachTestINst(LT_SortedTestInstIndex) + index, 10).value = LT_FT_TestNumber(LT_SortedTestInstIndex, index)
            End If
            
    '''        FlowT_Sheet.Cells(FlowT_FindoutInstRowsIndex + 1 + LT_CP_PinsCounterForEachTestINst(LT_SortedTestInstIndex) + Index, 11).Value = LT_FT_LowLimit(LT_SortedTestInstIndex, Index)
    '''        FlowT_Sheet.Cells(FlowT_FindoutInstRowsIndex + 1 + LT_CP_PinsCounterForEachTestINst(LT_SortedTestInstIndex) + Index, 12).Value = LT_FT_HighLimit(LT_SortedTestInstIndex, Index)
            FlowT_Sheet.Cells(FlowT_FindoutInstRowsIndex + 1 + LT_CP_PinsCounterForEachTestINst(LT_SortedTestInstIndex) + index, 11).value = LT_FT_LL_Formula(LT_SortedTestInstIndex, index)
            FlowT_Sheet.Cells(FlowT_FindoutInstRowsIndex + 1 + LT_CP_PinsCounterForEachTestINst(LT_SortedTestInstIndex) + index, 12).value = LT_FT_HL_Formula(LT_SortedTestInstIndex, index)
        Next index
        
    ElseIf UCase(LT_Sheet.Cells(1, 4).value) = UCase("Disable") And UCase(LT_Sheet.Cells(1, 6).value) = UCase("FT1_Func") Then
        For index = 0 To LT_FT_PinsCounterForEachTestINst(LT_SortedTestInstIndex) - 1
            
            If b_DisplayPinName Then
                FlowT_Sheet.Cells(FlowT_FindoutInstRowsIndex + 1 + index, 9).value = LT_FT_TestName(LT_SortedTestInstIndex, index)
            End If
            
            If b_DisplayTestNum Then
                FlowT_Sheet.Cells(FlowT_FindoutInstRowsIndex + 1 + index, 10).value = LT_FT_TestNumber(LT_SortedTestInstIndex, index)
            End If
            
            FlowT_Sheet.Cells(FlowT_FindoutInstRowsIndex + 1 + index, 11).value = LT_FT_LL_Formula(LT_SortedTestInstIndex, index)
            FlowT_Sheet.Cells(FlowT_FindoutInstRowsIndex + 1 + index, 12).value = LT_FT_HL_Formula(LT_SortedTestInstIndex, index)
        Next index
    
    End If
    
ElseIf UCase(CheckWord) = "CP" Then

    FlowT_Sheet.range("D" & (FlowT_FindoutInstRowsIndex + 1) & ":D" & FlowT_FindoutInstRowsIndex + InsertRowsMax) = "CP1"
    FlowT_Sheet.range("G" & (FlowT_FindoutInstRowsIndex + 1) & ":G" & FlowT_FindoutInstRowsIndex + InsertRowsMax) = "Use-Limit"
    FlowT_Sheet.range("H" & (FlowT_FindoutInstRowsIndex + 1) & ":H" & FlowT_FindoutInstRowsIndex + InsertRowsMax) = LT_SortedTestInstanceName
    
  
    For index = 0 To LT_CP_PinsCounterForEachTestINst(LT_SortedTestInstIndex) - 1
'''        FlowT_Sheet.Cells(FlowT_FindoutInstRowsIndex + 1 + Index, 4).Value = "CP1"
'''        FlowT_Sheet.Cells(FlowT_FindoutInstRowsIndex + 1 + Index, 7).Value = "Use-Limit"
'''        FlowT_Sheet.Cells(FlowT_FindoutInstRowsIndex + 1 + Index, 8).Value = LT_SortedTestInstanceName
        
        If b_DisplayPinName Then
            FlowT_Sheet.Cells(FlowT_FindoutInstRowsIndex + 1 + index, 9).value = LT_CP_TestName(LT_SortedTestInstIndex, index)
        End If
        If b_DisplayTestNum Then
            FlowT_Sheet.Cells(FlowT_FindoutInstRowsIndex + 1 + index, 10).value = LT_CP_TestNumber(LT_SortedTestInstIndex, index)
        End If
        
''        FlowT_Sheet.Cells(FlowT_FindoutInstRowsIndex + 1 + Index, 11).Value = LT_CP_LowLimit(LT_SortedTestInstIndex, Index)
''        FlowT_Sheet.Cells(FlowT_FindoutInstRowsIndex + 1 + Index, 12).Value = LT_CP_HighLimit(LT_SortedTestInstIndex, Index)
        FlowT_Sheet.Cells(FlowT_FindoutInstRowsIndex + 1 + index, 11).value = LT_CP_LL_Formula(LT_SortedTestInstIndex, index)
        FlowT_Sheet.Cells(FlowT_FindoutInstRowsIndex + 1 + index, 12).value = LT_CP_HL_Formula(LT_SortedTestInstIndex, index)

    Next index

ElseIf UCase(CheckWord) = "FT" Then
    FlowT_Sheet.range("D" & (FlowT_FindoutInstRowsIndex + 1) & ":D" & FlowT_FindoutInstRowsIndex + InsertRowsMax) = "FT1"
    FlowT_Sheet.range("G" & (FlowT_FindoutInstRowsIndex + 1) & ":G" & FlowT_FindoutInstRowsIndex + InsertRowsMax) = "Use-Limit"
    FlowT_Sheet.range("H" & (FlowT_FindoutInstRowsIndex + 1) & ":H" & FlowT_FindoutInstRowsIndex + InsertRowsMax) = LT_SortedTestInstanceName
   
 
    For index = 0 To LT_CP_PinsCounterForEachTestINst(LT_SortedTestInstIndex) - 1
'''        FlowT_Sheet.Cells(FlowT_FindoutInstRowsIndex + 1 + Index, 4).Value = "FT1"
'''        FlowT_Sheet.Cells(FlowT_FindoutInstRowsIndex + 1 + Index, 7).Value = "Use-Limit"
'''        FlowT_Sheet.Cells(FlowT_FindoutInstRowsIndex + 1 + Index, 8).Value = LT_SortedTestInstanceName
        
        If b_DisplayPinName Then
            FlowT_Sheet.Cells(FlowT_FindoutInstRowsIndex + 1 + index, 9).value = LT_FT_TestName(LT_SortedTestInstIndex, index)
        End If
        
        If b_DisplayTestNum Then
            FlowT_Sheet.Cells(FlowT_FindoutInstRowsIndex + 1 + index, 10).value = LT_FT_TestNumber(LT_SortedTestInstIndex, index)
        End If
        
'''        FlowT_Sheet.Cells(FlowT_FindoutInstRowsIndex + 1 + Index, 11).Value = LT_FT_LowLimit(LT_SortedTestInstIndex, Index)
'''        FlowT_Sheet.Cells(FlowT_FindoutInstRowsIndex + 1 + Index, 12).Value = LT_FT_HighLimit(LT_SortedTestInstIndex, Index)
        FlowT_Sheet.Cells(FlowT_FindoutInstRowsIndex + 1 + index, 11).value = LT_FT_LL_Formula(LT_SortedTestInstIndex, index)
        FlowT_Sheet.Cells(FlowT_FindoutInstRowsIndex + 1 + index, 12).value = LT_FT_HL_Formula(LT_SortedTestInstIndex, index)
    Next index

End If

End Sub



Sub SortLimitTable(ByRef LT_TestInstName() As String, LT_Sheet As Worksheet, _
                             ByRef LT_CP_TestName() As String, ByRef LT_CP_TestNumber() As Long, ByRef LT_CP_LowLimit() As Double, ByRef LT_CP_HighLimit() As Double, ByRef LT_CP_PinsCounterForEachTestINst() As Integer, _
                             ByRef LT_CP_LL_Formula() As String, ByRef LT_CP_HL_Formula() As String, _
                             ByRef LT_FT_TestName() As String, ByRef LT_FT_TestNumber() As Long, ByRef LT_FT_LowLimit() As Double, ByRef LT_FT_HighLimit() As Double, ByRef LT_FT_PinsCounterForEachTestINst() As Integer, _
                             ByRef LT_FT_LL_Formula() As String, ByRef LT_FT_HL_Formula() As String)

Dim LT_TestInstMax As Integer
Dim LT_RowsMax As Integer
LT_TestInstMax = UBound(LT_TestInstName)
LT_RowsMax = LT_Sheet.Cells(Rows.Count, 1).End(xlUp).Row - LT_RowsShift

ReDim LT_CP_TestName(LT_TestInstMax, LT_RowsMax - 1) As String
ReDim LT_CP_TestNumber(LT_TestInstMax, LT_RowsMax - 1) As Long
ReDim LT_CP_LowLimit(LT_TestInstMax, LT_RowsMax - 1) As Double
ReDim LT_CP_HighLimit(LT_TestInstMax, LT_RowsMax - 1) As Double

ReDim LT_CP_LL_Formula(LT_TestInstMax, LT_RowsMax - 1) As String
ReDim LT_CP_HL_Formula(LT_TestInstMax, LT_RowsMax - 1) As String

ReDim LT_FT_TestName(LT_TestInstMax, LT_RowsMax - 1) As String
ReDim LT_FT_TestNumber(LT_TestInstMax, LT_RowsMax - 1) As Long
ReDim LT_FT_LowLimit(LT_TestInstMax, LT_RowsMax - 1) As Double
ReDim LT_FT_HighLimit(LT_TestInstMax, LT_RowsMax - 1) As Double

ReDim LT_FT_LL_Formula(LT_TestInstMax, LT_RowsMax - 1) As String
ReDim LT_FT_HL_Formula(LT_TestInstMax, LT_RowsMax - 1) As String

Dim LT_TestInstCounter As Integer
Dim IndexCounter As Integer

Dim LT_CP_RowsCounter As Integer
Dim LT_FT_RowsCounter As Integer

'''Dim LT_CP_PinsCounterForEachTestINst() As Integer
ReDim LT_CP_PinsCounterForEachTestINst(LT_TestInstMax) As Integer

'''Dim LT_FT_PinsCounterForEachTestINst() As Integer
ReDim LT_FT_PinsCounterForEachTestINst(LT_TestInstMax) As Integer

Dim LT_Address_CP_LL As String
Dim LT_Address_CP_HL As String
Dim LT_Address_FT_LL As String
Dim LT_Address_FT_HL As String

For LT_TestInstCounter = 0 To LT_TestInstMax

    LT_CP_RowsCounter = 0
    LT_FT_RowsCounter = 0
    
    For IndexCounter = 0 To LT_RowsMax - 1
    
        If LT_Sheet.Cells(IndexCounter + 1 + LT_RowsShift, 1).value = LT_TestInstName(LT_TestInstCounter) Then
            If UCase(LT_Sheet.Cells(IndexCounter + 1 + LT_RowsShift, 8).value) = "T" Or UCase(LT_Sheet.Cells(IndexCounter + 1 + LT_RowsShift, 8).value) = "" Then
                If LT_Sheet.Cells(IndexCounter + 1 + LT_RowsShift, 4).value <> "" And LT_Sheet.Cells(IndexCounter + 1 + LT_RowsShift, 5).value <> "" Then  ''CP
                
                    LT_CP_TestName(LT_TestInstCounter, LT_CP_RowsCounter) = LT_Sheet.Cells(IndexCounter + 1 + LT_RowsShift, 2).value
                    LT_CP_TestNumber(LT_TestInstCounter, LT_CP_RowsCounter) = LT_Sheet.Cells(IndexCounter + 1 + LT_RowsShift, 3).value
                    LT_CP_LowLimit(LT_TestInstCounter, LT_CP_RowsCounter) = LT_Sheet.Cells(IndexCounter + 1 + LT_RowsShift, 4).value
                    LT_CP_HighLimit(LT_TestInstCounter, LT_CP_RowsCounter) = LT_Sheet.Cells(IndexCounter + 1 + LT_RowsShift, 5).value
                    
                    LT_Address_CP_LL = LT_Sheet.Cells(IndexCounter + 1 + LT_RowsShift, 4).Address(RowAbsolute:=False, ColumnAbsolute:=False)
                    LT_Address_CP_HL = LT_Sheet.Cells(IndexCounter + 1 + LT_RowsShift, 5).Address(RowAbsolute:=False, ColumnAbsolute:=False)
                    
                    LT_CP_LL_Formula(LT_TestInstCounter, LT_CP_RowsCounter) = "=OFFSET(A_Limits_Table!" & LT_Address_CP_LL & ",0,_Limit_Offset)"
                    LT_CP_HL_Formula(LT_TestInstCounter, LT_CP_RowsCounter) = "=OFFSET(A_Limits_Table!" & LT_Address_CP_HL & ",0,_Limit_Offset)"
    
                    LT_CP_RowsCounter = LT_CP_RowsCounter + 1
                
                End If
                
                If LT_Sheet.Cells(IndexCounter + 1 + LT_RowsShift, 6).value <> "" And LT_Sheet.Cells(IndexCounter + 1 + LT_RowsShift, 7).value <> "" Then  ''FT
                    
                    LT_FT_TestName(LT_TestInstCounter, LT_FT_RowsCounter) = LT_Sheet.Cells(IndexCounter + 1 + LT_RowsShift, 2).value
                    LT_FT_TestNumber(LT_TestInstCounter, LT_FT_RowsCounter) = LT_Sheet.Cells(IndexCounter + 1 + LT_RowsShift, 3).value
                    LT_FT_LowLimit(LT_TestInstCounter, LT_FT_RowsCounter) = LT_Sheet.Cells(IndexCounter + 1 + LT_RowsShift, 6).value
                    LT_FT_HighLimit(LT_TestInstCounter, LT_FT_RowsCounter) = LT_Sheet.Cells(IndexCounter + 1 + LT_RowsShift, 7).value
    
                    LT_Address_FT_LL = LT_Sheet.Cells(IndexCounter + 1 + LT_RowsShift, 6).Address(RowAbsolute:=False, ColumnAbsolute:=False)
                    LT_Address_FT_HL = LT_Sheet.Cells(IndexCounter + 1 + LT_RowsShift, 7).Address(RowAbsolute:=False, ColumnAbsolute:=False)
                    
                    LT_FT_LL_Formula(LT_TestInstCounter, LT_FT_RowsCounter) = "=OFFSET(A_Limits_Table!" & LT_Address_FT_LL & ",0,_Limit_Offset)"
                    LT_FT_HL_Formula(LT_TestInstCounter, LT_FT_RowsCounter) = "=OFFSET(A_Limits_Table!" & LT_Address_FT_HL & ",0,_Limit_Offset)"
                    
                    LT_FT_RowsCounter = LT_FT_RowsCounter + 1
                
                End If
            End If
        End If
    
    Next IndexCounter
    
    LT_CP_PinsCounterForEachTestINst(LT_TestInstCounter) = LT_CP_RowsCounter
    LT_FT_PinsCounterForEachTestINst(LT_TestInstCounter) = LT_FT_RowsCounter
    
Next LT_TestInstCounter

End Sub

Sub CPFT_SortLimitTable(ByRef LT_TestInstName() As String, LT_Sheet As Worksheet, _
                             ByRef LT_CPFT_TestName() As String, ByRef LT_CPFT_TestNumber() As String, ByRef LT_CPFT_LowLimit() As String, ByRef LT_CPFT_HighLimit() As String, ByRef LT_CPFT_PinsCounterForEachTestINst() As Integer, _
                             ByRef LT_CPFT_LL_Formula() As String, ByRef LT_CPFT_HL_Formula() As String, ByRef LT_CPFT_JobGate() As String)
 

Dim LT_TestInstMax As Integer
Dim LT_RowsMax As Integer
LT_TestInstMax = UBound(LT_TestInstName)
LT_RowsMax = LT_Sheet.Cells(Rows.Count, 1).End(xlUp).Row - LT_RowsShift

ReDim LT_CP_TestName(LT_TestInstMax, LT_RowsMax - 1) As String
ReDim LT_CP_TestNumber(LT_TestInstMax, LT_RowsMax - 1) As String
ReDim LT_CP_LowLimit(LT_TestInstMax, LT_RowsMax - 1) As String
ReDim LT_CP_HighLimit(LT_TestInstMax, LT_RowsMax - 1) As String

ReDim LT_CP_LL_Formula(LT_TestInstMax, LT_RowsMax - 1) As String
ReDim LT_CP_HL_Formula(LT_TestInstMax, LT_RowsMax - 1) As String

ReDim LT_FT_TestName(LT_TestInstMax, LT_RowsMax - 1) As String
ReDim LT_FT_TestNumber(LT_TestInstMax, LT_RowsMax - 1) As String
ReDim LT_FT_LowLimit(LT_TestInstMax, LT_RowsMax - 1) As String
ReDim LT_FT_HighLimit(LT_TestInstMax, LT_RowsMax - 1) As String

ReDim LT_FT_LL_Formula(LT_TestInstMax, LT_RowsMax - 1) As String
ReDim LT_FT_HL_Formula(LT_TestInstMax, LT_RowsMax - 1) As String

Dim LT_TestInstCounter As Integer
Dim IndexCounter As Integer

Dim LT_CP_RowsCounter As Integer
Dim LT_FT_RowsCounter As Integer

'''Dim LT_CP_PinsCounterForEachTestINst() As Integer
ReDim LT_CP_PinsCounterForEachTestINst(LT_TestInstMax) As Integer

'''Dim LT_FT_PinsCounterForEachTestINst() As Integer
ReDim LT_FT_PinsCounterForEachTestINst(LT_TestInstMax) As Integer

Dim LT_Address_CP_LL As String
Dim LT_Address_CP_HL As String
Dim LT_Address_FT_LL As String
Dim LT_Address_FT_HL As String

For LT_TestInstCounter = 0 To LT_TestInstMax

    LT_CP_RowsCounter = 0
    LT_FT_RowsCounter = 0
    
    For IndexCounter = 0 To LT_RowsMax - 1
    
        If LT_Sheet.Cells(IndexCounter + 1 + LT_RowsShift, 1).value = LT_TestInstName(LT_TestInstCounter) Then
            If UCase(LT_Sheet.Cells(IndexCounter + 1 + LT_RowsShift, 8).value) = "T" Or UCase(LT_Sheet.Cells(IndexCounter + 1 + LT_RowsShift, 8).value) = "" Then
''                If LT_Sheet.Cells(IndexCounter + 1 + LT_RowsShift, 4).Value <> "" And LT_Sheet.Cells(IndexCounter + 1 + LT_RowsShift, 5).Value <> "" Then  ''CP
                
                    LT_CP_TestName(LT_TestInstCounter, LT_CP_RowsCounter) = LT_Sheet.Cells(IndexCounter + 1 + LT_RowsShift, 2).value
                    LT_CP_TestNumber(LT_TestInstCounter, LT_CP_RowsCounter) = LT_Sheet.Cells(IndexCounter + 1 + LT_RowsShift, 3).value
                    LT_CP_LowLimit(LT_TestInstCounter, LT_CP_RowsCounter) = LT_Sheet.Cells(IndexCounter + 1 + LT_RowsShift, 4).value
                    LT_CP_HighLimit(LT_TestInstCounter, LT_CP_RowsCounter) = LT_Sheet.Cells(IndexCounter + 1 + LT_RowsShift, 5).value
                    
                    LT_Address_CP_LL = LT_Sheet.Cells(IndexCounter + 1 + LT_RowsShift, 4).Address(RowAbsolute:=False, ColumnAbsolute:=False)
                    LT_Address_CP_HL = LT_Sheet.Cells(IndexCounter + 1 + LT_RowsShift, 5).Address(RowAbsolute:=False, ColumnAbsolute:=False)
                    
                    LT_CP_LL_Formula(LT_TestInstCounter, LT_CP_RowsCounter) = "=OFFSET(A_Limits_Table!" & LT_Address_CP_LL & ",0,_Limit_Offset)"
                    LT_CP_HL_Formula(LT_TestInstCounter, LT_CP_RowsCounter) = "=OFFSET(A_Limits_Table!" & LT_Address_CP_HL & ",0,_Limit_Offset)"
    
                    LT_CP_RowsCounter = LT_CP_RowsCounter + 1
                
''                End If
                
''                If LT_Sheet.Cells(IndexCounter + 1 + LT_RowsShift, 6).Value <> "" And LT_Sheet.Cells(IndexCounter + 1 + LT_RowsShift, 7).Value <> "" Then  ''FT
                    
                    LT_FT_TestName(LT_TestInstCounter, LT_FT_RowsCounter) = LT_Sheet.Cells(IndexCounter + 1 + LT_RowsShift, 2).value
                    LT_FT_TestNumber(LT_TestInstCounter, LT_FT_RowsCounter) = LT_Sheet.Cells(IndexCounter + 1 + LT_RowsShift, 3).value
                    LT_FT_LowLimit(LT_TestInstCounter, LT_FT_RowsCounter) = LT_Sheet.Cells(IndexCounter + 1 + LT_RowsShift, 6).value
                    LT_FT_HighLimit(LT_TestInstCounter, LT_FT_RowsCounter) = LT_Sheet.Cells(IndexCounter + 1 + LT_RowsShift, 7).value
    
                    LT_Address_FT_LL = LT_Sheet.Cells(IndexCounter + 1 + LT_RowsShift, 6).Address(RowAbsolute:=False, ColumnAbsolute:=False)
                    LT_Address_FT_HL = LT_Sheet.Cells(IndexCounter + 1 + LT_RowsShift, 7).Address(RowAbsolute:=False, ColumnAbsolute:=False)
                    
                    LT_FT_LL_Formula(LT_TestInstCounter, LT_FT_RowsCounter) = "=OFFSET(A_Limits_Table!" & LT_Address_FT_LL & ",0,_Limit_Offset)"
                    LT_FT_HL_Formula(LT_TestInstCounter, LT_FT_RowsCounter) = "=OFFSET(A_Limits_Table!" & LT_Address_FT_HL & ",0,_Limit_Offset)"
                    
                    LT_FT_RowsCounter = LT_FT_RowsCounter + 1
                
''                End If
            End If
        End If
    
    Next IndexCounter
    
    LT_CP_PinsCounterForEachTestINst(LT_TestInstCounter) = LT_CP_RowsCounter
    LT_FT_PinsCounterForEachTestINst(LT_TestInstCounter) = LT_FT_RowsCounter
    
Next LT_TestInstCounter


ReDim LT_CPFT_TestName(LT_TestInstMax, LT_RowsMax - 1) As String
ReDim LT_CPFT_TestNumber(LT_TestInstMax, LT_RowsMax - 1) As String
ReDim LT_CPFT_LowLimit(LT_TestInstMax, LT_RowsMax - 1) As String
ReDim LT_CPFT_HighLimit(LT_TestInstMax, LT_RowsMax - 1) As String

ReDim LT_CPFT_LL_Formula(LT_TestInstMax, LT_RowsMax - 1) As String
ReDim LT_CPFT_HL_Formula(LT_TestInstMax, LT_RowsMax - 1) As String

ReDim LT_CPFT_JobGate(LT_TestInstMax, LT_RowsMax - 1) As String

ReDim LT_CPFT_PinsCounterForEachTestINst(LT_TestInstMax) As Integer

Dim LT_CPFT_ArrayIndex As Integer

For LT_TestInstCounter = 0 To LT_TestInstMax
    LT_CPFT_ArrayIndex = 0
    
    For IndexCounter = 0 To LT_CP_PinsCounterForEachTestINst(LT_TestInstCounter) - 1
  
        If LT_CP_LowLimit(LT_TestInstCounter, IndexCounter) = LT_FT_LowLimit(LT_TestInstCounter, IndexCounter) And LT_CP_HighLimit(LT_TestInstCounter, IndexCounter) = LT_FT_HighLimit(LT_TestInstCounter, IndexCounter) Then
            
            LT_CPFT_TestName(LT_TestInstCounter, LT_CPFT_ArrayIndex) = LT_CP_TestName(LT_TestInstCounter, IndexCounter)
            
            LT_CPFT_TestNumber(LT_TestInstCounter, LT_CPFT_ArrayIndex) = LT_CP_TestNumber(LT_TestInstCounter, IndexCounter)
            
            LT_CPFT_LowLimit(LT_TestInstCounter, LT_CPFT_ArrayIndex) = LT_CP_LowLimit(LT_TestInstCounter, IndexCounter)
            
            LT_CPFT_HighLimit(LT_TestInstCounter, LT_CPFT_ArrayIndex) = LT_CP_HighLimit(LT_TestInstCounter, IndexCounter)
            
            LT_CPFT_LL_Formula(LT_TestInstCounter, LT_CPFT_ArrayIndex) = LT_CP_LL_Formula(LT_TestInstCounter, IndexCounter)
            
            LT_CPFT_HL_Formula(LT_TestInstCounter, LT_CPFT_ArrayIndex) = LT_CP_HL_Formula(LT_TestInstCounter, IndexCounter)
           
            LT_CPFT_ArrayIndex = LT_CPFT_ArrayIndex + 1
        
        Else
            
            If LT_CP_LowLimit(LT_TestInstCounter, IndexCounter) <> "" And LT_CP_HighLimit(LT_TestInstCounter, IndexCounter) <> "" And _
               LT_FT_LowLimit(LT_TestInstCounter, IndexCounter) <> "" And LT_FT_HighLimit(LT_TestInstCounter, IndexCounter) <> "" Then

                 LT_CPFT_TestName(LT_TestInstCounter, LT_CPFT_ArrayIndex) = LT_CP_TestName(LT_TestInstCounter, IndexCounter)
                 LT_CPFT_TestName(LT_TestInstCounter, LT_CPFT_ArrayIndex + 1) = LT_FT_TestName(LT_TestInstCounter, IndexCounter)

                 LT_CPFT_TestNumber(LT_TestInstCounter, LT_CPFT_ArrayIndex) = LT_CP_TestNumber(LT_TestInstCounter, IndexCounter)
                 LT_CPFT_TestNumber(LT_TestInstCounter, LT_CPFT_ArrayIndex + 1) = LT_FT_TestNumber(LT_TestInstCounter, IndexCounter)

                 LT_CPFT_LowLimit(LT_TestInstCounter, LT_CPFT_ArrayIndex) = LT_CP_LowLimit(LT_TestInstCounter, IndexCounter)
                 LT_CPFT_LowLimit(LT_TestInstCounter, LT_CPFT_ArrayIndex + 1) = LT_FT_LowLimit(LT_TestInstCounter, IndexCounter)

                 LT_CPFT_HighLimit(LT_TestInstCounter, LT_CPFT_ArrayIndex) = LT_CP_HighLimit(LT_TestInstCounter, IndexCounter)
                 LT_CPFT_HighLimit(LT_TestInstCounter, LT_CPFT_ArrayIndex + 1) = LT_FT_HighLimit(LT_TestInstCounter, IndexCounter)

                 LT_CPFT_LL_Formula(LT_TestInstCounter, LT_CPFT_ArrayIndex) = LT_CP_LL_Formula(LT_TestInstCounter, IndexCounter)
                 LT_CPFT_LL_Formula(LT_TestInstCounter, LT_CPFT_ArrayIndex + 1) = LT_FT_LL_Formula(LT_TestInstCounter, IndexCounter)

                 LT_CPFT_HL_Formula(LT_TestInstCounter, LT_CPFT_ArrayIndex) = LT_CP_HL_Formula(LT_TestInstCounter, IndexCounter)
                 LT_CPFT_HL_Formula(LT_TestInstCounter, LT_CPFT_ArrayIndex + 1) = LT_FT_HL_Formula(LT_TestInstCounter, IndexCounter)

                 LT_CPFT_JobGate(LT_TestInstCounter, LT_CPFT_ArrayIndex) = "CP1"
                 LT_CPFT_JobGate(LT_TestInstCounter, LT_CPFT_ArrayIndex + 1) = "FT1"

                 LT_CPFT_ArrayIndex = LT_CPFT_ArrayIndex + 2

            ElseIf LT_CP_LowLimit(LT_TestInstCounter, IndexCounter) <> "" And LT_CP_HighLimit(LT_TestInstCounter, IndexCounter) <> "" And _
                     LT_FT_LowLimit(LT_TestInstCounter, IndexCounter) = "" And LT_FT_HighLimit(LT_TestInstCounter, IndexCounter) = "" Then

                LT_CPFT_TestName(LT_TestInstCounter, LT_CPFT_ArrayIndex) = LT_CP_TestName(LT_TestInstCounter, IndexCounter)
                LT_CPFT_TestNumber(LT_TestInstCounter, LT_CPFT_ArrayIndex) = LT_CP_TestNumber(LT_TestInstCounter, IndexCounter)
                LT_CPFT_LowLimit(LT_TestInstCounter, LT_CPFT_ArrayIndex) = LT_CP_LowLimit(LT_TestInstCounter, IndexCounter)
                LT_CPFT_HighLimit(LT_TestInstCounter, LT_CPFT_ArrayIndex) = LT_CP_HighLimit(LT_TestInstCounter, IndexCounter)
                LT_CPFT_LL_Formula(LT_TestInstCounter, LT_CPFT_ArrayIndex) = LT_CP_LL_Formula(LT_TestInstCounter, IndexCounter)
                LT_CPFT_HL_Formula(LT_TestInstCounter, LT_CPFT_ArrayIndex) = LT_CP_HL_Formula(LT_TestInstCounter, IndexCounter)
                LT_CPFT_JobGate(LT_TestInstCounter, LT_CPFT_ArrayIndex) = "CP1"
                
                LT_CPFT_ArrayIndex = LT_CPFT_ArrayIndex + 1

            ElseIf LT_CP_LowLimit(LT_TestInstCounter, IndexCounter) = "" And LT_CP_HighLimit(LT_TestInstCounter, IndexCounter) = "" And _
                     LT_FT_LowLimit(LT_TestInstCounter, IndexCounter) <> "" And LT_FT_HighLimit(LT_TestInstCounter, IndexCounter) <> "" Then

                LT_CPFT_TestName(LT_TestInstCounter, LT_CPFT_ArrayIndex) = LT_FT_TestName(LT_TestInstCounter, IndexCounter)
                LT_CPFT_TestNumber(LT_TestInstCounter, LT_CPFT_ArrayIndex) = LT_FT_TestNumber(LT_TestInstCounter, IndexCounter)
                LT_CPFT_LowLimit(LT_TestInstCounter, LT_CPFT_ArrayIndex) = LT_FT_LowLimit(LT_TestInstCounter, IndexCounter)
                LT_CPFT_HighLimit(LT_TestInstCounter, LT_CPFT_ArrayIndex) = LT_FT_HighLimit(LT_TestInstCounter, IndexCounter)
                LT_CPFT_LL_Formula(LT_TestInstCounter, LT_CPFT_ArrayIndex) = LT_FT_LL_Formula(LT_TestInstCounter, IndexCounter)
                LT_CPFT_HL_Formula(LT_TestInstCounter, LT_CPFT_ArrayIndex) = LT_FT_HL_Formula(LT_TestInstCounter, IndexCounter)
                LT_CPFT_JobGate(LT_TestInstCounter, LT_CPFT_ArrayIndex) = "FT1"
                
                LT_CPFT_ArrayIndex = LT_CPFT_ArrayIndex + 1

             End If
        End If
        
        LT_CPFT_PinsCounterForEachTestINst(LT_TestInstCounter) = LT_CPFT_ArrayIndex
        
    Next IndexCounter
    
    
Next LT_TestInstCounter

End Sub





Sub GetFlowInfo()

Dim sourcesheet As Worksheet
Dim list As Worksheet
Dim i As Integer
Dim j As Integer
Dim TempValue As String
Dim TempSheet As String
Dim maxcol As Integer
Dim MaxRow As Integer
Dim FlowSheetCountMax As Integer
Dim rowIndex As Integer
Dim TempInst As String
Dim TempBin As Long
Dim TempSoft As Long
Dim TempLo As Double
Dim TempHi As Double
Dim TempTestName As String
Dim TempTestNum As Long
Dim k As Integer
Dim ADD_L As String
Dim ADD_H As String
Dim FormulaStrL As String
Dim FlowTableName As String
Dim InstMaxRow As Integer

Set list = Sheets("InstanceList")

list.Cells.Clear
list.Cells(2, 1) = "Test Instance"
list.Cells(2, 2) = "Test Flow"

Call ExtractDuplicate

InstMaxRow = list.Cells(Rows.Count, 1).End(xlUp).Row


FlowSheetCountMax = ActiveWorkbook.Sheets.Count

rowIndex = 3

'List.Cells(2, 1).Value = "Flow"
'List.Cells(2, 2).Value = "TestInstance"
'List.Cells(2, 3).Value = "TestName"
'List.Cells(2, 4).Value = "TestNum"
'List.Cells(2, 5).Value = "Low Limit"
'List.Cells(2, 6).Value = "High Limit

'''For j = 3 To InstMaxRow
'''    List.Cells(j, 1).Value = ""
'''    List.Cells(j, 2).Value = ""
'''Next j


For j = rowIndex To InstMaxRow  ''' InstanceList Table

    TempInst = list.Cells(j, 1).value

    For k = 2 To FlowSheetCountMax
    
        Set sourcesheet = Sheets(k)
      
        If UCase(sourcesheet.name) Like "*FLOW*" And Not UCase(sourcesheet.name) Like "*CHAR*" Then
      
            MaxRow = sourcesheet.Cells(Rows.Count, 7).End(xlUp).Row
              
            FlowTableName = sourcesheet.name
              
            For i = 5 To MaxRow
              
                TempTestName = sourcesheet.Cells(i, 8).value
                              
                If UCase(TempInst) = UCase(TempTestName) Then
                  
                    If list.Cells(j, 2).value = "" Then
                        list.Cells(j, 2).value = FlowTableName
                    Else
                        list.Cells(j, 2).value = list.Cells(j, 2).value & "," & FlowTableName
                    End If
                    
                    i = MaxRow
                
                End If
              
            Next i
        
        End If
    
    Next k

Next j

''' Check Flow whether empty: If empty, use Test Instance combine _CP_NV, _CP_HV, _CP_LV, _FT_NV, _FT_HV, _FT_LV to seach each Flow Table
Dim CombineInstEnd() As Variant
CombineInstEnd = Array("_CP_NV", "_CP_HV", "_CP_LV", "_FT_NV", "_FT_HV", "_FT_LV")
Dim index As Integer
Dim b_CheckflowDuplicate As Boolean
Dim CellShiftCounter As Integer
Dim b_CheckCellShiftCounter As Boolean

For j = rowIndex To InstMaxRow  ''' InstanceList Table

    CellShiftCounter = 3
    
    If list.Cells(j, 2).value = "" Then
        
        For k = 2 To FlowSheetCountMax
        
            b_CheckflowDuplicate = False
            
            
            Set sourcesheet = Sheets(k)
            
            If UCase(sourcesheet.name) Like "*FLOW*" And Not UCase(sourcesheet.name) Like "*CHAR*" Then
            
                MaxRow = sourcesheet.Cells(Rows.Count, 7).End(xlUp).Row
                FlowTableName = sourcesheet.name
                
                b_CheckCellShiftCounter = False
                
                For index = 0 To UBound(CombineInstEnd)
    
                    TempInst = list.Cells(j, 1).value & CombineInstEnd(index)
                
                    For i = 5 To MaxRow
                      
                        TempTestName = sourcesheet.Cells(i, 8).value
                                      
                        If UCase(TempInst) = UCase(TempTestName) Then
                          
                            If list.Cells(j, 2).value = "" Then
                                If b_CheckflowDuplicate = False Then
                                    list.Cells(j, 2).value = FlowTableName
                                    b_CheckflowDuplicate = True
                                End If
                            Else
                                If b_CheckflowDuplicate = False Then
                                    list.Cells(j, 2).value = list.Cells(j, 2).value & "," & FlowTableName
                                    b_CheckflowDuplicate = True
                                End If
                            End If
                            
                            If list.Cells(j, CellShiftCounter).value = "" Then
                            
                                list.Cells(j, CellShiftCounter).value = TempTestName
                                b_CheckCellShiftCounter = True
                            Else
                            
                                list.Cells(j, CellShiftCounter).value = list.Cells(j, CellShiftCounter).value & "," & TempTestName
                                b_CheckCellShiftCounter = True
                            End If
                            
                            i = MaxRow
                        
                        End If
                      
                    Next i
                    
                Next index
                If b_CheckCellShiftCounter = True Then
                    CellShiftCounter = CellShiftCounter + 1
                End If
              End If



        Next k

    End If
Next j

'''Call CreateListBox

End Sub

Sub ExtractDuplicate() 'Remove repeat Instance Name

Dim start As Integer
Dim stoppo As Integer
Dim TempName As String
Dim AfterSortInstance As String
Dim i As Integer
Dim Count As Integer
Dim Repeat As Boolean
Dim j As Integer
Dim TargetSheet As Worksheet
Dim OutputSheet As Worksheet

Set TargetSheet = Sheets("A_Limits_Table")
Set OutputSheet = Sheets("InstanceList")

start = 3
stoppo = TargetSheet.Cells(Rows.Count, 1).End(xlUp).Row

For i = start To stoppo

    If i = start Then
     
        TempName = UCase(TargetSheet.Cells(i, 1).value)
        AfterSortInstance = TempName
        OutputSheet.Cells(i, 1).value = UCase(AfterSortInstance)
        Count = 1
     
    Else
        Repeat = False
        TempName = UCase(TargetSheet.Cells(i, 1).value)
         
        For j = 0 To Count
         
            If TempName = OutputSheet.Cells(start + j, 1).value Then
                Repeat = True
                j = Count
            End If

        Next j
        
        If Repeat <> True Then
            
            AfterSortInstance = TempName
            OutputSheet.Cells(start + Count, 1).value = UCase(AfterSortInstance)
            Count = Count + 1
            
        End If
     
    End If

Next i

End Sub
Sub CreateListBox()
    Dim LT_Sheet As Worksheet
    Set LT_Sheet = Worksheets("A_Limits_Table")
    
    
    LT_Sheet.Cells(1, 4).Clear
    LT_Sheet.Cells(1, 6).Clear
    
    With LT_Sheet.Cells(1, 4).Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .Color = 49407
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
    
    With LT_Sheet.Cells(1, 6).Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .Color = 49407
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
    
    
    With LT_Sheet.Cells(1, 4).Validation
        .delete
        .Add type:=xlValidateList, AlertStyle:=xlValidAlertStop, operator:= _
        xlBetween, Formula1:="CP1_Func, Disable"
        .IgnoreBlank = True
        .InCellDropdown = True
        .InputTitle = "CP update"
        .ErrorTitle = ""
        .InputMessage = ""
        .ErrorMessage = ""
        .IMEMode = xlIMEModeNoControl
        .ShowInput = True
        .ShowError = True
    End With
    
        With LT_Sheet.Cells(1, 6).Validation
        .delete
        .Add type:=xlValidateList, AlertStyle:=xlValidAlertStop, operator:= _
        xlBetween, Formula1:="FT1_Func, Disable"
        .IgnoreBlank = True
        .InCellDropdown = True
        .InputTitle = "FT update"
        .ErrorTitle = ""
        .InputMessage = ""
        .ErrorMessage = ""
        .IMEMode = xlIMEModeNoControl
        .ShowInput = True
        .ShowError = True
    End With
End Sub



