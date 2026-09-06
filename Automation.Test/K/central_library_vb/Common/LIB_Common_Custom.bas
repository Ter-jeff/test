Attribute VB_Name = "LIB_Common_Custom"
Option Explicit
'Revision History:
'V0.0 initial bring up

'variable declaration
Public Const Version_Lib_Common_Custom = "0.1"  'lib version

Function max(lng1 As Double, lng2 As Double) As Double
On Error GoTo errHandler 'Add ErrHandler 2023/08/18
    max = lng1
    If lng2 >= lng1 Then max = lng2
Exit Function 'Add ErrHandler 2023/08/18
errHandler: 'Add ErrHandler 2023/08/18
    Call Print_Error_Message(Error_Info, "LIB_Common_Custom", "max") 'Add ErrHandler 2023/08/18
    If AbortTest Then Exit Function Else Resume Next 'Add ErrHandler 2023/08/18
End Function


Function min(lng1 As Double, lng2 As Double) As Double
On Error GoTo errHandler 'Add ErrHandler 2023/08/18
    min = lng1
    If lng2 <= lng1 Then min = lng2
Exit Function 'Add ErrHandler 2023/08/18
errHandler: 'Add ErrHandler 2023/08/18
    Call Print_Error_Message(Error_Info, "LIB_Common_Custom", "Min") 'Add ErrHandler 2023/08/18
    If AbortTest Then Exit Function Else Resume Next 'Add ErrHandler 2023/08/18
End Function


Sub sbHideASheet(Sheet As String)
On Error GoTo errHandler 'Add ErrHandler 2023/08/18
    Sheets(Sheet).Visible = False
Exit Sub 'Add ErrHandler 2023/08/18
errHandler: 'Add ErrHandler 2023/08/18
    Call Print_Error_Message(Error_Info, "LIB_Common_Custom", "sbHideASheet") 'Add ErrHandler 2023/08/18
    If AbortTest Then Exit Sub Else Resume Next 'Add ErrHandler 2023/08/18
End Sub


Sub sbUnhideASheet(Sheet As String)
On Error GoTo errHandler 'Add ErrHandler 2023/08/18
        '$$ManifestSheet
    Sheets(Sheet).Visible = True
Exit Sub 'Add ErrHandler 2023/08/18
errHandler: 'Add ErrHandler 2023/08/18
    Call Print_Error_Message(Error_Info, "LIB_Common_Custom", "sbUnhideASheet") 'Add ErrHandler 2023/08/18
    If AbortTest Then Exit Sub Else Resume Next 'Add ErrHandler 2023/08/18
End Sub


Function RepeatChr(str As String, repeat As Long) As String
On Error GoTo errHandler 'Add ErrHandler 2023/08/18
    Dim i As Long
    Dim sa_TempStr(1) As String
    RepeatChr = vbNullString
    sa_TempStr(1) = str
    For i = 0 To repeat - 1
        sa_TempStr(0) = RepeatChr
        RepeatChr = Join(sa_TempStr, "")
    Next i
Exit Function 'Add ErrHandler 2023/08/18
errHandler: 'Add ErrHandler 2023/08/18
    Call Print_Error_Message(Error_Info, "LIB_Common_Custom", "RepeatChr") 'Add ErrHandler 2023/08/18
    If AbortTest Then Exit Function Else Resume Next 'Add ErrHandler 2023/08/18
End Function

Public Function FormatNumericDatalog(num As Variant, length As Long, LeftZero As Boolean) As String
On Error GoTo errHandler 'Add ErrHandler 2023/08/18

    Dim funcName As String:: funcName = "FormatNumericDatalog"
    
    ''''Example
    ''''----------------------------------------
    '''' length > 0  is to right shift
    '''' length < 0  is to left  shift
    ''''----------------------------------------
    ''''FormatNumeric(123456, 8) + "...end"
    ''''  123456...end
    ''''
    ''''FormatNumeric(123456,-8) + "...end"
    ''''123456  ...end
    ''''
    ''''----------------------------------------
    
    Dim numStr As String
    Dim tmpLen As Long
    Dim spcLen As Long
    
    numStr = CStr(num)
    tmpLen = Len(numStr)
    
    If (tmpLen > Abs(length)) Then
        spcLen = 0
    Else
        spcLen = Abs(length) - tmpLen
    End If
    
    If (length < 0) Then   ''''number shift to the very left
        FormatNumericDatalog = CStr(num) + Space(spcLen)
    ElseIf LeftZero Then ''''default: shift to the very right
        FormatNumericDatalog = Space(spcLen) + CStr(num)
    Else
        FormatNumericDatalog = CStr(num) + Space(spcLen)
    End If
    
Exit Function
errHandler: 'Add ErrHandler 2023/08/18
    Call Print_Error_Message(Error_Info, "LIB_Common_Custom", "FormatNumericDatalog") 'Add ErrHandler 2023/08/18
    If AbortTest Then Exit Function Else Resume Next 'Add ErrHandler 2023/08/18
End Function
