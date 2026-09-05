VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "eFuseCfgField"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
' From the table
Public msb As Long
Public LSB As Long
Private m_size As Long
Private m_hBytes As Long
Public Msb1st As Boolean

Public name As String
Public DefaulString As String
Public FieldDefault As String
Public FieldActive As String
Public Trimmed As Boolean
Public DefaultOrReal As DeaultReal

Public BlowLocation As String
Public ReadOnly As Boolean
Public Default As String
Public isEarly As Boolean
Public Default_IsReadOnly As Boolean
Public Actived As Boolean
Public CfgFields As Dictionary     ' List of values keyed by column header name indicating if field is read-only
Public CfgFields_IsReadOnly As Dictionary     ' List of values keyed by column header name indicating if field is read-only
' Non-table data
Public DataRowNum As Long
Public myUtility As Utility

Private Sub Class_Initialize()
On Error GoTo errHandler
    Set myUtility = New Utility
    myUtility.IniDictionary CfgFields
    myUtility.IniDictionary CfgFields_IsReadOnly
    Msb1st = True
    FieldDefault = "0x0"
    ReadOnly = True
    Actived = False
    isEarly = False
    DefaultOrReal = DeaultReal.dr_default
    Default = "0x0"

Exit Sub
errHandler:
    Call Print_Error_Message(Error_Info, "eFuseCfgField", "Class_Initialize")
    If AbortTest Then Exit Sub Else Resume Next
End Sub

Public Property Get size() As Long
    size = m_size
End Property

Public Property Let size(v As Long)
    m_size = v
    m_hBytes = (m_size - 1) \ 4 + 1
End Property

Public Property Get hBytes() As Long
    hBytes = m_hBytes
End Property

Public Function IsOverflow(v As SiteVariant, Optional Force As Boolean = False) As Boolean
On Error GoTo errHandler
Dim max1Loc As Long, binstr As String
Dim site As Variant

    binstr = vbNullString
    If myUtility.IsHex(v) Then
        'If Not force Then IsOverflow = False: Exit Function ' skip it to save test time
        For Each site In TheExec.sites.Selected
            binstr = myUtility.xHex2BinStr(v(site))
            max1Loc = InStr(1, binstr, "1")
            If max1Loc <> 0 Then
                max1Loc = Len(binstr) - InStr(1, binstr, "1") + 1
            End If
            If max1Loc > m_size Then
                Call Print_Error_Message(Error_Info, "eFuseCfgField", "IsOverflow", "eFuse CFG Field is Overflow [" + name + "] @ row " & DataRowNum)
                IsOverflow = True: Exit Function
            End If
        Next
        IsOverflow = False: Exit Function
    Else
        For Each site In TheExec.sites.Selected
              If v(site) > 2 ^ m_size Then
                Call Print_Error_Message(Error_Info, "eFuseCfgField", "IsOverflow", "eFuse CFG Field is Overflow [" + name + "] @ row " & DataRowNum)
                IsOverflow = True: Exit Function
              End If
        Next
        IsOverflow = False
    End If

Exit Function
errHandler:
    Call Print_Error_Message(Error_Info, "eFuseCfgField", "IsOverflow")
    If AbortTest Then Exit Function Else Resume Next
End Function
