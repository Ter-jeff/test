VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "LVCC_VminBoundary_List"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private Const SHEET_DFC = "DFC_List"
Private Const ENABLE_DFC = "Enable_DFTLHFC_Debug"
Private Const ENABLE_FAILLOG = "Enable_Faillog_Debug"
Private instanceName_ As String
Private itemEnabledList_ As Dictionary

Sub Class_Initialize()
On Error GoTo errHandler
    instanceName_ = Trim(UCase(TheExec.DataManager.instancename))
    Set itemEnabledList_ = initialItemList
Exit Sub
errHandler:
    Call Print_Error_Message(Error_Info, "LVCC_VminBoundary_List", "Class_Initialize")
    If AbortTest Then Exit Sub Else Resume Next
End Sub

''' ===== write only for below ===== '''
''' =============================== '''

''' ===== read only for below ===== '''
Public Property Get IsEnableDFC() As Boolean
    IsEnableDFC = TheExec.enableWord(ENABLE_DFC)
End Property
Public Property Get IsEnableFAILLOG() As Boolean
    IsEnableFAILLOG = TheExec.enableWord(ENABLE_FAILLOG)
End Property
Public Property Get IsItemEnabled() As Boolean
    IsItemEnabled = itemEnabledList_.Exists(instanceName_)
End Property
''' =============================== '''
Private Function IsWorksheetExist(name As String) As Boolean
    On Error Resume Next
    IsWorksheetExist = (ActiveWorkbook.Sheets(name).index > 0)
End Function
Private Function initialItemList() As Dictionary
On Error GoTo errHandler
    ''' initial enabled items list '''
    Dim dic As Dictionary
    Dim sht As Worksheet
    Dim Row As Integer
    Dim val As String
    
    Set dic = New Dictionary
    If Not IsWorksheetExist(SHEET_DFC) Then GoTo endFunc
    
    Set sht = Sheets(SHEET_DFC)
    For Row = 1 To sht.UsedRange.Rows.Count
        val = Trim(UCase(sht.Cells(Row, 1).value))
        If Not dic.Exists(val) Then dic.Add val, True
    Next Row
endFunc:
    Set initialItemList = dic
    Set dic = Nothing
    Set sht = Nothing
Exit Function
errHandler:
    Call Print_Error_Message(Error_Info, "LVCC_VminBoundary_List", "initialItemList")
    If AbortTest Then Exit Function Else Resume Next
End Function
