VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Class_MTR_Coefficient_Parsing"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private MTR_COEFF_INFO_DICT As New Dictionary
Private BKM_Check As Boolean

Sub Class_Initialize()
    'Parsing "MTR_Coefficient_Table" to get all of parameter for post calculate function use
    'Use function as below:
    '-- Calc_MetrologyTMPS_Coefficients
    '-- Calc_MetrologyTMPS_Coefficients_N3
    '-- Calc_MetrologyBTS_Coefficient
    '-- Calc_MetrologyBTS_Coefficient_10b_N3
    '-- Calc_MetrologyBTS_Coefficient_10b_N3
    '-- Calc_MetrologyDCTS_C0_Coefficients  (New)
    
    On Error GoTo errHandler:
    If glb_MTR_COEFF_PARSING_Flag = True Then Exit Sub                          'Update parsing one times ------20230525

    Dim sheetName As String
    Dim ws_def As Worksheet
    Dim wb As Workbook
    Dim MaxRow As Long
    Dim maxcol As Long
    Dim isSheetFound As Boolean
    Dim iRow As Long
    Dim iCol As Long
    Dim i As Long
    Dim BKM_DataAry As Variant
    Dim vsite As Variant
    Dim BKM_CODE As String
    Dim MTR_TYPE As String
    
    Dim Parsing_Indx As Long: Parsing_Indx = 0
    
    '''*****************************************************************'''
    sheetName = "MTR_Coefficient_Table"
    Set wb = Application.ActiveWorkbook
    Call check_Sheet_Range(sheetName, wb, ws_def, MaxRow, maxcol, isSheetFound, False)           'Max row reference Column:G (Ratio_of_RAK)
    '''*****************************************************************'''
    If isSheetFound Then
        Application.ScreenUpdating = False
        
        BKM_DataAry = ws_def.range(Cells(1, 1), Cells(MaxRow, maxcol))

        MTR_COEFF_INFO_DICT.RemoveAll       'Reset Dictionary information
        For iCol = 1 To maxcol
            For iRow = 1 To MaxRow
                If BKM_DataAry(iRow, iCol) = "" Then
                    Exit For
                Else
                    If iRow = 1 Then
                        MTR_TYPE = BKM_DataAry(iRow, iCol)
                    ElseIf iRow = 2 Then
                        BKM_CODE = BKM_DataAry(iRow, iCol + 1)
                        Call MTR_COEFF_INFO_DICT.Add(MTR_TYPE & "_" & BKM_CODE, "Dummy")
                    Else
                        Call MTR_COEFF_INFO_DICT.Add(UCase(MTR_TYPE & "_" & BKM_CODE & "_" & UCase(BKM_DataAry(iRow, iCol))), BKM_DataAry(iRow, iCol + 1))      'Store BKM to Dictionary
                    End If
                End If
            Next iRow
        Next iCol
        Application.ScreenUpdating = True
        glb_MTR_COEFF_PARSING_Flag = True   'Update parsing one times -- 20230214
    End If
   
    Exit Sub
errHandler:
    TheExec.Datalog.WriteComment "<Error> " + "Class_MTR_Coefficient_Parsing" + ":: please check it out."
    If AbortTest Then Exit Sub Else Resume Next
End Sub
Property Get BKM(ByVal key As String)
    BKM_Check = MTR_COEFF_INFO_DICT.Exists(key)
End Property
Property Get CoefficientRead(ByVal key As String)
    If MTR_COEFF_INFO_DICT.Exists(key) Then
        CoefficientRead = MTR_COEFF_INFO_DICT(key)
    Else
        TheExec.Datalog.WriteComment "<Error> " + "Class_MTR_Coefficient_Parsing" + " No such key : " + key + ":: please check it out."
    End If
End Property
