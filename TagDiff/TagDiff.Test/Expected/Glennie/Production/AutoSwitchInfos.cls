VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "AutoSwitchInfos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Const SelSrmShtName As String = "SELSRM_Mapping_Table"
Private blocks_ As Scripting.Dictionary
Private patterns_ As Scripting.Dictionary
Private selSrmInfos_() As SelSrmInfos
Private selSrmSht_ As Worksheet

Sub Class_Initialize()
    Set blocks_ = New Scripting.Dictionary
    Set patterns_ = New Scripting.Dictionary
End Sub

Public Property Get patterns() As Scripting.Dictionary
    Set patterns = patterns_
End Property

Public Property Get Blocks() As Scripting.Dictionary
    Set Blocks = blocks_
End Property

Private Function IsExistSheet() As Boolean
    On Error Resume Next
    IsExistSheet = (ActiveWorkbook.Sheets(SelSrmShtName).index > 0)
    If IsExistSheet Then
        Set selSrmSht_ = Sheets(SelSrmShtName)
    Else
        TheExec.ErrorLogMessage ("Error! Can not find select sram mapping table: " + SelSrmShtName)
    End If
End Function

Private Function getMaxBlockCount() As Long
    Dim Row As Integer
    Dim Col As Integer
    Dim blk As Integer: blk = -1
    Dim Block As String
    Dim content As String
    
    For Col = 1 To selSrmSht_.Cells(1, Columns.Count).End(xlToLeft).Column
        content = UCase(Trim(selSrmSht_.Cells(1, Col).value))
        If content = "BLOCK" Then
            blk = Col
            Exit For
        End If
    Next Col
    
    If blk = -1 Then
        getMaxBlockCount = -1
        Exit Function
    End If
    
    For Row = 1 To selSrmSht_.Cells(Rows.Count, blk).End(xlUp).Row
        content = UCase(Trim(selSrmSht_.Cells(Row, blk).value))
        If content = "END" Then Exit Function
        If content <> "BLOCK" Then
'            If content = "" Or content <> Block Then
            If Not content = "" And content <> Block Then
                getMaxBlockCount = getMaxBlockCount + 1
                Block = content
            End If
        End If
    Next Row

End Function

Private Function getMaxBitsCount() As Long
 
    Dim Row As Integer
    Dim Col As Integer
    Dim bit As Integer:: bit = -1
    Dim head As String
    Dim Block As String
    Dim content As String
    Dim bitCount As String
    For Col = 1 To selSrmSht_.Cells(1, Columns.Count).End(xlToLeft).Column
        head = LCase(selSrmSht_.Cells(1, Col).value)
        If head = "bits" Then
            bit = Col
            Exit For
        End If
    Next Col
    If bit = -1 Then
        getMaxBitsCount = bit
        Exit Function
    End If
    For Row = 1 To selSrmSht_.Cells(Rows.Count, bit).End(xlUp).Row
        content = UCase(Trim(selSrmSht_.Cells(Row, 1).value))
        bitCount = Trim(selSrmSht_.Cells(Row, bit).value)
        If content = "END" Then Exit Function
        If IsNumeric(bitCount) Then
            If CLng(bitCount) > getMaxBitsCount Then
                getMaxBitsCount = CLng(bitCount)
            End If
        End If
    Next Row
End Function
Public Function Initialize() As Boolean
    
On Error GoTo errhandle

    Dim Row As Long
    Dim Col As Long
    Dim hIndex As Integer
    Dim bIndex As Integer
    Dim Block As String
    Dim header As String
    Dim content As String
    Dim Count As Long
    Dim PinCount As Long
    Dim maxCount As Long
    Dim maxBitsCount As Long
    Dim maxBlockCount As Long
    
    ''' initial sheet
    Initialize = IsExistSheet
    If Not Initialize Then Exit Function
    
    ''' initial
    Count = -1
    maxBitsCount = getMaxBitsCount
    maxBlockCount = getMaxBlockCount - 1
    If maxBitsCount = -1 Or maxBlockCount = -1 Then GoTo errhandle
    ReDim selSrmInfos_(maxBlockCount, maxBitsCount)
    
'    For Row = 1 To selSrmSht_.Cells(Rows.Count, 1).END(xlUp).Row
    For Row = 1 To selSrmSht_.UsedRange.Rows.Count
        For Col = 1 To selSrmSht_.Cells(Row, Columns.Count).End(xlToLeft).Column
'        For col = 1 To selSrmSht_.UsedRange.Columns.Count
            content = Trim(selSrmSht_.Cells(Row, Col).value)
            If UCase(content) = "END" Then Initialize = True: Exit Function
            
            If content = "" Then GoTo nextCol
            If UCase(content) = "BLOCK" And hIndex = 0 Then hIndex = Row: GoTo nextCol
            If UCase(content) = "BITS" And bIndex = 0 Then bIndex = Col: GoTo nextRow
    
            If hIndex <> 0 Then header = selSrmSht_.Cells(hIndex, Col).value ''' get header
            If bIndex <> 0 Then PinCount = CLng(selSrmSht_.Cells(Row, bIndex).value)
            Select Case UCase(header)
                Case "BLOCK":
                    If content <> Block Then
                        ''' store index for block
                        Count = Count + 1
                        Block = content
                        If Not blocks_.Exists(Block) And Not Block = "" Then blocks_.Add UCase(Block), Count
                    End If
                Case "PATTERN":
                    If Not patterns_.Exists(Block) And Block <> "" Then patterns_.Add Block, content
                Case "LOGIC PINS":
                    selSrmInfos_(Count, PinCount).SelSrmPin.logicPin = content
                Case "SRAM PINS":
                    selSrmInfos_(Count, PinCount).SelSrmPin.SramPin = content
                Case "SELSRM1":
                    selSrmInfos_(Count, PinCount).SelSrmCode.SelSrm1 = content
                Case "SELSRM0":
                    selSrmInfos_(Count, PinCount).SelSrmCode.SelSrm0 = content
            End Select
nextCol:
        Next Col
nextRow:
    Next Row
    Exit Function
errhandle:
    TheExec.ErrorLogMessage "Parsing table: " & selSrmSht_.name & " fail, Please check it!!"
    Initialize = False
End Function

Public Function getSelectSramBit(blockName As String, vrsPinsVol As PinListData, index As Integer) As String
    
    Dim logicPin As String
    Dim SramPin As String
    Dim site As Variant
    
    ''' get logic and sram pin
    logicPin = selSrmInfos_(blocks_(blockName), index).SelSrmPin.logicPin
    SramPin = selSrmInfos_(blocks_(blockName), index).SelSrmPin.SramPin

    ''' judge whether logic higher sram or not
    For Each site In TheExec.sites.Active
        If vrsPinsVol.pins(logicPin).value(site) > vrsPinsVol.pins(SramPin).value Then
            getSelectSramBit = selSrmInfos_(blocks_(blockName), index).SelSrmCode.SelSrm0
        Else
            getSelectSramBit = selSrmInfos_(blocks_(blockName), index).SelSrmCode.SelSrm1
        End If
    Next site

End Function

Public Function getVbumpSelectSramBit(blockName As String, vrsPinsVol As PinListData, index As Integer, site As Variant) As String
    
    Dim logicPin As String
    Dim SramPin As String
'    Dim site As Variant
    
    ''' get logic and sram pin
    logicPin = selSrmInfos_(blocks_(blockName), index).SelSrmPin.logicPin
    SramPin = selSrmInfos_(blocks_(blockName), index).SelSrmPin.SramPin

    ''' judge whether logic higher sram or not
'    For Each site In theexec.sites.Active
        If vrsPinsVol.pins(logicPin).value(site) > vrsPinsVol.pins(SramPin).value(site) Then
            getVbumpSelectSramBit = selSrmInfos_(blocks_(blockName), index).SelSrmCode.SelSrm0
        Else
            getVbumpSelectSramBit = selSrmInfos_(blocks_(blockName), index).SelSrmCode.SelSrm1
        End If
'    Next site

End Function

Public Function getMaxPinCount() As Long
    getMaxPinCount = UBound(selSrmInfos_, 2)
End Function
