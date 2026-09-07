VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "eFuseCfgParser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Public sheetName As String
Public DoubleBits As Boolean
Public Fields As Dictionary   ' Database Return
Public FullSize As Long
Public pgmMode As ProgramMode
Public CfgValue As Dictionary
Public CfgCmpValue As Dictionary
Public CfgValueFutureStages  As Dictionary
Public CfgCmpValueLumpStages As Dictionary

Private bEarlyStage As Boolean
Private bLumpStage As Boolean
Private HeaderRow As Integer
Private HeaderStCol As Integer
Private HeaderEndCol As Integer
Private CfgCollection As Dictionary 'ex,  3 =>"A00", Column location => Config Name
Private FuseCol As Integer, DescCol As Integer, DefaultCol As Integer, BlowLocCol As Integer
Private ws As Worksheet
Public myUtility As Utility
Public CateNameMaxLen As Long
Public DicCfgList As Dictionary
Public DicFuseStages As Dictionary
Public BlankMaskDspWave As DSPWave
Public BlankMaskDspWaveFutureStages As DSPWave
Private blankMaskData() As Double
Private blankMaskDataFutureStages() As Double
Const MaxColRang = 99
Private ParsingArrayData() As Variant

Public Sub Initialize()
On Error GoTo errHandler
        HeaderRow = -1
        Set myUtility = New Utility
        FullSize = 0
        CateNameMaxLen = 0
        myUtility.IniDictionary CfgCollection
        myUtility.IniDictionary Fields
        myUtility.IniDictionary DicFuseStages
        myUtility.IniDictionary CfgValue
        myUtility.IniDictionary CfgCmpValue
        myUtility.IniDictionary CfgValueFutureStages
        myUtility.IniDictionary CfgCmpValueLumpStages
        myUtility.IniDictionary DicCfgList
        bEarlyStage = False
Exit Sub
errHandler:
    Call Print_Error_Message(Error_Info, "eFuseCfgParser", "Initialize")
    
        If AbortTest Then Exit Sub Else Resume Next

End Sub

Public Sub eFuseCfgInitialize(opbank As eFuseBdfBank, currStage As String, Optional selCfg As String = vbNullString)
On Error GoTo errHandler
Dim myfield As eFuseCfgField, fieldStr As Variant
Dim field As eFuseBdfField, bdfFieldStr As Variant
Dim tmp As String: tmp = vbNullString

    myUtility.Initialize (currStage)
    If (myUtility.testedStages.Count > 0 And myUtility.testedStages.Exists(currStage)) Then bEarlyStage = False
    For Each fieldStr In Fields.Keys
        Set myfield = Fields(fieldStr)
        myfield.FieldActive = "0x0"
        If Not myfield.CfgFields.Exists(selCfg) And Not bAllowNoCfgSel Then
           myUtility.MessageBox "Config didn't define! " & selCfg:
           Exit Sub
        End If
            'If fieldStr = "crc_disable_7" Then Stop
            '---------------------------------------
                    '- True case
                    '   Haven't choosed any CFG_condition.
                    '   Set default CFG_condition(CFG_A00).
                    
                    '- False case
                    '   One CFG_condition is enable.
                    '---------------------------------------
                    'If myfield.name = "CFG_condition_77" Then Stop
                    If selCfg = "" Then

                                  myfield.DefaultOrReal = IIf(myfield.Default_IsReadOnly, dr_default, dr_real)
                                  myfield.ReadOnly = myfield.Default_IsReadOnly
                    Else
                    
                                  myfield.DefaultOrReal = IIf(myfield.CfgFields_IsReadOnly(selCfg), dr_default, dr_real)
                                  myfield.ReadOnly = myfield.CfgFields_IsReadOnly(selCfg)
                    End If
                    
                    
                    If myUtility.testedStages.Exists(BdfDataBase.GetRealStage(myfield.BlowLocation)) Then
                        myfield.Trimmed = True
                        If selCfg = "" Then
                            myfield.FieldDefault = myfield.Default
                        Else
                            myfield.FieldDefault = myfield.CfgFields(selCfg)
                        End If
                    End If
                    'Port value to BDFField if the fused stage is the same with current job.
                    If myUtility.currStage = myfield.BlowLocation Or bLumpStages Then
                        If selCfg = "" Then
                                myfield.FieldActive = myfield.Default
                                If myUtility.currStage = myfield.BlowLocation Then
                                    myfield.Actived = True
                                End If
                        Else
                                myfield.FieldDefault = myfield.CfgFields(selCfg)
                                If (myUtility.currStage = myfield.BlowLocation And Not bEarlyStage) _
                                   Or (myfield.isEarly And bEarlyStage) Then
                                    myfield.FieldActive = myfield.CfgFields(selCfg)
                                    myfield.Actived = True
                                End If
                        End If
                        
                    End If
                    ''202004xx for ap
                    Set field = opbank.Fields(fieldStr)
 '20200519
'                    field.BlowLocation = myfield.BlowLocation
'
'                    field.IsEarlyFuse = myfield.IsEarly
                    If BdfDataBase.CFG_Only Then
                            'Set field = opbank.Fields(fieldStr)
                            field.ReadOnly = myfield.ReadOnly
                    End If
              Next
         Call CfgValue.RemoveAll
         Call CfgCmpValue.RemoveAll
         Call CfgValueFutureStages.RemoveAll
         Call CfgCmpValueLumpStages.RemoveAll
   
   For Each bdfFieldStr In opbank.DicCondTables.Keys
         Set field = opbank.Fields(bdfFieldStr)
          ''20200519
          'If field.name = "CFG_condition_77" Then Stop
          If field.IsEarlyFuse = False Then
            CfgValue.Add bdfFieldStr, GetCfgValue(field.name, field.LSB, field.msb, FuseWrite, , True)
            CfgCmpValue.Add bdfFieldStr, GetCfgValue(field.name, field.LSB, field.msb, ReadOutCheck, , True)
            tmp = GetCfgValue(field.name, field.LSB, field.msb, ReadOutCheck, , True)
         Else
            CfgValue.Add bdfFieldStr, GetCfgValue(field.name, field.LSB, field.msb, FuseWrite)
            CfgCmpValue.Add bdfFieldStr, GetCfgValue(field.name, field.LSB, field.msb, ReadOutCheck)
            tmp = GetCfgValue(field.name, field.LSB, field.msb, ReadOutCheck)
         End If
'         CfgValue.Add bdfFieldStr, GetCfgValue(field.LSB, field.msb, FuseWrite)
'         CfgCmpValue.Add bdfFieldStr, GetCfgValue(field.LSB, field.msb, ReadOutCheck)
         CfgValueFutureStages.Add bdfFieldStr, GetCfgValue(field.name, field.LSB, field.msb, FuseWrite, bLumpStage, True)
         CfgCmpValueLumpStages.Add bdfFieldStr, GetCfgValue(field.name, field.LSB, field.msb, ReadOutCheck, bLumpStage, True)

         ''202004xx for ap
         If field.size < 32 Then
            field.Llimit = myUtility.String2Dbl(tmp)
            field.Hlimit = field.Llimit
        Else
            field.xLlimit = tmp
            field.xHlimit = field.xLlimit
        End If

    Next
         Call SetBlankMaskData
Exit Sub
errHandler:
    Call Print_Error_Message(Error_Info, "eFuseCfgParser", "eFuseCfgInitialize")
   
        If AbortTest Then Exit Sub Else Resume Next

End Sub

Public Sub SetCfgBlankMaskDspWave()
On Error GoTo errHandler
    Set BlankMaskDspWave = New DSPWave
    Set BlankMaskDspWaveFutureStages = New DSPWave
    BlankMaskDspWave.data = blankMaskData
    BlankMaskDspWaveFutureStages.data = blankMaskDataFutureStages
Exit Sub
errHandler:
    Call Print_Error_Message(Error_Info, "eFuseCfgParser", "SetCfgBlankMaskDspWave")
 
        If AbortTest Then Exit Sub Else Resume Next

End Sub

''20201007 add parameter "fieldStr"
Public Function GetCfgValue(fieldStr As Variant, stBit As Long, spBit As Long, Optional myproc As DataProc = FuseWrite, Optional data4LumpStage As Boolean = False, Optional consider_early As Boolean = False) As String
On Error GoTo errHandler
          Dim myfield As eFuseCfgField ', fieldStr As Variant
          Dim binstr As String, HexStr As String, value As String
                    
                    binstr = vbNullString
                                    HexStr = vbNullString
                                    value = vbNullString
                'For Each fieldStr In Fields.Keys   ''20201007
                        Set myfield = Fields(fieldStr)
                        If (myfield.LSB >= stBit And myfield.msb <= spBit) Or _
                           (myfield.LSB >= spBit And myfield.msb <= stBit) Then
                            Select Case myproc:
                            Case blankCheck
                                    value = IIf(myfield.Trimmed, myfield.FieldDefault, "0x0")
                            Case FuseWrite
                                    value = "0x0"
                                    If data4LumpStage Then
                                            If myUtility.IsFutureStage(myfield.BlowLocation) Then
                                                    value = myfield.FieldDefault
                                            ElseIf consider_early And bEarlyStage And LCase(myUtility.currStage) = LCase(myfield.BlowLocation) And Not myfield.isEarly Then
                                                    value = myfield.FieldDefault
                                            End If
                                    ElseIf consider_early And bEarlyStage Then
                                            If LCase(myUtility.currStage) = LCase(myfield.BlowLocation) And Not myfield.isEarly Then
                                                    value = myfield.FieldDefault
                                            End If
                                    Else
                                        If (myUtility.currStage = myfield.BlowLocation And Not bEarlyStage) _
                                   Or (myfield.isEarly And bEarlyStage) Then
                                           value = myfield.FieldActive
                                        End If
                                    End If
                            Case ReadOutCheck
                                     value = IIf(myfield.Trimmed Or data4LumpStage _
                                     Or (myfield.isEarly And bEarlyStage) _
                                     Or (consider_early And bEarlyStage And LCase(myUtility.currStage) = LCase(myfield.BlowLocation)) _
                                     , myfield.FieldDefault, myfield.FieldActive)
                            End Select
                            binstr = myUtility.xHex2BinStr(value, myfield.size) & binstr
                        End If
                'Next
                HexStr = myUtility.Bin2HexStr(binstr, (spBit - stBit + 1 - 1) \ 4 + 1)

          GetCfgValue = HexStr
Exit Function
errHandler:
    Call Print_Error_Message(Error_Info, "eFuseCfgParser", "GetCfgValue")
 
        If AbortTest Then Exit Function Else Resume Next

End Function

Public Function Parse(ByVal ShtName As String, lumpStage As Boolean) As Boolean
On Error GoTo errHandler
    sheetName = ShtName
    bLumpStage = lumpStage
    Set ws = Worksheets(sheetName)
    
    '' True Case-----------------------------------------------------------------------------------
    '' Test plan has more than one config table.(EFUSE_CONFIG_MAIN_A, EFUSE_CONFIG_MAIN_B, ......)
    '' AutoGen merges these config table to one.
    ''---------------------------------------------------------------------------------------------
    If MergeFormatCheck Then
        Parse = WorkFlow_MergeFormat(sheetName)
    Else
        Parse = WorkFlow_OriginalFormat
    End If
Exit Function
errHandler:
    Call Print_Error_Message(Error_Info, "eFuseCfgParser", "Parse")
  
        If AbortTest Then Exit Function Else Resume Next

End Function

Private Sub SetBlankMaskData()
On Error GoTo errHandler
    Dim eFuseWave As New DSPWave, eFuseWaveFutureStages As New DSPWave, fieldStr As Variant, myfield As eFuseCfgField, iBit As Long
    Dim stBit As Long, spBit As Long, bitStep As Long, binA() As New SiteLong, bitDealt As Long, iSite As Long
    eFuseWave.CreateConstant 0, FullSize
    eFuseWaveFutureStages.CreateConstant 0, FullSize
    For Each fieldStr In Fields.Keys
        Set myfield = Fields(fieldStr)
        stBit = IIf(myfield.LSB < myfield.msb, myfield.LSB, myfield.msb)
        spBit = IIf(myfield.LSB < myfield.msb, myfield.msb, myfield.LSB)
        
        If myfield.Actived Or Not myUtility.testedStages.Exists(myfield.BlowLocation) Then
                        For iBit = stBit To spBit
                            eFuseWave.Element(iBit) = 1
                        Next
        End If
        If myUtility.IsFutureStage(myfield.BlowLocation) Then
                        For iBit = stBit To spBit
                            eFuseWaveFutureStages.Element(iBit) = 1
                        Next
        End If

    Next
    blankMaskData = eFuseWave.data
    blankMaskDataFutureStages = eFuseWaveFutureStages.data
Exit Sub
errHandler:
    Call Print_Error_Message(Error_Info, "eFuseCfgParser", "SetBlankMaskData")
  
        If AbortTest Then Exit Sub Else Resume Next

End Sub

Private Function MergeFormatCheck() As Boolean
On Error GoTo errHandler
    MergeFormatCheck = IIf(UCase(ws.Cells(1, 1).value) Like "EFUSE_CONFIG_MAIN*", True, False)
Exit Function
errHandler:
    Call Print_Error_Message(Error_Info, "eFuseCfgParser", "MergeFormatCheck")
 
        If AbortTest Then Exit Function Else Resume Next

End Function

Private Function WorkFlow_OriginalFormat() As Boolean
On Error GoTo errHandler
    WorkFlow_OriginalFormat = True
    If Not HeaderCheck Then
        WorkFlow_OriginalFormat = False
    Else
        Call ObtainFields
    End If
Exit Function
errHandler:
    Call Print_Error_Message(Error_Info, "eFuseCfgParser", "WorkFlow_OriginalFormat")

        If AbortTest Then Exit Function Else Resume Next

End Function

''Config table parse to "eFuseCfgField"
Private Function WorkFlow_MergeFormat(TableName As String) As Boolean
On Error GoTo errHandler
Dim endCol As Long, endRow As Long, iCol As Long, iRow As Long
Dim msb As Variant, LSB As Variant
Dim myCfgField As eFuseCfgField, cellVal As String, tmp As Variant
Dim fieldStr As Variant
Dim tmpHeaderDic As New Dictionary: tmpHeaderDic.compareMode = TextCompare
Dim tmpstr As Variant
Dim msg As String: msg = vbNullString

    myUtility.Get_Sheet_Info TableName, endRow, endCol, ParsingArrayData
    'header collect and check
    For iCol = 1 To endCol
        If ParsingArrayData(1, iCol) Like "EFUSE_CONFIG_MAIN_*" Then
            tmpstr = right(ParsingArrayData(1, iCol), 1)
            If Not tmpHeaderDic.Exists(tmpstr) Then tmpHeaderDic.Add tmpstr, True
        End If
    Next

    Call CollectCfgColumns(2, endCol)
    
    If tmpHeaderDic.Count = 0 Then
        msg = "can't find any sheet definition (EFUSE_CONFIG_MAIN_*) in first row, please check !!!"
        GoTo skipParser
    Else
        If CfgCollection.Count = 0 Then
            msg = "can't find any header (*00-*12) in second row, please check !!!"
            GoTo skipParser
        Else
            For Each tmpstr In CfgCollection.Items
                If Not tmpHeaderDic.Exists(left(tmpstr, 1)) Then
                    msg = "the header " + tmpstr + " sheet definition missing, please check !!!"
                    GoTo skipParser
                End If
            Next tmpstr
        End If
    End If

    ''Puts CFG table info into myCfgField.(cls object)
    ''----------------------------------------------------------------
    '' Col_1 | Col_2 | Col_3 | Col_4    | Col_5         | After Col_6
    '' Name  | MSB   | LSB   | BitWidth | Pgm Insertion | CFG Value
    ''----------------------------------------------------------------
    
    For iRow = 3 To endRow 'Fixed format
        Set myCfgField = New eFuseCfgField
        Set myCfgField.myUtility = myUtility
        myCfgField.DataRowNum = iRow
        
        'LSB/MSB/Size
        msb = ParsingArrayData(iRow, 2)
        LSB = ParsingArrayData(iRow, 3)
        If CStr(msb) = "" Or CStr(LSB) = "" Then
            msg = "MSB or LSB is empty, please check!!! (row:" + CStr(iRow) + ")"
            GoTo skipParser
        End If

        'Name
        cellVal = ParsingArrayData(iRow, 1)
        myCfgField.name = Replace(cellVal, " ", "_")

        If Len(Trim(cellVal)) <> 0 Then
            Call FieldLsbMsb(myCfgField, CLng(msb), CLng(LSB))
            If Len(cellVal) > CateNameMaxLen Then CateNameMaxLen = Len(cellVal)

            'Fuse BlowLocation
            cellVal = ParsingArrayData(iRow, 5)
            '20230606
            If EFUSE_REPLACE_SLT_TO_FT3 And UCase(Trim(cellVal)) = "SLT" Then
                cellVal = "FT3"
            ElseIf EFUSE_REPLACE_WLFT1_TO_FT1 And UCase(Trim(cellVal)) = "WLFT1" Then
                cellVal = "FT1"
            End If
            myCfgField.BlowLocation = Replace(cellVal, " ", "")
            If myCfgField.BlowLocation = "" Then
                msg = "BlowLocation is empty, please check!!! (row:" + CStr(iRow) + " col:5)"
                GoTo skipParser
            Else
                If Not DicFuseStages.Exists(myCfgField.BlowLocation) Then DicFuseStages.Add myCfgField.BlowLocation, True
            End If

            ''Enable early fusing flag
            If myUtility.IsStrMatch(myCfgField.BlowLocation, "_early") Then
                bEarlyStage = True
                myCfgField.BlowLocation = CurrentStage(myCfgField.BlowLocation)
                myCfgField.isEarly = True
            End If

            'A00-A12
            ''collect values of config table by row
            If CollectCfg(myCfgField, iRow) = False Then
                msg = "fuse value has format issue, please check!!! (row:" + CStr(iRow) + ")"
                GoTo skipParser
            End If
            ''Gets whole field name from CFG table.
            Fields.Add myCfgField.name, myCfgField
        Else
            msg = " category name is empty (row:" + CStr(iRow) + " col:1)"
            GoTo skipParser
        End If
    Next
    
    For Each fieldStr In Fields.Keys
        Set myCfgField = Fields(fieldStr)
        myCfgField.FieldDefault = myCfgField.CfgFields("A00")
    Next
    WorkFlow_MergeFormat = True

Exit Function
skipParser:
    Call Print_Error_Message(Error_Info, "eFuseCfgParser", "ObtainFields_MergeFormat", "The " + TableName + msg)
    TheExec.flow.TestLimit resultVal:=999, lowVal:=1, hiVal:=1, Tname:="Parser_Config_table"
    WorkFlow_MergeFormat = False
    Exit Function
errHandler:
    Call Print_Error_Message(Error_Info, "eFuseCfgParser", "ObtainFields_MergeFormat")
  
        If AbortTest Then Exit Function Else Resume Next

End Function

Private Sub FieldLsbMsb(myCfgField As eFuseCfgField, msb As Long, LSB As Long)
On Error GoTo errHandler
          If CLng(msb) >= CLng(LSB) Then
                myCfgField.msb = CLng(msb)
                myCfgField.LSB = CLng(LSB)
          Else
                myCfgField.msb = CLng(LSB)
                myCfgField.LSB = CLng(msb)
          End If
         '20230605,Will modify for last field 1 bit table bug
          If myCfgField.msb >= FullSize Then FullSize = myCfgField.msb + 1
          myCfgField.size = myCfgField.msb - myCfgField.LSB + 1
Exit Sub
errHandler:
    Call Print_Error_Message(Error_Info, "eFuseCfgParser", "FieldLsbMsb")
 
        If AbortTest Then Exit Sub Else Resume Next

End Sub

Private Sub ObtainFields()
On Error GoTo errHandler
Dim endRow As Long, iRow As Long
Dim myCfgField As eFuseCfgField, cellVal As String, tmp As Variant

    FuseCol = GetColInfoLoc("fuse", HeaderRow, 1, HeaderRow, 5)
    DescCol = GetColInfoLoc("Description", HeaderRow, 1, HeaderRow, MaxColRang)
    DefaultCol = GetColInfoLoc("1st*Silicon", HeaderRow, 1, HeaderRow, MaxColRang)
    BlowLocCol = GetColInfoLoc("Fuse*Blow*Location", HeaderRow, 1, HeaderRow, MaxColRang)

    endRow = ws.Cells(HeaderRow, HeaderStCol).End(xlDown).Row
    
    For iRow = HeaderRow + 1 To endRow
        Set myCfgField = New eFuseCfgField
        Set myCfgField.myUtility = myUtility
        myCfgField.DataRowNum = iRow
          
          'LSB/MSB/Size
          cellVal = ws.Cells(iRow, FuseCol).value
          tmp = Split(Replace(Replace(Replace(cellVal, "]", ""), "[", ""), " ", ""), ":")
          Call FieldLsbMsb(myCfgField, CLng(tmp(0)), CLng(tmp(1)))
          'Description
          cellVal = ws.Cells(iRow, DescCol).value
          myCfgField.name = Replace(cellVal, " ", "_")
          If Len(Trim(cellVal)) = 0 Then Exit For
          If Len(cellVal) > CateNameMaxLen Then CateNameMaxLen = Len(cellVal)
          'Analog Fuse Requirement1st Silicon
          cellVal = ws.Cells(iRow, DefaultCol).value
          myCfgField.Default = myUtility.String2Hex(cellVal, -1, myCfgField.hBytes)
          If myCfgField.Default = "-1" Then
                myCfgField.Default_IsReadOnly = False
                myCfgField.Default = "0x0"
          Else
                myCfgField.Default_IsReadOnly = True
          End If
          myCfgField.FieldDefault = myCfgField.Default
          myCfgField.DefaulString = cellVal
          'Fuse BlowLocation
          cellVal = ws.Cells(iRow, BlowLocCol).value
          myCfgField.BlowLocation = Replace(cellVal, " ", "")
          If myUtility.IsStrMatch(myCfgField.BlowLocation, "_early") Then
                bEarlyStage = True
                myCfgField.BlowLocation = CurrentStage(myCfgField.BlowLocation)
                myCfgField.isEarly = True
          End If
          If Not DicFuseStages.Exists(myCfgField.BlowLocation) Then DicFuseStages.Add myCfgField.BlowLocation, True
           'A00-A12
          Call CollectCfg(myCfgField, iRow)
          Fields.Add myCfgField.name, myCfgField
    Next
Exit Sub
errHandler:
    Call Print_Error_Message(Error_Info, "eFuseCfgParser", "ObtainFields")
  
        If AbortTest Then Exit Sub Else Resume Next

End Sub

Private Function CollectCfg(myCfgField As eFuseCfgField, iRow As Long) As Boolean
On Error GoTo errHandler
Dim item As Variant, fieldName As String, val As String, cellVal As String, chkVal As New SiteVariant
Dim EnCollectCfg As Boolean: EnCollectCfg = False
    If DicCfgList.Count = 0 Then EnCollectCfg = True
    For Each item In CfgCollection.Keys
        cellVal = ParsingArrayData(iRow, item)
        fieldName = CfgCollection(item)
        val = myUtility.String2Hex(cellVal, -1, myCfgField.hBytes)
        chkVal = val
        If myCfgField.IsOverflow(chkVal, True) = True Then
            CollectCfg = False
            Exit Function
        End If
        
        If val = "-1" Then
            Call Print_Error_Message(Error_Info, "eFuseCfgParser", "CollectCfg", "invalid format (" + cellVal + ") (row:" + CStr(iRow) + " col:" + CStr(item) + ")")
            CollectCfg = False
            Exit Function
        Else
            myCfgField.CfgFields_IsReadOnly.Add fieldName, True
        End If
        myCfgField.CfgFields.Add fieldName, val
        If (EnCollectCfg) Then DicCfgList.Add fieldName, "CFG_" + fieldName
    Next
    CollectCfg = True
Exit Function
errHandler:
    Call Print_Error_Message(Error_Info, "eFuseCfgParser", "CollectCfg")

        If AbortTest Then Exit Function Else Resume Next

End Function

Private Sub CollectCfgColumns(stRow As Long, Optional maxcol As Long = MaxColRang)
On Error GoTo errHandler
     Dim reParse As New RegExp, cellStr As String, iCol As Long
     ' Parse A00->A12 Columns
     reParse.IgnoreCase = True
     reParse.Pattern = "^Condition|^MSB|^LSB|^Bit|^Program|^Comment|^End|^Fuse|^Description|^Fuse"
     CfgCollection.RemoveAll
     For iCol = 1 To maxcol
     'For iCol = 1 To MaxColRang
     cellStr = Trim(ParsingArrayData(stRow, iCol))
'     cellStr = Trim(ws.Cells(stRow, iCol).value)
         If Not reParse.test(cellStr) And Len(cellStr) > 0 Then
            CfgCollection.Add iCol, cellStr
         End If
     Next
Exit Sub
errHandler:
    Call Print_Error_Message(Error_Info, "eFuseCfgParser", "CollectCfgColumns")
  
        If AbortTest Then Exit Sub Else Resume Next

End Sub

Private Function HeaderCheck() As Boolean
On Error GoTo errHandler
Dim stRow As Integer, stCol As Integer, bitData As String
Dim reParse As New RegExp, cellStr As String: cellStr = vbNullString

    'Get Single or Double Bits, Header Col Start/End
     bitData = "Access Mode (2-bit data)"
     stRow = GetRowInfoLoc(bitData)
     stCol = GetRowInfoLoc(bitData)
     If stRow <> -1 Then
        DoubleBits = True
     Else
        bitData = "Access Mode (1-bit data)"
        stRow = GetRowInfoLoc(bitData)
        stCol = GetRowInfoLoc(bitData)
        If stRow <> -1 Then
            DoubleBits = False
        Else
            HeaderCheck = False
            MsgBox "[Fatal Error] Wrong Config Table Format"
            Exit Function
        End If
     End If
     HeaderRow = stRow + 1: HeaderStCol = stCol
     cellStr = ws.Cells(stRow, stCol).value
     reParse.Pattern = "JTAG"
     pgmMode = IIf(reParse.test(cellStr), pgm_JTAG, pgm_DAA)
     
     Call CollectCfgColumns(stRow + 1)
     HeaderCheck = True

Exit Function
errHandler:
    Call Print_Error_Message(Error_Info, "eFuseCfgParser", "HeaderCheck")
    If AbortTest Then Exit Function Else Resume Next
End Function

Private Function GetRowInfoLoc(ByVal searchStr As String, Optional stRow As Integer = 1, Optional stCol As Integer = 1, Optional spRow As Integer = 10, Optional spCol As Integer = 10) As Integer ' in 10*10 range
On Error GoTo errHandler
Dim c As Variant

    ws.Select
    With ws.range(Cells(stRow, stCol), Cells(spRow, spCol))
        Set c = .Find(searchStr, LookIn:=xlValues)
        If c Is Nothing Then GetRowInfoLoc = -1: Exit Function
        GetRowInfoLoc = c.Row
    End With
Exit Function
errHandler:
    Call Print_Error_Message(Error_Info, "eFuseCfgParser", "GetRowInfoLoc")

        If AbortTest Then Exit Function Else Resume Next

End Function

Private Function GetColInfoLoc(ByVal searchStr As String, Optional stRow As Integer = 1, Optional stCol As Integer = 1, Optional spRow As Integer = 10, Optional spCol As Integer = 10) As Integer ' in 10*10 range
On Error GoTo errHandler
Dim c As Variant

    ws.Select
    With ws.range(Cells(stRow, stCol), Cells(spRow, spCol))
        Set c = .Find(searchStr, LookIn:=xlValues)
        If c Is Nothing Then GetColInfoLoc = -1: Exit Function
        GetColInfoLoc = c.Column
    End With

Exit Function
errHandler:
    Call Print_Error_Message(Error_Info, "eFuseCfgParser", "GetColInfoLoc")
    If AbortTest Then Exit Function Else Resume Next
End Function

Private Function CurrentStage(ByVal stagestr As String)
On Error GoTo errHandler

    CurrentStage = Replace(stagestr, "_Early", "", 1, -1, VbCompareMethod.vbTextCompare)

Exit Function
errHandler:
    Call Print_Error_Message(Error_Info, "eFuseCfgParser", "CurrentStage")
    If AbortTest Then Exit Function Else Resume Next
End Function

Public Function EarlyStage() As Boolean
On Error GoTo errHandler
    EarlyStage = bEarlyStage
Exit Function
errHandler:
    Call Print_Error_Message(Error_Info, "eFuseCfgParser", "EarlyStage")
 
        If AbortTest Then Exit Function Else Resume Next

End Function


