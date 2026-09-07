Attribute VB_Name = "VBT_ZeFuseOthers"
Option Explicit
'20230605 add for gating enable word
Public Enum EnableWdGatingPhase
    unknown = 0
    Disable
    OnProgValidation
    OnInitEnableWord
    NotExistStp
End Enum

Public Type EnableWdGatingTable
    ParserTableDone As Boolean
    TableExist As Boolean
    enableWdDic As Dictionary  'key:checkPhase; value:(key:enable word, value: tp name keyword)
End Type

Public enableWdGatingTableInfo As EnableWdGatingTable
Public SkipStep1_Flag As Boolean 'For JudgeDRAMType_T

Public Function PseudoFuse_WriteToFile(FilePath As String)
On Error GoTo errHandler
Dim OutputFile As String
Dim m_FileTmpName As String
Dim m_FileType As String: m_FileType = "csv"
Dim m_FileName As String
Dim FolderEmpty_Falg As Boolean: FolderEmpty_Falg = True

    OutputFile = FilePath & "\" & BdfDataBase.PseudoFuseFileName & "." & m_FileType
    
    If Not (CheckFileExist(FilePath, BdfDataBase.PseudoFuseFileName, m_FileType, m_FileName, FolderEmpty_Falg:=FolderEmpty_Falg)) Then
        If (FolderEmpty_Falg = True) Then
            theexec.Flow.TestLimit resultVal:=1, lowVal:=0, hiVal:=0, Tname:="Pesudo WriteToFile is not Exist"
            Exit Function
        End If
        CreateFile (OutputFile)
    End If
    
    BdfDataBase.DumpFusedDataToFile (OutputFile)

Exit Function
errHandler:
    Call Print_Error_Message(Error_Info, "VBT_ZeFuseOthers", "PseudoFuse_WriteToFile")
    If AbortTest Then Exit Function Else Resume Next
End Function

Public Function PseudoFuse_ReadFromFile(FilePath As String)
On Error GoTo errHandler
Dim m_FileTmpName As String
Dim m_FileType As String: m_FileType = "csv"
Dim m_FileName As String
Dim m_file As String
Dim m_site As Variant
Dim bankstr As Variant
Dim bank As New eFuseBdfBank
Dim keyname As String
Dim DataExist As New SiteLong

    If (ParsePseudoFuseFile = False) Then
        If Not (CheckFileExist(FilePath, BdfDataBase.PseudoFuseFileName, m_FileType, m_FileName, m_file)) Then
            Call Print_Error_Message(Error_Info, "VBT_ZeFuseOthers", "PseudoFuse_ReadFromFile", "File:: " + BdfDataBase.PseudoFuseFileName + " doesn't exist, please check it!")
            'GoTo errHandler
            theexec.Flow.TestLimit resultVal:=0, lowVal:=1, hiVal:=1, Tname:="PesudoFileExist"
            Exit Function
        End If
        Open m_file For Input As #1
            Do Until EOF(1)
                ''Create Efuse Dictionary
                BdfDataBase.GetFileData BdfDataBase.DicPseudoFuseData
            Loop
        Close #1
        ParsePseudoFuseFile = True
    End If

    DataExist = 1   '1:Pass, 0:Fail
    Call GetPseudoFuseData(DataExist)

    theexec.Flow.TestLimit resultVal:=DataExist, lowVal:=1, hiVal:=1, Tname:="PesudoFileExist"

Exit Function
errHandler:
    Call Print_Error_Message(Error_Info, "VBT_ZeFuseOthers", "PseudoFuse_ReadFromFile")
    If AbortTest Then Exit Function Else Resume Next
End Function

Private Sub CreateFile(Output As String)
On Error GoTo errHandler
Dim bankstr As Variant
Dim fieldStr As Variant
Dim BankHeaderStr As String: BankHeaderStr = "(X_Y)"
Dim bank As New eFuseBdfBank
Dim field As New eFuseBdfField

    For Each bankstr In BdfDataBase.Banks.Keys
        If bankstr Like "*cmp*" Then GoTo skiploop
        Set bank = BdfDataBase.Banks(bankstr)
        BankHeaderStr = BankHeaderStr & "," & CStr(bankstr) & "_Start"
        For Each fieldStr In bank.Fields
            Set field = bank.Fields(fieldStr)
            BankHeaderStr = BankHeaderStr & "," & field.name
        Next fieldStr
skiploop:
    Next bankstr
    
    Open Output For Append As #41
        Print #41, BankHeaderStr & "," & "EFUSE_END" & "," & "Pgm_Name" & "," & "Is_OI_Running"
    Close #41
    
Exit Sub
errHandler:
    Call Print_Error_Message(Error_Info, "VBT_ZeFuseOthers", "CreateFile")
    If AbortTest Then Exit Sub Else Resume Next
End Sub

Private Sub GetPseudoFuseData(SiteDataExist As SiteLong)
On Error GoTo errHandler
Dim headerKey As String: headerKey = "(X_Y)"
Dim headerStr As String
Dim headerArr() As String
Dim keyname As String
Dim dataStr As String
Dim dataArr() As String
Dim offset As Long:
Dim bank As New eFuseBdfBank
Dim bankstr As Variant
Dim m_site As Variant

    headerStr = GlbUtility.GetDictionary(headerKey, BdfDataBase.DicPseudoFuseData)
    headerArr = Split(headerStr, ",")

    For Each m_site In theexec.sites
        offset = 0
        If UCase(theexec.CurrentJob) Like "*FT*" Then
            '20211210, Modify for FT pseudo fuse with real package
            If gB_Package_PsudoFuse = True Then
                keyname = HramLotId & "_" & HramWaferId & "_" & HramXCoord & "_" & HramYCoord 'HramXCoord(m_site) & "_ " & HramYCoord(m_site)
            Else
                keyname = XCoord(m_site) & "_" & YCoord(m_site)
            End If
        Else
            keyname = theexec.Datalog.Setup.WaferSetup.GetXCoord(m_site) & "_" & _
                      theexec.Datalog.Setup.WaferSetup.GetYCoord(m_site)
        End If
        
        dataStr = GlbUtility.GetDictionary(keyname, BdfDataBase.DicPseudoFuseData)
        If dataStr = "" Then
            SiteDataExist(m_site) = 0
            GoTo skipsite
        End If
        
        dataArr = Split(dataStr, ",")

        For Each bankstr In BdfDataBase.Banks.Keys
            theexec.Datalog.WriteComment "************************"
            theexec.Datalog.WriteComment "[ " & bankstr & " ]"
            theexec.Datalog.WriteComment "************************"
            If bankstr Like "*cmp*" Then GoTo skiploop
            Set bank = BdfDataBase.Banks(bankstr)
            offset = offset + 1
            bank.StoreDataToBank keyname, dataArr, headerArr, offset
skiploop:
        Next bankstr

skipsite:
    Next m_site

Exit Sub
errHandler:
    Call Print_Error_Message(Error_Info, "VBT_ZeFuseOthers", "GetPseudoFuseData")
    If AbortTest Then Exit Sub Else Resume Next
End Sub

'''Public Function Bank_CompareWave_APvsObj(bank As String, ConditionStr As String, earlyfuse As Boolean)
'''On Error GoTo errHandler
'''Dim funcName As String: funcName = "Bank_CompareWave_APvsObj"
'''
'''    Dim RWCond As Long
'''    Dim opbank As eFuseBdfBank
''''    Dim Obj_ReadWave As New DSPWave
''''    Dim Obj_WriteWave As New DSPWave
''''    Dim AP_ReadWave As New DSPWave
''''    Dim AP_WriteWave As New DSPWave
'''    Dim MaskWave As New DSPWave
'''    Dim Cmpwave As New DSPWave
'''    Dim result As New SiteLong
'''    Dim m_Fusetype As eFuseBlockType
'''    Dim m_InWave As New DSPWave
'''    Dim Rvenable As Boolean: Rvenable = False
'''
''''If bank = "UDRP" Or bank = "UDR_P" Then Stop
'''    Set opbank = GetBdfBank(bank)
'''
'''    If earlyfuse Then
'''        MaskWave = opbank.EarlyStageMask
'''    Else
'''        MaskWave = opbank.StageMask
'''    End If
'''
'''    If bank = "CFG" Then
'''        m_Fusetype = eFuse_CFG
'''        Rvenable = g_Rvenable
'''    ElseIf bank = "ECID" Then
'''        m_Fusetype = eFuse_ECID
'''    ElseIf bank = "UDR_P" Or bank = "UDRP" Then
'''        m_Fusetype = eFuse_UDRP
'''    ElseIf bank = "UDR_E" Or bank = "UDRE" Then
'''        m_Fusetype = eFuse_UDRE
'''    ElseIf bank = "MON" Then
'''        m_Fusetype = eFuse_MON
'''    End If
'''
'''
'''    If ConditionStr = "WRITE" Then
'''        RWCond = 2
'''        m_InWave = opbank.DsscWave_Eff
'''        'Call rundsp.eFuse_compWave(bank, RWCond, opbank.DsscWave_Eff, maskWave, Cmpwave, result)
'''    ElseIf ConditionStr = "READ" Then
'''        RWCond = 1
'''        If opbank.pgmMode = pgm_JTAG Then
'''            m_InWave = opbank.JtagCapturedSerial
'''            'Call rundsp.eFuse_compWave(m_Fusetype, RWCond, opbank.JtagCapturedSerial, maskWave, Cmpwave, result)
'''        ElseIf opbank.pgmMode = pgm_DAA Then
'''             m_InWave = opbank.DaaCapWaveSerial
'''            'Call rundsp.eFuse_compWave(m_Fusetype, RWCond, opbank.DaaCapWaveSerial, maskWave, Cmpwave, result)
'''        End If
'''    Else
'''
'''    End If
'''    Call rundsp.eFuse_compWave(m_Fusetype, RWCond, m_InWave, MaskWave, Cmpwave, result, Rvenable)
'''    TheExec.Flow.TestLimit resultVal:=result, lowVal:=0, hiVal:=0, Tname:="compare_" + ConditionStr
'''
'''Exit Function
'''errHandler:
'''    GlbUtility.WriteDlg "<Error> " + funcName + ":: please check it out."
'''    If AbortTest Then Exit Function Else Resume Next
'''End Function

Public Sub SwitchFlag()
On Error GoTo errHandler
Dim bank As New eFuseBdfBank
Dim bankstr As Variant

    If PseudoFuseEnable = True Then
        If MixPseudoFuseEnable = True Then
            MixPseudoFuseEnable = False
        Else
            MixPseudoFuseEnable = True
        End If
    End If

    If BdfDataBase.Banks.Exists(Empty) Then BdfDataBase.Banks.Remove (Empty)
    For Each bankstr In BdfDataBase.Banks.Keys
        Set bank = BdfDataBase.Banks(bankstr)
        ''20221004, Add for multi blank check
        Call bank.InitialMultiCheckVarVariable
    Next

Exit Sub
errHandler:
    Call Print_Error_Message(Error_Info, "VBT_ZeFuseOthers", "SwitchFlag")
    If AbortTest Then Exit Sub Else Resume Next
End Sub

Public Sub SwitchEfuseDecodeFlag()
On Error GoTo errHandler

    If gB_eFuse_Disable_DecodeDataPrint_Flag = True Or gB_eFuse_Disable_DSPwavePrint_Flag = True Then
        If ForceDecodeEnable = True Then
            ForceDecodeEnable = False
        Else
            ForceDecodeEnable = True
        End If
    End If

Exit Sub
errHandler:
    Call Print_Error_Message(Error_Info, "VBT_ZeFuseOthers", "SwitchEfuseDecodeFlag")
    If AbortTest Then Exit Sub Else Resume Next
End Sub

Public Function CheckValueToDec(m_defval As Variant, convertValue As Variant)
On Error GoTo errHandler

    If LCase(m_defval) Like "x*" Or LCase(m_defval) Like "0x*" Then
        convertValue = GlbUtility.Hex2Dec(CStr(m_defval))
    ElseIf LCase(m_defval) Like "b*" Then
        convertValue = GlbUtility.Bin2Dec(CStr(m_defval))
    Else
        convertValue = CStr(m_defval)
    End If

Exit Function
errHandler:
    Call Print_Error_Message(Error_Info, "VBT_ZeFuseOthers", "CheckValueToDec")
    If AbortTest Then Exit Function Else Resume Next
End Function

Public Function ParseBinChk(TableName As String) As Boolean
On Error GoTo errHandler
Dim SheetMaxCol As Long
Dim SheetMaxRow As Long
Dim ParsingArr() As Variant
Dim Col As Long, Row As Long, i As Long, j As Long, k As Long, x As Long, sb As Long, eb As Long
Dim startRow As Long: startRow = 1
Dim flagIdx As Variant
Dim startBinCheckCol As Long: startBinCheckCol = 999
Dim headerName As String: headerName = vbNullString
Dim headerDic As Dictionary: GlbUtility.IniDictionary headerDic
Dim Specs As String: Specs = vbNullString
Dim bankNameCol As Long, cateNameCol As Long, cateIdx As Long, endCod As Long
Dim bitwidth As Long, sameKeyCnt As Long, bytes As Long
Dim opbank As eFuseBdfBank, field As eFuseBdfField
Dim rangeParse As New RegExp
Dim regexCombine As New RegExp
Dim SpecArr() As String
Dim tmpArr() As String
Dim tmpArr2() As String
Dim tmpCombineArr() As String
Dim tmpStr As String: tmpStr = vbNullString
Dim value As Variant
Dim loLimit As String, hiLimit As String
Dim key As String
Dim msg As String: msg = vbNullString
Dim job As Variant
Dim item As Variant, condItem As Variant
Dim tmpCombineNames As String: tmpCombineNames = vbNullString
Dim tmpFieldName As String: tmpFieldName = vbNullString
Dim tempStr As String: tempStr = vbNullString
Dim dicConditionRecord As Dictionary
Dim listConditionRecord As Collection
Dim tmpLsb As Long, tmpMsb As Long, orderCnt As Long
Dim tmpCombineJob As String: tmpCombineJob = vbNullString
Dim tmpCateName As String: tmpCateName = vbNullString
Dim tmpRange As String: tmpRange = vbNullString
Dim dicCombineFields As Dictionary
Dim tmpSpecStr As String
Dim tmpSpecArr() As String
Dim allSpecArr() As String
Dim insertIdx As Integer
Dim stageColIdx As Integer

    GlbUtility.Get_Sheet_Info TableName, SheetMaxRow, SheetMaxCol, ParsingArr

    rangeParse.IgnoreCase = True
    rangeParse.Pattern = "\[(\d+|x[A-F0-9]+|0x[A-F0-9]+|b[0-1]+)-(x[A-F0-9]+|0x[A-F0-9]+|\d+|b[0-1]+)\]$"

    regexCombine.IgnoreCase = True
    regexCombine.Pattern = "(\[\d+\-\d+\])|&"  '"(\w+)?\[\d+-\d+\](\w+)?"
    
    loLimit = vbNullString: hiLimit = vbNullString
    '1.Collect header
    Set DictBinNameMapping = New Dictionary
    DictBinNameMapping.compareMode = TextCompare
    Set DictBinBankMapping = New Dictionary
    DictBinBankMapping.compareMode = TextCompare
    endCod = SheetMaxCol
    sameKeyCnt = 0
    stageColIdx = 0
    For Row = 1 To SheetMaxRow
        If Trim(ParsingArr(Row, 1)) <> "" And LCase(Trim(ParsingArr(Row, 1))) Like "bankname" Then
            For Col = 1 To SheetMaxCol
                headerName = Trim(ParsingArr(Row, Col))
                If LCase(headerName) Like "*comment*" Then
                    If Col = SheetMaxCol Then
                        endCod = Col - 1
                    End If
                    stageColIdx = Col
                ElseIf Not IsEmpty(headerName) Then
                    If Col > startBinCheckCol And Col <= endCod Then
                        If LCase(headerName) Like "bin*" Then
                            key = headerName
                        Else
                            key = "Bin" + headerName
                        End If
                        If DictBinNameMapping.Exists(key) Then
                            sameKeyCnt = sameKeyCnt + 1
                            key = key + "_" + CStr(sameKeyCnt)
                            Call Print_Error_Message(Warning_Info, "VBT_ZeFuseOthers", "ParseBinChk", "The header " + headerName + " exist duplicate in the " + TableName + " !!!")
                        End If
                        DictBinNameMapping.Add key, (Col - startBinCheckCol - 1)
                    End If
                    If Not headerDic.Exists(headerName) Then headerDic.Add headerName, Col
                    If LCase(headerName) Like "fuse\bin" Then
                        startBinCheckCol = Col
                    End If
                Else
                    Call Print_Error_Message(Warning_Info, "VBT_ZeFuseOthers", "ParseBinChk", "The " + TableName + " sheet exist empty cell in header row !!!")
                End If
            Next Col
        Else
            If headerDic.Count > 0 Then Exit For
        End If
    Next Row

    If startBinCheckCol <> 999 Then
        startBinCheckCol = startBinCheckCol + 1
        ReDim BinCheckData(endCod - startBinCheckCol)
    Else
        msg = "can not get header correctly (keyword: Fuse\Bin)!!!"
        GoTo skipParser
    End If

    startRow = startRow + 1
    bankNameCol = IIf(headerDic.Exists("bankName"), headerDic.item("bankname"), 0)
    cateNameCol = IIf(headerDic.Exists("Fuse\Bin"), headerDic.item("Fuse\Bin"), 0)
    If bankNameCol = 0 Or cateNameCol = 0 Then
        msg = "can not get header correctly (keyword: BankName or Fuse\Bin)!!!"
        GoTo skipParser
    End If

    '2.Collect the specs for inspection
    For Each flagIdx In DictBinNameMapping.Items
        Set dicConditionRecord = New Dictionary
        dicConditionRecord.compareMode = TextCompare
        
        Col = flagIdx + startBinCheckCol
        Erase BinCheckData(flagIdx).CateArr()
        For Row = startRow To SheetMaxRow
            If ParsingArr(Row, bankNameCol) <> "" Then
                cateIdx = Row - startRow
                ReDim Preserve BinCheckData(flagIdx).CateArr(cateIdx)
                tmpLsb = -1
                tmpMsb = -1
                With BinCheckData(flagIdx).CateArr(cateIdx)
                    .bankName = UCase(Replace(ParsingArr(Row, bankNameCol), "bank_", ""))
                    .cateName = UCase(Trim(ParsingArr(Row, cateNameCol)))
                    .cateNameOri = UCase(Trim(ParsingArr(Row, cateNameCol)))
                    .MergeBitCheck = False
                    .SplitBitCheck = False
                    .Bit_LSB = 0
                    .Bit_MSB = 0
                    .specSite = False
                    .CombineFieldCheck = False
                    .conditionCheck = False
                    .conditionResult = False
                    .skipCheck = IIf(stageColIdx > 0, CheckCurrentJob(Trim(ParsingArr(Row, stageColIdx))), False)
                End With
                
                Set opbank = GetBdfBank(BinCheckData(flagIdx).CateArr(cateIdx).bankName)
                If Not DictBinBankMapping.Exists(UCase(Replace(ParsingArr(Row, bankNameCol), "bank_", ""))) Then
                    DictBinBankMapping.Add UCase(Replace(ParsingArr(Row, bankNameCol), "bank_", "")), Empty
                End If
                
                If BinCheckData(flagIdx).CateArr(cateIdx).cateName Like "CFG_CONDITION_*..*" Then
                    BinCheckData(flagIdx).CateArr(cateIdx).MergeBitCheck = True
                    Erase tmpArr()
                    tmpArr = Split(Replace(Replace(BinCheckData(flagIdx).CateArr(cateIdx).cateName, "CFG_CONDITION_[", ""), "]", ""), "..")
                    BinCheckData(flagIdx).CateArr(cateIdx).Bit_MSB = IIf(opbank.Fields.Exists("CFG_condition_" & tmpArr(0)), opbank.Fields("CFG_condition_" & tmpArr(0)).msb, 0)
                    BinCheckData(flagIdx).CateArr(cateIdx).Bit_LSB = IIf(opbank.Fields.Exists("CFG_condition_" & tmpArr(1)), opbank.Fields("CFG_condition_" & tmpArr(1)).LSB, 0)
                    If BinCheckData(flagIdx).CateArr(cateIdx).Bit_MSB = 0 And BinCheckData(flagIdx).CateArr(cateIdx).Bit_LSB = 0 Then
                        msg = "CFG_Condition rang incorrect (row:" + CStr(Row) + " col:" + CStr(cateNameCol) + "), please check it!!!"
                        GoTo skipParser
                    End If
                ElseIf regexCombine.test(BinCheckData(flagIdx).CateArr(cateIdx).cateName) Then 'KKK
                    BinCheckData(flagIdx).CateArr(cateIdx).CombineFieldCheck = True
                    tmpCateName = BinCheckData(flagIdx).CateArr(cateIdx).cateName
                    tmpCombineJob = Empty
                    Erase tmpCombineArr()
                    Erase tmpArr()
                    tmpCombineArr = Split(tmpCateName, "&")
                    
                    Set BinCheckData(flagIdx).CateArr(cateIdx).listCombineFieldsGroup = New Collection
                    For k = 0 To UBound(tmpCombineArr)
                        If InStr(tmpCombineArr(k), "[") > 0 Then
                            tmpRange = mid(tmpCombineArr(k), InStr(tmpCombineArr(k), "["), InStr(tmpCombineArr(k), "]") - InStr(tmpCombineArr(k), "[") + 1)
                            tmpArr = Split(Replace(Replace(tmpRange, "[", ""), "]", ""), "-")
                            For i = tmpArr(0) To tmpArr(1)
                                If opbank.Fields.Exists(Replace(tmpCombineArr(k), tmpRange, i)) Then
                                    Set field = opbank.Fields(Replace(tmpCombineArr(k), tmpRange, i))
                                    If tmpCombineJob = Empty Then tmpCombineJob = field.BlowLocation
                                    If tmpCombineJob <> field.BlowLocation Then
                                        msg = "Combine fields: " + CStr(Replace(tmpCateName, tmpRange, tmpArr(i))) + " is in different stage  (row:" + CStr(Row) + " col:" + CStr(cateNameCol) + "), please check it!!!"
                                        GoTo skipParser
                                    End If
                                    Set dicCombineFields = New Dictionary
                                    dicCombineFields.compareMode = TextCompare
                                    dicCombineFields.Add "Name", field.name
                                    dicCombineFields.Add "LSB", field.LSB
                                    dicCombineFields.Add "MSB", field.msb
                                    BinCheckData(flagIdx).CateArr(cateIdx).listCombineFieldsGroup.Add dicCombineFields
                                    BinCheckData(flagIdx).CateArr(cateIdx).Width = (field.msb - field.LSB + 1) + BinCheckData(flagIdx).CateArr(cateIdx).Width
                                Else
                                    msg = Replace(tmpCombineArr(k), tmpRange, i) + " not exist in BDF  (row:" + CStr(Row) + " col:" + CStr(cateNameCol) + "), please check it!!!"
                                    GoTo skipParser
                                End If
                            Next i
                            bitwidth = BinCheckData(flagIdx).CateArr(cateIdx).Width
                        Else
                            If opbank.Fields.Exists(tmpCombineArr(k)) Then
                                Set field = opbank.Fields(tmpCombineArr(k))
                                If tmpCombineJob = Empty Then tmpCombineJob = field.BlowLocation
                                    If tmpCombineJob <> field.BlowLocation Then
                                        msg = "Combine fields: " + CStr(Replace(tmpCateName, tmpRange, tmpArr(i))) + " is in different stage  (row:" + CStr(Row) + " col:" + CStr(cateNameCol) + "), please check it!!!"
                                        GoTo skipParser
                                    End If
                                Set dicCombineFields = New Dictionary
                                dicCombineFields.compareMode = TextCompare
                                dicCombineFields.Add "Name", field.name
                                dicCombineFields.Add "LSB", field.LSB
                                dicCombineFields.Add "MSB", field.msb
                                BinCheckData(flagIdx).CateArr(cateIdx).listCombineFieldsGroup.Add dicCombineFields
                                BinCheckData(flagIdx).CateArr(cateIdx).Width = (field.msb - field.LSB + 1) + BinCheckData(flagIdx).CateArr(cateIdx).Width
                            End If
                            bitwidth = BinCheckData(flagIdx).CateArr(cateIdx).Width
                        End If
                    Next k
                    BinCheckData(flagIdx).CateArr(cateIdx).CombineFieldJob = tmpCombineJob
                ElseIf BinCheckData(flagIdx).CateArr(cateIdx).cateName Like "*[*:*]*" Then
                    BinCheckData(flagIdx).CateArr(cateIdx).SplitBitCheck = True
                    Erase tmpArr()
                    tmpArr = Split(BinCheckData(flagIdx).CateArr(cateIdx).cateName, "[")
                    'get real category name
                    BinCheckData(flagIdx).CateArr(cateIdx).cateName = Trim(tmpArr(0))
    
                    Erase tmpArr2()
                    tmpArr2 = Split(Replace(Replace(Replace(tmpArr(1), "[", ""), "]", ""), " ", ""), ":")
                    If UBound(tmpArr2) <> 1 Then
                        msg = "category bit range format incorrect (row:" + CStr(Row) + " col:" + CStr(cateNameCol) + "), please check it!!!"
                        GoTo skipParser
                    End If
                    sb = CLng(tmpArr2(0))
                    eb = CLng(tmpArr2(1))
                    
                    If sb > eb Then
                        BinCheckData(flagIdx).CateArr(cateIdx).Bit_LSB = eb
                        BinCheckData(flagIdx).CateArr(cateIdx).Bit_MSB = sb
                    Else
                        BinCheckData(flagIdx).CateArr(cateIdx).Bit_LSB = sb
                        BinCheckData(flagIdx).CateArr(cateIdx).Bit_MSB = eb
                    End If
                End If
                
                If opbank Is Nothing Then
                    msg = "has bank name not recognized issue (row:" + CStr(Row) + " col:" + CStr(bankNameCol) + "), please check it!!!"
                    GoTo skipParser
                End If
                
                If Not BinCheckData(flagIdx).CateArr(cateIdx).MergeBitCheck And Not BinCheckData(flagIdx).CateArr(cateIdx).CombineFieldCheck Then
                    If opbank.Fields.Exists(BinCheckData(flagIdx).CateArr(cateIdx).cateName) Then
                        Set field = opbank.Fields(BinCheckData(flagIdx).CateArr(cateIdx).cateName)
                    Else
                        msg = "has unknow category name(" + BinCheckData(flagIdx).CateArr(cateIdx).cateName + ") in bank " + opbank.name + " (row:" + CStr(Row) + " col:" + CStr(cateNameCol) + "), please check it!!!"
                        GoTo skipParser
                    End If
                End If
    
                Specs = CStr(Trim(ParsingArr(Row, Col)))
                If Specs = "" Then
                    msg = "not allow check spec is empty (row:" + Row + " col:" + Col + "), please check it!!!"
                    GoTo skipParser
                ElseIf LCase(Specs) Like "if(*" Then
                    BinCheckData(flagIdx).CateArr(cateIdx).conditionCheck = True
                    BinCheckData(flagIdx).CateArr(cateIdx).conditionInfo = mid(Specs, InStr(Specs, "(") + 1, InStr(Specs, ":") - (InStr(Specs, "(")) - 1)
                    
                    insertIdx = 0
                    tmpSpecStr = Empty
                    'True part
                    tmpSpecStr = mid(Specs, InStr(Specs, ":") + 1, InStr(Specs, ",") - InStr(Specs, ":") - 1)
                    SpecArr = Split(Replace(Replace(tmpSpecStr, "|", ",|,"), "&", ",&,"), ",")
                    
                    ReDim allSpecArr(UBound(SpecArr))
                    
                    For i = 0 To UBound(SpecArr)
                        allSpecArr(i) = SpecArr(i) + "$t"
                        insertIdx = insertIdx + 1
                    Next i
                    
                    'False part
                    tmpSpecStr = mid(Specs, InStr(Specs, ",") + 1, InStr(Specs, ")") - InStr(Specs, ",") - 1)
                    SpecArr = Split(Replace(Replace(tmpSpecStr, "|", ",|,"), "&", ",&,"), ",")
                    
                    ReDim Preserve allSpecArr(UBound(allSpecArr) + UBound(SpecArr) + 1)
                    
                    For i = 0 To UBound(SpecArr)
                        allSpecArr(insertIdx + i) = SpecArr(i) + "$f"
                    Next
                    SpecArr = allSpecArr
                ElseIf LCase(Specs) Like "*site*" Then
                    BinCheckData(flagIdx).CateArr(cateIdx).specSite = True
                    Specs = Replace(Replace(Specs, " ", ""), "),", ");")
                    SpecArr = Split(Replace(Replace(Specs, "|", ";|;"), "&", ";&;"), ";")
                Else
                    SpecArr = Split(Replace(Replace(Specs, "|", ",|,"), "&", ",&,"), ",")
                End If
                
                Erase BinCheckData(flagIdx).CateArr(cateIdx).checkRules()
                For i = 0 To UBound(SpecArr)
                    SpecArr(i) = LCase(Replace(SpecArr(i), " ", ""))
                    ReDim Preserve BinCheckData(flagIdx).CateArr(cateIdx).checkRules(i)
                    Set BinCheckData(flagIdx).CateArr(cateIdx).checkRules(i).siteInfo = New Dictionary
                    BinCheckData(flagIdx).CateArr(cateIdx).checkRules(i).siteInfo.compareMode = TextCompare
                    
                    If BinCheckData(flagIdx).CateArr(cateIdx).conditionCheck Then
                        If SpecArr(i) Like "*$t" Then
                            BinCheckData(flagIdx).CateArr(cateIdx).checkRules(i).ConditionChkTrue = True
                        ElseIf SpecArr(i) Like "*$f" Then
                            BinCheckData(flagIdx).CateArr(cateIdx).checkRules(i).ConditionChkTrue = False
                        End If
                        BinCheckData(flagIdx).CateArr(cateIdx).checkRules(i).Rule = Replace(Replace(SpecArr(i), "$t", ""), "$f", "")
                    ElseIf BinCheckData(flagIdx).CateArr(cateIdx).specSite And (Not (SpecArr(i) Like "|" Or SpecArr(i) Like "&")) Then
                        If Not SpecArr(i) Like "*site*" Then
                            msg = "missing site infomation (" + SpecArr(i) + ") (row:" + CStr(Row) + " col:" + CStr(Col) + "), please check it!!!"
                            GoTo skipParser
                        End If
                        
                        Erase tmpArr()
                        tmpArr = Split(SpecArr(i), "(")
                        BinCheckData(flagIdx).CateArr(cateIdx).checkRules(i).Rule = tmpArr(0)
                        Erase tmpArr2()
                        tmpArr2 = Split(Replace(Replace(Replace(tmpArr(1), "site:", ""), "(", ""), ")", ""), ",")
                        For j = 0 To UBound(tmpArr2)
                            If BinCheckData(flagIdx).CateArr(cateIdx).checkRules(i).siteInfo.Exists(tmpArr2(j)) Then
                                msg = "site infomation duplicate (" + SpecArr(i) + ") (row:" + CStr(Row) + " col:" + CStr(Col) + "), please check it!!!"
                                GoTo skipParser
                            End If
                            BinCheckData(flagIdx).CateArr(cateIdx).checkRules(i).siteInfo.Add tmpArr2(j), True
                        Next j
                    Else
                        BinCheckData(flagIdx).CateArr(cateIdx).checkRules(i).Rule = SpecArr(i)
                    End If
    
                    If BinCheckData(flagIdx).CateArr(cateIdx).SplitBitCheck Then
                        bitwidth = (BinCheckData(flagIdx).CateArr(cateIdx).Bit_MSB - BinCheckData(flagIdx).CateArr(cateIdx).Bit_LSB) + 1
                        bytes = Round(bitwidth / 4)
                    ElseIf BinCheckData(flagIdx).CateArr(cateIdx).CombineFieldCheck Then
                        bytes = Round(bitwidth / 4)
                    Else
                        bitwidth = field.size
                        bytes = field.hBytes
                    End If
    
                    If (BinCheckData(flagIdx).CateArr(cateIdx).checkRules(i).Rule Like "|") Or (BinCheckData(flagIdx).CateArr(cateIdx).checkRules(i).Rule Like "&") Then
                        BinCheckData(flagIdx).CateArr(cateIdx).checkRules(i).isOperator = True
                    ElseIf BinCheckData(flagIdx).CateArr(cateIdx).checkRules(i).Rule Like "walking-*" Then
                        BinCheckData(flagIdx).CateArr(cateIdx).checkRules(i).WalkingCheck = True
                        BinCheckData(flagIdx).CateArr(cateIdx).checkRules(i).Walkingvalue = CLng(Replace(BinCheckData(flagIdx).CateArr(cateIdx).checkRules(i).Rule, "walking-", ""))
                    ElseIf BinCheckData(flagIdx).CateArr(cateIdx).checkRules(i).Rule Like "two-1" Then
                        BinCheckData(flagIdx).CateArr(cateIdx).checkRules(i).TwoCheck = True
                        If (bitwidth Mod 2) <> 0 Then
                            msg = "bit width can not be divided with no remainder (row:" + CStr(Row) + " col:" + CStr(Col) + "), please check it!!!"
                            GoTo skipParser
                        End If
                        BinCheckData(flagIdx).CateArr(cateIdx).checkRules(i).GroupBits = bitwidth / 2
                        BinCheckData(flagIdx).CateArr(cateIdx).checkRules(i).GroupPick = 2
                        BinCheckData(flagIdx).CateArr(cateIdx).checkRules(i).FailBit = CLng(Replace(BinCheckData(flagIdx).CateArr(cateIdx).checkRules(i).Rule, "two-", ""))
                    ElseIf BinCheckData(flagIdx).CateArr(cateIdx).checkRules(i).Rule Like "ids*" Then
                        BinCheckData(flagIdx).CateArr(cateIdx).checkRules(i).IDSCheck = True
                        If Not opbank.Fields.Exists(BinCheckData(flagIdx).CateArr(cateIdx).checkRules(i).Rule) Then
                            msg = "can't find ids category name (" + BinCheckData(flagIdx).CateArr(cateIdx).checkRules(i).Rule + ") in BDF table (row:" + CStr(Row) + " col:" + CStr(Col) + "), please check it!!!"
                            GoTo skipParser
                        End If
                    ElseIf BinCheckData(flagIdx).CateArr(cateIdx).checkRules(i).Rule Like "s*-p*-f*" Then
                        Erase tmpArr()
                        BinCheckData(flagIdx).CateArr(cateIdx).checkRules(i).GroupCheck = True
                        tmpArr = Split(Replace(BinCheckData(flagIdx).CateArr(cateIdx).checkRules(i).Rule, " ", ""), "-")
                        If (bitwidth Mod CLng(Replace(tmpArr(0), "s", ""))) <> 0 Then
                            msg = "bit width can not be divided with no remainder (row:" + CStr(Row) + " col:" + CStr(Col) + "), please check it!!!"
                            GoTo skipParser
                        End If
                        BinCheckData(flagIdx).CateArr(cateIdx).checkRules(i).GroupBits = bitwidth / CLng(Replace(tmpArr(0), "s", ""))
                        BinCheckData(flagIdx).CateArr(cateIdx).checkRules(i).GroupPick = CLng(Replace(tmpArr(1), "p", ""))
                        BinCheckData(flagIdx).CateArr(cateIdx).checkRules(i).FailBit = CLng(Replace(tmpArr(2), "f", ""))
                    ElseIf rangeParse.test(BinCheckData(flagIdx).CateArr(cateIdx).checkRules(i).Rule) Then
                        BinCheckData(flagIdx).CateArr(cateIdx).checkRules(i).RangeCheck = True
                        Erase tmpArr()
                        tmpArr = Split(Replace(Replace(BinCheckData(flagIdx).CateArr(cateIdx).checkRules(i).Rule, "[", ""), "]", ""), "-")
                        If bitwidth > 32 Then
                            loLimit = UCase(GlbUtility.String2Hex(CStr(tmpArr(0)), hBytes:=bytes))
                            hiLimit = UCase(GlbUtility.String2Hex(CStr(tmpArr(1)), hBytes:=bytes))
                            If GlbUtility.xHexCompare(loLimit, hiLimit, ChkGreaterThan) Then
                                value = hiLimit
                                hiLimit = loLimit
                                loLimit = value
                            End If
                            BinCheckData(flagIdx).CateArr(cateIdx).checkRules(i).LowHexValue = loLimit
                            BinCheckData(flagIdx).CateArr(cateIdx).checkRules(i).HighHexValue = hiLimit
                        Else
                            Call CheckValueToDec(CStr(tmpArr(0)), value)
                            BinCheckData(flagIdx).CateArr(cateIdx).checkRules(i).LowDecValue = value
                            Call CheckValueToDec(CStr(tmpArr(1)), value)
                            BinCheckData(flagIdx).CateArr(cateIdx).checkRules(i).HighDecValue = value
                            If BinCheckData(flagIdx).CateArr(cateIdx).checkRules(i).HighDecValue < BinCheckData(flagIdx).CateArr(cateIdx).checkRules(i).LowDecValue Then
                                value = BinCheckData(flagIdx).CateArr(cateIdx).checkRules(i).HighDecValue
                                BinCheckData(flagIdx).CateArr(cateIdx).checkRules(i).HighDecValue = BinCheckData(flagIdx).CateArr(cateIdx).checkRules(i).LowDecValue
                                BinCheckData(flagIdx).CateArr(cateIdx).checkRules(i).LowDecValue = value
                            End If
                        End If

                    ElseIf GlbUtility.IsNumber(BinCheckData(flagIdx).CateArr(cateIdx).checkRules(i).Rule) Then
                        BinCheckData(flagIdx).CateArr(cateIdx).checkRules(i).ValueCheck = True
                        If bitwidth > 32 Then
                            BinCheckData(flagIdx).CateArr(cateIdx).checkRules(i).specificHexValue = GlbUtility.String2Hex(BinCheckData(flagIdx).CateArr(cateIdx).checkRules(i).Rule, hBytes:=bytes)
                        Else
                            Call CheckValueToDec(BinCheckData(flagIdx).CateArr(cateIdx).checkRules(i).Rule, value)
                            BinCheckData(flagIdx).CateArr(cateIdx).checkRules(i).specificDecValue = value
                        End If
                    ElseIf BinCheckData(flagIdx).CateArr(cateIdx).checkRules(i).Rule Like "x" Then
                        BinCheckData(flagIdx).CateArr(cateIdx).checkRules(i).SkipTest = True
                        'don't check
                    Else
                        'unexpect format
                        msg = "spec format invalid (" + BinCheckData(flagIdx).CateArr(cateIdx).checkRules(i).Rule + ") (row:" + CStr(Row) + " col:" + CStr(Col) + "), please check it!!!"
                        GoTo skipParser
                    End If
                Next i
            Else
                Call Print_Error_Message(Warning_Info, "VBT_ZeFuseOthers", "ParseBinChk", "The " + TableName + " sheet bank name is empty !!! (row:" + CStr(Row) + ")")
            End If
        Next Row
    Next flagIdx
    ParseBinChk = True

Exit Function
skipParser:
    ParseBinChk = False
    Call Print_Error_Message(Error_Info, "VBT_ZeFuseOthers", "ParseBinChk", "The sheet " + TableName + " " + msg)
    theexec.Flow.TestLimit resultVal:=999, lowVal:=1, hiVal:=1, Tname:="ParseBinChk"
    Exit Function
errHandler:
    Call Print_Error_Message(Error_Info, "VBT_ZeFuseOthers", "ParseBinChk")
    If AbortTest Then Exit Function Else Resume Next
End Function

'20210911, Add for cmplfuse create realfields' mask and do setefuse
Public Function cmplfuse_createrealfielmask_setfuse(ob As eFuseBdfBank)
On Error GoTo errHandler
Dim binstr As String, setfuseBinStr As String, MaskBinStr As String
Dim ReadValueBin As Variant, CmpValueBin As Variant
Dim MaskArr() As New SiteDouble
Dim errorbitscnt As New SiteLong, errflag As Boolean
Dim setFuseValue As New SiteVariant
Dim cmpfield As Variant
Dim i As Long
Dim site As Variant

    For Each cmpfield In ob.DicCmpFuse.Keys
        Dim cmp_field As New eFuseBdfField
        If ob.Fields.Exists(cmpfield) And Not cmpfield Like "*_mod" Then
            Set cmp_field = ob.Fields(cmpfield)
            If cmp_field.DefaultOrReal = dr_real Then
                ReDim MaskArr(cmp_field.size - 1)
                errorbitscnt = 0
                errflag = False
                
                For Each site In theexec.sites
                    binstr = GlbUtility.Hex2BinStr(cmp_field.DsscValue, cmp_field.size)
                    ReadValueBin = Split(StrConv(binstr, vbUnicode), vbNullChar)
                    binstr = ob.DicCmpFuse(cmpfield)
                    CmpValueBin = Split(StrConv(binstr, vbUnicode), vbNullChar)

                    setfuseBinStr = vbNullString
                    For i = cmp_field.size - 1 To 0 Step -1
                        If CmpValueBin(i) = "N" Then   '0/1->N : Don't Care, keep original
                            setfuseBinStr = ReadValueBin(i) + setfuseBinStr
                            MaskArr((cmp_field.size - 1) - i) = 0
                            MaskBinStr = CStr(MaskArr((cmp_field.size - 1) - i)) + MaskBinStr
                        ElseIf ReadValueBin(i) = "0" And CmpValueBin(i) = "1" Then   '0->1 : mask set true
                            setfuseBinStr = CmpValueBin(i) + setfuseBinStr
                            MaskArr((cmp_field.size - 1) - i) = 1
                            MaskBinStr = CStr(MaskArr((cmp_field.size - 1) - i)) + MaskBinStr
                        ElseIf ReadValueBin(i) = "1" And CmpValueBin(i) = "1" Then   '1->:1 : mask set false
                            setfuseBinStr = CmpValueBin(i) + setfuseBinStr
                            MaskArr((cmp_field.size - 1) - i) = 0
                            MaskBinStr = CStr(MaskArr((cmp_field.size - 1) - i)) + MaskBinStr
                        ElseIf ReadValueBin(i) = "0" And CmpValueBin(i) = "0" Then   '0->:0 : mask set true
                            setfuseBinStr = CmpValueBin(i) + setfuseBinStr
                            MaskArr((cmp_field.size - 1) - i) = 1
                            MaskBinStr = CStr(MaskArr((cmp_field.size - 1) - i)) + MaskBinStr
                        Else    '1->:0 :Don't be allowed
                            MaskBinStr = "E" + MaskBinStr
                            GlbUtility.WriteDlg "Site:" + CStr(site) + ", " + cmp_field.name & " had error comoplement bits(1 to 0)! Please check Table again!"
                            errorbitscnt = errorbitscnt + 1
                            errflag = True
                        End If
                    Next i
                    GlbUtility.WriteDlg "Site:" + CStr(site) + ", <CmpFuseMask> " + cmp_field.name + ":" + MaskBinStr
                    setFuseValue = GlbUtility.Bin2Dec(setfuseBinStr)
                Next
                
                If errflag = True Then
                    theexec.Flow.TestLimit resultVal:=errorbitscnt, lowVal:=0, hiVal:=0, Tname:="OverwriteErrorBitCnt", PinName:="Value"
                Else
                    ob.SetEfuse cmp_field.name, setFuseValue, , , , , True
                End If
                
                If ob.DicCmpFuse.Exists(cmpfield + "_mod") Then ob.DicCmpFuse.Remove (cmpfield + "_mod")
                ob.DicCmpFuse.Add cmpfield + "_mod", MaskArr
            End If
        End If
    Next

Exit Function
errHandler:
     Call Print_Error_Message(Error_Info, "VBT_ZeFuseOthers", "cmplfuse_createrealfielmask_setfuse")
     If AbortTest Then Exit Function Else Resume Next
End Function

'20210911, Add for cmplfuse modify dspwave for avoiding double fusing
Public Function cmplfuse_modifydspwave_avoiddoublefusing(ob As eFuseBdfBank, dcTrimmed As Dictionary)
On Error GoTo errHandler
Dim cmpfield As Variant
Dim stBit As Long
Dim spBit As Long
Dim iBit As Long
Dim site As Variant

    For Each cmpfield In ob.DicCmpFuse.Keys
        If dcTrimmed.Exists(cmpfield) And Not cmpfield Like "*_mod" Then
            stBit = ob.Fields(cmpfield).LSB
            spBit = ob.Fields(cmpfield).msb
            
            If ob.Fields(cmpfield).DefaultOrReal = dr_real Then
                If ob.DicCmpFuse.Exists(cmpfield + "_mod") Then
                    For Each site In theexec.sites
                        For iBit = stBit To spBit
                            If ob.DicCmpFuse(cmpfield + "_mod")(iBit - stBit) = 0 Then
                                ob.SingleBitWave.Element(iBit) = 0
                                GlbUtility.WriteDlg "Site:" + CStr(site) + ", Bit: " + CStr(iBit) + ",Set write bit to 0."
                            End If
                        Next
                    Next
                    ob.DicCmpFuse.Remove (cmpfield + "_mod")
                Else
                    'error situation, don't fuse anything
                    For iBit = stBit To spBit
                        For Each site In theexec.sites
                            ob.SingleBitWave.Element(iBit) = 0
                            GlbUtility.WriteDlg "Site:" + CStr(site) + ", Bit: " + CStr(iBit) + ",Set write bit to 0."
                        Next
                    Next
                End If
            Else
                For iBit = stBit To spBit
                    If ob.DicCmpFuse(cmpfield)(iBit - stBit) = 0 Then
                        For Each site In theexec.sites
                             ob.SingleBitWave.Element(iBit) = 0
                             GlbUtility.WriteDlg "Site:" + CStr(site) + ", Bit: " + CStr(iBit) + ",Set write bit to 0."
                        Next
                    End If
                Next
            End If
        End If
    Next

Exit Function
errHandler:
     Call Print_Error_Message(Error_Info, "VBT_ZeFuseOthers", "cmplfuse_modifydspwave_avoiddoublefusing")
     If AbortTest Then Exit Function Else Resume Next
End Function

'20211110,Add for two type bincut table syntax check
Public Function ParseBincut(EfuseBinCut() As EFUSE_BINCUT_TYPE, Optional AdditionalTableName As String = vbNullString) As Boolean
On Error GoTo errHandler
Dim wb As Workbook
Dim ws_def As Worksheet
Dim MaxRow As Long
Dim maxcol As Long
Dim isSheetFound As Boolean
Dim sheetName As String
Dim passBinCut As Variant
Dim col_binned As Integer
Dim col_domain As Integer
Dim col_mode As Integer
Dim col_eqn As Integer
Dim col_id As Integer
Dim col_c As Integer
Dim col_m As Integer
Dim col_cp_vmax As Integer
Dim col_cp_vmin As Integer
Dim col_cpgb As Integer
Dim main_p_mode As Integer
Dim powerDomain As String
Dim msg As String: msg = vbNullString
Dim p_mode As Long
Dim Row As Long
Dim Col As Long
Dim Row_of_Title As Long
Dim enableRowParsing As Boolean
Dim strAry_Temp As Variant
Dim idx_step As Long
    
    For Each passBinCut In PassBinCut_ary
        If AdditionalTableName = "" Then
            sheetName = "Vdd_Binning_Def_appA_" & passBinCut
            Set wb = Application.ActiveWorkbook
            Call check_Sheet_Range(sheetName, wb, ws_def, MaxRow, maxcol, isSheetFound)
        Else
            sheetName = "Vdd_Binning_Def_appA_" & passBinCut & "_" & AdditionalTableName
            Set wb = Application.ActiveWorkbook
            Call check_Sheet_Range(sheetName, wb, ws_def, MaxRow, maxcol, isSheetFound, False)
            If isSheetFound = False Then
                sheetName = "Vdd_Binning_Def_appA_" & passBinCut
                Set wb = Application.ActiveWorkbook
                Call check_Sheet_Range(sheetName, wb, ws_def, MaxRow, maxcol, isSheetFound)
            End If
        End If

        '''*****************************************************************'''
        If isSheetFound = True Then
            '''//init
            For p_mode = 0 To MaxPerformanceModeCount - 1           'initilize the MODE_STEP and ExcludedPmode
                EfuseBinCut(p_mode, passBinCut).Mode_Step = -99
            Next p_mode
        
            For Row = 1 To MaxRow
                For Col = 1 To maxcol
                    '''******************************************************************************************************************'''
                    '''//If CorePower and OtherRail are in the same table (only Vdd_Binning_Def), 1st column is "Binned".
                    '''//If CorePower and OtherRail are in the different tables (Vdd_Binning_Def and Other_Rail), 1st column is "Domain".
                    '''******************************************************************************************************************'''
                    '''If 1st column 1 of the header is "Binned", split the line and find out the keyword column.
                    If LCase(ws_def.Cells(Row, Col).value) Like "binned" Then
                        col_binned = Col
                        Row_of_Title = Row
                    End If
            
                    If Row_of_Title > 0 Then
                        If LCase(ws_def.Cells(Row_of_Title, Col).value) = "domain" Then
                            col_domain = Col
                        ElseIf LCase(ws_def.Cells(Row_of_Title, Col).value) = "mode" Then
                            col_mode = Col
                        ElseIf LCase(ws_def.Cells(Row_of_Title, Col).value) = "id" Then
                            col_id = Col
                        ElseIf LCase(ws_def.Cells(Row_of_Title, Col).value) = "eqn" Then
                            col_eqn = Col
                        ElseIf LCase(ws_def.Cells(Row_of_Title, Col).value) = "c" Then
                            col_c = Col
                        ElseIf LCase(ws_def.Cells(Row_of_Title, Col).value) = "m" Then
                            col_m = Col
                        ElseIf LCase(ws_def.Cells(Row_of_Title, Col).value) = "cpvmax" Or LCase(ws_def.Cells(Row_of_Title, Col).value) = LCase("BinningVmax") Then
                            col_cp_vmax = Col
                        ElseIf LCase(ws_def.Cells(Row_of_Title, Col).value) = "cpvmin" Or LCase(ws_def.Cells(Row_of_Title, Col).value) = LCase("BinningVmin") Then
                            col_cp_vmin = Col
                        ElseIf LCase(ws_def.Cells(Row_of_Title, Col).value) = "cpgb" Or LCase(ws_def.Cells(Row_of_Title, Col).value) = LCase("BinningGB") Then
                            col_cpgb = Col
                        End If
                    End If
                
                    '''//Check if all columns of the header exist...
                    '''Note: col_comment should be checked as the last available column of the table.
                    If col_domain > 0 And col_mode > 0 And col_id > 0 And col_eqn > 0 And col_c > 0 And col_m > 0 _
                    And col_cp_vmax > 0 And col_cp_vmin > 0 And col_cpgb > 0 Then
                        enableRowParsing = True
                    End If
                Next Col
            
                '''//If all columns of the header are found, skip the loop and start parsing each row.
                If enableRowParsing = True Then
                    Exit For
                End If
            
                If Row = MaxRow And (col_domain = 0 Or col_mode = 0) Then
                    enableRowParsing = False
                    msg = "Columns of header in " & sheetName & " are incorrect, please check it!!!"
                    GoTo skip
                End If
            Next Row
        
            If enableRowParsing = True And Row_of_Title + 1 <= MaxRow Then
                For Row = Row_of_Title + 1 To MaxRow
                    '''//If first word in the mode column is M(ex: MC601).
                    '''//If column "Binned" is "true", it means that performance_mode of powerDomain is the binning mode of CorePower.
                    If ws_def.Cells(Row, col_mode).value Like "M*" And LCase(ws_def.Cells(Row, col_binned).value) = "true" Then
                        main_p_mode = VddBinStr2Enum(ws_def.Cells(Row, col_mode))
                        If UCase(ws_def.Cells(Row, col_domain).value) Like "VDD*" Then
                            powerDomain = UCase(Trim(ws_def.Cells(Row, col_domain)))
                        ElseIf UCase(ws_def.Cells(Row, col_domain).value) <> "" Then
                            powerDomain = "VDD_" & UCase(Trim(ws_def.Cells(Row, col_domain)))
                        Else
                            theexec.Datalog.WriteComment ws_def.Cells(Row, col_domain) & " doesn't have the correct Domain cell in sheet " & sheetName & ". Error!!!"
                            theexec.ErrorLogMessage ws_def.Cells(Row, col_domain) & " doesn't have the correct Domain cell in sheet " & sheetName & ". Error!!!"
                        End If
                    
                        If ws_def.Cells(Row, col_eqn).value Like "E#*" Then '''read the E1 ~ En
                            strAry_Temp = Split(ws_def.Cells(Row, col_eqn), "E") '''ex: array(0)=E ; array(1)=1
                            idx_step = CLng(strAry_Temp(1)) - 1 '''step: the address for store the EQ number, ex: BinCut(P_mode,passbinnum).EQ_Num(0)=1, step = 0, EQ = 1
                        
                            EfuseBinCut(main_p_mode, passBinCut).c(idx_step) = CDbl(ws_def.Cells(Row, col_c).value)
                            EfuseBinCut(main_p_mode, passBinCut).m(idx_step) = CDbl(ws_def.Cells(Row, col_m).value)
                            
                            '''*************************************************************************************'''
                            '''//Check if CPVmin, CPVmax, and CPGB are multiple of Step Size voltage.
                            EfuseBinCut(main_p_mode, passBinCut).CP_Vmax(idx_step) = CDbl(ws_def.Cells(Row, col_cp_vmax).value)
                            
                            If EfuseBinCut(main_p_mode, passBinCut).CP_Vmax(idx_step) <> (Floor(EfuseBinCut(main_p_mode, passBinCut).CP_Vmax(idx_step) / BV_StepVoltage) * BV_StepVoltage) Then
                                theexec.Datalog.WriteComment sheetName & ", p_mode:" & ws_def.Cells(Row, col_mode).value & ", EQN:" & ws_def.Cells(Row, col_eqn).value & ", CPVmax:" & EfuseBinCut(main_p_mode, passBinCut).CP_Vmax(idx_step) & " should be multiple of 3.125. Error!!!"
                                theexec.ErrorLogMessage sheetName & ", p_mode:" & ws_def.Cells(Row, col_mode).value & ", EQN:" & ws_def.Cells(Row, col_eqn).value & ", CPVmax:" & EfuseBinCut(main_p_mode, passBinCut).CP_Vmax(idx_step) & " should be multiple of 3.125. Error!!!"
                            End If
                        
                            EfuseBinCut(main_p_mode, passBinCut).CP_Vmin(idx_step) = CDbl(ws_def.Cells(Row, col_cp_vmin).value)
                            If EfuseBinCut(main_p_mode, passBinCut).CP_Vmin(idx_step) <> (Floor(EfuseBinCut(main_p_mode, passBinCut).CP_Vmin(idx_step) / BV_StepVoltage) * BV_StepVoltage) Then
                                theexec.Datalog.WriteComment sheetName & ", p_mode:" & ws_def.Cells(Row, col_mode).value & ", EQN:" & ws_def.Cells(Row, col_eqn).value & ", CPVmin:" & EfuseBinCut(main_p_mode, passBinCut).CP_Vmin(idx_step) & " should be multiple of 3.125. Error!!!"
                                theexec.ErrorLogMessage sheetName & ", p_mode:" & ws_def.Cells(Row, col_mode).value & ", EQN:" & ws_def.Cells(Row, col_eqn).value & ", CPVmin:" & EfuseBinCut(main_p_mode, passBinCut).CP_Vmin(idx_step) & " should be multiple of 3.125. Error!!!"
                            End If
                        
                            EfuseBinCut(main_p_mode, passBinCut).CP_GB(idx_step) = CDbl(ws_def.Cells(Row, col_cpgb).value)
    
                            If EfuseBinCut(main_p_mode, passBinCut).CP_GB(idx_step) <> (Floor(EfuseBinCut(main_p_mode, passBinCut).CP_GB(idx_step) / BV_StepVoltage) * BV_StepVoltage) Then
                                theexec.Datalog.WriteComment sheetName & ", p_mode:" & ws_def.Cells(Row, col_mode).value & ", EQN:" & ws_def.Cells(Row, col_eqn).value & ", CPGB:" & EfuseBinCut(main_p_mode, passBinCut).CP_GB(idx_step) & " should be multiple of 3.125. Error!!!"
                                theexec.ErrorLogMessage sheetName & ", p_mode:" & ws_def.Cells(Row, col_mode).value & ", EQN:" & ws_def.Cells(Row, col_eqn).value & ", CPGB:" & EfuseBinCut(main_p_mode, passBinCut).CP_GB(idx_step) & " should be multiple of 3.125. Error!!!"
                            End If
                            '''*************************************************************************************'''
                            EfuseBinCut(main_p_mode, passBinCut).Mode_Step = idx_step
                        End If
                    End If
                Next Row
            Else
                msg = "Columns of the header in the sheet " & sheetName & " might be incorrect, please check it!!!"
                GoTo skip
            End If
        End If '''If isSheetFound = True
    
        '''set the last Step to the error value
        For p_mode = 0 To MaxPerformanceModeCount - 1
            EfuseBinCut(p_mode, passBinCut).c(MaxEqnNum) = 0
            EfuseBinCut(p_mode, passBinCut).m(MaxEqnNum) = 0
            EfuseBinCut(p_mode, passBinCut).CP_Vmax(MaxEqnNum) = 0
            EfuseBinCut(p_mode, passBinCut).CP_Vmin(MaxEqnNum) = 0
            EfuseBinCut(p_mode, passBinCut).CP_GB(MaxEqnNum) = 0
        Next p_mode
    Next passBinCut
    ParseBincut = True

Exit Function
skip:
    ParseBincut = False
    Call Print_Error_Message(Error_Info, "VBT_ZeFuseOthers", "ParseBincut", msg)
    theexec.Flow.TestLimit resultVal:=999, lowVal:=1, hiVal:=1, Tname:="ParseBincut"
    Exit Function
errHandler:
    Call Print_Error_Message(Error_Info, "VBT_ZeFuseOthers", "ParseBincut")
    If AbortTest Then Exit Function Else Resume Next
End Function

Public Function ParserEnableWdGatingTable(sheetName As String, realPhase As EnableWdGatingPhase) As Boolean
On Error GoTo errHandler
    Dim tmpUtility As New Utility
    Dim m_SheetMaxRow As Long, m_SheetMaxCol As Long
    Dim m_ParsingArr() As Variant
    Dim rowNum As Variant, phase As Variant
    Dim enableWd As String: enableWd = vbNullString
    Dim checkPhase As EnableWdGatingPhase
    Dim dic As New Dictionary
    Dim tpKeyword As String: tpKeyword = vbNullString
    Dim checkPhaseList() As String
    
    'Get enable word list and check phase
    'enableWdGatingTableInfo.enableWdDic:  key:checkPhase; value:(key:enable word, value: tp name keyword)
    ParserEnableWdGatingTable = True
    If Not enableWdGatingTableInfo.ParserTableDone Then
        enableWdGatingTableInfo.ParserTableDone = True
        If Not Evaluate("ISREF('" & sheetName & "'!A1)") Then
            enableWdGatingTableInfo.TableExist = False
        Else
            enableWdGatingTableInfo.TableExist = True
            Application.ScreenUpdating = False
            tmpUtility.Get_Sheet_Info sheetName, m_SheetMaxRow, m_SheetMaxCol, m_ParsingArr
            Application.ScreenUpdating = True
            Set enableWdGatingTableInfo.enableWdDic = New Dictionary
            enableWdGatingTableInfo.enableWdDic.compareMode = TextCompare
            For rowNum = 1 To m_SheetMaxRow
                enableWd = Trim(m_ParsingArr(rowNum, 1))
                tpKeyword = UCase(Trim(m_ParsingArr(rowNum, 2)))
                If Not (LCase(enableWd) Like "enable*word") Then 'skip header
                    checkPhaseList = Split(LCase(Trim(m_ParsingArr(rowNum, 3))), ",")
                    For Each phase In checkPhaseList
                        Set dic = New Dictionary: dic.compareMode = TextCompare
                        checkPhase = GetCheckPhaseType(Trim(phase))
                        If enableWd <> "" Then
                            If checkPhase = unknown Then
                                enableWdGatingTableInfo.ParserTableDone = False
                                If realPhase = OnProgValidation Then theexec.AddOutput "Please perform Validate Job again!!!", vbRed, True
                                ParserEnableWdGatingTable = False
                                Exit For
                            Else
                                If enableWdGatingTableInfo.enableWdDic.Exists(checkPhase) Then
                                    Set dic = enableWdGatingTableInfo.enableWdDic(checkPhase)
                                    If dic.Exists(enableWd) Then
                                        MsgBox "Duplicate enable word!!!Please check (" + enableWd + ")", vbOKOnly, "EnableWdGatingTable format incorrect"
                                        If realPhase = OnProgValidation Then theexec.AddOutput "Please perform Validate Job again!!!", vbRed, True
                                        enableWdGatingTableInfo.ParserTableDone = False
                                        ParserEnableWdGatingTable = False
                                        Exit For
                                    Else
                                        dic.Add enableWd, tpKeyword
                                        Set enableWdGatingTableInfo.enableWdDic(checkPhase) = dic
                                    End If
                                Else
                                    dic.Add enableWd, tpKeyword
                                    enableWdGatingTableInfo.enableWdDic.Add checkPhase, dic
                                End If
                            End If
                        End If
                    Next phase
                End If
            Next rowNum
        End If
    End If
Exit Function
errHandler:
    Call Print_Error_Message(Error_Info, "VBT_ZeFuseOthers", "ParserEnableWdGatingTable")
    If AbortTest Then Exit Function Else Resume Next
End Function

Public Function EnableWdGating(realPhase As EnableWdGatingPhase)
On Error GoTo errHandler
    Dim enableWd As Variant, colNum As Variant, checkPhase As Variant
    Dim enableWdState As Boolean, T0TxIsEnabled As Boolean
    Dim tpName As String: tpName = vbNullString
    Dim tpKeyword As String: tpKeyword = vbNullString
    Dim answer As Integer
    Dim msg As String: msg = vbNullString
    Dim m_site As Variant
    Dim i As Integer
    Dim StpFolderPath As String: StpFolderPath = "C:\Flex\Applications\"
    Dim m_FileType As String: m_FileType = "stp"
    Dim m_FileName As String
    Dim m_file As String
    Dim EnableWDArr() As String
    Dim ParserStpFlag As Boolean

    If realPhase = OnInitEnableWord Then
        If theexec.Flow.enableWord("EnableWdCheckFail") Then
            GoTo binout
        End If
    Else
        'init enable word state
        theexec.Flow.enableWord("EnableWdCheckFail") = False
    End If
    
    tpName = UCase(theexec.TestProgram.name)
    If theexec.RunMode = runModeProduction Then
        '1. Parsing enable word gating table and save infomation to global struct "enableWdGatingTableInfo"
        If Not ParserEnableWdGatingTable("EnableWdGatingTable", realPhase) Then
            If realPhase = OnInitEnableWord Then GoTo binout
        Else
            '2. check enable word status
            If Not enableWdGatingTableInfo.TableExist Then
                If realPhase = OnProgValidation Then
                    theexec.AddOutput "<Warning> Sheet EnableWdGatingTable not exist in TP!!", vbBlue, True
                Else
                    Call Print_Error_Message(Warning_Info, "VBT_ZeFuseOthers", "EnableWdGating", "Sheet EnableWdGatingTable not exist in TP!!")
                End If
            Else
                If enableWdGatingTableInfo.enableWdDic.Exists(OnProgValidation) Then
                    For Each enableWd In enableWdGatingTableInfo.enableWdDic.item(OnProgValidation)
                        If tpName Like enableWdGatingTableInfo.enableWdDic.item(OnProgValidation).item(enableWd) Then
                            ParserStpFlag = True
                        End If
                    Next enableWd
                End If
                
                If realPhase = OnProgValidation And ParserStpFlag Then
                    Erase EnableWDArr
                    If Not (CheckFileExist(StpFolderPath, "", m_FileType, m_FileName, m_file)) Then
                        If m_FileName = Empty Then
                            enableWdGatingTableInfo.enableWdDic.Add NotExistStp, enableWdGatingTableInfo.enableWdDic.item(OnProgValidation)
                            enableWdGatingTableInfo.enableWdDic.Remove (OnProgValidation)
                            theexec.AddOutput "<Warning> Can't find stp file. Change check phase to init enableWd sheet!!", vbBlue, True
                        Else
                            Call MsgBox("Open OI config file error!!!", vbOKOnly, "Enable Word Gating Fail")
                            GoTo binout
                        End If
                    Else
                        Open m_file For Input As #1
                        Do Until EOF(1)
                            Call GattingOIConfigEnableWD(EnableWDArr())
                        Loop
                        Close #1
                    End If
                End If
                
                T0TxIsEnabled = False
                For Each checkPhase In enableWdGatingTableInfo.enableWdDic
                    If checkPhase = realPhase Or (realPhase = OnInitEnableWord And checkPhase = NotExistStp) Then
                        For Each enableWd In enableWdGatingTableInfo.enableWdDic.item(checkPhase)
                            If checkPhase = OnProgValidation Then
                                enableWdState = False
                                
                                For i = 0 To UBound(EnableWDArr)
                                    If EnableWDArr(i) = enableWd Then
                                        enableWdState = True
                                        Exit For
                                    End If
                                Next
                            Else
                                enableWdState = theexec.Flow.enableWord(enableWd)
                            End If
                            
                            tpKeyword = enableWdGatingTableInfo.enableWdDic.item(checkPhase).item(enableWd)
                            msg = "Enable word (" + enableWd + ") is not selected, please check !!!"
                            If LCase(enableWd) Like "efuse_all_enable" And tpName Like tpKeyword Then
                                If Evaluate("ISREF('" & BdfSheetName & "'!A1)") And enableWdState = False Then
                                    Call MsgBox(msg, vbOKOnly, "Enable Word Gating Fail")
                                    GoTo binout
                                End If
                            ElseIf LCase(enableWd) Like "enableotp" And tpName Like tpKeyword Then
                                If Evaluate("ISREF('" & "otp_register_map" & "'!A1)") And enableWdState = False Then
                                    Call MsgBox(msg, vbOKOnly, "Enable Word Gating Fail")
                                    GoTo binout
                                End If
                            ElseIf enableWd Like "T0Tx_Hot" Or enableWd Like "T0Tx_Room" Then
                                If T0TxIsEnabled Then
                                    If enableWdState Then
                                        msg = "Both T0Tx_Hot and T0Tx_Room are selected, please check !!!"
                                        Call MsgBox(msg, vbOKOnly, "Enable Word Gating Fail")
                                        GoTo binout
                                    End If
                                Else
                                    If enableWdState Then
                                        If tpName Like tpKeyword Then
                                            T0TxIsEnabled = True
                                        Else
                                            msg = "Test Program Name keyword (" + tpKeyword + ") with enable word (" + enableWd + ") status mismatch, please check !!!"
                                            Call MsgBox(msg, vbOKOnly, "Enable Word Gating Fail")
                                            GoTo binout
                                        End If
                                    Else
                                        If tpName Like tpKeyword Then
                                            Call MsgBox(msg, vbOKOnly, "Enable Word Gating Fail")
                                            GoTo binout
                                        End If
                                    End If
                                End If
                            Else
                                If ((Not enableWdState) And (tpName Like tpKeyword)) Then
                                    Call MsgBox(msg, vbOKOnly, "Enable Word Gating Fail")
                                    GoTo binout
                                End If
                            End If
                        Next
                    End If
                Next
            End If
        End If
    End If

Exit Function
binout:
    If realPhase = OnInitEnableWord Then
        For Each m_site In theexec.sites
            theexec.Flow.TestLimit resultVal:=1, lowVal:=0, hiVal:=0, Tname:="EnableWdGating_Fail"
        Next
    ElseIf realPhase = OnProgValidation Then
        theexec.AddOutput "Please check enable word status and perform Validate Job again!!!", vbRed, True
        theexec.Flow.enableWord("EnableWdCheckFail") = True
    End If
    Exit Function
errHandler:
    Call Print_Error_Message(Error_Info, "VBT_ZeFuseOthers", "EnableWdGating")
    If AbortTest Then Exit Function Else Resume Next
End Function

Private Function GetCheckPhaseType(phase As String) As EnableWdGatingPhase
On Error GoTo errHandler

    If phase = 1 Or phase Like "disable" Then
        GetCheckPhaseType = Disable
    ElseIf phase Like "onprogvalidation" Or phase = 2 Or phase = "" Then
        GetCheckPhaseType = OnProgValidation
    ElseIf phase Like "oninitenableword" Or phase = 3 Then
        GetCheckPhaseType = OnInitEnableWord
    Else
        'unknow type
        GetCheckPhaseType = unknown
        MsgBox "Unknow type!!! Please check phase type (" + phase + ")", vbOKOnly, "EnableWdGatingTable format incorrect"
    End If

Exit Function
errHandler:
    Call Print_Error_Message(Error_Info, "VBT_ZeFuseOthers", "GetCheckPhaseType")
    If AbortTest Then Exit Function Else Resume Next
End Function

Public Sub GattingOIConfigEnableWD(ByRef EnableWDArr() As String)
On Error GoTo errHandler
Dim m_lineStr As String: m_lineStr = vbNullString
Dim key As String: key = vbNullString
Dim addr As String: addr = vbNullString
Dim data As String: data = vbNullString

    Line Input #1, m_lineStr
    m_lineStr = Trim(m_lineStr)
    addr = InStr(1, m_lineStr, ",")
    key = mid(m_lineStr, 1, addr - 1)
    data = Trim(Replace(mid(m_lineStr, addr + 1), vbTab, ""))
 
    If key Like "EnableWords" Then
        EnableWDArr = Split(data, " ")
    End If

Exit Sub
errHandler:
    Call Print_Error_Message(Error_Info, "VBT_ZeFuseOthers", "GattingOIConfigEnableWD")
    If AbortTest Then Exit Sub Else Resume Next
End Sub

Public Function JudgeDRAMType_T()
' THIS FUNCTION IS RESPONSIBLE FOR IDENTIFYING DRAM VENDOR
' BASED ON DRAM VENDOR SPECIFY RELATED ENABLE WORD TO
' ENABLE CORRELATIVE USE-LIMIT FOR IDS TEST
' 202405 JackChou new request version DRAM
' 3 ways for raise DRAM enable word :
' [Step2 only excute when enableWord "T0TX_HOT" and "T0TX_ROOM" are False]
' [Step3 only excute when enableWord "T0TX_HOT" or "T0TX_ROOM" is True]
' -Step1: For OSAT, raise DRAM enableWord by OI.
' -Step2: For OTC (InHouse), raise DRAM enableWord by DRAM_Mapping file.
' -Step3: Raise DRAM enableWord by DRAM fields value.

On Error GoTo errHandler

' INITIALIZATION
Dim LotInfo_FilePath As String: LotInfo_FilePath = "C:\Flex\Applications\lotinfo.txt"
'Dim DramMap_FilePath As String: DramMap_FilePath = "X:\Production\TMQG36\DRAM_Mapping\"
Dim DramMap_FilePath As String: DramMap_FilePath = theexec.TestProgram.path & "\DRAM_Mapping\"
Dim obj_fso As Object: Set obj_fso = CreateObject("Scripting.FileSystemObject")
Dim txtStream As TextStream
Dim LineTxt As String
Dim lineNum As Long: lineNum = 1
Dim lotId As String: lotId = theexec.Datalog.Setup.LotSetup.lotId
Dim Dram_PartName As String: Dram_PartName = vbNullString
Dim Dram_Mapping_Path As String: Dram_Mapping_Path = vbNullString
Dim Dram_Mapping_File As String: Dram_Mapping_File = vbNullString
Dim Dram_Type_Ieda_Temp() As String
Dim Dram_Type As String: Dram_Type = vbNullString
Dim Dram_Ieda As String: Dram_Ieda = vbNullString
Dim Tname As String: Tname = vbNullString
Dim ProgramName As String: ProgramName = theexec.TestProgram.name
Dim Project_PartName As String: Project_PartName = vbNullString
''''''''''''''''''''''''JackChou Dram offline test''''''''''''''''''''''''
'LotInfo_FilePath = "D:\_Hid\DRAM_TEST\lotinfo.txt"
'DramMap_FilePath = "D:\_Hid\DRAM_TEST\DRAM_Mapping\"
''''''''''''''''''''''''JackChou Dram offline test''''''''''''''''''''''''
Dim enableWord As Variant, siteControlArray As Variant
Dim k As Integer
Dim msg As String: msg = vbNullString
Dim ChkStepFlag As Boolean: ChkStepFlag = False
Dim opbank As eFuseBdfBank
Dim field As eFuseBdfField
Dim m_loop As Long
Dim m_catename As String
Dim m_value As New SiteVariant
Dim site As Variant
Dim FieldChkCnt As New SiteLong
Dim siteVal As New SiteBoolean: siteVal = False
Dim DRAMEnableWordStr As String: DRAMEnableWordStr = vbNullString
Dim AllSite_DRAM_IEDA As String
Dim Ary_Dram_IEDA() As String
Dim oValue() As Double
Dim siteResult As New SiteBoolean: siteResult = False
Dim resultCnt As Integer: resultCnt = 0

    ' Step1
    If Not SkipStep1_Flag Then
        Call BdfDataBase.ProcessHipDram(ChkStepFlag)
        If ChkStepFlag Then
            theexec.Flow.TestLimit resultVal:=1, lowVal:=1, hiVal:=1, Tname:="Judge_Dram_Type"
        End If
    End If
    If Not ChkStepFlag Then
        For Each enableWord In BdfDataBase.DicDramMap.Keys
            theexec.Flow.enableWord(enableWord) = False
        Next enableWord
        'step2
        If (theexec.Flow.enableWord("T0TX_HOT") = False) And (theexec.Flow.enableWord("T0TX_ROOM") = False) Then
            Call Print_Error_Message(Warning_Info, "VBT_ZeFuseOthers", "JudgeDRAMType_T", "Did not turn on any enable word from dram table!!!!")

            ' BINOUT: LOT INFO FILE NOT EXISTS
            If Not obj_fso.FileExists(LotInfo_FilePath) Then
                Call Print_Error_Message(Error_Info, "VBT_ZeFuseOthers", "JudgeDRAMType_T", "If you are OSAT: Please check DRAM Enable word!!!!")
                Call Print_Error_Message(Error_Info, "VBT_ZeFuseOthers", "JudgeDRAMType_T", "If you are InHouse: Please check lotinfo.txt File!!!!")
                theexec.Flow.TestLimit resultVal:=0, lowVal:=1, hiVal:=1, Tname:="lotinfo.txt File Not Found"
                Exit Function
            Else
                SkipStep1_Flag = True '''for InHouse case 2nd/3th/4th... TouchDown keep using "Step2"
                ' IF LOT INFO FILE EXISTS
                Set txtStream = obj_fso.OpenTextFile(LotInfo_FilePath, ForReading, False)
                
                ' LOOP LOT INFRO FILE LINE BY LINE
                Do While Not txtStream.AtEndOfStream
                    LineTxt = txtStream.ReadLine
                    
                    ' READ FIRST LINE, EX: LotID:L906H4.E8
                    If lineNum = 1 Then
                        ' BINOUT: LOTID MISMATCHES

                        theexec.Datalog.WriteComment vbCrLf & LineTxt & " (lotinfo.txt)" & vbCrLf & "LotID:" & lotId & " (OI)"
                        'theexec.Datalog.WriteComment "LotID:" & lotId & " (OI)"
                        If Not (LCase(LineTxt) Like "*lotid*") Or UCase(lotId) <> Split(LineTxt, ":")(1) Then
                            theexec.Flow.TestLimit resultVal:=0, lowVal:=1, hiVal:=1, Tname:="LotID Mismatch"
                            Exit Function
                        End If
                        lineNum = lineNum + 1
                    ' READ SECOND LINE, EX: Part32:TMIT78C-141C5L1T1D5CGAPL
                    ElseIf lineNum = 2 Then
                        ' BINOUT: SECOND LINE CONTENT MISMATCHES
                        If Not (LCase(LineTxt) Like "*part32*") Or LCase(mid(LineTxt, 10, 4)) <> LCase(mid(ProgramName, 1, 4)) Then
                            theexec.Datalog.WriteComment "Program part: " & mid(ProgramName, 1, 4) & vbCrLf & "LotInfo part: " & mid(LineTxt, 10, 4)
                            theexec.Flow.TestLimit resultVal:=0, lowVal:=1, hiVal:=1, Tname:="Part32 Mismatch"
                            Exit Function
                        Else
                            'LineTxt Content Part32:TMIT78C-141C5L1T1D5CGAPL
                            Dram_PartName = mid(LineTxt, 19, 2)
                            Project_PartName = mid(LineTxt, 8, 6)
                        End If
                        theexec.Datalog.WriteComment LineTxt
                        lineNum = lineNum + 1
                    End If
                Loop ' LOOP LOT INFRO FILE LINE BY LINE
                
                txtStream.Close

                Dram_Mapping_Path = DramMap_FilePath & Project_PartName & "_" & Dram_PartName & "_" & "*" & ".txt"
                Dram_Mapping_File = Dir(Dram_Mapping_Path)
                theexec.Datalog.WriteComment "Dram Mapping File :" & " " & Dram_Mapping_File
                
                If Dram_Mapping_File = "" Then
                    theexec.Flow.TestLimit resultVal:=0, lowVal:=1, hiVal:=1, Tname:="Dram Mapping File Not Found"
                    Exit Function
                ElseIf (Dir() <> "") Then   'More than one matched files
                    theexec.Flow.TestLimit resultVal:=0, lowVal:=1, hiVal:=1, Tname:="Dram Mapping File has more than one match files"
                    Exit Function
                Else
                    Dram_Type_Ieda_Temp = Split(Replace(Dram_Mapping_File, ".txt", ""), "_")
                    DRAMEnableWordStr = Dram_Type_Ieda_Temp(3) & "_" & Dram_Type_Ieda_Temp(4) & "_" & Dram_Type_Ieda_Temp(5) & "_" & Dram_Type_Ieda_Temp(6) & "_" & Dram_Type_Ieda_Temp(7)
                    Dram_Ieda = Dram_Type_Ieda_Temp(2)
                    
                    'Raise Enable Word by Dram_Mapping_File name
                    If CheckDRAMEnableWordExist(DRAMEnableWordStr) Then
                         theexec.Flow.enableWord(DRAMEnableWordStr) = True
                         Tname = Dram_Type_Ieda_Temp(7) & Dram_Type_Ieda_Temp(4)
                         enableWord = DRAMEnableWordStr
                    Else
                         ' BINOUT: The enable word written in Dram_Mapping_File is not exist
                         theexec.Flow.TestLimit resultVal:=0, lowVal:=1, hiVal:=1, Tname:="Dram Mapping File Name is not match DRAM EnableWord"
                         Exit Function
                    End If

                    Tname = Tname & " " & "En_Word"
                    theexec.Flow.TestLimit resultVal:=1, lowVal:=1, hiVal:=1, Tname:=Tname

                    ReDim Ary_Dram_IEDA(theexec.sites.Existing.Count - 1) As String
                    
                    For Each site In theexec.sites.Active
                        Ary_Dram_IEDA(site) = Dram_Ieda
                    Next site
                    
                    AllSite_DRAM_IEDA = Join(Ary_Dram_IEDA(), ",")
                    AllSite_DRAM_IEDA = auto_checkIEDAString(AllSite_DRAM_IEDA)
                    Call RegKeySave("DRAM_vendor", AllSite_DRAM_IEDA)
                    theexec.Datalog.WriteComment vbCrLf & "DRAM Type IEDA: " & AllSite_DRAM_IEDA & vbCrLf
                           
                    Call BdfDataBase.ProcessHipDram(ChkStepFlag, enableWord)
                    If ChkStepFlag Then
                        theexec.Flow.TestLimit resultVal:=1, lowVal:=1, hiVal:=1, Tname:="Judge_Dram_Type"
                    Else
                        Call Print_Error_Message(Error_Info, "VBT_ZeFuseOthers", "JudgeDRAMType_T", "Can't find any dram enableword!!!")
                        theexec.Flow.TestLimit resultVal:=0, lowVal:=1, hiVal:=1, Tname:="Judge_Dram_Type"
                    End If
                End If
            End If
        ElseIf Not ((theexec.Flow.enableWord("T0TX_HOT") = True) And (theexec.Flow.enableWord("T0TX_ROOM") = True)) Then
            'step3
            SkipStep1_Flag = True '''for 2nd/3th/4th... TouchDown keep using "Step3"
            BdfDataBase.DicDramHipMap.RemoveAll
            Set opbank = GetBdfBank("CFG")
            ReDim Preserve oValue(theexec.sites.Existing.Count - 1)
            For Each enableWord In BdfDataBase.DicDramMap.Keys
                ChkStepFlag = False
                If Not (LCase(enableWord) Like "fieldname") Then
                    For Each site In theexec.sites
                        FieldChkCnt(site) = 0
                        siteVal(site) = False
                    Next site
    
                    For m_loop = 0 To BdfDataBase.DramFieldSize - 1
                        m_catename = BdfDataBase.DicDramMap.item("fieldname")(m_loop)
                        Set field = opbank.Fields(m_catename)
    
                        For Each site In theexec.sites
                            m_value(site) = CStr(BdfDataBase.DicDramMap.item(enableWord)("fuseData")(m_loop))
                            If GlbUtility.IsNumber(m_value(site)) Then
                                If field.size < 32 Then
                                    m_value(site) = GlbUtility.String2Dbl(CStr(m_value(site)))
                                    If field.DsscDecValue(site) = m_value(site) Then
                                        FieldChkCnt(site) = FieldChkCnt(site) + 1
                                    End If
                                Else
                                    m_value(site) = GlbUtility.String2Hex(CStr(m_value(site)))
                                    If GlbUtility.xHexCompare(field.DsscValue, m_value(site), ChkEqualTo) Then
                                        FieldChkCnt(site) = FieldChkCnt(site) + 1
                                    End If
                                End If
                            Else
                                theexec.Flow.TestLimit resultVal:=0, lowVal:=1, hiVal:=1, Tname:="Dram_Table fuse data format incorrect"
                                Exit Function
                            End If
                        Next site
                    Next m_loop
    
                    For Each site In theexec.sites
                        If FieldChkCnt(site) = BdfDataBase.DramFieldSize Then
                            siteVal(site) = True
                            siteResult(site) = True
                            theexec.Flow.enableWord(enableWord) = True
                            theexec.Datalog.WriteComment "Site: " & site & ", Judge Device Dram EnableWord: " & enableWord
                        End If
                    Next site
    
                    If siteVal.Any(True) Then
                        Call BdfDataBase.ProcessHipDram(ChkStepFlag, enableWord, True, siteVal)
                    End If
                End If
            Next enableWord

            If Not siteResult.All(False) Then
                For Each site In theexec.sites
                    If siteResult(site) Then
                        oValue(site) = 1
                    Else
                        oValue(site) = 0
                    End If
                Next site
                theexec.Flow.TestLimit resultVal:=oValue, lowVal:=1, hiVal:=1, Tname:="Judge_Dram_Type"
            Else
                Call Print_Error_Message(Error_Info, "VBT_ZeFuseOthers", "JudgeDRAMType_T", "Can't find any dram enableword!!!")
                theexec.Flow.TestLimit resultVal:=0, lowVal:=1, hiVal:=1, Tname:="Judge_Dram_Type"
            End If
        End If
    End If
Exit Function
skip:
    Call Print_Error_Message(Error_Info, "VBT_ZeFuseOthers", "JudgeDRAMType_T")
    theexec.Flow.TestLimit resultVal:=999, lowVal:=1, hiVal:=1, Tname:="JudgeDRAMType_T"
    Exit Function
errHandler:
    Call Print_Error_Message(Error_Info, "VBT_ZeFuseOthers", "JudgeDRAMType_T")
    If AbortTest Then Exit Function Else Resume Next
End Function

Public Function CheckDRAMEnableWordExist(DRAMEnableWordStr As String) As Boolean
On Error GoTo errHandler
Dim enableWord As Variant
    CheckDRAMEnableWordExist = False
    For Each enableWord In BdfDataBase.DicDramMap.Keys
     If enableWord = DRAMEnableWordStr Then
        CheckDRAMEnableWordExist = True
        Exit For
     End If
    Next enableWord
Exit Function
errHandler:
    Call Print_Error_Message(Error_Info, "VBT_ZeFuseOthers", "CheckDRAMEnableWordExist")
    If AbortTest Then Exit Function Else Resume Next
End Function

Public Function CheckCurrentJob(job As String) As Boolean
On Error GoTo errHandler
Dim temp() As String
Dim item As Variant
    
    If job = "" Then
        CheckCurrentJob = False
    Else
        For Each item In Split(job, ",")
            If UCase(Trim(item)) = UCase(GlbUtility.currStage) Then
                CheckCurrentJob = False
                Exit Function
            End If
        Next item
        CheckCurrentJob = True
    End If
Exit Function
errHandler:
    Call Print_Error_Message(Error_Info, "VBT_ZeFuseOthers", "CheckCurrentJob")
    If AbortTest Then Exit Function Else Resume Next
End Function

Public Function UpdateReadPatternResult(activeSite As Dictionary, patResult As SiteBoolean, ByRef result As SiteBoolean)
On Error GoTo errHandler
Dim site As Variant

    For Each site In activeSite.Keys
        If patResult(site) = False Then
            result(site) = False
        End If
    Next site
    
Exit Function
errHandler:
    Call Print_Error_Message(Error_Info, "VBT_ZeFuseOthers", "UpdateReadPatternResult")
    If AbortTest Then Exit Function Else Resume Next
End Function

