Attribute VB_Name = "VBT_Correlation_AP"
Option Explicit
Public ParseIDSCorrelationFile As Boolean
Public DicIDSCorrelationData As Dictionary
Public DicIDSCor_CP1_Delta_limit As Dictionary
Public DicIDSCor_CP2_Delta_limit As Dictionary
Public DicIDSCor_CP2_Ratio_limit_low As Dictionary
Public DicIDSCor_CP2_Ratio_limit_high As Dictionary
Public DicIDSCor_CP2_Maping As Dictionary

Public Function ParsingCorrelationIDSTableSheet()
On Error GoTo errHandler
Dim sheetName As String
Dim maxrow As Long
Dim maxcol As Long
Dim isSheetFound As Boolean
Dim m_ParsingArr() As Variant
Dim i As Long, j As Long
Dim Strtmp As String, key As String, data As String

    sheetName = "Delta_IDS_Limit"
    
    If ParseIDSCorrelationFile = False Then
        If Evaluate("ISREF('" & sheetName & "'!A1)") = True Then
            GlbUtility.Get_Sheet_Info "Delta_IDS_Limit", maxrow, maxcol, m_ParsingArr
            
            GlbUtility.IniDictionary DicIDSCorrelationData, True
            
            For i = 1 To maxrow
                data = ""
                If i = 1 Then
                    key = "header"
                Else
                    key = m_ParsingArr(i, 1)
                End If
                
                For j = 2 To maxcol
                    data = data & m_ParsingArr(i, j) & ","
                Next j
                If Not DicIDSCorrelationData.Exists(key) Then
                    GlbUtility.AddDictionary key, data, DicIDSCorrelationData
                End If
            Next i
        Else
            Call Print_Error_Message(Error_Info, "VBT_Correlation_AP", "ParsingCorrelationIDSTableSheet", "Sheet:: " + " Delta_IDS_Limit" + " doesn't exist, please check it!")
            TheExec.Flow.TestLimit resultVal:=0, lowVal:=1, hiVal:=1, Tname:="Delta_IDS_Limit"
        End If
        
        Call GetIDSCorrelationData
        ParseIDSCorrelationFile = True
    End If

    Exit Function
errHandler:
    Call Print_Error_Message(Error_Info, "VBT_Correlation_AP", "ParsingCorrelationIDSTableSheet")
    If AbortTest Then Exit Function Else Resume Next
End Function


Private Sub GetIDSCorrelationData()
On Error GoTo errHandler
Dim headerKey As String: headerKey = "header"
Dim headerStr As String
Dim headerArr() As String
Dim col_job As Long, col_cp1_delta As Long, col_cp2_delta As Long, col_cp2_ratio_l As Long, col_cp2_ratio_h As Long
Dim i As Long, j As Long
Dim key As Variant
Dim keyname As String
Dim dataStr As String
Dim dataArr() As String
Dim cp1ids As String, cp2ids As String, cp1idsArr() As String, cp2idsArr() As String
Dim cp1idstmp As String

    headerStr = GlbUtility.GetDictionary(headerKey, DicIDSCorrelationData)
    headerArr = Split(headerStr, ",")
    
    For i = 0 To UBound(headerArr)
        If LCase(headerArr(i)) Like "*stage*" Then
            col_job = i
        ElseIf LCase(headerArr(i)) Like "*cp1_delta*" Then
            col_cp1_delta = i
        ElseIf LCase(headerArr(i)) Like "*cp2/cp1_ratio_low*" Then
            col_cp2_ratio_l = i
        ElseIf LCase(headerArr(i)) Like "*cp2/cp1_ratio_high*" Then
            col_cp2_ratio_h = i
        ElseIf LCase(headerArr(i)) Like "*cp2_delta*" Then
            col_cp2_delta = i
        End If
    Next i
    
    GlbUtility.IniDictionary DicIDSCor_CP1_Delta_limit, True
    GlbUtility.IniDictionary DicIDSCor_CP2_Delta_limit, True
    GlbUtility.IniDictionary DicIDSCor_CP2_Ratio_limit_low, True
    GlbUtility.IniDictionary DicIDSCor_CP2_Ratio_limit_high, True
    GlbUtility.IniDictionary DicIDSCor_CP2_Maping, True

    cp1ids = ""
    cp2ids = ""
    For Each key In DicIDSCorrelationData
        keyname = CStr(key)
        If LCase(keyname) <> "header" Then
            dataStr = GlbUtility.GetDictionary(keyname, DicIDSCorrelationData)
            dataArr = Split(dataStr, ",")
            If LCase(dataArr(col_job)) = "cp1" Then
                cp1ids = cp1ids + keyname + ","
                If Not DicIDSCor_CP1_Delta_limit.Exists(keyname) Then
                    GlbUtility.AddDictionary keyname, dataArr(col_cp1_delta), DicIDSCor_CP1_Delta_limit
                End If
            ElseIf LCase(dataArr(col_job)) = "cp2" Then
                cp2ids = cp2ids + keyname + ","
                If Not DicIDSCor_CP2_Delta_limit.Exists(keyname) Then
                    GlbUtility.AddDictionary keyname, dataArr(col_cp2_delta), DicIDSCor_CP2_Delta_limit
                End If
                
                If Not DicIDSCor_CP2_Ratio_limit_low.Exists(keyname) Then
                    GlbUtility.AddDictionary keyname, dataArr(col_cp2_ratio_l), DicIDSCor_CP2_Ratio_limit_low
                End If
                
                If Not DicIDSCor_CP2_Ratio_limit_high.Exists(keyname) Then
                    GlbUtility.AddDictionary keyname, dataArr(col_cp2_ratio_h), DicIDSCor_CP2_Ratio_limit_high
                End If
            End If
        End If
    Next key
    
    cp1idsArr = Split(cp1ids, ",")
    cp2idsArr = Split(cp2ids, ",")
    
    
    For i = 0 To UBound(cp2idsArr)
        For j = 0 To UBound(cp1idsArr)
            cp1idstmp = LCase(cp1idsArr(j))
            cp1idstmp = Replace(cp1idstmp, "_25c", "")
            cp1idstmp = Replace(cp1idstmp, "_25", "")
            
            If cp1idsArr(j) <> "" And cp2idsArr(i) <> "" Then
                If cp2idsArr(i) Like cp1idstmp & "*" Then
                    If Not DicIDSCor_CP2_Maping.Exists(cp2idsArr(i)) Then
                        GlbUtility.AddDictionary cp2idsArr(i), cp1idsArr(j), DicIDSCor_CP2_Maping
                        Exit For
                    End If
                End If
            End If
        Next j
    Next i
    
    
Exit Sub
errHandler:
    Call Print_Error_Message(Error_Info, "VBT_Correlation_AP", "GetIDSCorrelationData")
    If AbortTest Then Exit Sub Else Resume Next
End Sub

Public Function CorrelationBankBlankChk(bank As String)
On Error GoTo errHandler
Dim bankarr As Variant
Dim opbank As New eFuseBdfBank
Dim m_siteVar As String
Dim i As Integer
Dim dicTrimmed As Dictionary
Dim FBC As New SiteLong

    bankarr = Split(bank, ",")
    
    For i = 0 To UBound(bankarr)
        Set opbank = GetBdfBank(CStr(bankarr(i)))
        
        Set dicTrimmed = opbank.DicOthers
        If dicTrimmed.Count <> 0 Then
            Call CheckBankBlank(opbank, dicTrimmed, FBC)
            TheExec.Flow.TestLimit FBC, lowVal:=1, hiVal:=opbank.FullSize, Tname:=opbank.name, PinName:="blank check", unit:=unitNone
        End If
    Next i
    
    Exit Function
    
errHandler:
    Call Print_Error_Message(Error_Info, "VBT_Correlation_AP", "CorrelationBankBlankChk")
    If AbortTest Then Exit Function Else Resume Next
End Function

Public Function CorrelationBankVarChk(bank As String)
On Error GoTo errHandler
Dim bankarr As Variant
Dim m_siteVar As String
Dim i As Integer
Dim site As Variant

    bankarr = Split(bank, ",")
    
    For i = 0 To UBound(bankarr)
        For Each site In TheExec.sites
            m_siteVar = bankarr(i) + "Chk_Var"
            If TheExec.sites(site).SiteVariableValue(m_siteVar) = 1 Then
                'Chk_Var = 1 mean flesh deid, should be binout when correlation
                TheExec.Flow.TestLimit resultVal:=0, lowVal:=1, hiVal:=1, Tname:=m_siteVar, PinName:=m_siteVar
                TheExec.sites.item(site).FlagState("F_Corr_ChkVar_Chk") = logicFalse
            Else
                'Chk_Var = -1 : no run any efuse flow, Chk_Var = 0 : Only Run Bank read, Chk_Var = 2: Already fused
                TheExec.Flow.TestLimit resultVal:=1, lowVal:=1, hiVal:=1, Tname:=m_siteVar, PinName:=m_siteVar
            End If
        Next site
    Next i
    
    Exit Function
    
errHandler:
    Call Print_Error_Message(Error_Info, "VBT_Correlation_AP", "CorrelationBankVarChk")
    If AbortTest Then Exit Function Else Resume Next
End Function

Public Function CheckBankBlank(bank As eFuseBdfBank, indicTrimmed As Dictionary, ByRef result As SiteLong)
On Error GoTo errHandler
Dim fieldStr As Variant
Dim field As New eFuseBdfField
Dim dicCared As New DSPWave
Dim inwave As New DSPWave, outwave As New DSPWave
Dim sampleSize As Long
Dim stBit As Long, spBit As Long, iBit As Long
Dim site As Variant

    dicCared.CreateConstant 0, bank.FullSize

    For Each fieldStr In bank.Fields.Keys
            Set field = bank.Fields(fieldStr)
            bank.GetBitStSp field, stBit, spBit

            If GlbUtility.currStage = field.BlowLocation Then
                If indicTrimmed Is Nothing Then
                    For iBit = stBit To spBit
                        dicCared.Element(iBit) = 1
                    Next
                ElseIf indicTrimmed.Exists(fieldStr) Then
                    If indicTrimmed(fieldStr) = bank.name Then
                          For iBit = stBit To spBit
                              dicCared.Element(iBit) = 1
                          Next
                    End If
                End If
            End If
        Next
        
        If bank.pgmMode = pgm_DAA Then
            Set inwave = bank.DaaCapWaveSerial
        Else
            Set inwave = bank.JtagCapturedSerial
        End If
        sampleSize = bank.FullSize
        outwave.CreateConstant 0, sampleSize
        For Each site In TheExec.sites
            outwave = inwave.ConvertDataTypeTo(DspDouble).Multiply(dicCared)
            result = outwave.CountElements(GreaterThan, 0)
            If result = 0 Then
                TheExec.sites.item(site).FlagState("F_Corr_Blank_Chk") = logicFalse
            End If
        Next
Exit Function
errHandler:
    Call Print_Error_Message(Error_Info, "VBT_Correlation_AP", "CheckBankBlank")
    If AbortTest Then Exit Function Else Resume Next
End Function

Public Function IDS_Correlation_Delta(bank As String)
On Error GoTo errHandler
    Dim bankarr As Variant
    Dim opbank As New eFuseBdfBank
    Dim field As New eFuseBdfField, cp1_field As New eFuseBdfField
    Dim fieldStr As Variant
    Dim i As Long, j As Long, k As Long
    Dim IDS_from_Efuse As New SiteDouble
    Dim IDS_from_DCVS As New SiteDouble
    Dim IDS_Delta As New SiteDouble
    Dim IDS_Ratio As New SiteDouble
    Dim IDS_PwrName As String
    Dim HiLimit_IDS_Delta As Double
    Dim HiLimit_IDS_Ratio As Double
    Dim LoLimit_IDS_Delta As Double
    Dim LoLimit_IDS_Ratio As Double
    Dim site As Variant
    
    bankarr = Split(bank, ",")
    
    For i = 0 To UBound(bankarr)
        Set opbank = GetBdfBank(CStr(bankarr(i)))
        
        For Each fieldStr In opbank.Fields
            Set field = opbank.Fields(fieldStr)
            If field.Algorithm = alg_ids Then
                If UCase(GlbUtility.currStage) = "CP1" And UCase(field.BlowLocation) = "CP1" Then
                    For Each site In TheExec.sites
                        If field.DsscDecValue <> 0 And field.TrimAteDecValue <> 0 Then 'it means fused
                            IDS_from_Efuse = (GlbUtility.Hex2Dbl(field.DsscValue) * field.Resolution * 0.001)
                            IDS_from_DCVS = (GlbUtility.Hex2Dbl(field.TrimAteValue) * field.Resolution * 0.001)
                            IDS_Delta = IDS_from_DCVS.Subtract(IDS_from_Efuse)
                            IDS_Delta = IDS_Delta.Abs
                            If DicIDSCor_CP1_Delta_limit(fieldStr) = "" Or LCase(DicIDSCor_CP1_Delta_limit(fieldStr)) Like "*need*update*" Then
                                TheExec.Flow.TestLimit IDS_Delta, Tname:=fieldStr & " Delta", PinName:=fieldStr & " Delta", unit:=unitAmp, scaletype:=scaleMilli
                            Else
                                HiLimit_IDS_Delta = CDbl(DicIDSCor_CP1_Delta_limit(fieldStr))
                                TheExec.Flow.TestLimit IDS_Delta, lowVal:=0, hiVal:=HiLimit_IDS_Delta, Tname:=fieldStr & " Delta", PinName:=fieldStr & " Delta", unit:=unitAmp, scaletype:=scaleMilli
                            End If
                            TheExec.Datalog.WriteComment "site:" & CStr(site) & ", " & fieldStr & " from Efuse is " & CStr(IDS_from_Efuse * 1000) & "mA"
                            TheExec.Datalog.WriteComment "site:" & CStr(site) & ", " & fieldStr & " from DVCS is " & CStr(IDS_from_DCVS * 1000) & "mA"
                        Else
                            TheExec.Flow.TestLimit resultVal:=0, lowVal:=1, hiVal:=1, Tname:=fieldStr, PinName:=fieldStr
                            Call Print_Error_Message(Error_Info, "VBT_Correlation_AP", "IDS_Correlation_Delta", "site:" & CStr(site) & ", " & fieldStr & _
                                        ", ReadBuffer: " & CStr(field.DsscDecValue) & ", WriteBuffer: " & CStr(field.TrimAteDecValue) & ", have zero data")
                        End If
                    Next site
                ElseIf UCase(GlbUtility.currStage) = "CP2" And UCase(field.BlowLocation) = "CP2" Then
                    If DicIDSCor_CP2_Maping.Exists(fieldStr) Then
                        Set cp1_field = opbank.Fields(DicIDSCor_CP2_Maping(fieldStr))
                        
                        For Each site In TheExec.sites
                            If cp1_field.DsscDecValue <> 0 And field.DsscDecValue <> 0 And field.TrimAteDecValue <> 0 Then 'it means fused
                                IDS_from_Efuse = (GlbUtility.Hex2Dbl(field.DsscValue) * field.Resolution * 0.001)
                                IDS_from_DCVS = (GlbUtility.Hex2Dbl(field.TrimAteValue) * field.Resolution * 0.001)
                                IDS_Delta = IDS_from_DCVS.Subtract(IDS_from_Efuse)
                                IDS_Delta = IDS_Delta.Abs
                                If DicIDSCor_CP2_Delta_limit(fieldStr) = "" Or LCase(DicIDSCor_CP2_Delta_limit(fieldStr)) Like "*need*update*" Then
                                    TheExec.Flow.TestLimit IDS_Delta, Tname:=fieldStr & " Delta", PinName:=fieldStr & " Delta", unit:=unitAmp, scaletype:=scaleMilli
                                Else
                                    HiLimit_IDS_Delta = CDbl(DicIDSCor_CP2_Delta_limit(fieldStr))
                                    TheExec.Flow.TestLimit IDS_Delta, lowVal:=0, hiVal:=HiLimit_IDS_Delta, Tname:=fieldStr & " Delta", PinName:=fieldStr & " Delta", unit:=unitAmp, scaletype:=scaleMilli
                                End If
                                
                                IDS_from_Efuse = (GlbUtility.Hex2Dbl(cp1_field.DsscValue) * cp1_field.Resolution * 0.001)
                                IDS_Ratio = IDS_from_DCVS.divide(IDS_from_Efuse)
                                If (DicIDSCor_CP2_Ratio_limit_high(fieldStr) = "" Or LCase(DicIDSCor_CP2_Ratio_limit_high(fieldStr)) Like "*need*update*") _
                                    And (DicIDSCor_CP2_Ratio_limit_low(fieldStr) = "" Or LCase(DicIDSCor_CP2_Ratio_limit_low(fieldStr)) Like "*need*update*") Then
                                    TheExec.Flow.TestLimit IDS_Ratio, Tname:=fieldStr & " Ratio", PinName:=fieldStr & " Ratio", unit:=unitNone, scaletype:=scaleNoScaling
                                ElseIf (DicIDSCor_CP2_Ratio_limit_high(fieldStr) <> "" And Not LCase(DicIDSCor_CP2_Ratio_limit_high(fieldStr)) Like "*need*update*") _
                                    And (DicIDSCor_CP2_Ratio_limit_low(fieldStr) = "" Or LCase(DicIDSCor_CP2_Ratio_limit_low(fieldStr)) Like "*need*update*") Then
                                    HiLimit_IDS_Ratio = CDbl(DicIDSCor_CP2_Ratio_limit_high(fieldStr))
                                    TheExec.Flow.TestLimit IDS_Ratio, hiVal:=HiLimit_IDS_Ratio, Tname:=fieldStr & " Ratio", PinName:=fieldStr & " Ratio", unit:=unitNone, scaletype:=scaleNoScaling
                                ElseIf (DicIDSCor_CP2_Ratio_limit_high(fieldStr) = "" Or LCase(DicIDSCor_CP2_Ratio_limit_high(fieldStr)) Like "*need*update*") _
                                    And (DicIDSCor_CP2_Ratio_limit_low(fieldStr) <> "" And Not LCase(DicIDSCor_CP2_Ratio_limit_low(fieldStr)) Like "*need*update*") Then
                                    LoLimit_IDS_Ratio = CDbl(DicIDSCor_CP2_Ratio_limit_low(fieldStr))
                                    TheExec.Flow.TestLimit IDS_Ratio, lowVal:=LoLimit_IDS_Ratio, Tname:=fieldStr & " Ratio", PinName:=fieldStr & " Ratio", unit:=unitNone, scaletype:=scaleNoScaling
                                Else
                                    HiLimit_IDS_Ratio = CDbl(DicIDSCor_CP2_Ratio_limit_high(fieldStr))
                                    LoLimit_IDS_Ratio = CDbl(DicIDSCor_CP2_Ratio_limit_low(fieldStr))
                                    TheExec.Flow.TestLimit IDS_Ratio, lowVal:=LoLimit_IDS_Ratio, hiVal:=HiLimit_IDS_Ratio, Tname:=fieldStr & " Ratio", PinName:=fieldStr & " Ratio", unit:=unitNone, scaletype:=scaleNoScaling
                                End If
                                TheExec.Datalog.WriteComment "site:" & CStr(site) & ", " & cp1_field.name & " from Efuse is " & CStr(IDS_from_Efuse * 1000) & "mA"
                                TheExec.Datalog.WriteComment "site:" & CStr(site) & ", " & fieldStr & " from DVCS is " & CStr(IDS_from_DCVS * 1000) & "mA"
                            Else
                                TheExec.Flow.TestLimit resultVal:=0, lowVal:=1, hiVal:=1, Tname:=fieldStr, PinName:=fieldStr
                                Call Print_Error_Message(Error_Info, "VBT_Correlation_AP", "IDS_Correlation_Delta", "site:" & CStr(site) & ", " & fieldStr & _
                                        ", ReadBuffer: " & CStr(field.DsscDecValue) & ", WriteBuffer: " & CStr(field.TrimAteDecValue) & _
                                        ", " & cp1_field.name & ", ReadBuffer: " & CStr(cp1_field.DsscDecValue))
                            End If
                        Next site
                    Else
                        Call Print_Error_Message(Error_Info, "VBT_Correlation_AP", "IDS_Correlation_Delta", "field:" & fieldStr & ", didn't have mapping cp1 ids field, please check it!!! ")
                    End If
                End If
            End If
        Next
    Next i
    

    Exit Function
    
errHandler:
    Call Print_Error_Message(Error_Info, "VBT_Correlation_AP", "IDS_Correlation_Delta")
    If AbortTest Then Exit Function Else Resume Next
End Function

Public Sub SwitchPseudoFuseEnableFlag()
On Error GoTo errHandler
   
    If PseudoFuseEnable = True Then
        PseudoFuseEnable = False
    Else
        PseudoFuseEnable = True
    End If
   
Exit Sub

errHandler:
    Call Print_Error_Message(Error_Info, "VBT_Correlation_AP", "SwitchPseudoFuseEnableFlag")
    If AbortTest Then Exit Sub Else Resume Next
End Sub


'Public Function ParsingCorrelationIDSTable(FilePath As String)
'    Dim m_FileTmpName As String
'    Dim m_FileType As String: m_FileType = "csv"
'    Dim m_FileName As String, filenametmp As String
'    Dim m_file As String
'    Dim m_site As Variant
'    Dim bankstr As Variant
'    Dim bank As New eFuseBdfBank
'    Dim keyname As String
'    Dim DataExist As New SiteLong
'
'On Error GoTo errHandler
'
'    If (ParseIDSCorrelationFile = False) Then
'        GlbUtility.IniDictionary DicIDSCorrelationData, True
'
'        filenametmp = "Delta_IDS_Limit"
'        If Not (CheckFileExist(FilePath, filenametmp, m_FileType, m_FileName, m_file)) Then
'            TheExec.Datalog.WriteComment "<Error> File:: " + " Delta_IDS_Limit.csv " + " doesn't exist, please check it!"
'            'GoTo errHandler
'            TheExec.Flow.TestLimit resultVal:=0, lowVal:=1, hiVal:=1, Tname:="Delta_IDS_Limit"
'            Exit Function
'        End If
'        Open m_file For Input As #1
'            Do Until EOF(1)
'                GetFileRowData DicIDSCorrelationData
'            Loop
'        Close #1
'
'        DataExist = 1   '1:Pass, 0:Fail
'
'        Call GetIDSCorrelationData
'        TheExec.Flow.TestLimit resultVal:=DataExist, lowVal:=1, hiVal:=1, Tname:="PesudoFileExist"
'
'        ParseIDSCorrelationFile = True
'    End If
'    Exit Function
'
'errHandler:
'    TheExec.Datalog.WriteComment "error in ParsingCorrelationIDSTable"
'    If AbortTest Then Exit Function Else Resume Next
'End Function
'
'
'Private Sub GetFileRowData(dic As Dictionary)
'On Error GoTo errHandler
'Dim funcName As String: funcName = "GetFileRowData"
'
'    Dim m_lineStr As String
'    Dim key As String
'    Dim addr As Variant
'    Dim data As Variant
'    Dim bitdef As Variant
'
'    Line Input #1, m_lineStr
'    m_lineStr = Trim(m_lineStr)
'    addr = InStr(1, m_lineStr, ",")
'    key = mid(m_lineStr, 1, addr - 1)
'    If key = "" Then key = "header"
'    data = mid(m_lineStr, addr + 1)
'
'    If Not dic.Exists(key) Then
'        GlbUtility.AddDictionary key, data, dic
'    End If
'Exit Sub
'errHandler:
'    GlbUtility.WriteDlg "<Error> " + funcName + ":: please check it out."
'    If AbortTest Then Exit Sub Else Resume Next
'End Sub
