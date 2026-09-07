Attribute VB_Name = "LIB_Common_AP"
#Const isUFP = True
Option Explicit
Function IEDA_GetString(ByRef Inputstr As String, RegistryName As String)
On Error GoTo errHandler 'Add ErrHandler 2023/08/18
'(ByRef InputStr As String, FuseCategory As String, CategoryIndex As Integer)
    Dim funcName As String:: funcName = "IEDA_GetString"
    Dim site As Variant
    Dim TmpString As String
    Dim field As eFuseBdfField ''20210406 Modify for Obj Efuse

    For Each site In THEEXEC.sites
        Select Case RegistryName
        'ECID IEDA
            Case "eFuseLotNumber"
                TmpString = BdfDataBase.Bank_Ecid.DsscLotStr    '20210406 Modify for new Efuse
                'TmpString = ECIDFuse.category(ECIDIndex("Lot_ID")).Read.ValStr(site)
            Case "eFuseWaferID"
                TmpString = BdfDataBase.Bank_Ecid.DsscWfrStr    '20210406 Modify for new Efuse
                'TmpString = ECIDFuse.category(ECIDIndex("Wafer_ID")).Read.ValStr(site)
            Case "eFuseDieX"
                TmpString = BdfDataBase.Bank_Ecid.DsscXcoor '20210406 Modify for new Efuse
                'TmpString = ECIDFuse.category(ECIDIndex("X_Coordinate")).Read.ValStr(site)
            Case "eFuseDieY"
                TmpString = BdfDataBase.Bank_Ecid.DsscYcoor '20210406 Modify for new Efuse
                'TmpString = ECIDFuse.category(ECIDIndex("Y_Coordinate")).Read.ValStr(site)
            Case "Hram_ECID_53bit"
                TmpString = BdfDataBase.Bank_Ecid.EcidBinStr    '20210406 Modify for new Efuse
                'TmpString = ECIDFuse.category(gI_Index_DEID).Read.BitStrL(site)

''            If CategoryIndex = gI_Index_53bits Then
''                TmpString = ECIDFuse.Category(CategoryIndex).Read.BitStrL(Site)
''            Else
''                TmpString = ECIDFuse.Category(CategoryIndex).Read.ValStr(Site)
''            End If
            'UID IEDA
            Case "Prov_Code"
                Call IEDA_UID_Decode
                '20210406 Modify for new Efuse
                Set field = BdfDataBase.Bank_Cfg.Fields(LCase("UID_Code"))
                If CStr(field.DsscDecValue) = "" Then
                    TmpString = vbNullString  'site not enable
                ElseIf field.DsscDecValue = 0 Then
                    TmpString = "0"
                ElseIf (field.Llimit < field.DsscDecValue) And (field.DsscDecValue < field.Hlimit) Then
                    TmpString = "1"
                End If
'                If UIDFuse.category(UIDIndex("UID_Code")).Read.ValStr(site) = "" Then
'                    TmpString = ""  'site not enable
'                ElseIf CDbl(UIDFuse.category(UIDIndex("UID_Code")).Read.ValStr(site)) = 0 Then
'                    TmpString = "0"
'                ElseIf CDbl(UIDFuse.category(UIDIndex("UID_Code")).LoLMT) < CDbl(UIDFuse.category(UIDIndex("UID_Code")).Read.ValStr(site)) And CDbl(UIDFuse.category(UIDIndex("UID_Code")).Read.ValStr(site)) < CDbl(UIDFuse.category(UIDIndex("UID_Code")).HiLMT) Then
'                    TmpString = "1"
'                End If
            'CFG IEDA
            Case "SVM_CFuse"
                TmpString = BdfDataBase.Bank_Cfg.DsscCFGCondStr '20210406 Modify for new Efuse
            'TmpString = CFGFuse.category(gI_CFG_firstbits_index).Read.BitStrM(site)
''            If CategoryIndex = gI_CFG_firstbits_index Then
''                TmpString = CFGFuse.Category(CategoryIndex).Read.BitStrM(Site)
''            Else
''                TmpString = CFGFuse.Category(CategoryIndex).Read.ValStr(Site)
''            End If
            Case "TMPS1_Untrim"
                TmpString = gS_TMPS1_Untrim(site)
            Case "TMPS2_Untrim"
                TmpString = gS_TMPS2_Untrim(site)
            Case "TMPS3_Untrim"
                TmpString = gS_TMPS3_Untrim(site)
            Case "TMPS4_Untrim"
                TmpString = gS_TMPS4_Untrim(site)
            Case "TMPS5_Untrim"
                TmpString = gS_TMPS5_Untrim(site)
            Case "TMPS6_Untrim"
                TmpString = gS_TMPS6_Untrim(site)
            Case "TMPS7_Untrim"
                TmpString = gS_TMPS7_Untrim(site)
            Case "TMPS8_Untrim"
                TmpString = gS_TMPS8_Untrim(site)
            Case "TMPS9_Untrim"
                TmpString = gS_TMPS9_Untrim(site)
            Case "TMPS10_Untrim"
                TmpString = gS_TMPS10_Untrim(site)
            Case "TMPS11_Untrim"
                TmpString = gS_TMPS11_Untrim(site)
            Case "TMPS12_Untrim"
                TmpString = gS_TMPS12_Untrim(site)
            Case "TMPS13_Untrim"
                TmpString = gS_TMPS13_Untrim(site)
            Case "TMPS14_Untrim"
                TmpString = gS_TMPS14_Untrim(site)
            
            
            Case "TMPS1_Trim"
                TmpString = gS_TMPS1_Trim(site)
            Case "TMPS2_Trim"
                TmpString = gS_TMPS2_Trim(site)
            Case "TMPS3_Trim"
                TmpString = gS_TMPS3_Trim(site)
            Case "TMPS4_Trim"
                TmpString = gS_TMPS4_Trim(site)
            Case "TMPS5_Trim"
                TmpString = gS_TMPS5_Trim(site)
            Case "TMPS6_Trim"
                TmpString = gS_TMPS6_Trim(site)
            Case "TMPS7_Trim"
                TmpString = gS_TMPS7_Trim(site)
            Case "TMPS8_Trim"
                TmpString = gS_TMPS8_Trim(site)
            Case "TMPS9_Trim"
                TmpString = gS_TMPS9_Trim(site)
            Case "TMPS10_Trim"
                TmpString = gS_TMPS10_Trim(site)
            Case "TMPS11_Trim"
                TmpString = gS_TMPS11_Trim(site)
            Case "TMPS12_Trim"
                TmpString = gS_TMPS12_Trim(site)
            Case "TMPS13_Trim"
                TmpString = gS_TMPS13_Trim(site)
            Case "TMPS14_Trim"
                TmpString = gS_TMPS14_Trim(site)
            
            
            Case "TMPS1"
                TmpString = gS_TMPS1(site)
            Case "TMPS2"
                TmpString = gS_TMPS2(site)
            Case "TMPS3"
                TmpString = gS_TMPS3(site)
            Case "TMPS4"
                TmpString = gS_TMPS4(site)
            Case "TMPS5"
                TmpString = gS_TMPS5(site)
            Case "TMPS6"
                TmpString = gS_TMPS6(site)
            Case "TMPS7"
                TmpString = gS_TMPS7(site)
            Case "TMPS8"
                TmpString = gS_TMPS8(site)
            Case "TMPS9"
                TmpString = gS_TMPS9(site)
            Case "TMPS10"
                TmpString = gS_TMPS10(site)
            Case "TMPS11"
                TmpString = gS_TMPS11(site)
            Case "TMPS12"
                TmpString = gS_TMPS12(site)
            Case "TMPS13"
                TmpString = gS_TMPS13(site)
            Case "TMPS14"
                TmpString = gS_TMPS14(site)
            Case "BKM"
                '20210406 Modify for new Efuse
                Set field = BdfDataBase.Bank_Cfg.Fields(LCase("bkm_package"))
                TmpString = field.TrimAteDecValue
                'TmpString = CFGFuse.category(CFGIndex("bkm_package")).Write.Decimal(site)
            Case "BKM_Fuse"
                '20210406 Modify for new Efuse
                Set field = BdfDataBase.Bank_Cfg.Fields(LCase("bkm_package"))
                TmpString = field.DsscDecValue
                'TmpString = CFGFuse.category(CFGIndex("bkm_package")).Read.Decimal(site)
            
''            Case "UDR"
''                TmpString = UDRFuse.Category(CategoryIndex).Read.ValStr(Site)
''            Case "SEN"
''                TmpString = SENFuse.Category(CategoryIndex).Read.ValStr(Site)
            
            Case Else
                Call Print_Error_Message(Warning_Info, "LIB_Common_AP", "IEDA_GetString", "print: warnining, no suitable registry choosed in VBT 'IEDA_GetString'.") 'Add ErrHandler 2023/08/18
        End Select
        If (site = THEEXEC.sites.Existing.Count - 1) Then
            Inputstr = Inputstr + TmpString
        Else
            Inputstr = Inputstr + TmpString + ","
        End If

    Next site

Exit Function 'Add ErrHandler 2023/08/18
errHandler: 'Add ErrHandler 2023/08/18
    Call Print_Error_Message(Error_Info, "LIB_Common_AP", "IEDA_GetString") 'Add ErrHandler 2023/08/18
    If AbortTest Then Exit Function Else Resume Next 'Add ErrHandler 2023/08/18
End Function


Function IEDA_AutoCheck_Print(ByRef Inputstr As String, RegistryName As String, DebugPrint As Boolean)
On Error GoTo errHandler 'Add ErrHandler 2023/08/18

    Dim funcName As String:: funcName = "IEDA_AutoCheck_Print"
    Dim TmpString As String

    Inputstr = auto_checkIEDAString(Inputstr)
    If DebugPrint Then THEEXEC.Datalog.WriteComment "print: Set IEDA registry ( " & RegistryName & " ) = " & Inputstr

Exit Function 'Add ErrHandler 2023/08/18
errHandler: 'Add ErrHandler 2023/08/18
    Call Print_Error_Message(Error_Info, "LIB_Common_AP", "IEDA_AutoCheck_Print") 'Add ErrHandler 2023/08/18
    If AbortTest Then Exit Function Else Resume Next 'Add ErrHandler 2023/08/18
End Function


Public Function IEDA_UID_Decode(Optional InitPinsHi As PinList, Optional InitPinsLo As PinList, Optional InitPinsHiZ As PinList)
On Error GoTo errHandler 'Add ErrHandler 2023/08/18
'''
'''On Error GoTo errHandler
'''    Dim funcName As String:: funcName = "IEDA_UID_Decode"
'''
'''    Dim site As Variant
'''    Dim SingleBitArray() As Long, SingleBitSum As Long
'''    Dim DoubleBitArray() As Long, DoubleBitSum As Long
'''    Dim m_siteVar As String
'''    m_siteVar = "UIDChk_Var"
'''
'''    For Each site In TheExec.sites
'''        '''''initialize per Site
'''        ReDim SingleBitArray(UIDTotalBits - 1)
'''        ReDim DoubleBitArray(UIDBitsPerBlockUsed - 1)
'''        SingleBitSum = 0
'''        DoubleBitSum = 0
'''
'''            Call auto_OR_2Blocks("UID", gS_SingleStrArray, SingleBitArray, DoubleBitArray)  ''''to get gL_UID_FBC()
'''
'''            If (DisplayUID = True) Then
'''                TheExec.Datalog.WriteComment ""
'''                TheExec.Datalog.WriteComment "Read AES/UID data from DSSC at Site (" + CStr(site) + ")"
'''                Call auto_PrintAllBitbyDSSC(SingleBitArray, UIDReadCycle, UIDTotalBits, UIDReadBitWidth)
'''            End If
'''
'''            Call auto_Decode_UIDBinary_Data(DoubleBitArray)  ''''20150616 New
'''
'''    Next site
'''
'''    'TheExec.Flow.TestLimit resultVal:=gL_AES_FBC, lowVal:=0, hiVal:=0, Tname:="FailBitCount"
'''
'''Exit Function
'''
'''errHandler:
'''    TheExec.Datalog.WriteComment "<Error> " + funcName + ":: please check it out."
'''    If AbortTest Then Exit Function Else Resume Next
'''
Exit Function 'Add ErrHandler 2023/08/18
errHandler: 'Add ErrHandler 2023/08/18
    Call Print_Error_Message(Error_Info, "LIB_Common_AP", "IEDA_UID_Decode") 'Add ErrHandler 2023/08/18
    If AbortTest Then Exit Function Else Resume Next 'Add ErrHandler 2023/08/18
End Function

Public Function CheckHaveBinCutSheet(nBinCutNumber As Long, Optional isUse_Product_Identifier As Boolean = False) As Long
On Error GoTo errHandler 'Add ErrHandler 2023/08/18
    'if have no BinCut sheet can reference
    Dim passBinCut As Variant
    Dim nTempBinCutNum As Long
    
    CheckHaveBinCutSheet = -999   '-999 = wrong bincut number
    
    If isUse_Product_Identifier And UCase(currentJobName) Like "*CP1*" Then
        CheckHaveBinCutSheet = 999
        Exit Function
    End If
    
    For Each passBinCut In PassBinCut_ary
        If isUse_Product_Identifier Then
            If nBinCutNumber = passBinCut Then     '1 = Bin1, 2 = BinX, 3 = BinY
                CheckHaveBinCutSheet = nBinCutNumber
                Exit For
            End If
        Else
            If nBinCutNumber = passBinCut Or nBinCutNumber = 999 Then     '1 = Bin1, 2 = BinX, 3 = BinY, 999 = Max(default)
                CheckHaveBinCutSheet = nBinCutNumber
                Exit For
            End If
        End If
    Next passBinCut
    
Exit Function 'Add ErrHandler 2023/08/18
errHandler: 'Add ErrHandler 2023/08/18
    Call Print_Error_Message(Error_Info, "LIB_Common_AP", "CheckHaveBinCutSheet") 'Add ErrHandler 2023/08/18
    If AbortTest Then Exit Function Else Resume Next 'Add ErrHandler 2023/08/18
End Function

' [20231003][All][Tank] modify after Chihome review
Public Function Compare_ForceVal_BV(pin As String, val As String, Optional HLV As Boolean = True, Optional isUse_Product_Identifier As Boolean = False, Optional nBincutNum As Long = 999) As Variant
On Error GoTo errHandler 'Add ErrHandler 2023/08/18
    
    
    Dim funcName As String:: funcName = "Compare_ForceVal_BV"
    
    Dim i As Long
    Dim Pin_Cnt As Long
    
    Dim Hilimit_IDS As Double
    Dim Lolimit_IDS As Double
    Dim powerDomain As String '''pinGroup
    Dim powerPin As String '''pin
    Dim Pin_Ary() As String
    Dim highest_Pmode_Name As String
        
'    TheExec.DataManager.DecomposePinList Pin, Pin_ary, Pin_cnt
    
    If val = "" Then                        'Use BinCut limit need to set blank in flow
        Hilimit_IDS = 0
        Lolimit_IDS = 0
        If domain2pinDict.Exists(UCase(Trim(pin))) Then
            powerPin = VddbinDomain2Pin(Get1stPinFromPingroup(UCase(Trim(pin))))
            powerDomain = UCase(Trim(pin))
        ElseIf pin2domainDict.Exists(UCase(Trim(pin))) Then
            powerPin = UCase(Trim(pin))
            powerDomain = VddbinPin2Domain(UCase(Trim(pin)))
        Else
            powerPin = UCase(Trim(pin))
            powerDomain = vbNullString
        End If
        
        If (UCase(currentJobName) Like "*CP1*") And isUse_Product_Identifier Then
            nBincutNum = 999
        End If

        If (powerDomain <> "") And (nBincutNum <> -999) Then
            Call Get_BV_Limit(Hilimit_IDS, Lolimit_IDS, FullCorePowerinFlowSheet, powerDomain, nBincutNum)
        Else
            Hilimit_IDS = -999
            Lolimit_IDS = -999
            THEEXEC.Datalog.WriteComment "Get wrong Product Identifier!!!"
        End If
    ElseIf val <> "" And val <> -999 Then   'Use flow limit
        Hilimit_IDS = CDbl(val)
        Lolimit_IDS = CDbl(val)
    ElseIf val = -999 Then                  'Use -999 will don't compare(w/o limit)
        Hilimit_IDS = 0
        Lolimit_IDS = 0
    End If
    
    If HLV = True Then
        Compare_ForceVal_BV = Hilimit_IDS
    Else
        Compare_ForceVal_BV = Lolimit_IDS
    End If
    
Exit Function 'Add ErrHandler 2023/08/18
errHandler: 'Add ErrHandler 2023/08/18
    Call Print_Error_Message(Error_Info, "LIB_Common_AP", "Compare_ForceVal_BV") 'Add ErrHandler 2023/08/18
    If AbortTest Then Exit Function Else Resume Next 'Add ErrHandler 2023/08/18
End Function

Public Function DTS_GetStoredData_Compare()
On Error GoTo errHandler 'Add ErrHandler 2023/08/18
    Dim funcName As String:: funcName = "DTS_GetStoredData_Compare"
    '20200408 Oliver
    Dim check_id As String, tmpLotStr() As String, tmpWaferStr() As String, tmpX_Coor_Str() As String, tmpY_Coor_Str() As String
    Dim MyArray() As String, DTSBarCode() As String, tmp_check_id As String, SiteCount As Integer
    
    Dim lotStr As String
    Dim Waferstr As String
    Dim X_Coor_Str As String
    Dim Y_Coor_Str As String
    Dim site As Variant 'Carter, 20240304
    lotStr = vbNullString 'reset id
    Waferstr = vbNullString
    X_Coor_Str = vbNullString
    Y_Coor_Str = vbNullString
    tmp_check_id = vbNullString
    
'    '20210406 Modify for new Efuse
'    For Each site In TheExec.sites
'        If (site = TheExec.sites.Existing.Count - 1) Then
'            lotStr = lotStr + BdfDataBase.Bank_Ecid.DsscLotStr
'            Waferstr = Waferstr + BdfDataBase.Bank_Ecid.DsscWfrStr
'            X_Coor_Str = X_Coor_Str + BdfDataBase.Bank_Ecid.DsscXcoor
'            Y_Coor_Str = Y_Coor_Str + BdfDataBase.Bank_Ecid.DsscYcoor
''            lotStr = lotStr + CStr(ECIDFuse.category(ECIDIndex("Lot_ID")).Read.ValStr(site))
''            Waferstr = Waferstr + CStr(ECIDFuse.category(ECIDIndex("Wafer_ID")).Read.ValStr(site))
''            X_Coor_Str = X_Coor_Str + CStr(ECIDFuse.category(ECIDIndex("X_Coordinate")).Read.ValStr(site))
''            Y_Coor_Str = Y_Coor_Str + CStr(ECIDFuse.category(ECIDIndex("Y_Coordinate")).Read.ValStr(site))
'        Else
'            lotStr = lotStr + BdfDataBase.Bank_Ecid.DsscLotStr + ","
'            Waferstr = Waferstr + BdfDataBase.Bank_Ecid.DsscWfrStr + ","
'            X_Coor_Str = X_Coor_Str + BdfDataBase.Bank_Ecid.DsscXcoor + ","
'            Y_Coor_Str = Y_Coor_Str + BdfDataBase.Bank_Ecid.DsscYcoor + ","
''            lotStr = lotStr + CStr(ECIDFuse.category(ECIDIndex("Lot_ID")).Read.ValStr(site)) + ","
''            Waferstr = Waferstr + CStr(ECIDFuse.category(ECIDIndex("Wafer_ID")).Read.ValStr(site)) + ","
''            X_Coor_Str = X_Coor_Str + CStr(ECIDFuse.category(ECIDIndex("X_Coordinate")).Read.ValStr(site)) + ","
''            Y_Coor_Str = Y_Coor_Str + CStr(ECIDFuse.category(ECIDIndex("Y_Coordinate")).Read.ValStr(site)) + ","
'        End If
'    Next site
    
    '20240708 michael fixed for disable site issue======
    
    ReDim tmpLotStr(THEEXEC.sites.Existing.Count - 1)
    ReDim tmpWaferStr(THEEXEC.sites.Existing.Count - 1)
    ReDim tmpX_Coor_Str(THEEXEC.sites.Existing.Count - 1)
    ReDim tmpY_Coor_Str(THEEXEC.sites.Existing.Count - 1)
    
    For Each site In THEEXEC.sites
        tmpLotStr(site) = BdfDataBase.Bank_Ecid.DsscLotStr
        tmpWaferStr(site) = BdfDataBase.Bank_Ecid.DsscWfrStr
        tmpX_Coor_Str(site) = BdfDataBase.Bank_Ecid.DsscXcoor
        tmpY_Coor_Str(site) = BdfDataBase.Bank_Ecid.DsscYcoor
    Next site
    
    lotStr = Join(tmpLotStr, ",")
    Waferstr = Join(tmpWaferStr, ",")
    X_Coor_Str = Join(tmpX_Coor_Str, ",")
    Y_Coor_Str = Join(tmpY_Coor_Str, ",")
    
    lotStr = auto_checkIEDAString(lotStr) 'get current id data from IEDA
    Waferstr = auto_checkIEDAString(Waferstr)
    X_Coor_Str = auto_checkIEDAString(X_Coor_Str)
    Y_Coor_Str = auto_checkIEDAString(Y_Coor_Str)
    
    '20240708 michael fixed for disable site issue======
    
    
'    tmpLotStr = Split(lotStr, ",")
'    tmpWaferStr = Split(Waferstr, ",")
'    tmpX_Coor_Str = Split(X_Coor_Str, ",")
'    tmpY_Coor_Str = Split(Y_Coor_Str, ",")
    
    THEEXEC.Datalog.WriteComment ("******************************")
    THEEXEC.Datalog.WriteComment ("*  Print out OCR data start  *")
    THEEXEC.Datalog.WriteComment ("******************************")
    For Each site In THEEXEC.sites
        SiteCount = THEEXEC.sites.siteNumber(site)
        If Len(tmpWaferStr(site)) < 2 Then                  'wafer id change format
            tmpWaferStr(site) = "0" + tmpWaferStr(site)
        Else
            tmpWaferStr(site) = tmpWaferStr(site)
        End If
        If Len(tmpX_Coor_Str(site)) < 2 Then                'x change format
            tmpX_Coor_Str(site) = "0" + tmpX_Coor_Str(site)
        Else
            tmpX_Coor_Str(site) = tmpX_Coor_Str(site)
        End If
        If Len(tmpY_Coor_Str(site)) < 2 Then                'y change format
            tmpY_Coor_Str(site) = "0" + tmpY_Coor_Str(site)
        Else
            tmpY_Coor_Str(site) = tmpY_Coor_Str(site)
        End If
        check_id = tmpLotStr(site) + tmpWaferStr(site) + tmpX_Coor_Str(site) + tmpY_Coor_Str(site)
        
        If DictOCR.Exists(check_id) Then
            THEEXEC.Datalog.WriteComment ("<@OCR_Data=" & site & "|" & DictOCR.item(check_id) & ">")

            tmp_check_id = tmp_check_id & DictOCR.item(check_id) + ","
        Else
            tmp_check_id = tmp_check_id + ","
        End If
    Next site
    THEEXEC.Datalog.WriteComment ("******************************")
    THEEXEC.Datalog.WriteComment ("*   Print out OCR data end   *")
    THEEXEC.Datalog.WriteComment ("******************************")
    tmp_check_id = left(tmp_check_id, Len(tmp_check_id) - 1)
    Call RegKeySave("DTSBarCode", tmp_check_id)
    
Exit Function 'Add ErrHandler 2023/08/18
errHandler: 'Add ErrHandler 2023/08/18
    Call Print_Error_Message(Error_Info, "LIB_Common_AP", "DTS_GetStoredData_Compare") 'Add ErrHandler 2023/08/18
    If AbortTest Then Exit Function Else Resume Next 'Add ErrHandler 2023/08/18
End Function

Public Function CloseFunction()
On Error GoTo errHandler 'Add ErrHandler 2023/08/18
    THEEXEC.Datalog.WriteComment ("******************************")
    THEEXEC.Datalog.WriteComment ("* OCR function already print *")
    THEEXEC.Datalog.WriteComment ("******************************")
Exit Function 'Add ErrHandler 2023/08/18
errHandler: 'Add ErrHandler 2023/08/18
    Call Print_Error_Message(Error_Info, "LIB_Common_AP", "CloseFunction") 'Add ErrHandler 2023/08/18
    If AbortTest Then Exit Function Else Resume Next 'Add ErrHandler 2023/08/18
End Function

Public Function IDS_ExtraLimit(Pins_Extra_Limit As String)
On Error GoTo errHandler 'Add ErrHandler 2023/08/18
    Dim JobPinSplit() As String
    Dim JobSplit() As String
    Dim CellJob As String
    Dim PinSplit() As String
    Dim pininfo() As String
    Dim JobValue As String
    Dim i, j As Integer
    
    ''---------------------------------------------------------------------------------------------------
    ''Pins_Extra_Limit ex:
    ''FT1:VDD_CPU_SRAM*1.2,VDD_DCS_DDR*1.2,VDD_DISP*1.2|FT2:VDD_CPU_SRAM*1.2,VDD_DCS_DDR*1.2,VDD_DISP*1.2
    ''---------------------------------------------------------------------------------------------------
    JobSplit = Split(Pins_Extra_Limit, "|")
    For i = 0 To UBound(JobSplit)
        JobPinSplit = Split(JobSplit(i), ":")
        CellJob = JobPinSplit(0)
        If (THEEXEC.CurrentJob Like "*" & CellJob & "*") Then
            gb_IDSLimit_Speical_Handle = True
            JobValue = JobPinSplit(1)
            PinSplit = Split(JobValue, ",")
            For j = 0 To UBound(PinSplit)
                pininfo = Split(PinSplit(j), "*")
                ''PinInfo(0) = PinName,PinInfo(1) = Value
                gDict_IDSLimit_Special_Handle.Add pininfo(0), pininfo(1)
            Next j
            Exit For
        End If
    Next i

Exit Function 'Add ErrHandler 2023/08/18
errHandler: 'Add ErrHandler 2023/08/18
    Call Print_Error_Message(Error_Info, "LIB_Common_AP", "IDS_ExtraLimit") 'Add ErrHandler 2023/08/18
    If AbortTest Then Exit Function Else Resume Next 'Add ErrHandler 2023/08/18
End Function


Public Function Limit_Special_Handle(Power_pin As String, Hilimit_IDS As Double)
On Error GoTo errHandler 'Add ErrHandler 2023/08/18

    If gDict_IDSLimit_Special_Handle.Exists(Power_pin) Then
        Hilimit_IDS = Hilimit_IDS * gDict_IDSLimit_Special_Handle.item(Power_pin)
    Else
        THEEXEC.Datalog.WriteComment "The Job: " & currentJobName & " Power_Pin: " & Power_pin & "is not defined in argement."
        THEEXEC.Datalog.WriteComment "Please Check it!!"
    End If

Exit Function 'Add ErrHandler 2023/08/18
errHandler: 'Add ErrHandler 2023/08/18
    Call Print_Error_Message(Error_Info, "LIB_Common_AP", "Limit_Special_Handle") 'Add ErrHandler 2023/08/18
    If AbortTest Then Exit Function Else Resume Next 'Add ErrHandler 2023/08/18
End Function

' [20231003][All][Tank] modify after Chihome review
Public Function Get_BV_Limit(Hilimit_IDS As Double, Lolimit_IDS As Double, FullCorePowerinFlowSheet As String, powerDomain As String, nBincutNum As Long)
    
    Dim s_highest_Pmode_Name As String
    
On Error GoTo errHandler
    
    If UCase("*," & FullCorePowerinFlowSheet & ",*") Like UCase("*," & powerDomain & ",*") Then
      
        s_highest_Pmode_Name = Get_Highest_Pmode_Name(powerDomain)
        
        Call Get_BincutSheetLimit(Hilimit_IDS, Lolimit_IDS, UCase(currentJobName), s_highest_Pmode_Name, nBincutNum)
        
    ElseIf UCase("*," & FullOtherRailinFlowSheet & ",*") Like UCase("*," & powerDomain & ",*") Then
        
        Call Get_BincutSheetLimit(Hilimit_IDS, Lolimit_IDS, UCase(currentJobName), powerDomain, nBincutNum, True)
        
    Else
''''    '                theexec.Datalog.WriteComment Pin & "is not BinCut CorePower or OtherRail. Please check DCVS_IDS_main_auto_range_and_measure.Error!!!"
''''    '                theexec.ErrorLogMessage Pin & "is not BinCut CorePower or OtherRail. Please check DCVS_IDS_main_auto_range_and_measure.Error!!!"
''''    '                theexec.Flow.TestLimit resultVal:=999, lowval:=1, hiVal:=1, Tname:=Pin & "is not BinCut CorePower or OtherRail. Please check DCVS_IDS_main_auto_range_and_measure.Error!!!"
    End If
    
    If gb_IDSLimit_Speical_Handle Then
        Limit_Special_Handle powerDomain, Hilimit_IDS
    End If
Exit Function

errHandler:
    Call Print_Error_Message(Error_Warning_Info.Error_Info, "LIB_Common_AP", "Get_BV_Limit")
    If AbortTest Then Exit Function Else Resume Next
End Function

' [20231003][All][Tank] modify after Chihome review
Public Function Get_Highest_Pmode_Name(s_powerDomain As String) As String
On Error GoTo errHandler
    '20211214,Modify for sram pmode is core power
    If dict_IsCorePowerInBinCutFlowSheet.item(s_powerDomain) = True And UCase(s_powerDomain) Like "*VDD*_SRAM*" Then
        Get_Highest_Pmode_Name = BinCut_Sram_Power_Seq(BinCut_Sram_Power_KeyMapping(s_powerDomain))(UBound(BinCut_Sram_Power_Seq(BinCut_Sram_Power_KeyMapping(s_powerDomain))))
    Else
        '''ex: "MC610" is the highest p_mode of "VDD_PCPU" in BinCut flow table "Non_Binning_Raill"
        Get_Highest_Pmode_Name = BinCut_Power_Seq(VddBinStr2Enum(s_powerDomain)).Power_Seq(UBound(BinCut_Power_Seq(VddBinStr2Enum(s_powerDomain)).Power_Seq))
    End If
Exit Function

errHandler:
    Call Print_Error_Message(Error_Warning_Info.Error_Info, "LIB_Common_AP", "Get_Highest_Pmode_Name")
    If AbortTest Then Exit Function Else Resume Next
End Function

' [20231003][All][Tank] modify after Chihome review
Public Function Get_BincutSheetLimit(hiLimit As Double, loLimit As Double, Stage As String, pmode As String, Optional bincutNum As Long = 999, Optional isOtherPin As Boolean = False)
On Error GoTo errHandler

    hiLimit = -999
    loLimit = -999
    Dim s_TempString As String
    
    If bincutNum = 999 Then
        If ((Stage Like "*CP1*") Or (Stage Like "*FT1*") Or (Stage Like "*WLFT*")) Then
            hiLimit = AllBinCut(VddBinStr2Enum(pmode)).IDS_CP_LIMIT / 1000
            loLimit = 0.01 * AllBinCut(VddBinStr2Enum(pmode)).IDS_CP_LIMIT / 1000
        ElseIf ((Stage Like "*CP2*") Or (Stage Like "*FT2*")) Then
            hiLimit = AllBinCut(VddBinStr2Enum(pmode)).IDS_FT_LIMIT / 1000
            loLimit = 0.01 * AllBinCut(VddBinStr2Enum(pmode)).IDS_FT_LIMIT / 1000
        Else
            s_TempString = "Get Wrong Stage = " & Stage
            Call Print_Error_Message(Error_Warning_Info.Warning_Info, "LIB_Common_AP", "Get_BincutSheetLimit", s_TempString)
        End If
        Exit Function
    End If
    
    If isOtherPin = False Then
        If ((Stage Like "*CP1*") Or (Stage Like "*FT1*") Or (Stage Like "*WLFT*")) Then
            hiLimit = BinCut(VddBinStr2Enum(pmode), bincutNum).IDS_CP_LIMIT(0) / 1000
            loLimit = 0.01 * BinCut(VddBinStr2Enum(pmode), bincutNum).IDS_CP_LIMIT(0) / 1000
        ElseIf ((Stage Like "*CP2*") Or (Stage Like "*FT2*")) Then
            hiLimit = BinCut(VddBinStr2Enum(pmode), bincutNum).IDS_FT_LIMIT(0) / 1000
            loLimit = 0.01 * BinCut(VddBinStr2Enum(pmode), bincutNum).IDS_FT_LIMIT(0) / 1000
        Else
            s_TempString = "Get Wrong Stage = " & Stage
            Call Print_Error_Message(Error_Warning_Info.Warning_Info, "LIB_Common_AP", "Get_BincutSheetLimit", s_TempString)
        End If
        Exit Function
    Else
        If ((Stage Like "*CP1*") Or (Stage Like "*FT1*") Or (Stage Like "*WLFT*")) Then
            hiLimit = CPIDS_Spec(VddBinStr2Enum(pmode), bincutNum) / 1000
            loLimit = 0.01 * CPIDS_Spec(VddBinStr2Enum(pmode), bincutNum) / 1000
        ElseIf ((Stage Like "*CP2*") Or (Stage Like "*FT2*")) Then
            hiLimit = FTIDS_Spec(VddBinStr2Enum(pmode), bincutNum) / 1000
            loLimit = 0.01 * FTIDS_Spec(VddBinStr2Enum(pmode), bincutNum) / 1000
        Else
            s_TempString = "Get Wrong Stage = " & Stage
            Call Print_Error_Message(Error_Warning_Info.Warning_Info, "LIB_Common_AP", "Get_BincutSheetLimit", s_TempString)
        End If
        Exit Function
    End If
Exit Function

errHandler:
    Call Print_Error_Message(Error_Warning_Info.Error_Info, "LIB_Common_AP", "Get_BincutSheetLimit")
    If AbortTest Then Exit Function Else Resume Next
End Function


' Read worksheet and saves info to ScanPatternsList
' [20231228][T-All][Tank] Read "SFCPatterns" sheet
Public Function Read_SFC_Table() As Long
On Error GoTo errHandler
    Dim n_MaxRow As Long
    Dim n_MaxColumn As Long
    Dim v_SFCPatternsSheetInfo() As Variant
    Dim n_Index As Long
    Dim sEnable_SFC_Record() As String
    Dim SFCPatternsSheet As String
    SFCPatternsSheet = "SFCPatterns"

    If glb_isParsingSFC = False Then
        If GetSheetInfo(SFCPatternsSheet, n_MaxRow, n_MaxColumn, v_SFCPatternsSheetInfo) Then
        
            ReDim ScanPatternsList(n_MaxRow - 1)
            
            For n_Index = 0 To n_MaxRow - 1
            
                sEnable_SFC_Record = Split(v_SFCPatternsSheetInfo(n_Index + 1, 1), ",")
            
                ScanPatternsList(n_Index).FailCycleCount = sEnable_SFC_Record(2) 'ScanFailPatterns.Cells(n_Index, 1).Value
            
                ScanPatternsList(n_Index).ScanPatternName = sEnable_SFC_Record(1) & "*" 'ScanFailPatterns.Cells(n_Index, 2).Value
            
                ScanPatternsList(n_Index).InstanceName = sEnable_SFC_Record(0) & "*" 'ScanFailPatterns.Cells(n_Index, 2).Value
            
            Next n_Index
            
            glb_isParsingSFC = True
            
            SFC_Show
        Else
            THEEXEC.AddOutput "Reading SFC Input ... FAILED"
    
            THEEXEC.AddOutput "        SFC       ... DISABLED"
    
            Read_SFC_Table = 0

            Exit Function
        End If
    End If
Exit Function
errHandler:
    Call Print_Error_Message(Error_Info, "LIB_Common_AP", "Read_SFC_Table") 'Add ErrHandler 2023/08/18
    If AbortTest Then Exit Function Else Resume Next 'Add ErrHandler 2023/08/18
    
End Function


' [20231228][T-All][Tank] Show "SFCPatterns" sheet log
Public Function SFC_Show()
On Error GoTo errHandler
    Dim n_Index As Long

    For n_Index = 0 To UBound(ScanPatternsList)

        THEEXEC.AddOutput " ->Instance[" & CStr(n_Index) & "]: " & ScanPatternsList(n_Index).InstanceName
        
        THEEXEC.AddOutput " ->Pattern[" & CStr(n_Index) & "]: " & ScanPatternsList(n_Index).ScanPatternName
        
        THEEXEC.AddOutput " ->CycleCount[" & CStr(n_Index) & "]: " & ScanPatternsList(n_Index).FailCycleCount

    Next n_Index
Exit Function
errHandler:
    Call Print_Error_Message(Error_Info, "LIB_Common_AP", "SFC_Show") 'Add ErrHandler 2023/08/18
    If AbortTest Then Exit Function Else Resume Next 'Add ErrHandler 2023/08/18
End Function


' [20231228][T-All][Tank] Check instance need to do SFC or not
Public Function SFC_Main(ByRef StartOfBodyF As InterposeName, ByRef PostPatF As InterposeName, ByRef PostPatFArg As String)
On Error GoTo errHandler
    ''''>>>>> SFC
    Dim s_SFC_StartOfBody As String
    Dim s_SFC_PostPat As String

    If glb_isSFC_Enabled Then

        If SFC_InstanceMatch(s_SFC_StartOfBody, s_SFC_PostPat, PostPatFArg) = True Then
            PostPatF = s_SFC_PostPat
        End If

    End If
Exit Function
errHandler:
    Call Print_Error_Message(Error_Info, "LIB_Common_AP", "SFC_Main") 'Add ErrHandler 2023/08/18
    If AbortTest Then Exit Function Else Resume Next 'Add ErrHandler 2023/08/18
End Function


' [20231228][T-All][Tank] Check instance need to do SFC or not
Public Function SFC_InstanceMatch(ByRef PrePatF As String, ByRef PostPatF As String, ByRef PostPatFArg As String) As Boolean
On Error GoTo errHandler
    SFC_InstanceMatch = False
    
    glb_SFC_Scan_Check = False

    Dim n_Index As Long

    Dim s_instance_name As String

    s_instance_name = THEEXEC.DataManager.InstanceName

    For n_Index = 0 To UBound(ScanPatternsList)

'        If ScanPatternsList(n_Index).InstanceName Like Instance_Name Then

        If s_instance_name Like ScanPatternsList(n_Index).InstanceName Then

            PostPatFArg = ScanPatternsList(n_Index).FailCycleCount & "," & ScanPatternsList(n_Index).ScanPatternName

            PostPatF = "SFC_PostPatF"

            'PrePatF = "SFC_PrePatF"

            SFC_InstanceMatch = True
            
            glb_SFC_Scan_Check = True

        End If

    Next n_Index
Exit Function
errHandler:
    Call Print_Error_Message(Error_Info, "LIB_Common_AP", "SFC_InstanceMatch") 'Add ErrHandler 2023/08/18
    If AbortTest Then Exit Function Else Resume Next 'Add ErrHandler 2023/08/18
End Function


' [20231228][T-All][Tank] When pattern test end call post pattern interpose function process SFC
Public Function SFC_PostPatF(n_pmCnt As Long, sa_pmNames() As String)
On Error GoTo errHandler

    If glb_isSFC_Enabled = False Then Exit Function
    If glb_TesterType = "UltraFLEXplus" Then
        ifcmemGetFailCycles_UP n_pmCnt, sa_pmNames
    Else
        ifcmemGetFailCycles_UF n_pmCnt, sa_pmNames
    End If
    
Exit Function
errHandler:
    Call Print_Error_Message(Error_Info, "LIB_Common_AP", "SFC_PostPatF") 'Add ErrHandler 2023/08/18
    If AbortTest Then Exit Function Else Resume Next 'Add ErrHandler 2023/08/18
End Function



' [20231228][T-All][Tank] Process SFC function
Public Function ifcmemGetFailCycles_UF(pmCnt As Long, pmNames() As String)
On Error GoTo errHandler

    Dim v_site As Variant
    Dim i As Long
    Dim s_ModName As String
    Dim s_InsName As String
    Dim PatIdx() As Long
    Dim PatNames() As String
    Dim Sbln_PatternPass As New SiteBoolean
    Dim patternNames() As String
    Dim dbl_CmemCycleDataArr() As Double
    Dim dbl_CmemVectorDataArr() As Double
    Dim SiteIndexData As New SiteVariant
    Dim SitePinData  As New SiteVariant
    Dim indexArr() As Long
    Dim pinDataArr() As Double
    Dim Failpin_temp() As String
    Dim FailedPin_Get_Index As String
    Dim FailedPins As String
    Dim s_InputData() As String
    Dim n_CycleCount As Long
    Dim FlowName As String
    
    Dim FailPinListData As New PinListData
    Dim dic_FailPinCycle As New Dictionary
    Dim j As Long
    Dim s_tempFailPin As String
    Dim s_tempFailCycle As String
    
    'Read back the central data.
    Sbln_PatternPass = TheHdw.Digital.Patgen.PatternBurstPassedPerSite
    If Sbln_PatternPass.Any(False) Then
        dbl_CmemCycleDataArr = TheHdw.Digital.CMEM.PatGenInfo(tlCMEMModCycle, -1)
        dbl_CmemVectorDataArr = TheHdw.Digital.CMEM.PatGenInfo(tlCMEMVMVectorOffset, -1)
        
        s_InputData = Split(pmNames(0), ",")
        n_CycleCount = CLng(s_InputData(0))
    End If
       
    For Each v_site In THEEXEC.sites
        If Sbln_PatternPass(v_site) = False Then
       
            'Get fail pin
            Failpin_temp = TheHdw.Digital.FailedPins(v_site)
            If UBound(Failpin_temp) >= 0 Then
                FailedPin_Get_Index = Join(Failpin_temp, ",")
                FailedPins = Join(Failpin_temp, "/")
            End If
            
          
            
            'Get fail pin index
            Call TheHdw.Digital.Pins(FailedPin_Get_Index).CMEM.StoredCycleData(SiteIndexData, SitePinData, -1, True)
            
            indexArr = SiteIndexData
            
            'Get flow ,pattern, insname
            FlowName = THEEXEC.Flow.CurrentFlowSheetName
            Call TheHdw.Digital.CMEM.PatternName(PatIdx, PatNames)
            s_ModName = PatNames(0)
            s_InsName = THEEXEC.DataManager.InstanceName
            
            If n_CycleCount > 1 Then
                '==== Get all vector/cycle fail pins ====
                FailPinListData = TheHdw.Digital.Pins(FailedPin_Get_Index).CMEM.FailIndexList(0, UBound(dbl_CmemCycleDataArr) + 1)
                
                For i = 0 To FailPinListData.Pins.Count - 1
                    For j = 0 To UBound(FailPinListData.Pins(i).value(v_site))
                        s_tempFailCycle = CStr(dbl_CmemCycleDataArr(FailPinListData.Pins(i).value(v_site)(j)))
                        If dic_FailPinCycle.Exists(s_tempFailCycle) Then
                            s_tempFailPin = CombineStringList(dic_FailPinCycle(s_tempFailCycle), FailPinListData.Pins(i).name, "/")
                            dic_FailPinCycle.Remove (s_tempFailCycle)
                        Else
                            s_tempFailPin = FailPinListData.Pins(i).name
                        End If
                        dic_FailPinCycle.Add s_tempFailCycle, s_tempFailPin
                    Next j
                Next i
                '==== Get all vector/cycle fail pins ====
                For i = 0 To n_CycleCount - 1
                    THEEXEC.Datalog.WriteComment "SFC,1," & CStr(v_site) & "," & mid(FlowName, 6, Len(FlowName)) & "," & s_InsName & "," & s_ModName _
                                                & "," & CStr(dbl_CmemVectorDataArr(indexArr(i))) & "," & CStr(dbl_CmemCycleDataArr(indexArr(i))) & "," & dic_FailPinCycle(CStr(dbl_CmemCycleDataArr(indexArr(i))))
                Next i
            Else
                THEEXEC.Datalog.WriteComment "SFC,1," & CStr(v_site) & "," & mid(FlowName, 6, Len(FlowName)) & "," & s_InsName & "," & s_ModName _
                                            & "," & CStr(dbl_CmemVectorDataArr(indexArr(0))) & "," & CStr(dbl_CmemCycleDataArr(indexArr(0))) & "," & FailedPins
            End If
        End If

    Next v_site

Exit Function
errHandler:
    THEEXEC.AddOutput ("Error in ifcmemGetFailCycles_UF function")
    Call Print_Error_Message(Error_Info, "LIB_Common_AP", "ifcmemGetFailCycles_UF") 'Add ErrHandler 2023/08/18
    If AbortTest Then Exit Function Else Resume Next 'Add ErrHandler 2023/08/18
End Function


' [20231228][T-All][Tank] Process SFC function
Public Function ifcmemGetFailCycles_UP(pmCnt As Long, pmNames() As String)
On Error GoTo errHandler

    Dim v_site As Variant
    Dim i As Long
    Dim s_ModName As String
    Dim s_InsName As String
    Dim PatIdx() As Long
    Dim PatNames() As String
    Dim Sbln_PatternPass As New SiteBoolean
    Dim patternNames() As String
    Dim dbl_CmemCycleDataArr As New SiteVariant
    Dim dbl_CmemVectorDataArr As New SiteVariant
    Dim SiteIndexData As New SiteVariant
    Dim SitePinData As New SiteVariant
    Dim indexArr() As Long
    Dim pinDataArr() As Double
    Dim Failpin_temp() As String
    Dim FailedPin_Get_Index As String
    Dim FailedPins As String
    Dim s_InputData() As String
    Dim n_CycleCount As Long
    Dim FlowName As String
    
    Dim FailPinListData As New PinListData
    Dim dic_FailPinCycle As New Dictionary
    Dim j As Long
    Dim s_tempFailPin As String
    
    Dim s_tempFailCycle As String
    
    'Read back the central data.
    Sbln_PatternPass = TheHdw.Digital.Patgen.PatternBurstPassedPerSite
    If Sbln_PatternPass.Any(False) Then
        s_InputData = Split(pmNames(0), ",")
        n_CycleCount = CLng(s_InputData(0))
        
        For Each v_site In THEEXEC.sites
            If Sbln_PatternPass(v_site) = False Then
           
                dbl_CmemCycleDataArr(v_site) = TheHdw.Digital.CMEM.PatGenInfo(tlCMEMModCycle, -1)
                dbl_CmemVectorDataArr(v_site) = TheHdw.Digital.CMEM.PatGenInfo(tlCMEMVMVectorOffset, -1)
           
                'Get fail pin
                Failpin_temp = TheHdw.Digital.FailedPins(v_site)
                If UBound(Failpin_temp) >= 0 Then
                    FailedPin_Get_Index = Join(Failpin_temp, ",")
                    FailedPins = Join(Failpin_temp, "/")
                End If
                
               
                
                'Get fail pin index
                Call TheHdw.Digital.Pins(FailedPin_Get_Index).CMEM.StoredCycleData(SiteIndexData, SitePinData, -1, True)
                
                indexArr = SiteIndexData
                
                'Get flow ,pattern, insname
                FlowName = THEEXEC.Flow.CurrentFlowSheetName
                Call TheHdw.Digital.CMEM.PatternName(PatIdx, PatNames)
                s_ModName = PatNames(0)
                s_InsName = THEEXEC.DataManager.InstanceName
                
                If n_CycleCount > 1 Then
                    '==== Get all vector/cycle fail pins ====
                    FailPinListData = TheHdw.Digital.Pins(FailedPin_Get_Index).CMEM.FailIndexList(0, UBound(dbl_CmemCycleDataArr(v_site)) + 1)
                    
                    For i = 0 To FailPinListData.Pins.Count - 1
                        For j = 0 To UBound(FailPinListData.Pins(i).value(v_site))
                            s_tempFailCycle = CStr(dbl_CmemCycleDataArr(v_site)(FailPinListData.Pins(i).value(v_site)(j)))
                            If dic_FailPinCycle.Exists(s_tempFailCycle) Then
                                s_tempFailPin = CombineStringList(dic_FailPinCycle(s_tempFailCycle), FailPinListData.Pins(i).name, "/")
                                dic_FailPinCycle.Remove (s_tempFailCycle)
                            Else
                                s_tempFailPin = FailPinListData.Pins(i).name
                            End If
                            dic_FailPinCycle.Add s_tempFailCycle, s_tempFailPin
                        Next j
                    Next i
                    '==== Get all vector/cycle fail pins ====
                    For i = 0 To n_CycleCount - 1
                        THEEXEC.Datalog.WriteComment "SFC,1," & CStr(v_site) & "," & mid(FlowName, 6, Len(FlowName)) & "," & s_InsName & "," & s_ModName _
                                                    & "," & CStr(dbl_CmemVectorDataArr(v_site)(indexArr(i))) & "," & CStr(dbl_CmemCycleDataArr(v_site)(indexArr(i))) & "," & dic_FailPinCycle(CStr(dbl_CmemCycleDataArr(indexArr(i))))
                    Next i
                Else
                    THEEXEC.Datalog.WriteComment "SFC,1," & CStr(v_site) & "," & mid(FlowName, 6, Len(FlowName)) & "," & s_InsName & "," & s_ModName _
                                                & "," & CStr(dbl_CmemVectorDataArr(v_site)(indexArr(0))) & "," & CStr(dbl_CmemCycleDataArr(v_site)(indexArr(0))) & "," & FailedPins
                End If
                dic_FailPinCycle.RemoveAll
            End If

        Next v_site
    End If
    
Exit Function
errHandler:
    THEEXEC.AddOutput ("Error in ifcmemGetFailCycles_UP function")
    Call Print_Error_Message(Error_Info, "LIB_Common_AP", "ifcmemGetFailCycles_UP") 'Add ErrHandler 2023/08/18
    If AbortTest Then Exit Function Else Resume Next 'Add ErrHandler 2023/08/18
End Function


' [20231228][T-All][Tank] Reset SFC CMEM
Public Function SFC_CMEM_Stop()
On Error GoTo errHandler

    If THEEXEC.TesterMode = testModeOnline Then
        ' Use this statement to turn off the Pin CMEM capture.
        Call TheHdw.Digital.CMEM.SetCaptureConfig(0, CmemCaptNone) ' Resets CMEM
        ' Use this statement to turn off the Central CMEM capture.
        TheHdw.Digital.CMEM.CentralFields = tlCMEMNone
    End If
    
     ''''' New request for pattern pin group '''''
    If glb_TesterType = "UltraFLEXplus" Then
        TheHdw.Digital.Patgen.ScanBurstEnabled = False
        THEEXEC.Datalog.Setup.ScanSetup.EnableScanLogging = False
    End If
    ''''''''''''''''''''''''''''''''''''''''''''''

Exit Function 'Add ErrHandler 2023/08/18
errHandler: 'Add ErrHandler 2023/08/18
    Call Print_Error_Message(Error_Info, "LIB_Common_AP", "SFC_CMEM_Stop") 'Add ErrHandler 2023/08/18
    If AbortTest Then Exit Function Else Resume Next 'Add ErrHandler 2023/08/18
End Function
