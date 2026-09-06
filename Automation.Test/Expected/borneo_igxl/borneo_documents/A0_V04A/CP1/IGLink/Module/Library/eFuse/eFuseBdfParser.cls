VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "eFuseBdfParser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Public sheetName As String
'Once apear in BDF, then new the memory for it
Public Bank_Cfg As eFuseBdfBank
Public Bank_UdrP As eFuseBdfBank
Public Bank_UdrPcmp As eFuseBdfBank
''''20210607,Add for Rhod''''
Public Bank_UdrP0 As eFuseBdfBank
Public Bank_UdrP0cmp As eFuseBdfBank
Public Bank_UdrP1 As eFuseBdfBank
Public Bank_UdrP1cmp As eFuseBdfBank
'''''''''''''''''''''''''''''
Public Bank_UdrE As eFuseBdfBank
Public Bank_UdrEcmp As eFuseBdfBank
Public Bank_Udr As eFuseBdfBank
Public Bank_Udrcmp As eFuseBdfBank
Public Bank_Mon As eFuseBdfBank
Public Bank_Uid As eFuseBdfBank
Public Bank_Ecid As eFuseBdfBank
Public Bank_Ecid1 As eFuseBdfBank
Public Bank_Sw0 As eFuseBdfBank
Public Bank_Sw1 As eFuseBdfBank

''202004xx for ap
Public PseudoFuseFileName As String
Public Fuse_Revision As Long
Public Chip_rev_minor As Long
Public Chip_rev_major As Long
Public DicPseudoFuseData As Dictionary

Public BinCutNumBank As eFuseBdfBank ' one of above bank, which contains the field "product_identifier"
Public FieldBankMap As Dictionary ' FieldName => eFuseBdfField
Public Banks As Dictionary
Public BaseVoltage As Double      'the algorithm is alg_base
Public BaseResolution As Double 'the algorithm is alg_base
Public DatabaseVersion As Long  'the field named "EFUSE_DATABASE_REVISION"
Public DicVddBinPmodeMap As Dictionary 'assume that will be unique from all banks, otherwise move this to bank class
Public DicVddBinBankMap As Dictionary 'assume that will be unique from all banks, otherwise move this to bank class
Public DicVddBinIdsFieldMap As Dictionary 'assume that will be unique from all banks, otherwise move this to bank class
Public DicBKMMap As Dictionary 'parse BKMinfo Sheet ''202004xx for ap
Public ReadRealBKMdone As Boolean
Public DicDramMap As Dictionary
Public DicDramHipMap As Dictionary
Public DicCPMMap As Dictionary
Public DramFieldSize As Long
Public DramSiteContrl As Boolean
Private RefCfgBase As eFuseCfgParser
Private bAutoCfgAssign As Boolean
Private ws As Worksheet
Private myUtility As Utility
Const MaxColRang = 99
'These 2 columns are used to check the format should be same as proposed one, sequence  didn't allow any change
Const header = "* eFuse Bit Def,MSB BIT,LSB BIT,Bit Width,N/A,N/A,N/A,N/A,Programming Stage,Low Limit,High Limit,Resolution,Algorithm,Description,Default or Real,Default Value"
Const header2 = "* eFuse Bit Def,MSB BIT,LSB BIT,Bit Width,USI MSB-Bit Cycle,USI LSB-Bit Cycle,USO MSB-Bit Cycle,USO LSB-Bit Cycle,Programming Stage,Low Limit,High Limit,Resolution,Algorithm,Description,Default or Real,Default Value"
Const headerSimple = "* eFuse Bit Def,MSB BIT,LSB BIT,Bit Width,Programming Stage,Low Limit,High Limit,Resolution,Algorithm,Description,Default or Real,Default Value"
Private ColShift As Long
Private stage As String
Public CFG_Only As Boolean
Private dicStageCheck As Dictionary 'Just for msg check, no use for others
Private ParsingArrayData() As Variant

Public Sub Initialize(currStage As String)
On Error GoTo errHandler
    stage = currStage
    Set myUtility = New Utility
    myUtility.Initialize (stage)
    Set Banks = New Dictionary
    Banks.compareMode = TextCompare
    CFG_Only = False
    ReadRealBKMdone = False
    BaseVoltage = 0
    BaseResolution = 25
    DatabaseVersion = -1
    myUtility.IniDictionary DicVddBinPmodeMap
    myUtility.IniDictionary DicVddBinIdsFieldMap
    myUtility.IniDictionary DicVddBinBankMap
    myUtility.IniDictionary dicStageCheck
    myUtility.IniDictionary FieldBankMap
    myUtility.IniDictionary DicBKMMap ''202004xx for ap
    myUtility.IniDictionary DicPseudoFuseData ''202004xx for ap
    myUtility.IniDictionary DicDramMap
    myUtility.IniDictionary DicDramHipMap
    myUtility.IniDictionary DicCPMMap
    myUtility.IniDictionary DicVddPmode
    myUtility.IniDictionary DicVddPmodeShadow
Exit Sub
errHandler:
    Call Print_Error_Message(Error_Info, "eFuseBdfParser", "Initialize")
    If AbortTest Then Exit Sub Else Resume Next
End Sub

Public Function Parse(ByVal ShtName As String, ByRef myCfg As eFuseCfgParser, Optional autoCfgAssign As Boolean = False) As Boolean
On Error GoTo errHandler
    sheetName = ShtName
    Set RefCfgBase = myCfg
    bAutoCfgAssign = autoCfgAssign
    Set ws = Worksheets(sheetName)
    Parse = WorkFlow(sheetName)
Exit Function
errHandler:
    Call Print_Error_Message(Error_Info, "eFuseBdfParser", "Parse")
    If AbortTest Then Exit Function Else Resume Next
End Function

Private Function IsCfgField(ByRef myfield As eFuseBdfField) As Boolean
On Error GoTo errHandler
Dim cfgFiledStr As Variant, cfgfield As New eFuseCfgField

    If Not bAutoCfgAssign Then IsCfgField = False: Exit Function
    If myfield.Algorithm <> alg_app Then IsCfgField = False: Exit Function ' only allow algorithm is app to be changed!

    For Each cfgFiledStr In RefCfgBase.Fields.Keys
        Set cfgfield = RefCfgBase.Fields(cfgFiledStr)
        ''Some field of BDF table can corresponds to CFG table which algorithm isn't "cond".
        If LCase(cfgfield.name) = LCase(myfield.name) Then
            IsCfgField = True
            If cfgfield.msb <> myfield.msb Or cfgfield.LSB <> myfield.LSB Then
                myUtility.MessageBox myfield.name & "> The LSB or MSB isn't equal, they both are defined on BDF and CFG sheets. "
                Call Print_Error_Message(Error_Info, "eFuseBdfParser", "IsCfgField", myfield.name & "> The LSB or MSB isn't equal, they both are defined on BDF and CFG sheets. ")
            End If
            If myfield.Algorithm <> alg_cond Then  'even it's written "scan", will also be changed
                myfield.Algorithm = alg_cond
            End If
            Exit Function
        End If
    Next
    IsCfgField = False

Exit Function
errHandler:
    Call Print_Error_Message(Error_Info, "eFuseBdfParser", "IsCfgField")
    If AbortTest Then Exit Function Else Resume Next
End Function

Private Sub AutoDefaultCheck(ByRef myfield As eFuseBdfField)
On Error GoTo errHandler

    If Not bAutoDefaultCheck Or LCase(myfield.name) = "rf_id_radio_ad" Or myfield.Llimit = -1 Then  ' "rf_id_radio_ad" special specify for Gloay by customer
        Exit Sub
    Else
        If myfield.Llimit = myfield.Hlimit Then
            If myfield.DefaultOrReal <> dr_default Then
                myfield.DefaultOrReal = dr_default
                myfield.ReadOnly = True
                myUtility.WriteDlg myUtility.TxtFmt(myfield.name, 48, left:=True) & " => The field's DefaultOrReal is changed to ""default""."
            End If
            If myfield.Llimit <> myUtility.String2Dbl(myfield.DefaulString) Then
                myUtility.MessageBox myfield.name & " => The field default value is not within H/L limit => Value " & myfield.DefaulString & ", Limits " & myfield.Llimit & "/" & myfield.Hlimit
                Call Print_Error_Message(Error_Info, "eFuseBdfParser", "AutoDefaultCheck", myfield.name & " => The field default value is not within H/L limit => Value " & myfield.DefaulString & ", Limits " & myfield.Llimit & "/" & myfield.Hlimit)
            End If
        End If
    End If

Exit Sub
errHandler:
    Call Print_Error_Message(Error_Info, "eFuseBdfParser", "AutoDefaultCheck")
    If AbortTest Then Exit Sub Else Resume Next
End Sub

Private Function WorkFlow(sheetName As String) As Boolean
On Error GoTo errHandler
Dim endRow As Integer, iRow As Integer
Dim lEfuseBank As eFuseBdfBank
Dim lEfuseField As eFuseBdfField
Dim cellvalue As String, bCfgField As Boolean
Dim m_SheetMaxRow As Long, m_SheetMaxCol As Long
Dim myCfgField As eFuseCfgField
Dim defVal As Double
Dim msg As String: msg = vbNullString

    DicVddPmode.RemoveAll
    DicVddPmodeShadow.RemoveAll
    
    endRow = ws.Cells(65536, 1).End(xlUp).Row
    myUtility.Get_Sheet_Info sheetName, m_SheetMaxRow, m_SheetMaxCol, ParsingArrayData

    For iRow = 1 To m_SheetMaxRow
        ''Check bank header.
        If IsBankStart(iRow, ParsingArrayData) Then
            Set lEfuseBank = InitialBank(iRow)
            If lEfuseBank Is Nothing Then
                msg = "Failed to initialize bank information, please check it!!!"
                GoTo skip
            End If
            lEfuseBank.MultiTypeJobList.RemoveAll
            If Banks.Exists(lEfuseBank.name) Then
                msg = "Duplicate Bank name, " + lEfuseBank.name + " Please check BDF!"
                GoTo skip
            End If
            If lEfuseBank.name <> "" Then
                Banks.Add lEfuseBank.name, lEfuseBank
                ''202005xx fo ap
                If lEfuseBank.name Like "*cmp*" Then lEfuseBank.cmpFlag = True
                ''Check header sequence which is the same with "const header define"
                If Not HeaderSeqCheck(iRow + 1) Then
                    msg = "Error format! Please check it again!"
                    GoTo skip
                End If
            End If
            iRow = iRow + 1
            GoTo nextstep
        Else
            cellvalue = ParsingArrayData(iRow, 1)
        End If
        
        ''Get Field data row by row.
        Set lEfuseField = GetFieldRow(iRow, lEfuseBank.SeparateEarly)
        If Not lEfuseField Is Nothing Then
            If lEfuseField.BlowLocation Like "*_*" And Not LCase(lEfuseField.BlowLocation) Like "*_early" Then
                If Not lEfuseBank.MultiTypeJobList.Exists(lEfuseField.BlowLocation) Then lEfuseBank.MultiTypeJobList.Add lEfuseField.BlowLocation, True
            End If
            
            'if all field name of all bank is unique, then can enable the function.
            'AP can't use it. Some field name is the same with different bank.
            If AllFieldUnique Then FieldBankMap.Add lEfuseField.name, lEfuseField
            If lEfuseBank.name = CfgBankName Then
                bCfgField = IsCfgField(lEfuseField)
                ''---------------------------------------------------------------------------------------
                ''For RF some proj
                ''The the cell info of default or real cell is always "real" in BDF table.
                ''No matter the value depends on table(Default value) or test item(Real value).
                ''Check which field's high/low limit is the same and then modify the "real"==> "default".
                ''---------------------------------------------------------------------------------------
                Call AutoDefaultCheck(lEfuseField)
            End If
            lEfuseField.bankName = lEfuseBank.name

            lEfuseBank.AddField lEfuseField

            'Check stage if the field exists in CFG table.
            'Port some CFGfield info to BDFfield first because will run with BDFfield obj only.
            
            '20220210,Modify for Fixed SVM Enable issue
            If Not CfgDataBaseAct Is Nothing And UCase(lEfuseBank.name) = UCase(CfgBankName) Then
                If CfgDataBaseAct.Fields.Exists(lEfuseField.name) Then
                    Set myCfgField = CfgDataBaseAct.Fields(lEfuseField.name)
                    lEfuseField.IsEarlyFuse = myCfgField.isEarly
                    lEfuseField.BlowLocation = myCfgField.BlowLocation
                    If Not (LCase(lEfuseField.BlowLocation) = LCase(myUtility.currStage)) Then lEfuseField.IsEarlyFuse = False
                End If
            End If
            Call Add2Category(lEfuseBank, lEfuseField)
            Select Case lEfuseField.Algorithm
                Case alg_force:
                    lEfuseBank.HadForceFuse = True
                    AddItem lEfuseBank.DicForceFields, lEfuseField, lEfuseBank.name
                    Call Print_Error_Message(Warning_Info, "eFuseBdfParser", "WorkFlow", "The BDF contains the [force] algorithm~ Please check and confirm it, THANKS!")
                Case alg_crc:
                    If lEfuseField.crc_type = crc_normal Then
                       lEfuseBank.HadCrcFuse = True
                    ElseIf lEfuseField.crc_type = crc_onecomp Then
                       lEfuseBank.HadCrcCompFuse = True
                    End If
                Case alg_base:
                    defVal = myUtility.Hex2Dec(lEfuseField.FieldDefault)
                    'Get resolution.
                    If lEfuseField.Resolution = 0 Then
                        lEfuseField.Resolution = BaseResolution
                    Else
                        If BaseResolution <> lEfuseField.Resolution And BaseResolution <> 25 Then
                            msg = "The base voltage values are different between banks, please confirm!"
                            GoTo skip
                        End If
                        BaseResolution = lEfuseField.Resolution
                    End If
                    'Get base voltage.
                    If defVal > 2 ^ lEfuseField.size And lEfuseField.DefaultOrReal <> dr_decimal Then 'mV
                        lEfuseBank.Base = defVal
                        lEfuseField.Default = myUtility.Dec2HexStr(defVal \ lEfuseField.Resolution - 1, lEfuseField.hBytes)
                        lEfuseField.FieldDefault = lEfuseField.Default
                    Else
                        lEfuseBank.Base = (myUtility.Hex2Dec(lEfuseField.FieldDefault) + 1) * lEfuseField.Resolution
                    End If
                    If BaseVoltage = 0 Then
                            BaseVoltage = lEfuseBank.Base
                    ElseIf BaseVoltage <> lEfuseBank.Base Then
                            msg = "The base voltage values are different between banks, please confirm!"
                            GoTo skip
                    End If
                Case alg_vddbin:
                    'Enable Fused flag for bincut field.
                    'Add field to dictionary.
                    If lEfuseField.DefaultOrReal = dr_real Then
                        lEfuseBank.HadVddBinFuse = True
                        AddItem lEfuseBank.DicBinCut, lEfuseField, lEfuseBank.name
                    End If
                    If LCase(lEfuseField.name) Like "*_shadow" Then
                        If Not DicVddPmodeShadow.Exists(lEfuseField.name) Then
                            DicVddPmodeShadow.Add lEfuseField.name, lEfuseField.BlowLocation
                        Else
                            msg = "Duplicate VddBin field name (" & lEfuseField.name & ") in BDF"
                            GoTo skip
                        End If
                    Else
                        If Not DicVddPmode.Exists(lEfuseField.name) Then
                            DicVddPmode.Add lEfuseField.name, lEfuseField.BlowLocation
                        Else
                            msg = "Duplicate VddBin field name (" & lEfuseField.name & ") in BDF"
                            GoTo skip
                        End If
                    End If
                Case alg_ids:
                    'Set the low limit from "0" to the ids resolution. Set high limit. Set resolution.
                    'Enable Fused Flag for ids field.
                    'Add field to dictionary.
                    If lEfuseField.Llimit = 0 Then lEfuseField.Llimit = (lEfuseField.Resolution) * 0.001 'to A
                    lEfuseField.Hlimit = lEfuseField.Hlimit * 0.001 'to A
                    lEfuseField.Resolution = lEfuseField.Resolution  'keep it mA
                    If lEfuseField.DefaultOrReal = dr_real Then
                        lEfuseBank.HadIdsFuse = True
                        AddItem lEfuseBank.DicIds, lEfuseField, lEfuseBank.name
                    End If
                    lEfuseBank.DicAllIdsList.Add lEfuseField, lEfuseBank.name
                Case alg_uid:
                    If lEfuseField.DefaultOrReal = dr_real Then
                        lEfuseBank.HadRandomFuse = True
                        AddItem lEfuseBank.DicRandomFused, lEfuseField, lEfuseBank.name
                    End If
                Case alg_lotid, alg_numeric:
                    lEfuseBank.EcidFieldSize = lEfuseBank.EcidFieldSize + lEfuseField.size
                Case Else
            End Select
    
            If lEfuseField.Algorithm <> alg_vddbin Then
                If lEfuseField.IsOverflow(lEfuseField.FieldDefault, True) Then
                    msg = "The field " + lEfuseField.name + "fuse value is overflow, please check it!!!"
                    GoTo skip
                End If
            End If
            If (LCase(stage) = LCase(lEfuseField.BlowLocation) Or bLumpStages) Or (lEfuseField.IsEarlyFuse And (Not myUtility.testedStages.Exists(stage))) Then
                If lEfuseField.ReadOnly Or lEfuseField.IsEarlyFuse Then
                'Add field to dictionary with default case or early case.
                    lEfuseBank.DicDefaultIncludeEarly.Add lEfuseField.name, lEfuseBank.name
                Else
                'Add field to dictionary with non-default case and non-early case.
                    lEfuseBank.DicNonDefaultNonEarly.Add lEfuseField.name, lEfuseBank.name
                End If
            End If
            
            If LCase(lEfuseField.name) = BinCutIdentifierField Then Set BinCutNumBank = lEfuseBank
        End If
nextstep:
     Next

     'set MSB / LSB / MSBFirst to DSPwave.
     Call CollectFieldIdx
     'Check safe voltage is all the same.
     Call CheckVddBinSafeVoltage
     'Calc RV segment(for Partial fuse, DVRV).
     Call AddRVSegment
     '20221117, support multi-type program stage
     Call AddMultiTypeRVSegment
     Call MultiTypeFieldClassify
     
     WorkFlow = True

Exit Function
skip:
    Call Print_Error_Message(Error_Info, "eFuseBdfParser", "WorkFlow", msg)
    WorkFlow = False
    TheExec.Flow.TestLimit resultVal:=999, lowVal:=1, hiVal:=1, Tname:="BDF_Parser"
    Exit Function
errHandler:
    Call Print_Error_Message(Error_Info, "eFuseBdfParser", "WorkFlow")
    If AbortTest Then Exit Function Else Resume Next
End Function

Private Sub CheckVddBinSafeVoltage()
On Error GoTo errHandler
Dim bank As eFuseBdfBank, bankstr As Variant
Dim fieldStr As Variant, field As eFuseBdfField
Dim code As String: code = vbNullString

    If BaseVoltage = 0 And BaseVoltage <> VddbinningBaseVoltage Then  '' Leon 2021 1206 _  to solve "No define base voltage at efuse table" case
        BaseVoltage = VddbinningBaseVoltage
    End If

    For Each bankstr In Banks.Keys
        Set bank = Banks(CStr(bankstr))
        For Each fieldStr In bank.Fields.Keys
            Set field = bank.Fields(fieldStr)
            If field.Algorithm = alg_vddbin And field.DefaultOrReal <> dr_real Then
                field.Default = Trim(Replace(LCase(field.Default), "mv", ""))
                code = CStr(myUtility.CeilingValue((field.Default - BaseVoltage) / field.Resolution, 1))
                field.Default = myUtility.String2Hex(code, 0, field.hBytes)
                field.FieldDefault = field.Default
                field.Llimit = ((code * field.Resolution) + BaseVoltage) * 0.001 'to V
                field.Hlimit = ((code * field.Resolution) + BaseVoltage) * 0.001
            End If
        Next
    Next

Exit Sub
errHandler:
    Call Print_Error_Message(Error_Info, "eFuseBdfParser", "CheckVddBinSafeVoltage")
    If AbortTest Then Exit Sub Else Resume Next
End Sub

Private Sub CollectFieldIdx()
On Error GoTo errHandler
Dim bank As eFuseBdfBank, bankstr As Variant
Dim fieldStr As Variant, site As Variant, field As eFuseBdfField
Dim iIdx As Long

    For Each bankstr In Banks.Keys
        Set bank = Banks(bankstr)
        Set bank.FieldIdxSt = New DSPWave: bank.FieldIdxSt.CreateConstant 0, bank.Fields.Count
        Set bank.FieldIdxSp = New DSPWave: bank.FieldIdxSp.CreateConstant 0, bank.Fields.Count
        Set bank.FieldIdxMsb = New DSPWave: bank.FieldIdxMsb.CreateConstant 0, bank.Fields.Count
        
        For Each site In TheExec.sites
            iIdx = 0
            For Each fieldStr In bank.Fields.Keys
                Set field = bank.Fields(fieldStr)
                bank.FieldIdxSt.ElementLite(iIdx) = IIf(field.LSB < field.msb, field.LSB, field.msb)
                bank.FieldIdxSp.ElementLite(iIdx) = IIf(field.LSB < field.msb, field.msb, field.LSB)
                bank.FieldIdxMsb.ElementLite(iIdx) = IIf(field.Msb1st, 1, 0)
                iIdx = iIdx + 1
            Next
        Next
    Next

Exit Sub
errHandler:
    Call Print_Error_Message(Error_Info, "eFuseBdfParser", "CollectFieldIdx")
    If AbortTest Then Exit Sub Else Resume Next
End Sub

Private Sub Add2Category(lEfuseBank As eFuseBdfBank, lEfuseField As eFuseBdfField)
On Error GoTo errHandler

    'Early Fused
    If lEfuseField.IsEarlyFuse Then
        lEfuseBank.HadEarlyFuse = True
        lEfuseBank.DicEarlyFused.Add lEfuseField.name, lEfuseBank.name
    Else
        'case A - The fused stage equals to current job.
        'case B - LumpStage = True, it means the field's fused stage is after CP1(CP2,WLFT1,FT2...) and change to CP1 to fuse.
        If LCase(stage) = LCase(lEfuseField.BlowLocation) Or bLumpStages Then
            lEfuseBank.DicOthers.Add lEfuseField.name, lEfuseBank.name
        ElseIf LCase(stage) = LCase(GetRealStage(lEfuseField.BlowLocation)) Then
            lEfuseBank.DicAllTypefields.Add lEfuseField.name, lEfuseField.BlowLocation
        End If
    End If
    
    'Config Table field
    If lEfuseField.Algorithm = alg_cond Then lEfuseBank.DicCondTables.Add lEfuseField.name, lEfuseBank.name
    If UCase(lEfuseField.name) Like "*EFUSE_DATABASE_REVISION*" Then DatabaseVersion = myUtility.Hex2Dec(lEfuseField.Default)

Exit Sub
errHandler:
    Call Print_Error_Message(Error_Info, "eFuseBdfParser", "Add2Category")
    If AbortTest Then Exit Sub Else Resume Next
End Sub

Private Sub AddItem(dic As Dictionary, field As eFuseBdfField, bank As String, Optional skipStageChk As Boolean = False)
On Error GoTo errHandler

    If (LCase(stage) = LCase(BdfDataBase.GetRealStage(field.BlowLocation)) Or bLumpStages) Or skipStageChk Then
        dic.Add field.name, bank
    End If

Exit Sub
errHandler:
    Call Print_Error_Message(Error_Info, "eFuseBdfParser", "AddItem")
    If AbortTest Then Exit Sub Else Resume Next
End Sub

Private Function IsBankStart(searStRow As Integer, ParsingArray() As Variant) As Boolean
On Error GoTo errHandler
Dim cellStr As String: cellStr = vbNullString

    cellStr = ParsingArray(searStRow, 1)
    IsBankStart = myUtility.IsStrMatch(cellStr, "Access.*Mode")

Exit Function
errHandler:
    Call Print_Error_Message(Error_Info, "eFuseBdfParser", "IsBankStart")
    If AbortTest Then Exit Function Else Resume Next
End Function

Private Function InitialBank(searStRow As Integer) As eFuseBdfBank
On Error GoTo errHandler
Dim DoubleBits As Boolean
Dim cellStr  As String, bankstr As String
Dim lBank As eFuseBdfBank
Dim bank As String: bank = vbNullString

    cellStr = ParsingArrayData(searStRow, 1)  ''''Head Str
    bankstr = ParsingArrayData(searStRow + 1, 1)  ''' Bank Table Header

    ''' bank instance
    Set lBank = Nothing
    bank = Replace(Split(LCase(bankstr), " ")(0), "bank_", "")
    Select Case bank
        Case "config", "CFG"
              Set Bank_Cfg = New eFuseBdfBank
              Set lBank = Bank_Cfg
              bank = CfgBankName
        Case "udr_e"
              Set Bank_UdrE = New eFuseBdfBank
              Set lBank = Bank_UdrE
        Case "cmp_e"
              Set Bank_UdrEcmp = New eFuseBdfBank
              Set lBank = Bank_UdrEcmp
        Case "udr_p"
              Set Bank_UdrP = New eFuseBdfBank
              Set lBank = Bank_UdrP
        Case "cmp_p"
              Set Bank_UdrPcmp = New eFuseBdfBank
              Set lBank = Bank_UdrPcmp
        ''''20210607,Add for Rhod''''
        Case "udr_p0"
              Set Bank_UdrP0 = New eFuseBdfBank
              Set lBank = Bank_UdrP0
        Case "cmp_p0"
              Set Bank_UdrP0cmp = New eFuseBdfBank
              Set lBank = Bank_UdrP0cmp
        Case "udr_p1"
              Set Bank_UdrP1 = New eFuseBdfBank
              Set lBank = Bank_UdrP1
        Case "cmp_p1"
              Set Bank_UdrP1cmp = New eFuseBdfBank
              Set lBank = Bank_UdrP1cmp
        '''''''''''''''''''''''''''''
        Case "udr"
              Set Bank_Udr = New eFuseBdfBank
              Set lBank = Bank_Udr
        Case "cmp"
              Set Bank_Udrcmp = New eFuseBdfBank
              Set lBank = Bank_Udrcmp
        Case "mon"
              Set Bank_Mon = New eFuseBdfBank
              Set lBank = Bank_Mon
        Case "uid", "uid(aes)"
              Set Bank_Uid = New eFuseBdfBank
              Set lBank = Bank_Uid
        Case "ecid", "ecid_0"
              Set Bank_Ecid = New eFuseBdfBank
              Set lBank = Bank_Ecid
        Case "ecid_1"
              Set Bank_Ecid1 = New eFuseBdfBank
              Set lBank = Bank_Ecid1
        Case "sw0"
              Set Bank_Sw0 = New eFuseBdfBank
              Set lBank = Bank_Sw0
        Case "sw1"
              Set Bank_Sw1 = New eFuseBdfBank
              Set lBank = Bank_Sw1
        Case Else
             Call Print_Error_Message(Error_Info, "eFuseBdfParser", "InitialBank", "No mapping bank - " & bank)
             Set InitialBank = lBank
             Exit Function
    End Select
    
    ''' check 1 bits/2 bits
    If myUtility.IsStrMatch(cellStr, "Access.*Mode.*(2-bit.*data)") Then
        DoubleBits = True
    ElseIf myUtility.IsStrMatch(cellStr, "Access.*Mode.*(1-bit.*data)") Then
        DoubleBits = False
    Else
        Call Print_Error_Message(Error_Info, "eFuseBdfParser", "InitialBank", "Missing Access Mode bit data")
        Set InitialBank = Nothing
        Exit Function
    End If
   
    If Not lBank Is Nothing Then
        lBank.name = bank
        Set lBank.myUtility = myUtility
        lBank.DoubleBits = DoubleBits
        lBank.sheetName = sheetName
        If myUtility.IsStrMatch(cellStr, "JTAG") Then
            lBank.pgmMode = pgm_JTAG
        Else
            If myUtility.IsStrMatch(cellStr, "Direct-Access") Then
                lBank.pgmMode = pgm_DAA
            Else
                Call Print_Error_Message(Error_Info, "eFuseBdfParser", "InitialBank", "Unrecognize access mode! neither DAA nor JTAG!")
                Set InitialBank = Nothing
                Exit Function
            End If
        End If
        Select Case bank 'Customize Setting at here
            Case "config", "CFG"
                  If RunPartType = PartType.Wireless Then lBank.SeparateEarly = False ' fuse at a time
            Case "ecid"
                  If RunPartType = PartType.Wireless Then lBank.SeparateEarly = False ' fuse at a time
            ''''20210607,Add for Rhod''''
            '''''''''''''''''''''''''''''
                  
            Case Else
        End Select
        If bForceJtagMode Then lBank.pgmMode = pgm_JTAG 'This is for Golay only
    End If
    Set InitialBank = lBank

Exit Function
errHandler:
    Call Print_Error_Message(Error_Info, "eFuseBdfParser", "InitialBank")
    If AbortTest Then Exit Function Else Resume Next
End Function

Private Function GetFieldRow(searStRow As Integer, SeparateEarly As Boolean) As eFuseBdfField
On Error GoTo errHandler
Dim cellStr As String
Dim myfield As eFuseBdfField
    
    'Name
    cellStr = Trim(ParsingArrayData(searStRow, 1))
    
    'catch "end" -> Sheet end
    If cellStr = "" Or LCase(cellStr) = "end" Then Set GetFieldRow = myfield: Exit Function
   
    Set myfield = New eFuseBdfField
    Set myfield.myUtility = myUtility
    myfield.name = cellStr
    
    'The field exists in which row of the table.
    myfield.DataRowNum = searStRow
    
    'MSB / LSB / SIZE
    cellStr = Trim(ParsingArrayData(searStRow, 2)): myfield.msb = CLng(cellStr)
    cellStr = Trim(ParsingArrayData(searStRow, 3)): myfield.LSB = CLng(cellStr)
    cellStr = Trim(ParsingArrayData(searStRow, 4)): myfield.size = CLng(cellStr)
    
    'For UDRXX bank
    'It's for UDR shift in pattern order and shift out pattern order
    If ColShift = 0 Then
        cellStr = Trim(ParsingArrayData(searStRow, 5)): myfield.UsiMsb = myUtility.String2Long(cellStr, -1)
        cellStr = Trim(ParsingArrayData(searStRow, 6)): myfield.UsiLsb = myUtility.String2Long(cellStr, -1)
        cellStr = Trim(ParsingArrayData(searStRow, 7)): myfield.UsoMsb = myUtility.String2Long(cellStr, -1)
        cellStr = Trim(ParsingArrayData(searStRow, 8)): myfield.UsoLsb = myUtility.String2Long(cellStr, -1)
    End If

    'Stage / Hlimit / Llimit / Resoultion / MSBFirst / Alogrithm
    cellStr = Trim(ParsingArrayData(searStRow, 9 - ColShift)): myfield.BlowLocation = GetBlowLocation(cellStr)
    cellStr = Trim(ParsingArrayData(searStRow, 10 - ColShift)): myfield.Llimit = myUtility.String2Dbl(cellStr, -1): myfield.xLlimit = myUtility.String2Hex(cellStr, , myfield.hBytes)
    cellStr = Trim(ParsingArrayData(searStRow, 11 - ColShift)): myfield.Hlimit = myUtility.String2Dbl(cellStr, -1): myfield.xHlimit = myUtility.String2Hex(cellStr, , myfield.hBytes)
    cellStr = Trim(ParsingArrayData(searStRow, 12 - ColShift)): myfield.Resolution = myUtility.String2Dbl(cellStr, 0)
    myfield.Msb1st = myfield.msb >= myfield.LSB
    cellStr = Trim(ParsingArrayData(searStRow, 13 - ColShift)): myfield.Algorithm = GetAlgorithm(cellStr, myfield.BlowLocation)
    
    'Set early fusing flag
    If EarlyFuseAlgorithm.Exists(myfield.Algorithm) And SeparateEarly Then
        myfield.IsEarlyFuse = True
    End If
    If LCase(myfield.BlowLocation) Like LCase(myfield.myUtility.currStage) + "_early" Then
        myfield.IsEarlyFuse = True
    End If
    
    'Description / Default Real
    cellStr = Trim(ParsingArrayData(searStRow, 14 - ColShift)): myfield.Description = cellStr
    cellStr = Trim(ParsingArrayData(searStRow, 15 - ColShift)): myfield.DefaultOrReal = GetDefaultOrReal(cellStr)

    'Value
    cellStr = Trim(ParsingArrayData(searStRow, 16 - ColShift)):
    If (myfield.Algorithm = alg_vddbin And myfield.DefaultOrReal = dr_default) Then
        myfield.Default = cellStr
    Else
        myfield.Default = myUtility.String2Hex(cellStr, 0, myfield.hBytes)
    End If
    myfield.FieldDefault = myfield.Default
    If myfield.Algorithm = alg_crc And myfield.DefaultOrReal = dr_real Then
        myfield.Llimit = 1
        myfield.Hlimit = 2 ^ myfield.size
    End If
    
    'Set default value's H/L limit
    If myfield.DefaultOrReal = dr_default Or myfield.DefaultOrReal = dr_decimal Then
            myfield.Llimit = myUtility.String2Dbl(myfield.FieldDefault, -1)
            myfield.Hlimit = myfield.Llimit
    End If
     If myfield.Hlimit = -1 Then myfield.Llimit = -1
    
    myfield.DefaulString = cellStr ' store the original string from table
    If myfield.DefaultOrReal <> dr_default And myfield.DefaultOrReal <> dr_decimal Then myfield.ReadOnly = False
    If myfield.Algorithm = alg_cond Or myfield.Algorithm = alg_scan Then myfield.ReadOnly = True
    Set GetFieldRow = myfield

Exit Function
errHandler:
    Call Print_Error_Message(Error_Info, "eFuseBdfParser", "GetFieldRow")
    If AbortTest Then Exit Function Else Resume Next
End Function

Private Function GetBlowLocation(v As String) As String
On Error GoTo errHandler

    If Not myUtility.dicStageList.Exists(GetRealStage(v)) And Not LCase(v) = "cfgtable" And Not dicStageCheck.Exists(v) Then
        If Not LCase(v) Like "*_early" Then
            dicStageCheck.Add v, True
            myUtility.MessageBox "Unrecognized stage!! >> " & v
            Call Print_Error_Message(Error_Info, "eFuseBdfParser", "GetBlowLocation", "Unrecognized stage while reading BDF sheet!")
        End If
    End If

    If EFUSE_REPLACE_SLT_TO_FT3 And UCase(Trim(v)) = "SLT" Then
        GetBlowLocation = "FT3"
    Else
        GetBlowLocation = v
    End If

Exit Function
errHandler:
    Call Print_Error_Message(Error_Info, "eFuseBdfParser", "GetBlowLocation")
    If AbortTest Then Exit Function Else Resume Next
End Function

Private Function GetAlgorithm(v As String, stage As String) As Algorithm
On Error GoTo errHandler
Dim reParse As New RegExp

    v = Trim(v)
    reParse.IgnoreCase = True
    reParse.Pattern = "app|mtr|revision":   If reParse.test(v) Then GetAlgorithm = alg_app:     Exit Function
    reParse.Pattern = "ids":                If reParse.test(v) Then GetAlgorithm = alg_ids:     Exit Function
    reParse.Pattern = "vddbin":             If reParse.test(v) Then GetAlgorithm = alg_vddbin:  Exit Function
    reParse.Pattern = "cond":               If reParse.test(v) Then GetAlgorithm = alg_cond:    Exit Function
    reParse.Pattern = "scan":
    
    If reParse.test(v) Then
        If LCase(stage) = LCase(myUtility.dicStageList.Keys(0)) Then
            GetAlgorithm = alg_scan:
        Else
            GetAlgorithm = alg_app:
        End If
        Exit Function
    End If
               
    reParse.Pattern = "force":      If reParse.test(v) Then GetAlgorithm = alg_force:   Exit Function
    reParse.Pattern = "crc":        If reParse.test(v) Then GetAlgorithm = alg_crc:     Exit Function
    reParse.Pattern = "uid":        If reParse.test(v) Then GetAlgorithm = alg_uid:     Exit Function
    reParse.Pattern = "lotid":      If reParse.test(v) Then GetAlgorithm = alg_lotid:   Exit Function
    reParse.Pattern = "numeric":    If reParse.test(v) Then GetAlgorithm = alg_numeric: Exit Function
    reParse.Pattern = "base":       If reParse.test(v) Then GetAlgorithm = alg_base:    Exit Function
    reParse.Pattern = "n/a|na":     If reParse.test(v) Then GetAlgorithm = alg_na:      Exit Function
    myUtility.MessageBox "Unrecognize algorithm! " & v
    Call Print_Error_Message(Error_Info, "eFuseBdfParser", "GetAlgorithm", "Unrecognize algorithm! " & v)

Exit Function
errHandler:
    Call Print_Error_Message(Error_Info, "eFuseBdfParser", "GetAlgorithm")
    If AbortTest Then Exit Function Else Resume Next
End Function

Private Function GetDefaultOrReal(v As String) As DeaultReal
On Error GoTo errHandler
Dim reParse As New RegExp

    v = Trim(v)
    reParse.IgnoreCase = True
    reParse.Pattern = "default":            If reParse.test(v) Then GetDefaultOrReal = dr_default:  Exit Function
    reParse.Pattern = "safe.*voltage":      If reParse.test(v) Then GetDefaultOrReal = dr_default:  Exit Function
    reParse.Pattern = "revision":           If reParse.test(v) Then GetDefaultOrReal = dr_default:  Exit Function
    reParse.Pattern = "real|bincut":        If reParse.test(v) Then GetDefaultOrReal = dr_real:     Exit Function
    reParse.Pattern = "decimal":            If reParse.test(v) Then GetDefaultOrReal = dr_decimal:  Exit Function
    reParse.Pattern = "n/a|na":             If reParse.test(v) Then GetDefaultOrReal = dr_na:       Exit Function
    myUtility.MessageBox "Unrecognize default or real! " & v
    Call Print_Error_Message(Error_Info, "eFuseBdfParser", "GetDefaultOrReal", "Unrecognize default or real! " & v)

Exit Function
errHandler:
    Call Print_Error_Message(Error_Info, "eFuseBdfParser", "GetDefaultOrReal")
    If AbortTest Then Exit Function Else Resume Next
End Function

Private Function HeaderSeqCheck(searStRow As Integer) As Boolean
On Error GoTo errHandler

    If CheckHeaders(searStRow, header) Then HeaderSeqCheck = True: ColShift = 0: Exit Function
    If CheckHeaders(searStRow, header2) Then HeaderSeqCheck = True: ColShift = 0: Exit Function
    If CheckHeaders(searStRow, headerSimple) Then HeaderSeqCheck = True: ColShift = 4: Exit Function
    HeaderSeqCheck = False

Exit Function
errHandler:
    Call Print_Error_Message(Error_Info, "eFuseBdfParser", "HeaderSeqCheck")
    If AbortTest Then Exit Function Else Resume Next
End Function

Private Function CheckHeaders(searStRow As Integer, headerStr As String) As Boolean
On Error GoTo errHandler
Dim iCol As Integer, Headers As Variant, correct As Boolean
Dim cellStr As String

    correct = True
    Headers = Split(headerStr, ",")
    For iCol = 1 To UBound(Headers) + 1
        cellStr = ParsingArrayData(searStRow, iCol)
        If Not LCase(cellStr) Like LCase(Headers(iCol - 1)) Then correct = False
    Next
    CheckHeaders = correct

Exit Function
errHandler:
    Call Print_Error_Message(Error_Info, "eFuseBdfParser", "CheckHeaders")
    If AbortTest Then Exit Function Else Resume Next
End Function

Public Sub LogBdfInfo()
On Error GoTo errHandler
Dim bank As Variant, field As Variant, item As eFuseBdfField 'bank name as key
Dim bankstr As String, mybank As eFuseBdfBank: bankstr = vbNullString
Dim EarlyFused As String, deid As String, crc As String, dbits As String, pgmMode As String, fieldAtThisStage As Long
Dim vddBin As String, containsVddBin As Boolean: containsVddBin = False
Dim ids As String, containsIds As Boolean: containsIds = False

 bankstr = bankstr + myUtility.TxtFmt("Bank", 12) + myUtility.TxtFmt("Size", 6) + myUtility.TxtFmt("Field-Count", 16) _
              & myUtility.TxtFmt("Field-Count(" + stage + ")", 20) & myUtility.TxtFmt("Early-Fused", 12) & myUtility.TxtFmt("DEID", 8) & myUtility.TxtFmt("CRC", 8) _
              & myUtility.TxtFmt("BinCut", 8) & myUtility.TxtFmt("Ids", 8) & myUtility.TxtFmt("Double-Bits", 12) & myUtility.TxtFmt("Pgm-Mode", 12) + vbCrLf
              bankstr = bankstr + String(132, "-") + vbCrLf

    For Each bank In Banks.Keys
        Set mybank = Banks(bank) ' lEfuseBank.DicBinCut
        vddBin = "-": EarlyFused = "-": deid = "-": crc = "-": dbits = "-": pgmMode = "DAA": ids = "-"
        fieldAtThisStage = mybank.DicEarlyFused.Count + mybank.DicOthers.Count
        If mybank.HadEarlyFuse Then EarlyFused = mybank.DicEarlyFused.Count
        If mybank.HadDeidFuse Then deid = "V"
        If mybank.HadCrcFuse Then crc = "V"
        If mybank.DicBinCut.Count > 0 Then
            vddBin = mybank.DicBinCut.Count
            containsVddBin = True
        End If
        If mybank.DicIds.Count > 0 Then
            ids = mybank.DicIds.Count
            containsIds = True
        End If
        If mybank.DoubleBits Then dbits = "V"
        If mybank.pgmMode = pgm_JTAG Then pgmMode = "JTAG"
        bankstr = bankstr + myUtility.TxtFmt(mybank.name, 12) + myUtility.TxtFmt(mybank.FullSize, 6) + myUtility.TxtFmt(mybank.Fields.Count, 16) _
            & myUtility.TxtFmt(fieldAtThisStage, 20) & myUtility.TxtFmt(EarlyFused, 12) & myUtility.TxtFmt(deid, 8) & myUtility.TxtFmt(crc, 8) _
            & myUtility.TxtFmt(vddBin, 8) & myUtility.TxtFmt(ids, 8) & myUtility.TxtFmt(dbits, 12) & myUtility.TxtFmt(pgmMode, 12) + vbCrLf
    Next

    myUtility.WriteDlg String(132, "=") & vbCrLf & "Config """ & SelectedCfg & """ is selected!" & vbCrLf & bankstr & String(132, "-")
    If containsVddBin Then myUtility.WriteDlg "Base Voltage = " & BaseVoltage & ", Resolution = " & BaseResolution
    If containsVddBin And BaseVoltage = 0 Then Call Print_Error_Message(Error_Info, "eFuseBdfParser", "LogBdfInfo", "0mV of Base Voltage = " & BaseVoltage)
    If containsVddBin And BaseResolution = 0 Then Call Print_Error_Message(Warning_Info, "eFuseBdfParser", "LogBdfInfo", "0mV of BaseResolution = " & BaseResolution)
    If (bLumpStages) Then myUtility.WriteDlg "Lump Stages is Enabled" & vbCrLf
    myUtility.WriteDlg "EFUSE_DATABASE_REVISION = " & DatabaseVersion & vbCrLf & String(132, "=") & vbCrLf

Exit Sub
errHandler:
    Call Print_Error_Message(Error_Info, "eFuseBdfParser", "LogBdfInfo")
    If AbortTest Then Exit Sub Else Resume Next
End Sub

Public Sub DumpBdfData(path As String, Optional dsscRt As Boolean = False)
On Error GoTo errHandler
Dim bank As Variant, field As Variant, item As eFuseBdfField 'bank name as key
Dim bankstr As String: bankstr = vbNullString
Dim dateStr As String, filename As String, fieldCnt As Long, padding As String: padding = " "
Dim Delimiter As String, stageTag As String, siteTpVal As String, trimInstance As String, lotStr As String: Delimiter = vbTab
Dim site As Variant 'Carter, 20240304

    path = Replace(path & "\" & dateStr, "\\", "\")
    If Dir(path, vbDirectory) = "" Then MkDir path
    filename = Replace(path & "\BDF_Dump", "\\", "\")
    Open filename & "_AllSite_" & myUtility.currStage & ".txt" For Output As #1

    For Each bank In Banks.Keys
        bankstr = "Bank - [" & bank & "]@" + TheExec.DataManager.instancename + vbCrLf
        bankstr = bankstr + myUtility.TxtFmt("Field Name" & Delimiter, 48) & myUtility.TxtFmt("Stage" & Delimiter, 8) & myUtility.TxtFmt("Length" & Delimiter, 8)
        
        siteTpVal = vbNullString:
        For Each site In TheExec.sites
                  siteTpVal = siteTpVal & myUtility.TxtFmt("Site#" & site & " (TP)-" + SelectedCfg & Delimiter, 24)
        Next
        bankstr = bankstr & siteTpVal _
        & myUtility.TxtFmt("FieldDefault(Table)" & Delimiter, 32) _
        & myUtility.TxtFmt("DefaultOrReal" & Delimiter, 15) & myUtility.TxtFmt("SetTrim" & Delimiter, 8) & myUtility.TxtFmt("Trimmed" & Delimiter, 8) & myUtility.TxtFmt("From-Instance", 28) + vbCrLf
        
        fieldCnt = 0
        For Each field In Banks(bank).Fields
            Set item = Banks(bank).Fields(field)
            stageTag = vbNullString
            If BdfDataBase.GetRealStage(item.BlowLocation) = myUtility.currStage Then stageTag = vbNullString '">>"
            'padding = IIf((fieldCnt + 1) Mod 10 = 0, "_", " ")
            bankstr = bankstr + myUtility.TxtFmt(item.name & Delimiter, 48, padding) & myUtility.TxtFmt(stageTag & item.BlowLocation & Delimiter, 8) & myUtility.TxtFmt(item.size & Delimiter, 8, padding)
            siteTpVal = vbNullString: trimInstance = vbNullString: lotStr = vbNullString
            For Each site In TheExec.sites
                If item.Algorithm = Algorithm.alg_lotid Then
                    lotStr = "(" & Banks(bank).Decode(item, dsscRt) & ")"
                End If
                siteTpVal = siteTpVal + myUtility.TxtFmt(lotStr & IIf(item.ReadOnly, item.FieldDefault, IIf(dsscRt, item.DsscValue, IIf(myUtility.testedStages.Exists(item.BlowLocation), item.DsscValue, item.TrimAteValue))) & Delimiter, 24, padding)
                If trimInstance = "" Then trimInstance = item.TrimTestInstance
            Next
            bankstr = bankstr & siteTpVal & myUtility.TxtFmt(item.DefaulString & Delimiter, 32, padding)
            '& myUtility.TxtFmt(item.DefaultOrReal & Delimiter, 15, padding) & myUtility.TxtFmt(IIf(item.TrimFlag, "O", "X") & Delimiter, 8, padding) & myUtility.TxtFmt(IIf(item.Trimmed, "O", "X") & Delimiter, 8, padding) & myUtility.TxtFmt(trimInstance, 28) + vbCrLf
            fieldCnt = fieldCnt + 1
        Next
        bankstr = bankstr + String(150, "=") + vbCrLf
        Print #1, bankstr
    Next
    Close #1

Exit Sub
errHandler:
    Call Print_Error_Message(Error_Info, "eFuseBdfParser", "DumpBdfData")
    If AbortTest Then Exit Sub Else Resume Next
End Sub

Public Sub PrintCrcField()
On Error GoTo errHandler
Dim bank As Variant, obj As Variant, item As eFuseBdfBank, field As eFuseBdfField 'bank name as key
Dim bankstr As String, crc As Dictionary

    If Not eFusePrinted Then Exit Sub
    bankstr = vbNullString
    For Each bank In Banks.Keys
        Set item = Banks(bank)
        If item.Crcs.Count > 0 Then
            For Each obj In item.Crcs.Keys
                Set crc = item.Crcs(obj)
                Set field = item.Fields(obj)
                bankstr = bankstr + myUtility.TxtFmt("Bank = [" & bank & "]", 20, left:=True) & myUtility.TxtFmt(", Field = " & field.name, 32, left:=True) _
                    & myUtility.TxtFmt(", Stage = " + field.BlowLocation, 20, left:=True) + myUtility.TxtFmt(GetCrcRegionsStr(crc), 36, left:=True) + vbCrLf
            Next
        End If
        If item.Crcs_Comp.Count > 0 Then
            For Each obj In item.Crcs_Comp.Keys
                Set field = item.Fields(obj)
                bankstr = bankstr + myUtility.TxtFmt("Bank = [" & bank & "]", 20, left:=True) & myUtility.TxtFmt(", Field = " & field.name, 32, left:=True) _
                & myUtility.TxtFmt(", Stage = " + field.BlowLocation, 20, left:=True) + myUtility.TxtFmt("one_complement_target : " & "[" & item.Crcs_Comp(field.name) & "]", 36, left:=True) + vbCrLf
            Next
        End If
    Next
    If Len(bankstr) > 0 Then myUtility.WriteDlg bankstr

Exit Sub
errHandler:
    Call Print_Error_Message(Error_Info, "eFuseBdfParser", "PrintCrcField")
    If AbortTest Then Exit Sub Else Resume Next
End Sub

Private Function GetCrcRegionsStr(dic As Dictionary) As String
On Error GoTo errHandler
Dim bit As Variant, regionStr As String, ideal As Long, bitSt As Long
Dim lastBit As Long

    regionStr = vbNullString: lastBit = -1: ideal = 0
    For Each bit In dic.Keys
            If lastBit = -1 Then
                bitSt = bit
           ElseIf lastBit <> bit + 1 Then
                If bitSt = lastBit Then
                    regionStr = regionStr & bitSt & " "
                Else
                    regionStr = regionStr & bitSt & ":" & lastBit & " "
                End If
                If ideal = dic.Count - 1 Then regionStr = regionStr & bit
                bitSt = bit
           ElseIf ideal = dic.Count - 1 Then
                If lastBit <> bit + 1 Then
                    regionStr = regionStr & bitSt & ":" & lastBit & " " & bit
                Else
                    regionStr = regionStr & bitSt & ":" & bit
                End If
           End If
           lastBit = bit: ideal = ideal + 1
    Next
    GetCrcRegionsStr = "[" + Replace(Trim(regionStr), " ", "], [") + "]"

Exit Function
errHandler:
    Call Print_Error_Message(Error_Info, "eFuseBdfParser", "GetCrcRegionsStr")
    If AbortTest Then Exit Function Else Resume Next
End Function

Public Sub eFuseBdfInitialize()
On Error GoTo errHandler
Dim bankstr As Variant, bank As eFuseBdfBank

    If Banks.Exists(Empty) Then Banks.Remove (Empty)
    For Each bankstr In Banks.Keys
        Set bank = Banks(bankstr)
        bank.AteTrimReset stage, myUtility.testedStages, CFG_Only
        ''20221004, Add for multi blank check
        Call bank.InitialMultiCheckVarVariable
        bank.SplitReadDspWaveFlag = False
        If UCase(bank.name) = "CFG" And EFUSE_CFG_READ_BYJTAG Then
            bank.NeedJTAGRead = EFUSE_CFG_READ_BYJTAG
        ElseIf UCase(bank.name) = "ECID" And EFUSE_ECID_READ_BYJTAG Then
            bank.NeedJTAGRead = EFUSE_ECID_READ_BYJTAG
        Else
            bank.NeedJTAGRead = False
        End If
        '20240918, Add for store read pattern results to determine whether IEDA is set
        Set bank.ReadPatResut = New SiteBoolean: bank.ReadPatResut = True
    Next

Exit Sub
errHandler:
    Call Print_Error_Message(Error_Info, "eFuseBdfParser", "eFuseBdfInitialize")
    If AbortTest Then Exit Sub Else Resume Next
End Sub

Public Sub MergeCfg2BdfTable(CfgDataBase As eFuseCfgParser)
On Error GoTo errHandler
Dim lEfuseBank As New eFuseBdfBank, myfield As eFuseBdfField
Dim cfgfieldstr As Variant, cfgfield As eFuseCfgField

    If Not CfgDataBase Is Nothing Then 'Only take MainA as the base
        lEfuseBank.name = CfgBankName
        For Each cfgfieldstr In CfgDataBase.Fields.Keys
            Set myfield = New eFuseBdfField
            Set myfield.myUtility = myUtility
            Set cfgfield = CfgDataBase.Fields(cfgfieldstr)
            myfield.name = cfgfield.name
            myfield.msb = cfgfield.msb
            myfield.LSB = cfgfield.LSB
            myfield.size = cfgfield.size
            myfield.UsiMsb = -1
            myfield.UsiLsb = -1
            myfield.UsoMsb = -1
            myfield.UsoLsb = -1
            myfield.BlowLocation = cfgfield.BlowLocation
            myfield.Llimit = 0
            myfield.Hlimit = 2# ^ cfgfield.size
            myfield.Resolution = -1
            myfield.Msb1st = myfield.msb >= myfield.LSB

            'if u need to have it as early fuse, then u should set it additionally
            myfield.Algorithm = alg_app ' set to app always
            myfield.IsEarlyFuse = False
            myfield.Description = cfgfield.name
            myfield.DefaultOrReal = cfgfield.DefaultOrReal
            myfield.FieldDefault = cfgfield.FieldDefault
            myfield.DefaulString = cfgfield.DefaulString
            myfield.DataRowNum = cfgfield.DataRowNum

            If myfield.DefaultOrReal = dr_default Or myfield.DefaultOrReal = dr_decimal Then
                myfield.Llimit = myUtility.String2Dbl(myfield.FieldDefault, -1)
                myfield.Hlimit = myfield.Llimit
            End If

            myfield.DefaulString = cfgfield.DefaulString ' store the original string from table
            myfield.ReadOnly = cfgfield.ReadOnly

            lEfuseBank.AddField myfield
            Call Add2Category(lEfuseBank, myfield)
            If myfield.IsEarlyFuse Then
                lEfuseBank.HadEarlyFuse = True
            End If
        Next
        lEfuseBank.DoubleBits = CfgDataBase.DoubleBits
        lEfuseBank.FullSize = CfgDataBase.FullSize
        lEfuseBank.pgmMode = CfgDataBase.pgmMode
        BdfDataBase.Banks.Add lEfuseBank.name, lEfuseBank
        Set BdfDataBase.Bank_Cfg = lEfuseBank
    End If
    Call CollectFieldIdx
    Set BdfDataBase.Bank_Cfg.myUtility = myUtility
    
Exit Sub
errHandler:
    Call Print_Error_Message(Error_Info, "eFuseBdfParser", "MergeCfg2BdfTable")
    If AbortTest Then Exit Sub Else Resume Next
End Sub

Public Function ParseBKMInfo(BKMSheetName As String) As Boolean
On Error GoTo errHandler
Dim m_LoopRow As Long
Dim BKM_DECODE As String
Dim BKM_Name As String
Dim m_SheetMaxRow As Long
Dim m_SheetMaxCol As Long
Dim m_ParsingArr() As Variant
Dim startRow As Long
Dim opbank As eFuseBdfBank
Dim field As eFuseBdfField
Dim fuseValue As New SiteVariant
Dim tmp As String: tmp = vbNullString
Dim msg As String: msg = vbNullString

    ParseBKMInfo = True
    startRow = 0
    myUtility.Get_Sheet_Info BKMSheetName, m_SheetMaxRow, m_SheetMaxCol, m_ParsingArr

    For m_LoopRow = 1 To m_SheetMaxRow
        If (UCase(m_ParsingArr(m_LoopRow, 1)) = UCase("Decimal")) Then
            startRow = m_LoopRow
            Exit For
        End If
    Next

    If startRow = 0 Then
        msg = "Header not found in sheet " + BKMSheetName + ", please check it!!!"
        GoTo skip
    End If
    Set opbank = GetBdfBank("CFG")
    If opbank.Fields.Exists("bkm_package") Then
        Set field = opbank.Fields("bkm_package")
    ElseIf opbank.Fields.Exists("bkm_process") Then
        Set field = opbank.Fields("bkm_process")
    Else
        Call Print_Error_Message(Error_Info, "eFuseBdfParser", "ParseBKMInfo", "The CFG bank not exist bkm_package/bkm_process category but TP has " + BKMSheetName + " sheet, please check it!!!")
        GoTo skip
    End If
    For m_LoopRow = startRow + 1 To m_SheetMaxRow
        If (m_ParsingArr(m_LoopRow, 3) <> "") Then
            BKM_DECODE = Trim(m_ParsingArr(m_LoopRow, 2))
            BKM_Name = UCase(m_ParsingArr(m_LoopRow, 3))
            tmp = BKM_DECODE
            If Not myUtility.IsStrMatch(BKM_DECODE, "^b[0|1]+") Then
                msg = "Fuse value (" + BKM_DECODE + ") must be in binary format (row:" + CStr(m_LoopRow) + " col:2), please check it!!!"
                GoTo skip
            Else
                fuseValue = myUtility.Bin2HexStr(tmp, Len(BKM_DECODE) - 1)
                If field.IsOverflow(fuseValue, bMsgBox:=False) Then
                    msg = "Fuse value (" + BKM_DECODE + ") is overflow (row:" + CStr(m_LoopRow) + " col:2), please check it!!!"
                    GoTo skip
                End If
            End If

            If (UCase(m_ParsingArr(m_LoopRow, 3)) Like UCase("*unknown*")) Then gS_BKM_Unknown = BKM_Name
            If Not DicBKMMap.Exists(BKM_Name) Then
                DicBKMMap.Add BKM_Name, BKM_DECODE
            Else
                msg = "BKM sheet(" + BKMSheetName + ") has duplicate BKM info (" + BKM_Name + ")(row:" + CStr(m_LoopRow) + " col:3), please check it!!!"
                GoTo skip
            End If
        Else
            Call Print_Error_Message(Warning_Info, "eFuseBdfParser", "ParseBKMInfo", "The " + BKMSheetName + " sheet BKM info is empty !!! (row:" + CStr(m_LoopRow) + " col:3)")
        End If
    Next

Exit Function
skip:
    ParseBKMInfo = False
    Call Print_Error_Message(Error_Info, "eFuseBdfParser", "ParseBKMInfo", msg)
    TheExec.Flow.TestLimit resultVal:=999, lowVal:=1, hiVal:=1, Tname:="ParseBKMInfo"
    Exit Function
errHandler:
    Call Print_Error_Message(Error_Info, "eFuseBdfParser", "ParseBKMInfo")
    If AbortTest Then Exit Function Else Resume Next
End Function

Public Sub ParseBKMInfoFromFile()
On Error GoTo errHandler
Dim m_file As String
Dim m_lineStr As String
Dim tok() As String
Dim i As Integer
Dim BKM_DECODE As String
Dim BKM_Nmae As String

    'User should maintain this path for each project, EX: X:\Production\TMNK84\BKM\BKM_Info.txt
    m_file = vbNullString
    
    'BKM_Info.txt content
    'Staten will allocate 4 bits for BKM information
    'This will use a lookup table with a maximum of 16 options
    '
    'Decimal Binary  BKM Comment
    '0   b0000 unknown
    '1   b0001   BKM6.0
    '2   b0010   BKM6.3
    '3   b0011   BKM6.5

    If Dir(m_file) = "" Or m_file = "" Then Exit Sub

    Open m_file For Input As #1
    Do Until EOF(1)
        Line Input #1, m_lineStr
        If InStr(m_lineStr, "Decimal") Then
            For i = 1 To 16
                Line Input #1, m_lineStr
                tok = Split(m_lineStr, vbTab)

'                'tok(0) '1, tok(1) 'b0001, tok(2) 'BKM6.0
                If (UCase(tok(2)) Like UCase("*unknown*")) Then
                    gS_BKM_Unknown = UCase(tok(2))
                    BKM_DECODE = Trim(tok(1))
                    If DicBKMMap.Exists(UCase(gS_BKM_Unknown)) = False Then Call DicBKMMap.Add(gS_BKM_Unknown, BKM_DECODE)
                Else
                    BKM_Nmae = UCase(tok(2))
                    BKM_DECODE = Trim(tok(1))
                    If DicBKMMap.Exists(BKM_Nmae) = False Then Call DicBKMMap.Add(BKM_Nmae, BKM_DECODE)
                End If
            Next
       End If
    Loop
    Close #1

Exit Sub
errHandler:
    Call Print_Error_Message(Error_Info, "eFuseBdfParser", "ParseBKMInfoFromFile")
    If AbortTest Then Exit Sub Else Resume Next
End Sub

Public Sub ObtainBankStageMaskWave()
On Error GoTo errHandler
Dim bank As eFuseBdfBank
Dim bankstr As Variant
Dim fieldStr As Variant
Dim field As eFuseBdfField
Dim stBit As Long, spBit As Long
Dim iBit As Variant
Dim MaskWave As New DSPWave
Dim EarlyTestedMaskWave As New DSPWave
Dim PostTestedMaskWave As New DSPWave
Dim MaskWaveMultiTyp As New Dictionary
Dim EarlyMaskWave As New DSPWave
Dim dvArray() As Long
Dim dvWave As New DSPWave
Dim tmp As Long
Dim fieldStage As String
Dim jobType As Variant
Dim i As Long
Dim bitStep As Long, binA() As New SiteLong, bitDealt As Long
Dim site As Variant 'Carter, 20240304

    If firstrun_2 = True Then
        For Each bankstr In Banks.Keys
            Set bank = Banks(CStr(bankstr))
            MaskWave.CreateConstant 0, bank.FullSize, DspLong
            EarlyTestedMaskWave.CreateConstant 0, bank.FullSize, DspLong
            PostTestedMaskWave.CreateConstant 0, bank.FullSize, DspLong
            myUtility.IniDictionary MaskWaveMultiTyp, True
    
            For Each jobType In bank.MultiTypeJobList.Keys
                MaskWaveMultiTyp.Add jobType, MaskWave.Copy
            Next jobType
            
            EarlyMaskWave.CreateConstant 0, bank.FullSize, DspLong
            dvWave.CreateConstant 0, bank.FullSize, DspLong
         
            
            For Each fieldStr In bank.Fields.Keys
                Set field = bank.Fields(fieldStr)
                fieldStage = BdfDataBase.GetRealStage(field.BlowLocation)
                fieldStage = Replace(UCase(fieldStage), "_EARLY", "")
                
                stBit = IIf(field.LSB < field.msb, field.LSB, field.msb)
                spBit = IIf(field.LSB < field.msb, field.msb, field.LSB)
                bitStep = 1
                    If stBit > spBit Then
                        bitStep = -1: 'LSB first, if doing CRC, it needs to use the MSB first
                    End If
                
                If field.IsEarlyFuse Then
                    For iBit = stBit To spBit
                        EarlyMaskWave.Element(iBit) = 1
                    Next
                Else
                    If field.BlowLocation = stage Then
                        For iBit = stBit To spBit
                            MaskWave.Element(iBit) = 1
                        Next
                    ElseIf bank.MultiTypeJobList.Exists(field.BlowLocation) Then
                        For iBit = stBit To spBit
                            MaskWaveMultiTyp(field.BlowLocation).Element(iBit) = 1
                        Next
                    End If
                End If
                If Not TheExec.CurrentJob Like "CP1" Then
                    If GlbUtility.testedStages.Exists(fieldStage) Then
                        For iBit = stBit To spBit
                            ' Before testedmask of early stage.
                            ' Ex: Programming stage -> TestedMask
                            ' 1.CP1_Early -> empty, 2.CP2_Early -> all CP1 mask
                            EarlyTestedMaskWave.Element(iBit) = 1
                        Next
                    End If
                End If
                bitDealt = 0
                binA = bank.GetBinA(DafaultValue, field, , True, Not field.IsEarlyFuse)     '20210202
                For iBit = stBit To spBit Step bitStep
                    For Each site In TheExec.sites.Selected
                        tmp = binA(bitDealt)
                        Exit For
                    Next site
                    dvWave.ElementLite(iBit) = tmp
                    bitDealt = bitDealt + 1
                Next
                
                '20210514, Add for bank only has CP1_early, no CP1 issue
                If (fieldStage = TheExec.CurrentJob) Then
                    bank.JobExistInBank = True
                End If
            Next
    
            Set PostTestedMaskWave = EarlyMaskWave.Copy
            ' Inculde testedmask of early stage for post fuse.
            ' Ex: Programming stage -> TestedMask
            ' 1.CP1 -> CP1_Early, 2.CP2 -> all CP1 + CP2_Early
            PostTestedMaskWave = PostTestedMaskWave.BitwiseOr(EarlyTestedMaskWave)
            
            For i = 0 To (bank.FullSize - 1)
                If PostTestedMaskWave.Element(i) <> 0 Then
                    bank.PostTestedStageMaskCheck = True
                End If
                
                If EarlyTestedMaskWave.Element(i) <> 0 Then
                    bank.EarlyTestedStageMaskCheck = True
                    bank.NonFirstfuse = True
                End If
                
                If bank.PostTestedStageMaskCheck And bank.EarlyTestedStageMaskCheck Then Exit For
            Next
    
            Set bank.EarlyStageMask = EarlyMaskWave.Copy
            Set bank.StageMask = MaskWave.Copy
            Set bank.DefaultWave = dvWave.Copy
            Set bank.StageMaskMultiTyp = MaskWaveMultiTyp
            Set bank.EarlyTestedstagemask = EarlyTestedMaskWave.Copy
            Set bank.PostTestedstagemask = PostTestedMaskWave.Copy
        Next
        firstrun_2 = False
    End If

Exit Sub
errHandler:
    Call Print_Error_Message(Error_Info, "eFuseBdfParser", "ObtainBankStageMaskWave")
    If AbortTest Then Exit Sub Else Resume Next
End Sub

Public Sub PseudoFuseSetUp()
On Error GoTo errHandler
Dim bank As eFuseBdfBank
Dim bankstr As Variant
Dim tmpDspWave As New DSPWave
Dim m_BC_arrary() As String

    ''202011xx
    PseudoFuseEnable = False
    If Not TheExec.enableWord("Pgm2File") Then Exit Sub
    PseudoFuseEnable = True
    
    PgmName_ForPseudoFuse = TheExec.TestProgram.name
    
    If TheExec.RunMode = runModeProduction Then
        IsOI_ForPseudoFuse = "YES"
    Else
        IsOI_ForPseudoFuse = "NO"
    End If
    
    m_BC_arrary = Split(Version_Vdd_Binning_Def, ".")
    
    For Each bankstr In Banks.Keys
        Set bank = Banks(CStr(bankstr))
                tmpDspWave.CreateConstant 0, bank.FullSize, DspLong
        Set bank.PseudoFuseDspWave = tmpDspWave.Copy
    Next
    
    '20211210, Modify for FT pseudo fuse with real package
    If UCase(TheExec.CurrentJob) Like "*FT*" And gB_Package_PsudoFuse = True Then
        PseudoFuseFileName = "TGold_" & LotID_ForPseudoFuse & "_" & _
                                  "M" & CStr(Format(BdfDataBase.Chip_rev_major, "0")) & "_" & _
                                  "m" & CStr(Format(BdfDataBase.Chip_rev_minor, "0")) & "_" & _
                                  "V" & CStr(Format(BdfDataBase.Fuse_Revision, "00"))
    Else
        PseudoFuseFileName = "TGold_" & LotID_ForPseudoFuse & "_" & _
                                  Format(WaferID_ForPseudoFuse, "00") & "_" & _
                                  "M" & CStr(Format(BdfDataBase.Chip_rev_major, "0")) & "_" & _
                                  "m" & CStr(Format(BdfDataBase.Chip_rev_minor, "0")) & "_" & _
                                  "V" & CStr(Format(BdfDataBase.Fuse_Revision, "00")) '& "_" & _
                                  '"BCV" & m_BC_arrary(0) '& _
                                  '"P" & m_BC_arrary(1)
    End If
    

Exit Sub
errHandler:
    Call Print_Error_Message(Error_Info, "eFuseBdfParser", "PseudoFuseSetUp")
    If AbortTest Then Exit Sub Else Resume Next
End Sub

Public Sub DumpFusedDataToFile(OutputFile As String)
On Error GoTo errHandler
Dim BankValueStr As String
Dim bank As New eFuseBdfBank
Dim field As New eFuseBdfField
Dim m_site As Variant
Dim bankstr As Variant
Dim fieldStr As Variant
Dim tmp As String: tmp = vbNullString

    Open OutputFile For Append As #42

    For Each m_site In TheExec.sites
        BankValueStr = vbNullString
        If UCase(TheExec.CurrentJob) Like "*FT*" Then
            If gB_Package_PsudoFuse = True Then
                BankValueStr = HramLotId & "_" & HramWaferId & "_" & HramXCoord & "_" & HramYCoord
            Else
                BankValueStr = TheExec.Datalog.Setup.WaferSetup.GetXCoord(m_site) & "_" & TheExec.Datalog.Setup.WaferSetup.GetYCoord(m_site)
            End If
        Else
            BankValueStr = TheExec.Datalog.Setup.WaferSetup.GetXCoord(m_site) & "_" & TheExec.Datalog.Setup.WaferSetup.GetYCoord(m_site)
        End If

        For Each bankstr In Banks.Keys
            If bankstr Like "*cmp*" Then GoTo skiploop
            Set bank = Banks(bankstr)
            BankValueStr = BankValueStr & "," & CStr(bankstr) & "_Start"
            For Each fieldStr In bank.Fields
                Set field = bank.Fields(fieldStr)
                If field.Algorithm = alg_ids Then
                    BankValueStr = BankValueStr & "," & CStr(field.FuseMeasureValue * 1000)
                ElseIf field.Algorithm = alg_vddbin Then
                    BankValueStr = BankValueStr & "," & field.FuseMeasureValue
                ElseIf field.Algorithm = alg_base Then
                    tmp = CStr(field.BaseValue_V * 1000)
                    BankValueStr = BankValueStr & "," & tmp
                ElseIf field.Algorithm = alg_lotid Then
                    BankValueStr = BankValueStr & "," & bank.Decode(field, True)
                ElseIf field.size >= 32 Or field.Algorithm = alg_crc Then
                    BankValueStr = BankValueStr & "," & field.DsscValue
                Else
                    BankValueStr = BankValueStr & "," & field.DsscDecValue
                End If
            Next fieldStr
skiploop:
        Next bankstr
        Print #42, BankValueStr & "," & "EFUSE_END" & "," & PgmName_ForPseudoFuse & "," & IsOI_ForPseudoFuse
    Next m_site
    Close #42
    
Exit Sub
errHandler:
    Call Print_Error_Message(Error_Info, "eFuseBdfParser", "DumpFusedDataToFile")
    If AbortTest Then Exit Sub Else Resume Next
End Sub


Public Sub GetFileData(dic As Dictionary)
On Error GoTo errHandler
Dim m_lineStr As String: m_lineStr = vbNullString
Dim key As String: key = vbNullString
Dim addr As Variant
Dim data As Variant
Dim bitdef As Variant
Dim IsOIRunning As Variant

    Line Input #1, m_lineStr
    m_lineStr = Trim(m_lineStr)
    addr = InStr(1, m_lineStr, ",")
    key = mid(m_lineStr, 1, addr - 1)
    data = mid(m_lineStr, addr + 1)
    IsOIRunning = Split(data, ",")
        
    If IsOI_ForPseudoFuse = "YES" Then
        If IsOIRunning(UBound(IsOIRunning)) = IsOI_ForPseudoFuse Or IsOIRunning(UBound(IsOIRunning)) = "Is_OI_Running" Then
            GlbUtility.AddDictionary key, data, dic
        End If
    Else
        GlbUtility.AddDictionary key, data, dic
    End If

Exit Sub
errHandler:
    Call Print_Error_Message(Error_Info, "eFuseBdfParser", "GetFileData")
    If AbortTest Then Exit Sub Else Resume Next
End Sub

Private Sub AddRVSegment()
On Error GoTo errHandler
Dim iBit As Long
Dim stBit As Long, spBit As Long, bitStep As Long, binA() As New SiteLong, bitDealt As Long
Dim bank As eFuseBdfBank, bankstr As Variant
Dim fieldStr As Variant, field As eFuseBdfField
Dim divider As Long: divider = 32
Dim jobName As String: jobName = vbNullString
Dim indexLoop As Long
Dim idx0 As Long, idx1 As Long

    bankstr = "CFG"
    If Banks.Exists(bankstr) Then
        Set bank = Banks(CStr(bankstr))
        If bank.DoubleBits Then divider = 16
        ReDim segFlagArr(bank.FullSize / divider - 1)
        
        For Each fieldStr In bank.Fields.Keys
            Set field = bank.Fields(fieldStr)
            jobName = field.BlowLocation
            If EFUSE_RV_PATTERN_FULL Then jobName = GetRealStage(field.BlowLocation)
            If jobName = TheExec.CurrentJob And field.DefaultOrReal <> dr_default And field.Algorithm <> alg_cond Then
                bitDealt = field.size - 1
                stBit = field.LSB: spBit = field.msb: bitStep = 1: ' MSB 1st
                If field.Msb1st Then
                     bank.GetBitStSp field, stBit, spBit
                ElseIf stBit > spBit Then
                     bitStep = -1: 'LSB first, if doing CRC, it needs to use the MSB first
                Else
                     myUtility.MessageBox "Something wrong in dealing BDF GetDspWave"
                End If
    
                idx0 = Floor(stBit / divider) 'please check floor
                idx1 = Floor(spBit / divider) 'please check floor
                segFlagArr(idx0) = 1
    
                If (idx0 <> idx1) Then
                    For indexLoop = 1 To idx1 - idx0
                        segFlagArr(idx0 + indexLoop) = 1
                    Next
                End If
            End If
        Next fieldStr
    Else
        ReDim segFlagArr(0)
    End If

Exit Sub
errHandler:
    Call Print_Error_Message(Error_Info, "eFuseBdfParser", "AddRVSegment")
    If AbortTest Then Exit Sub Else Resume Next
End Sub

'202405 Add DRAM table duplicate condition check
Public Function ParseDram(DramSheetName As String) As Boolean
On Error GoTo errHandler
Dim m_LoopRow As Long
Dim m_LoopCol As Long
Dim m_Loop As Long
Dim m_SheetMaxRow As Long
Dim m_SheetMaxCol As Long
Dim m_ParsingArr() As Variant
Dim enableWord As String: enableWord = vbNullString
Dim m_catename As String: m_catename = vbNullString
Dim m_mrrStr As String: m_mrrStr = vbNullString
Dim m_array() As String
Dim dicEnableWdData As Dictionary
Dim dicDramDescriptor As Dictionary
Dim dicCheckMrrData As Dictionary
Dim MrrEnableWd As Variant
Dim startRow As Long, startCol As Long
Dim enableWdStartRow As Long, enableWdStartCol As Long, enableWdEndCol As Long
Dim KeyStartCol As Long
Dim findHeader As Boolean: findHeader = False
Dim headerName As String: headerName = vbNullString
Dim startFuseRow As Integer
Dim msg As String: msg = vbNullString
Dim enableWordPrefix As String: enableWordPrefix = "DRAM_"
Dim opbank As eFuseBdfBank
Dim AllFuseDataArr() As Variant 'Check duplicate
Dim KeyOrder As Variant         'Check duplicate
Dim mrrStr As String: mrrStr = vbNullString
Dim tmpstr As String: tmpstr = vbNullString
Dim mrrVal As Double

    
    myUtility.Get_Sheet_Info DramSheetName, m_SheetMaxRow, m_SheetMaxCol, m_ParsingArr
    
    Set dicCheckMrrData = New Dictionary
    dicCheckMrrData.compareMode = TextCompare
    dicCheckMrrData.RemoveAll
            
    '1.find start row/col.
    For m_LoopRow = 1 To m_SheetMaxRow
        If UCase(m_ParsingArr(m_LoopRow, 1)) Like UCase("Category") Then
            startRow = m_LoopRow
        ElseIf startRow <> 0 And UCase(m_ParsingArr(m_LoopRow, 1)) Like UCase("Fuse") Then
            startFuseRow = m_LoopRow
            Exit For
        End If
    Next m_LoopRow
    
    For m_LoopCol = 1 To m_SheetMaxCol
        If UCase(m_ParsingArr(startRow, m_LoopCol)) Like UCase("Category") And startCol < 1 Then
            startCol = m_LoopCol
            KeyStartCol = m_LoopCol + 1
            findHeader = True
        ElseIf UCase(m_ParsingArr(startRow, m_LoopCol)) Like enableWordPrefix & "*" And enableWdStartCol < 1 Then
            enableWdStartRow = startRow
            enableWdStartCol = m_LoopCol
        ElseIf UCase(m_ParsingArr(startRow, m_LoopCol)) Like UCase("Notes") Or m_ParsingArr(startRow, m_LoopCol) Like Empty Then
            enableWdEndCol = m_LoopCol - 1
            Exit For
        End If
    Next m_LoopCol
    
    
    If Not findHeader Then
        msg = "The sheet (" + DramSheetName + ") could not find start row/col (keyword:Category), please check it!!!"
        GoTo skip
    End If
    
    ReDim AllFuseDataArr(enableWdEndCol - enableWdStartCol) 'Check duplicate
    
    '2.find all header col/row.
    For m_LoopCol = enableWdStartCol To enableWdEndCol
        KeyOrder = m_LoopCol - enableWdStartCol 'Check duplicate'
        Set dicEnableWdData = New Dictionary
        dicEnableWdData.compareMode = TextCompare
        Set dicDramDescriptor = New Dictionary
        dicDramDescriptor.compareMode = TextCompare
        ReDim m_array(m_SheetMaxRow - startFuseRow)
        m_mrrStr = vbNullString
        enableWord = CStr(m_ParsingArr(enableWdStartRow, m_LoopCol))
        
        For m_LoopRow = 1 To m_SheetMaxRow
            If m_ParsingArr(m_LoopRow, startCol) <> vbNullString Then
                headerName = LCase(m_ParsingArr(m_LoopRow, startCol))
                
                Select Case headerName
                    Case "descriptor":
                        If m_ParsingArr(m_LoopRow, KeyStartCol) = vbNullString Then
                            msg = "The catagory(" & headerName & ") could not find any key(Col:" & KeyStartCol & "), please check it!!!"
                            GoTo skip
                        Else
'                            If m_ParsingArr(m_LoopRow, m_LoopCol) = vbNullString Then
'                                msg = "The Key(" & CStr(m_ParsingArr(m_LoopRow, KeyStartCol)) & ") could not find any item(" & enableWord & "), please check it!!!"
'                                GoTo skip
'                            Else
                                dicDramDescriptor.Add CStr(m_ParsingArr(m_LoopRow, KeyStartCol)), CStr(m_ParsingArr(m_LoopRow, m_LoopCol))
'                            End If
                        End If
                        
                    Case "mrr results":
                        If m_ParsingArr(m_LoopRow, KeyStartCol) = vbNullString Then
                            msg = "The catagory(" & headerName & ") could not find any key(Col:" & KeyStartCol & "), please check it!!!"
                            GoTo skip
                        Else
                            If m_ParsingArr(m_LoopRow, m_LoopCol) = vbNullString Then
                                msg = "The Key(" & CStr(m_ParsingArr(m_LoopRow, KeyStartCol)) & ") could not find any item(" & enableWord & "), please check it!!!"
                                GoTo skip
                            Else
                                'Pending for afer disscussing.
'                                If UCase(m_ParsingArr(m_LoopRow, KeyStartCol)) Like "MRR*" Then
'                                    If m_mrrStr <> vbNullString Then m_mrrStr = m_mrrStr + "_"
'                                    m_mrrStr = m_mrrStr + Trim(CStr(m_ParsingArr(m_LoopRow, m_LoopCol)))
'                                End If
                                mrrStr = Trim(m_ParsingArr(m_LoopRow, m_LoopCol))
                                If myUtility.IsNumber(mrrStr) Then
                                    mrrVal = myUtility.String2Long(mrrStr)
                                    dicEnableWdData.Add CStr(m_ParsingArr(m_LoopRow, KeyStartCol)), mrrVal
                                Else
                                    tmpstr = Replace(Replace(Replace(Replace(mrrStr, ">", ""), "=", ""), "<", ""), " ", "")
                                    If myUtility.IsNumber(tmpstr) Then
                                        mrrStr = Replace(mrrStr, tmpstr, "," & myUtility.String2Long(tmpstr))
                                    Else
                                        msg = "MRR value(" & mrrStr & ") format need to check"
                                        GoTo skip
                                    End If
                                    dicEnableWdData.Add CStr(m_ParsingArr(m_LoopRow, KeyStartCol)), mrrStr
                                End If
                            End If
                        End If
                        
                    Case "fuse":
                        If m_ParsingArr(m_LoopRow, KeyStartCol) = vbNullString Then
                            msg = "The catagory(" + headerName + ") could not find any key(Col:" & KeyStartCol & "), please check it!!!"
                            GoTo skip
                        Else
                            If m_ParsingArr(m_LoopRow, m_LoopCol) = vbNullString Then
                                msg = "The Key(" & CStr(m_ParsingArr(m_LoopRow, KeyStartCol)) & ") could not find any item(" & enableWord & "), please check it!!!"
                                GoTo skip
                            Else
                                m_array(m_LoopRow - startFuseRow) = m_ParsingArr(m_LoopRow, m_LoopCol)
                                AllFuseDataArr(KeyOrder) = AllFuseDataArr(KeyOrder) + CStr(m_ParsingArr(m_LoopRow, m_LoopCol)) 'Check duplicate
                            End If
                        End If
                        
                    Case Else
                        If m_LoopRow > enableWdStartRow Then
                            msg = "The catagory(" & headerName & ") is not recognized, please check it!!!"
                            GoTo skip
                        End If
                End Select
            Else
                If m_ParsingArr(m_LoopRow, startCol) = vbNullString Then
                    msg = "The catagory is empty (row:" & m_LoopRow & "), please check it!!!"
                    GoTo skip
                End If
            End If
        Next m_LoopRow
        dicEnableWdData.Add "Descriptor", dicDramDescriptor
        dicEnableWdData.Add "fusedata", m_array
                
        If DicDramMap.Exists(enableWord) = True Then DicDramMap.Remove enableWord
        Call DicDramMap.Add(enableWord, dicEnableWdData)
        
        '3.Check each mrr data is unique or not.
'        If Not dicCheckMrrData.Exists(m_mrrStr) Then
'            dicCheckMrrData.Add m_mrrStr, enableWord
'        Else
'            msg = "The mrr data(" & m_mrrStr & ") is duplicate and the enableWd is " & enableWord & ", please check it!!!"
'            GoTo skip
'        End If
        AllFuseDataArr(KeyOrder) = Replace(AllFuseDataArr(KeyOrder), "_", "") 'Check duplicate
    Next m_LoopCol
    
    '''''''''''''''''Check duplicate''''''''''''''''''''
    Dim BeCompareOne As Variant
    Dim CompareOne As Variant
    'Stop
    For BeCompareOne = 0 To (enableWdEndCol - enableWdStartCol - 1)
        For CompareOne = BeCompareOne + 1 To (enableWdEndCol - enableWdStartCol)
            If AllFuseDataArr(BeCompareOne) = AllFuseDataArr(CompareOne) Then
                Call Print_Error_Message(Warning_Info, "eFuseBdfParser", "ParseDram", "DRAM_Table " & m_ParsingArr(1, BeCompareOne + 3) & " and " & m_ParsingArr(1, CompareOne + 3) & " fuse value are duplicate, please check!!!")
            End If
        Next CompareOne
    Next BeCompareOne
    '''''''''''''''''Check duplicate''''''''''''''''''''
    
    For m_LoopRow = startFuseRow To m_SheetMaxRow
        m_array(m_LoopRow - startFuseRow) = m_ParsingArr(m_LoopRow, KeyStartCol)
    Next m_LoopRow
    If DicDramMap.Exists("fieldname") = True Then DicDramMap.Remove ("fieldname")
    Call DicDramMap.Add("fieldname", m_array)
    DramFieldSize = UBound(DicDramMap("fieldname")) + 1
    
    '4.Check the fields of dram table definition in BDF table or not.
    Set opbank = GetBdfBank("CFG")
    For m_Loop = 0 To DramFieldSize - 1
        m_catename = DicDramMap.item("fieldname")(m_Loop)
        If Not opbank.Fields.Exists(m_catename) Then
            msg = "The fields(" & m_catename & ") of dram table does not exist in the BDF table, please check it!!!"
            GoTo skip
        End If
    Next m_Loop
    
    ParseDram = True

Exit Function
skip:
    Call Print_Error_Message(Error_Info, "eFuseBdfParser", "ParseDram", msg)
    ParseDram = False
    TheExec.Flow.TestLimit resultVal:=999, lowVal:=1, hiVal:=1, Tname:="ParseDram"
    Exit Function
errHandler:
    Call Print_Error_Message(Error_Info, "eFuseBdfParser", "ParseDram")
    If AbortTest Then Exit Function Else Resume Next
End Function


Private Function ProcessMrrData(m_ParsingArr As Variant, m_LoopRow As Long, startHipCol As Integer, endHipCol As Integer, mrr5Col As Integer, mrr8Col As Integer, mrr6Col As Integer) As String()
On Error GoTo errHandler
Dim m_LoopCol As Integer
Dim m_hiparray() As String
Dim i As Variant
Dim tmp As String: tmp = vbNullString
    
    ReDim m_hiparray(endHipCol - startHipCol)
    For m_LoopCol = startHipCol To endHipCol
        tmp = Trim(CStr(m_ParsingArr(m_LoopRow, m_LoopCol)))
        
        If m_LoopCol = mrr5Col Then
            m_hiparray(0) = tmp
        ElseIf m_LoopCol = mrr8Col Then
            m_hiparray(1) = tmp
        ElseIf m_LoopCol = mrr6Col Then
            m_hiparray(2) = tmp
        ElseIf tmp <> vbNullString Then
            For i = 2 To UBound(m_hiparray)
                If (m_hiparray(i) = vbNullString) Then
                    m_hiparray(i) = tmp
                    Exit For
                End If
            Next i
        End If
    Next

    ProcessMrrData = m_hiparray

Exit Function
errHandler:
    Call Print_Error_Message(Error_Info, "eFuseBdfParser", "ProcessMrrData")
    If AbortTest Then Exit Function Else Resume Next
End Function

'202405 Jack Add new agrument for JudgeDRAMType_T
Public Sub ProcessHipDram(ByRef ChkStepFlag As Boolean, Optional enableWd As Variant = vbNullString, Optional siteDiffFromDSSC As Boolean = False, Optional siteVal As SiteBoolean)
On Error GoTo errHandler
Dim enableWord As Variant, siteControlArray As Variant
Dim k As Integer
Dim msg As String: msg = vbNullString
Dim site As Variant 'Carter, 20240304

    If Not siteDiffFromDSSC Then
        DicDramHipMap.RemoveAll
    End If
    If enableWd = vbNullString Then
        For Each enableWord In DicDramMap.Keys
            If TheExec.Flow.enableWord(enableWord) Then
                ChkStepFlag = True
                If DicDramMap(enableWord)("sitecontrol") = vbNullString Then
                    For Each site In TheExec.sites
                        If Not DicDramHipMap.Exists(site) Then
                            DicDramHipMap.Add site, DicDramMap(enableWord)
                        Else
                            msg = "The enable word (" + enableWord + ") has duplicate site ( " + CStr(site) + " )"
                            GoTo skip
                        End If
                        
                    Next site
                Else
                    siteControlArray = Split(DicDramMap(enableWord)("sitecontrol"), ",")
                    For k = 0 To UBound(siteControlArray)
                        For Each site In TheExec.sites
                            If CInt(siteControlArray(k)) = CInt(site) Then
                                If Not DicDramHipMap.Exists(site) Then
                                    DicDramHipMap.Add CInt(siteControlArray(k)), DicDramMap(enableWord)
                                Else
                                    msg = "The enable word (" + enableWord + ") has duplicate site ( " + siteControlArray(k) + " )"
                                    GoTo skip
                                End If
                            End If
                        Next site
                    Next k
                End If
            End If
        Next enableWord
    Else
        If TheExec.Flow.enableWord(enableWd) Then
            ChkStepFlag = True
            If DicDramMap(enableWd)("sitecontrol") = vbNullString Then
                If siteDiffFromDSSC Then
                    For Each site In TheExec.sites
                        If siteVal(site) Then
                            If Not DicDramHipMap.Exists(site) Then
                                DicDramHipMap.Add site, DicDramMap(enableWd)
                            Else
                                msg = "The enable word (" + enableWd + ") has duplicate site ( " + CStr(site) + " )"
                                GoTo skip
                            End If
                        End If
                    Next site
                Else
                    For Each site In TheExec.sites
                        If Not DicDramHipMap.Exists(site) Then
                            DicDramHipMap.Add site, DicDramMap(enableWd)
                        Else
                            msg = "The enable word (" + enableWd + ") has duplicate site ( " + CStr(site) + " )"
                            GoTo skip
                        End If
                    Next site
                End If
            Else
                siteControlArray = Split(DicDramMap(enableWd)("sitecontrol"), ",")
                For k = 0 To UBound(siteControlArray)
                    If siteDiffFromDSSC Then
                        For Each site In TheExec.sites
                            If (CInt(siteControlArray(k)) = CInt(site)) And siteVal(site) Then
                                If Not DicDramHipMap.Exists(site) Then
                                    DicDramHipMap.Add CInt(siteControlArray(k)), DicDramMap(enableWd)
                                Else
                                    msg = "The enable word (" + enableWd + ") has duplicate site ( " + siteControlArray(k) + " )"
                                    GoTo skip
                                End If
                            Else
                                'do nothing
                            End If
                        Next site
                    Else
                        For Each site In TheExec.sites
                            If CInt(siteControlArray(k)) = CInt(site) Then
                                If Not DicDramHipMap.Exists(site) Then
                                    DicDramHipMap.Add CInt(siteControlArray(k)), DicDramMap(enableWd)
                                Else
                                    msg = "The enable word (" + enableWd + ") has duplicate site ( " + siteControlArray(k) + " )"
                                    GoTo skip
                                End If
                            End If
                        Next site
                    End If
                Next k
            End If
        End If
    End If
    
Exit Sub
skip:
    Call Print_Error_Message(Error_Info, "eFuseBdfParser", "ProcessHipDram", msg)
    TheExec.Flow.TestLimit resultVal:=999, lowVal:=1, hiVal:=1, Tname:="ProcessHipDram"
    Exit Sub
errHandler:
    Call Print_Error_Message(Error_Info, "eFuseBdfParser", "ProcessHipDram")
    If AbortTest Then Exit Sub Else Resume Next
End Sub

Public Function ParseCPM(CPMSheetName As String) As Boolean
On Error GoTo errHandler
Dim m_LoopRow As Long
Dim m_LoopCol As Long
Dim m_SheetMaxRow As Long
Dim m_SheetMaxCol As Long
Dim m_ParsingArr() As Variant
Dim startRow As Long, startFlagCol As Long, idx As Long
Dim Col As Variant
Dim tmpstr As String: tmpstr = vbNullString
Dim headerName As String: headerName = vbNullString
Dim fieldName As String, bankName As String
Dim fuseColDic As New Dictionary
Dim flagColDic As New Dictionary
Dim opbank As eFuseBdfBank
Dim field As eFuseBdfField
Dim keyword As String
Dim m_array() As String
Dim m_flagarray() As String
Dim msg As String: msg = vbNullString

    myUtility.Get_Sheet_Info CPMSheetName, m_SheetMaxRow, m_SheetMaxCol, m_ParsingArr

    For m_LoopRow = 1 To m_SheetMaxRow
        If startRow <> 0 Then Exit For
        For m_LoopCol = 1 To m_SheetMaxCol
            headerName = LCase(m_ParsingArr(m_LoopRow, m_LoopCol))
            If headerName Like "flag*" Then
                flagColDic.Add m_LoopCol, headerName
                If startRow = 0 And startFlagCol = 0 Then
                    startRow = m_LoopRow
                    startFlagCol = m_LoopCol
                End If
            ElseIf headerName <> "" Then
                fuseColDic.Add m_LoopCol, headerName
            End If
        Next
    Next
    If startRow = 0 Then
        msg = "The sheet (" + CPMSheetName + ") could not find start row/col (keyword:flag*), please check it!!!"
        GoTo skip
    ElseIf fuseColDic.Count = 0 Then
        msg = "The sheet (" + CPMSheetName + ") could not find any fuse column, please check it!!!"
        GoTo skip
    Else
        ReDim m_flagarray(flagColDic.Count - 1)
        ReDim m_array(fuseColDic.Count - 1)
    End If

    DicCPMMap.RemoveAll
    For m_LoopRow = startRow To m_SheetMaxRow
        If m_LoopRow = startRow Then
            idx = 0
            For Each Col In fuseColDic
                tmpstr = m_ParsingArr(m_LoopRow, Col)
                If tmpstr = "" Then
                    msg = "The sheet (" + CPMSheetName + ") can not exist empty bank name (row:" + CStr(m_LoopRow) + " col:" + CStr(Col) + "), please check it!!!"
                    GoTo skip
                Else
                    If GetBdfBank(tmpstr) Is Nothing Then
                        msg = "The sheet (" + CPMSheetName + ") has bank name not recognized issue (row:" + CStr(m_LoopRow) + " col:" + CStr(Col) + "), please check it!!!"
                        GoTo skip
                    Else
                        m_array(idx) = tmpstr
                        idx = idx + 1
                    End If
                End If
            Next
            keyword = "bank"
            If Not DicCPMMap.Exists(keyword) Then DicCPMMap.Add keyword, m_array
        ElseIf m_LoopRow = startRow + 1 Then
            idx = 0
            For Each Col In flagColDic
                tmpstr = m_ParsingArr(m_LoopRow, Col)
                If tmpstr = "" Then
                    msg = "The sheet (" + CPMSheetName + ") can not exist empty flag name (row:" + CStr(m_LoopRow) + " col:" + CStr(Col) + "), please check it!!!"
                    GoTo skip
                Else
                    m_flagarray(idx) = tmpstr
                    idx = idx + 1
                End If
            Next
            keyword = "flag"
            If Not DicCPMMap.Exists(keyword) Then DicCPMMap.Add keyword, m_flagarray

            idx = 0
            For Each Col In fuseColDic
                bankName = m_ParsingArr(startRow, Col)
                Set opbank = GetBdfBank(bankName)
                fieldName = m_ParsingArr(m_LoopRow, Col)
                If fieldName = "" Then
                    msg = "The sheet (" + CPMSheetName + ") can not exist empty category name (row:" + CStr(m_LoopRow) + " col:" + CStr(Col) + "), please check it!!!"
                    GoTo skip
                Else
                    If opbank.Fields.Exists(fieldName) Then
                        m_array(idx) = fieldName
                        idx = idx + 1
                    Else
                        msg = "The sheet (" + CPMSheetName + ") has unknow category name (" + fieldName + ") in bank " + bankName + ", please check it!!!"
                        GoTo skip
                    End If
                End If
            Next
            keyword = "field"
            If Not DicCPMMap.Exists(keyword) Then DicCPMMap.Add keyword, m_array
        Else
            keyword = vbNullString
            If m_ParsingArr(m_LoopRow, startFlagCol) <> "" Then
                For Each Col In flagColDic
                    tmpstr = LCase(m_ParsingArr(m_LoopRow, Col))
                    If tmpstr = "true" Or tmpstr = "false" Then
                        keyword = keyword + tmpstr + ","
                    Else
                        msg = "The sheet (" + CPMSheetName + ") has unknow flag status (" + tmpstr + ") (row:" + CStr(m_LoopRow) + " col:" + CStr(Col) + "), please check it!!!"
                        GoTo skip
                    End If
                Next
                idx = 0
                For Each Col In fuseColDic
                    tmpstr = m_ParsingArr(m_LoopRow, Col)
                    If tmpstr <> "" Then
                        m_array(idx) = tmpstr
                        idx = idx + 1
                    Else
                        msg = "The sheet (" + CPMSheetName + ") has empty fuse value (row:" + CStr(m_LoopRow) + " col:" + CStr(Col) + "), please check it!!!"
                        GoTo skip
                    End If
                Next
                
                If DicCPMMap.Exists(keyword) Then
                    msg = "The sheet (" + CPMSheetName + ") has duplicate flag status combinations (row:" + CStr(m_LoopRow) + "), please check it!!!"
                    GoTo skip
                Else
                    DicCPMMap.Add keyword, m_array
                End If
            Else
                Call Print_Error_Message(Warning_Info, "eFuseBdfParser", "ParseCPM", "The " + CPMSheetName + " sheet flag status is empty !!! (row:" + CStr(m_LoopRow) + ")")
            End If
        End If
    Next
    ParseCPM = True

Exit Function
skip:
    Call Print_Error_Message(Error_Info, "eFuseBdfParser", "ParseCPM", msg)
    ParseCPM = False
    TheExec.Flow.TestLimit resultVal:=999, lowVal:=1, hiVal:=1, Tname:="ParseCPM"
    Exit Function
errHandler:
    Call Print_Error_Message(Error_Info, "eFuseBdfParser", "ParseCPM")
    If AbortTest Then Exit Function Else Resume Next
End Function

'20210712, Add for CmpFusing
Public Function ParseCmpFuse(CmpFuseSheetName As String) As Boolean
On Error GoTo errHandler
Dim iRow As Integer, i As Integer
Dim lEfuseBank As eFuseBdfBank
Dim lEfuseField As eFuseBdfField
Dim cellvalue As String, bCfgField As Boolean
Dim m_SheetMaxRow As Long, m_SheetMaxCol As Long
Dim m_ParsingArr() As Variant
Dim bankName As String, MaskArrayStr As String
Dim MaskArray As Variant
Dim msg As String: msg = vbNullString

    bankName = vbNullString: MaskArrayStr = vbNullString
    ParseCmpFuse = True
    myUtility.Get_Sheet_Info CmpFuseSheetName, m_SheetMaxRow, m_SheetMaxCol, m_ParsingArr
    GlbUtility.WriteDlg "=============Parsing Cmp Fuse Table============="

    For iRow = 1 To m_SheetMaxRow
        If IsBankStart(iRow, m_ParsingArr) Then
            Set lEfuseBank = Nothing
            bankName = Replace(Split(LCase(m_ParsingArr(iRow + 1, 1)), " ")(0), "bank_", "")
            If LCase(bankName) = "config" Then bankName = CfgBankName
            If Banks.Exists(bankName) Then
                Set lEfuseBank = Banks(bankName)
            Else
                msg = "Bank name (" + bankName + ") must exist in BDF table, please check it !!!"
                GoTo skip
            End If
            lEfuseBank.DicCmpFuse.RemoveAll
            iRow = iRow + 1
            GoTo nextstep
        Else
            cellvalue = m_ParsingArr(iRow, 1)
        End If
        
        If cellvalue = "" Then
            Call Print_Error_Message(Warning_Info, "eFuseBdfParser", "ParseCmpFuse", CmpFuseSheetName + " exist blank cell in the first column !!!")
            GoTo nextstep
        ElseIf LCase(cellvalue) = "end" Then
            Exit For
        End If

        If lEfuseBank.Fields.Exists(cellvalue) Then
            Set lEfuseField = lEfuseBank.Fields(cellvalue)
        Else
            msg = "Field name (" + cellvalue + ") must exist in BDF table, please check it !!!"
            GoTo skip
        End If

        If Not lEfuseField Is Nothing Then
            MaskArrayStr = vbNullString
            If Not ReplaceCmpFuseFieldRow(m_ParsingArr, iRow, lEfuseField, lEfuseBank.SeparateEarly, MaskArray) Then
                msg = "Field name (" + lEfuseField.name + ") fuse value has incorrect format, please check it !!!"
                GoTo skip
            End If
            '20210910, Add for Crc overwrite
            If lEfuseBank.Crcs.Exists(lEfuseField.name) Then
                lEfuseBank.Crcs.Remove (lEfuseField.name)
                If lEfuseBank.Crcs.Count = 0 Then
                    lEfuseBank.HadCrcFuse = False
                End If
            End If

            If lEfuseField.DefaultOrReal = dr_default Then
                For i = 0 To lEfuseField.size - 1
                    MaskArrayStr = MaskArrayStr & CStr(MaskArray(lEfuseField.size - 1 - i))
                Next i
            Else
                MaskArrayStr = MaskArray
            End If
            If Not (lEfuseBank.DicCmpFuse.Exists(cellvalue)) Then
                lEfuseBank.DicCmpFuse.Add cellvalue, MaskArray
            End If
            GlbUtility.WriteDlg "<CmpFuseMask> " + cellvalue + ":" + MaskArrayStr
            
            Call ModifyCmpFuse2Category(lEfuseBank, lEfuseField)
        End If
nextstep:
     Next
     GlbUtility.WriteDlg "================================================"

Exit Function
skip:
    ParseCmpFuse = False
    Call Print_Error_Message(Error_Info, "eFuseBdfParser", "ParseCmpFuse", msg)
    TheExec.Flow.TestLimit resultVal:=999, lowVal:=1, hiVal:=1, Tname:="ParseCmpFuse"
    Exit Function
errHandler:
    Call Print_Error_Message(Error_Info, "eFuseBdfParser", "ParseCmpFuse")
    If AbortTest Then Exit Function Else Resume Next
End Function

Private Function ReplaceCmpFuseFieldRow(ParsingArr() As Variant, searStRow As Integer, ReplaceField As eFuseBdfField, SeparateEarly As Boolean, MaskArray As Variant) As Boolean
On Error GoTo errHandler
Dim funcName As String: funcName = "ReplaceCmpFuseFieldRow"
Dim cellStr As String: cellStr = vbNullString
Dim i As Long
Dim OrgValueBin As Variant, CmpValueBin As Variant
Dim MaskArr() As Double
Dim msg As String: msg = vbNullString

    'Name
    cellStr = Trim(ParsingArr(searStRow, 1))

    'Stage / Hlimit / Llimit / Resoultion / MSBFirst / Alogrithm
    cellStr = Trim(ParsingArr(searStRow, 9 - ColShift)): ReplaceField.BlowLocation = GetBlowLocation(cellStr)
    If Not ReplaceField.Algorithm = alg_crc Then
        cellStr = Trim(ParsingArr(searStRow, 10 - ColShift)): ReplaceField.Llimit = myUtility.String2Dbl(cellStr, -1): ReplaceField.xLlimit = myUtility.String2Hex(cellStr)
        cellStr = Trim(ParsingArr(searStRow, 11 - ColShift)): ReplaceField.Hlimit = myUtility.String2Dbl(cellStr, -1): ReplaceField.xHlimit = myUtility.String2Hex(cellStr)
    End If

    If ReplaceField.Algorithm = alg_crc Then
        cellStr = Trim(ParsingArr(searStRow, 13 - ColShift)): ReplaceField.Algorithm = GetAlgorithm(cellStr, ReplaceField.BlowLocation)
    End If
    
    'Set early fusing flag
    If EarlyFuseAlgorithm.Exists(ReplaceField.Algorithm) And SeparateEarly Then
        ReplaceField.IsEarlyFuse = True
    Else
        ReplaceField.IsEarlyFuse = False
    End If
    If LCase(ReplaceField.BlowLocation) Like LCase(myUtility.currStage) + "*_early" Then
        ReplaceField.IsEarlyFuse = True
    Else
        ReplaceField.IsEarlyFuse = False
    End If
    
    'Description / Default Real
    cellStr = Trim(ParsingArr(searStRow, 14 - ColShift)): ReplaceField.Description = cellStr

    'Value
    cellStr = Trim(ParsingArr(searStRow, 16 - ColShift)):

    If ReplaceField.DefaultOrReal = dr_default Then
        ReDim MaskArr(ReplaceField.size - 1)
        OrgValueBin = Split(StrConv(myUtility.Dec2Bin(myUtility.String2Dbl(ReplaceField.Default), ReplaceField.size), vbUnicode), vbNullChar)
        CmpValueBin = Split(StrConv(myUtility.Dec2Bin(myUtility.String2Dbl(cellStr), ReplaceField.size), vbUnicode), vbNullChar)
        For i = ReplaceField.size - 1 To 0 Step -1
            If OrgValueBin(i) = "1" And CmpValueBin(i) = "1" Then   '1->1 : Don't Care
                MaskArr((ReplaceField.size - 1) - i) = 0
            ElseIf OrgValueBin(i) = "1" And CmpValueBin(i) = "0" Then   '1->:0 : Don't be allowed
                msg = ReplaceField.name + " had error comoplement bits(1 to 0), please check it !!!"
                GoTo skip
            Else
                MaskArr((ReplaceField.size - 1) - i) = 1
            End If
        Next i
        MaskArray = MaskArr
    Else
        If Not LCase(cellStr) Like "b*" Then
            msg = "Real field (" + ReplaceField.name + ") must use binary format, please check it !!!"
            GoTo skip
        End If
        MaskArray = Replace(cellStr, "b", "")
    End If

    If (ReplaceField.Algorithm = alg_vddbin And ReplaceField.DefaultOrReal = dr_default) Then
        ReplaceField.Default = cellStr
    Else
        ReplaceField.Default = myUtility.String2Hex(cellStr, 0, ReplaceField.hBytes)
    End If
    ReplaceField.FieldDefault = ReplaceField.Default
    If ReplaceField.Algorithm = alg_crc And ReplaceField.DefaultOrReal = dr_real Then
        ReplaceField.Llimit = 1
        ReplaceField.Hlimit = 2 ^ ReplaceField.size
    End If
    
    'Set default value's H/L limit
    If ReplaceField.DefaultOrReal = dr_default Or ReplaceField.DefaultOrReal = dr_decimal Then
            ReplaceField.Llimit = myUtility.String2Dbl(ReplaceField.FieldDefault, -1)
            ReplaceField.Hlimit = ReplaceField.Llimit
    End If
    If ReplaceField.Hlimit = -1 Then ReplaceField.Llimit = -1
    
    ReplaceField.DefaulString = cellStr ' store the original string from table
    If ReplaceField.DefaultOrReal <> dr_default And ReplaceField.DefaultOrReal <> dr_decimal Then ReplaceField.ReadOnly = False
    If ReplaceField.Algorithm = alg_cond Or ReplaceField.Algorithm = alg_scan Then ReplaceField.ReadOnly = True
    ReplaceCmpFuseFieldRow = True

Exit Function
skip:
    ReplaceCmpFuseFieldRow = False
    Call Print_Error_Message(Error_Info, "eFuseBdfParser", "ReplaceCmpFuseFieldRow", msg)
    Exit Function
errHandler:
    Call Print_Error_Message(Error_Info, "eFuseBdfParser", "ReplaceCmpFuseFieldRow")
    If AbortTest Then Exit Function Else Resume Next
End Function

Private Sub ModifyCmpFuse2Category(lEfuseBank As eFuseBdfBank, lEfuseField As eFuseBdfField)
On Error GoTo errHandler

    'Early Fused
    If lEfuseField.IsEarlyFuse Then
        lEfuseBank.HadEarlyFuse = True
        If Not (lEfuseBank.DicEarlyFused.Exists(lEfuseField.name)) Then
            lEfuseBank.DicEarlyFused.Add lEfuseField.name, lEfuseBank.name
        End If
    Else
        If LCase(stage) = LCase(lEfuseField.BlowLocation) Or bLumpStages Then
            If Not (lEfuseBank.DicOthers.Exists(lEfuseField.name)) Then
                lEfuseBank.DicOthers.Add lEfuseField.name, lEfuseBank.name
                If lEfuseBank.DicEarlyFused.Exists(lEfuseField.name) Then
                    lEfuseBank.DicEarlyFused.Remove lEfuseField.name
                End If
            End If
        End If
    End If
    
    'Config Table field
    If lEfuseField.Algorithm = alg_cond Then
        If Not (lEfuseBank.DicCondTables.Exists(lEfuseField.name)) Then
            lEfuseBank.DicCondTables.Add lEfuseField.name, lEfuseBank.name
        End If
    End If

Exit Sub
errHandler:
    Call Print_Error_Message(Error_Info, "eFuseBdfParser", "ModifyCmpFuse2Category")
    If AbortTest Then Exit Sub Else Resume Next
End Sub

Public Function GetRealStage(bdfJobName As String) As String
On Error GoTo errHandler
Dim realJobName() As String

    GetRealStage = bdfJobName
    If Not LCase(bdfJobName) Like "*early" Then
        realJobName = Split(bdfJobName, "_")
        GetRealStage = realJobName(0)
    End If

Exit Function
errHandler:
    Call Print_Error_Message(Error_Info, "eFuseBdfParser", "GetRealStage")
    If AbortTest Then Exit Function Else Resume Next
End Function

Private Sub MultiTypeFieldClassify()
On Error GoTo errHandler
Dim bankstr As Variant, bank As eFuseBdfBank
Dim MultiType As Variant
Dim field As Variant
Dim fieldDic As New Dictionary

    If Banks.Exists(Empty) Then Banks.Remove (Empty)
    For Each bankstr In Banks.Keys
        Set bank = Banks(bankstr)
        If bank.MultiTypeJobList.Count <> 0 Then
            For Each MultiType In bank.MultiTypeJobList.Keys
                GlbUtility.IniDictionary fieldDic, True
                For Each field In bank.DicAllTypefields.Keys
                    If bank.DicAllTypefields(field) = MultiType Then fieldDic.Add field, bank.name
                Next field
                bank.DicMultiTypFields.Add MultiType, fieldDic
            Next MultiType
        End If
    Next bankstr

Exit Sub
errHandler:
    Call Print_Error_Message(Error_Info, "eFuseBdfParser", "MultiTypeFieldClassify")
    If AbortTest Then Exit Sub Else Resume Next
End Sub

Private Sub AddMultiTypeRVSegment()
On Error GoTo errHandler
Dim stBit As Long, spBit As Long, bitStep As Long
Dim bank As eFuseBdfBank, bankstr As Variant
Dim fieldStr As Variant, field As eFuseBdfField
Dim divider As Long: divider = 32
Dim indexLoop As Long
Dim idx0 As Long, idx1 As Long, finalIdx As Long
Dim jobType As Variant
Dim multiTypeRvArr As New Dictionary
Dim dicArrIndex As New Dictionary
Dim segArr() As Long
Dim i As Integer
Dim indexArr() As String

    bankstr = "CFG"
    If Banks.Exists(bankstr) Then
        Set bank = Banks(CStr(bankstr))
        If bank.DoubleBits Then divider = 16
        For Each jobType In bank.MultiTypeJobList.Keys
            multiTypeRvArr.Add jobType, ""
        Next jobType

        multiTypeRVseg.RemoveAll
        If EFUSE_RV_PATTERN_FULL Then
            For Each jobType In multiTypeRvArr.Keys
                multiTypeRVseg.Add jobType, segFlagArr
            Next jobType
        Else
            For Each fieldStr In bank.Fields.Keys
                Set field = bank.Fields(fieldStr)
                If GetRealStage(field.BlowLocation) = TheExec.CurrentJob And multiTypeRvArr.Exists(field.BlowLocation) And field.DefaultOrReal <> dr_default And field.Algorithm <> alg_cond Then
                    stBit = field.LSB: spBit = field.msb: bitStep = 1: ' MSB 1st
                    If field.Msb1st Then
                         bank.GetBitStSp field, stBit, spBit
                    ElseIf stBit > spBit Then
                         bitStep = -1: 'LSB first, if doing CRC, it needs to use the MSxB first
                    Else
                         myUtility.MessageBox "Something wrong in dealing BDF GetDspWave"
                    End If
        
                    idx0 = Floor(stBit / divider) 'please check floor
                    idx1 = Floor(spBit / divider) 'please check floor
                    If CStr(idx0) <> "" And Not (multiTypeRvArr(field.BlowLocation) Like CStr(idx0) Or multiTypeRvArr(field.BlowLocation) Like "*," + CStr(idx0) Or multiTypeRvArr(field.BlowLocation) Like CStr(idx0) + ",*" Or multiTypeRvArr(field.BlowLocation) Like "*," + CStr(idx0) + ",*") Then
                        multiTypeRvArr(field.BlowLocation) = IIf(multiTypeRvArr(field.BlowLocation) <> "", multiTypeRvArr(field.BlowLocation) + "," + CStr(idx0), CStr(idx0))
                    End If
        
                    If (idx0 <> idx1) Then
                        For indexLoop = 1 To idx1 - idx0
                            finalIdx = idx0 + indexLoop
                            If Not (multiTypeRvArr(field.BlowLocation) Like CStr(finalIdx) Or multiTypeRvArr(field.BlowLocation) Like "*," + CStr(finalIdx) Or multiTypeRvArr(field.BlowLocation) Like CStr(finalIdx) + ",*" Or multiTypeRvArr(field.BlowLocation) Like "*," + CStr(finalIdx) + ",*") Then
                                multiTypeRvArr(field.BlowLocation) = IIf(multiTypeRvArr(field.BlowLocation) <> "", multiTypeRvArr(field.BlowLocation) + "," + CStr(finalIdx), CStr(finalIdx))
                            End If
                        Next
                    End If
                End If
            Next fieldStr

            For Each jobType In multiTypeRvArr.Keys
                indexArr = Split(multiTypeRvArr(jobType), ",")
                ReDim segArr(bank.FullSize / divider - 1)
                For i = 0 To UBound(indexArr)
                    segArr(indexArr(i)) = 1
                Next i
                multiTypeRVseg.Add jobType, segArr
            Next jobType
        End If
    End If

Exit Sub
errHandler:
    Call Print_Error_Message(Error_Info, "eFuseBdfParser", "AddMultiTypeRVSegment")
    If AbortTest Then Exit Sub Else Resume Next
End Sub
