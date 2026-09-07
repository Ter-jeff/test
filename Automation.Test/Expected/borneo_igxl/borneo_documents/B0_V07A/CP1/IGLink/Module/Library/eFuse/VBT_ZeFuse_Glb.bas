Attribute VB_Name = "VBT_ZeFuse_Glb"
#Const isUFP = True
Option Explicit
Public Enum PartType
        AP = 0
        Wireless
        BDFonly
        CFGonly
End Enum

Public Enum ChkType
        ChkGreaterThan = 0
        ChkLessThan
        ChkEqualTo
        ChkGreaterEqualThan
        ChkLessEqualThan
End Enum

Public Enum Algorithm
    alg_app = 0 'v
    alg_ids '<<<<< ' Let user to put the values, all in "mA", or use the standalone function to do it
    alg_vddbin '<<<<< 'One function is ready for put all of them
    alg_cond 'v
    alg_scan 'v
    alg_base 'v
    alg_crc 'v
    alg_uid '<<<<<
    alg_lotid 'v
    alg_numeric 'v
    alg_force 'v
    alg_na 'v
End Enum

'20220602 Add for crc complement
Public Enum crctype
    crc_none = 0
    crc_normal
    crc_onecomp
    crc_twocomp
End Enum

Public Enum DeaultReal
    dr_default = 0
    dr_real
    dr_decimal
    dr_na
End Enum

Public Enum ProgramMode
    pgm_DAA = 0
    pgm_JTAG
End Enum

Public Enum DataRetrieve
    AteTrimData = 0
    DafaultValue
    DsscValue
    TrimmedData
End Enum

Public Enum DataProc
    blankCheck = 0
    FuseWrite
    ReadOutCheck
End Enum

Public Enum ReadTestType
    DefaultMode = 0
    JTAG_Or_Result
    JTAG_APB
End Enum

Public SelectedCfg As String
Public GlbUtility As Utility
Public firstRun As Boolean
Public DicEcidCat As Dictionary
Public firstrun_2 As Boolean
Public DicVddPmode As dictionary
Public DicVddPmodeShadow As dictionary
Public VddPmodeHadSameStage As Boolean

Public EarlyFuseAlgorithm As Dictionary
Public LotIdBits As Long
Public WfrIdBits As Long
Public XcoorBits As Long
Public YcoorBits As Long
Public stage As String
Private FileTrimSelectEver As Boolean
Public BdfDataBase  As eFuseBdfParser
Public CfgDataBase As eFuseCfgParser
Public CfgDataBaseSvm As eFuseCfgParser
Public CfgDataBaseAct As eFuseCfgParser 'This is the only actived database
Public Const BinCutIdentifierField = "product_identifier"
Public Const BankDumpPath = "D:\BDF_Dump"
Private Const bBankDump = False
Private LastRunCfg As String
Private site As Variant
Public Const MultiEcidBanks = False 'usually, it's false as default.
Public Const AllFieldUnique = False     ''202004xx disable for AP

''202004xx for ap
Public PseudoFuseEnable As Boolean
Public MixPseudoFuseEnable As Boolean
Public ParsePseudoFuseFile As Boolean
Public LotID_ForPseudoFuse As String
Public WaferID_ForPseudoFuse As String
Public DateTime_ForPseudoFuse As String
Public GetPseudoFuseFileOneTime As Boolean
Public PgmName_ForPseudoFuse As String
Public IsOI_ForPseudoFuse As String

Public gB_Dut2Db As Boolean

Public BlankCheck_capWave As New DSPWave
'<<<<<<<<<< Step1, Select the need type and feed the require consts below.
'#Const Golay = True
#Const AP = True        'BDF + CFG
'#Const BDFonly = True  'only BDF
'<<<<<<<<<< Step2, Setting specific declare for your project
Public Const FuseRevisionBankName = "CFG"  'Setting bank name of "fuse_revision" category
'<<<<<<<<<< Step3, check the ParseCfgTable function
Public Const DaaParaBits = 32 ' typically it's 32, but in Tesla is 8
Public Const MyBankWavePath = "\offlinewave\"
Private Const bBankDumpWave = False
Public StdfPRR As Boolean
Public Const Local_BKM_path = "D:\BKM\"
Public Const LotIdCharBits = 6
Public Const BKM_Main_Path = "X:\BKM\"
Public Const MONFullSize = 1024
Public Const PatRepeatCnt = 1

#If AP Then  ' This is the typical case, all banks are separated. BDF + Config_Table
'==============For AP, e.g. Cebu=========BDF+Config===
Public Const BdfSheetName = "EFUSE_BitDef_Table"
Private Const CfgSheetName = "Config_table"
Public Const DEID_Inst_Key = "_Deid"
Public Const NonDEID_Inst_Key = "_nonDeid"
Private Const CmpFuseSheetName = "Efuse_Cmp_Fuse_Table"
Private Const DramSheetName = "DRAM_Table"
Private Const CPMSheetName = "CPM_Table"
Private Const BKMSheetName = "BKM_Info"
Private Const BinChkSheetName = "FuseCheckTable"
Public Const stagesStr = "CP1,CP2,WLFT1,WLFT2,FT1,FT2,FT3,FTF,SLT" ' for AP
Public Const RunPartType = PartType.AP
Public Const bAutoDefaultCheck = False
Public Const bLumpStages = False
Public Const bForceJtagMode = False
Public Const bAllowNoCfgSel = False
#ElseIf Golay Then ' This is the special case, bank is shared. e.g. ECID contains UID.  BDF + Config_Table
'==============For Wireless, e.g. Golay====BDF+Config===
Public Const BdfSheetName = "EFUSE_BitDef_Table"
Private Const CfgSheetName = "Config_table"
Public Const DEID_Inst_Key = "ECID"
Public Const NonDEID_Inst_Key = "UID"
Public Const stagesStr = "CP1,FT1,FT2,FT3,FTF" ''BDF said that "ecid_crc" be put in "PgmFlow=nonDEID"
Public Const RunPartType = PartType.Wireless
Public Const bAutoDefaultCheck = True
Public Const bLumpStages = True  ' must put before the function [ParseCfgSheet]
Public Const bForceJtagMode = True  'It's DAA declaration on BDF but run JTAG pattern actually
Public Const bAllowNoCfgSel = False
Private Const CmpFuseSheetName = vbNullString
Private Const DramSheetName = vbNullString
Private Const CPMSheetName = vbNullString
Private Const BKMSheetName = vbNullString
Private Const BinChkSheetName = vbNullString
#ElseIf Tesla Then  ' This is the special case, no BDF but has Config_Table only, similar cases are S4E,Seagull
'==============For Wireless, e.g. Tesla=====Config Only==
Public Const BdfSheetName = vbNullString
Private Const CfgSheetName = "EFUSE_Config_Main_A_Tesla"
Public Const DEID_Inst_Key = "_Deid"
Public Const NonDEID_Inst_Key = "_nonDeid"
Public Const stagesStr = "CP1"
Public Const RunPartType = PartType.CFGonly
Public Const bAllowNoCfgSel = True
Public Const bAutoDefaultCheck = False
Public Const bLumpStages = False
Public Const bForceJtagMode = False
Private Const CmpFuseSheetName = vbNullString
Private Const DramSheetName = vbNullString
Private Const CPMSheetName = vbNullString
Private Const BKMSheetName = vbNullString
Private Const BinChkSheetName = vbNullString
#ElseIf BDFonly Then  ' This is the special case, no Config_Table but has BDF only
'==============For TestChip, e.g. Skua======BDF Only===
Public Const BdfSheetName = "EFUSE_BitDef_Table" '"EFUSE_BitDef_Table_MultiEcid"
Private Const CfgSheetName = vbNullString
Public Const DEID_Inst_Key = "_Deid"
Public Const NonDEID_Inst_Key = "_nonDeid"
Public Const stagesStr = "CP1,CP2,FT1,FT2"
Public Const RunPartType = PartType.BDFonly
Public Const bAutoDefaultCheck = False
Public Const bLumpStages = False
Public Const bForceJtagMode = False
Public Const bAllowNoCfgSel = False
Private Const CmpFuseSheetName = vbNullString
Private Const DramSheetName = vbNullString
Private Const CPMSheetName = vbNullString
Private Const BKMSheetName = vbNullString
Private Const BinChkSheetName = vbNullString
#End If

Public Const CheckTestTime = False
Public Const eFusePrinted = True    ' Set False, no printed
Public Const eFusePrintedCnt = -1 '20 , set -1 to listed all
Public CfgNeedReset As Boolean
Public Const CfgBankName = "CFG"

''202005xx for ap
Public Const EFUSE_REFUSE_FOR_PTE = False
Public Const EFUSE_BIN1_SETTING = False
Public Const EFUSE_POWER_OFF_SETTING = False
Public Const EFUSE_SETFUSE_SHOW_ERRMSX = False
Public Const EFUSE_REPLACE_SLT_TO_FT3 = False
Public Const EFUSE_ALWAYS_CHECK_CRC = False  '20220630, In crc field, executing syntax check at all job
Public Const EFUSE_RV_PATTERN_FULL = True    '20230131, Ture: Use one RV pattern execute all programming type (ex: CP1, CP1_1, CP1_2..)
Public Const EFUSE_PRINT_DOUBLE_HEXMAP = False '20230221, HexMap print double or single wave information
Public Const Flag_CFG_Multi_RV_Enable = False '20220215,Modify for Multi RV Pat, True: Support multiple RV write patterns, False: Support one RV write pattern
Public Const CFG_TotalWritePatCnt = 8 '20220215,Modify for Multi RV Pat, Define RV wirte pattern count
Public Const EFUSE_CHECK_VDDBIN_SAME_STAGE = True  '20240105, Check non shadow and shadow is same stage or not.

'20230530 ECID Sorting, Change it for ECID Sorting and defualt is False
'20230823 If fuse the CFG condition(ex: A09 - A12) will lock device in FT3, user can turn on "EFUSE_CFG_READ_BYJTAG" for "CFG" JTAG read to decode.
Public Const EFUSE_CFG_READ_BYJTAG = False
'20230823 If fuse the CFG condition(ex: A09 - A12) will lock device in FT3, user can turn on "EFUSE_ECID_READ_BYJTAG" for "ECID" JTAG read to decode.
Public Const EFUSE_ECID_READ_BYJTAG = False
Public EFUSE_ECID_SORTING_ENABLE As Boolean  '20230530 ECID Sorting, Add for Enable ECID Sorting
Public EFUSE_ECID_SORTING_2CMODE As Boolean  '20230530 ECID Sorting, Add for C2C project ECID Sorting

'20220215,Modify for Multi RV Pat
Public CFG_Multi_RV_PatSetName As String
Public CFG_Multi_RV_PatCnt As Long
Public CFG_Multi_RV_PatDsscCnt() As Long
Public CFG_Multi_RV_PatDsscTotalCnt As Long

Public gB_Obj_VBT_Enable As Boolean
Public segFlagArr() As Long
Public g_Rvenable As Boolean
'20221117, support multi-type program stage
Public multiTypeRVseg As New Dictionary
Public Dic_SplitDspWave As Dictionary

'20210713,Add for Cmp Fuse
Public gB_CmpFuseEnable As Boolean

Public Type EfuseBinCheckCase
    ValueCheck As Boolean
    specificDecValue As Double
    specificHexValue As String
    IDSCheck As Boolean
    RangeCheck As Boolean
    LowDecValue As Double
    HighDecValue As Double
    LowHexValue As String
    HighHexValue As String
    GroupCheck As Boolean
    GroupBits As Long
    GroupPick As Long
    FailBit As Long
    WalkingCheck As Boolean
    Walkingvalue As Long
    TwoCheck As Boolean
    isOperator As Boolean
    Rule As String
    SkipTest As Boolean
    siteInfo As New Dictionary
    ConditionChkTrue As Boolean
End Type

'20210511 Add for Bin Check Table instace
Public Type EFuseBinCheckData
    bankName As String
    cateName As String
    cateNameOri As String
    specSite As Boolean
    checkRules() As EfuseBinCheckCase
    MergeBitCheck As Boolean
    SplitBitCheck As Boolean
    Bit_LSB As Long
    Bit_MSB As Long
    Width As Long
    conditionCheck As Boolean
    conditionInfo As String
    conditionResult As Boolean
    CombineFieldCheck As Boolean
    CombineFieldJob As String
    listCombineFieldsGroup As Collection
    skipCheck As Boolean
End Type

Public Type EfuseBinCheckInfo
    CateArr() As EFuseBinCheckData
End Type
Public BinCheckData() As EfuseBinCheckInfo
Public DictBinNameMapping As New Dictionary
Public DictBinFlagMapping As New Dictionary
Public DictBinBankMapping As New Dictionary

Public Flag_Efuse_Config_Printed As Boolean

Public CurrentPassBinCutNum_additional As New SiteLong
Public CurrentPassBinCutNum_normal As New SiteLong

'20211110,Add for two type bincut table syntax check
Public Type EFUSE_BINCUT_TYPE
    c() As Double
    m() As Double
    CP_Vmax() As Double
    CP_Vmin() As Double
    CP_GB() As Double
    Mode_Step As Long
End Type
Public EfuseBinCut() As EFUSE_BINCUT_TYPE
Public EfuseBinCutAddition() As EFUSE_BINCUT_TYPE
Public Const BincutOriginalSheetName = vbNullString
Public Const BincutAdditionalSheetName = vbNullString

'20211210, Modify for FT pseudo fuse with real package
Public gB_Package_PsudoFuse As Boolean
'20230307, Modify for FT pseudo fuse with wafer
Public gB_FT_Wafer_PsudoFuse As Boolean

'20211227, Modify for multi config write patterns order control
Public Const ConfigWritePatOrder_F = "1"
Public Const ConfigWritePatOrder_O = "8,7,6,5,4,3,2,1"
Public Const ConfigWritePatOrder_Q = "4,3,2,1"
Public Const ConfigWritePatOrder_H = "16,15,14,13,12,11,10,9,8,7,6,5,4,3,2,1"

Private Function ParseCfgTable() As Boolean
On Error GoTo errHandler

    ParseCfgTable = True
    If RunPartType = PartType.BDFonly Then Exit Function

    If WorksheetExists(CfgSheetName) Then
        ParseCfgTable = ParseCfgSheet(CfgDataBase, CfgSheetName)
    Else
        GlbUtility.MessageBox "Sheet didn't exist!!! " & CfgSheetName
        ParseCfgTable = False
    End If

    If RunPartType = PartType.AP Then
        '20220210,Modify for Fixed SVM Enable issue
        If TheExec.enableWord("CFG_SVM") Then
            If WorksheetExists("Config_table_SVM") Then
                ParseCfgTable = ParseCfgSheet(CfgDataBaseSvm, "Config_table_SVM")
                Set CfgDataBaseAct = CfgDataBaseSvm
            Else
                GlbUtility.MessageBox "No SVM config, use org config! Please check it again!", "Warning"
                Set CfgDataBaseAct = CfgDataBase
            End If
        Else
            Set CfgDataBaseAct = CfgDataBase
        End If
    End If
    
Exit Function
errHandler:
    Call Print_Error_Message(Error_Info, "VBT_ZeFuse_Glb", "ParseCfgTable")
    If AbortTest Then Exit Function Else Resume Next
End Function

Public Function eFuseInitialization()
On Error GoTo errHandler

    Glb_Setup

    If Not Check1stRun Then
        firstRun = True
        firstrun_2 = True
        Exit Function
    End If

    GlbUtility.Timer = CheckTestTime
    StdfPRR = False
    BdfDataBase.ReadRealBKMdone = False
    'Clear Window system IEDA register
    'Call ClearRegKeys 'Have do it on the MainFlow
    If Dir(MyBankWavePath, vbDirectory) = "" And (bBankDumpWave Or Not GlbUtility.OnlineMode) Then MkDir MyBankWavePath
    'If Dir(MyBankWavePath, vbDirectory) = "" And (bBankDumpWave Or Not GlbUtility.IsOnline) Then MkDir MyBankWavePath
    If CfgNeedReset And (Not CfgDataBaseAct Is Nothing) Then
        'Port CFGFiled info to BDFBank / BDFField.
        Call CfgDataBaseAct.eFuseCfgInitialize(BdfDataBase.Bank_Cfg, GlbUtility.currStage, SelectedCfg)
        If BdfDataBase.CFG_Only Then DicRemoved BdfDataBase.Bank_Cfg.DicOthers, DicEcidCat
        CfgNeedReset = False
    End If
        
    If (Not CfgDataBaseAct Is Nothing) Then
        CfgDataBaseAct.SetCfgBlankMaskDspWave
        If BdfDataBase.CFG_Only Then
            Call UpdateFullCfg2BdfEfuse
        Else
            Dim field As eFuseBdfField, fieldStr As Variant
            Set BdfDataBase.Bank_Cfg.CfgCmpValue = CfgDataBaseAct.CfgCmpValue
            Set BdfDataBase.Bank_Cfg.CfgCmpValueLumpStages = CfgDataBaseAct.CfgCmpValueLumpStages
            For Each fieldStr In BdfDataBase.Bank_Cfg.DicCondTables.Keys
                Set field = BdfDataBase.Bank_Cfg.Fields(fieldStr)
                Call UpdateCfg2BdfEfuse(CfgDataBaseAct.CfgValue(fieldStr), DafaultValue, field.name)
                If bLumpStages And field.size < 32 Then
                    field.Llimit = GlbUtility.String2Dbl(BdfDataBase.Bank_Cfg.CfgCmpValueLumpStages(fieldStr))
                    field.Hlimit = GlbUtility.String2Dbl(BdfDataBase.Bank_Cfg.CfgCmpValueLumpStages(fieldStr))
                End If
            Next
        End If
    End If

    Call ObtainTrimSelectItems
    
    Call BdfDataBase.ObtainBankStageMaskWave
    
    Call BdfDataBase.eFuseBdfInitialize
    Call BdfDataBase.LogBdfInfo ''BDF info print out
    'Call BdfDataBase.ProcessHipDram 'move to JudgeDRAMType_T
    Call BdfDataBase.PseudoFuseSetUp

    GlbUtility.IniDictionary Dic_SplitDspWave
    Dic_SplitDspWave.RemoveAll
    
    #If Golay Then
        If GlbUtility.testedStages.Count = 0 And bLumpStages Then
            If BdfDataBase.Bank_Ecid.DicSetDefaultVauleToReal.Count = 0 Then
                BdfDataBase.Bank_Ecid.UpdatedDefaultVauleToReal ("bank_ecid_Spare0,model_revision,platform_version,fuse_ecid_indication,bank_ecid_Spare1,bank_ecid_Spare2,bank_ecid_Spare3,rev_id")
            End If
           ' Call ReadWaferData
            BdfDataBase.Bank_Ecid.PutDefaultVauleToReal
            BdfDataBase.Bank_Cfg.PutDefaultVauleToReal False
        End If
    #End If
    
    'If Not GlbUtility.IsOnline Then Call ReadWaferData
    'GlbUtility.UpdateDLogColumns (48)
    '202004xx for AP
    Call GetVddBinPmodeMap
    TheExec.Flow.TestLimit resultVal:=1, lowVal:=0, hiVal:=2, Tname:="eFuse_Initialize_" & GlbUtility.currStage, PinName:="Value"
    If Len(SelectedCfg) > 0 Then
        TheExec.Flow.TestLimit resultVal:=1, lowVal:=0, hiVal:=2, Tname:="CFG_Flag_Set_" & SelectedCfg, PinName:="Value"
    Else: TheExec.Flow.TestLimit resultVal:=1, lowVal:=0, hiVal:=2, Tname:="CFG_Flag_Set_NA", PinName:="Value"
    End If
    
    If bBankDump Or Not GlbUtility.OnlineMode Then BdfDataBase.DumpBdfData BankDumpPath
    'If bBankDump Or Not GlbUtility.IsOnline Then BdfDataBase.DumpBdfData BankDumpPath

    PrintEfuseConfigSetting

    If TheExec.enableWord("eFuse_Obj") = True Or _
       TheExec.enableWord("eFuse_PseudoFuse") = True Or _
       TheExec.enableWord("eFuse_Corr") = True Then gB_Obj_VBT_Enable = True

Exit Function
errHandler:
    Call Print_Error_Message(Error_Info, "VBT_ZeFuse_Glb", "eFuseInitialization")
    If AbortTest Then Exit Function Else Resume Next
End Function

Private Function Check1stRun() As Boolean
On Error GoTo errHandler

    stage = TheExec.CurrentJob '"CP1" 'Get it from JOB
    stage = Replace(stage, "_105C", "")
    Check1stRun = True
    If firstRun Then
        Set BdfDataBase = Nothing
        Set GlbUtility = Nothing
        Set CfgDataBase = Nothing
        VddPmodeHadSameStage = False
    End If
    If GlbUtility Is Nothing Then Set GlbUtility = New Utility

    If Not BdfDataBase Is Nothing Then
        If (Not CfgDataBase Is Nothing) Then
            If Not CfgEnableCheck Then GoTo skip
            If LastRunCfg <> SelectedCfg Then
                firstRun = True 'cfg enable word changes then reset all
                Set BdfDataBase = Nothing
            End If
        End If
    End If
    If Not GlbUtility.Initialize(stage) Then GoTo skip  ''Check the current job that exists in the stage list.
    
    If BdfDataBase Is Nothing Then
        firstRun = True
        firstrun_2 = True
        Application.ScreenUpdating = False
        If Not EfuseOnPgmValidate Then
            Application.ScreenUpdating = True
            GoTo skip
        End If
        Application.ScreenUpdating = True
    End If

    If Not CfgDataBase Is Nothing Then
        If Not CfgEnableCheck Then GoTo skip
    End If

    If Not firstRun Then
        #If BDFonly Then
            Exit Function
        #End If
    End If

    'Run below code if not 1st run------------------------------------
    If Not ObtainEcidCategories Then GoTo skip

    If (Not CfgDataBaseAct Is Nothing) Then
        LastRunCfg = SelectedCfg
        CfgNeedReset = True
    End If

    If firstRun Then
        Call CollectPrintIdx
        GlbUtility.UpdateDLogColumns (48)
    End If
    If EFUSE_CHECK_VDDBIN_SAME_STAGE Then
        If firstRun Or VddPmodeHadSameStage Then ShadowAndNonShadowSameStage
    End If
    firstRun = False
    
Exit Function
skip:
    Set BdfDataBase = Nothing
    GlbUtility.MessageBox "Fatal error! Check1stRun is failed!"
    TheExec.Flow.TestLimit resultVal:=999, lowVal:=1, hiVal:=1, Tname:="Check1stRun"
    Check1stRun = False
    Exit Function
errHandler:
    Call Print_Error_Message(Error_Info, "VBT_ZeFuse_Glb", "Check1stRun")
    If AbortTest Then Exit Function Else Resume Next
End Function

Public Function ObtainTrimOverWriteItems() As Long
On Error GoTo errHandler
Dim filename As String: filename = "./eFuseCtrl/User_Overwrite_Trim.csv"

    ReadTrimSelectOverWrite filename, True
    TheExec.Flow.TestLimit resultVal:=1, lowVal:=1, hiVal:=1, Tname:=filename + " Loaded", PinName:="Value"

Exit Function
errHandler:
    Call Print_Error_Message(Error_Info, "VBT_ZeFuse_Glb", "ObtainTrimOverWriteItems")
    If AbortTest Then Exit Function Else Resume Next
End Function

Public Sub DecideEarlyFuseAlgorithm()
On Error GoTo errHandler

    Set EarlyFuseAlgorithm = New Dictionary
    EarlyFuseAlgorithm.compareMode = TextCompare
    EarlyFuseAlgorithm.Add alg_cond, True
    EarlyFuseAlgorithm.Add alg_scan, True
        
Exit Sub
errHandler:
    Call Print_Error_Message(Error_Info, "VBT_ZeFuse_Glb", "DecideEarlyFuseAlgorithm")
    If AbortTest Then Exit Sub Else Resume Next
End Sub

Public Function eFuse_GetReadDecimal(ByVal bankstr As String, m_catename As String, Optional showPrint As Boolean = True) As SiteDouble
On Error GoTo errHandler
Dim opbank As eFuseBdfBank

    Set opbank = GetBdfBank(bankstr)
    Set eFuse_GetReadDecimal = opbank.GeteFuseValue(m_catename)
    If showPrint Then
        For Each site In TheExec.sites
            GlbUtility.WriteDlg vbTab & "Site(" & site & ")" & GlbUtility.TxtFmt(bankstr, 4) & GlbUtility.TxtFmt(m_catename, 24) _
            & GlbUtility.TxtFmt(eFuse_GetReadDecimal, 12)
        Next
    End If

Exit Function
errHandler:
    Call Print_Error_Message(Error_Info, "VBT_ZeFuse_Glb", "eFuse_GetReadDecimal")
    If AbortTest Then Exit Function Else Resume Next
End Function

Public Function EfuseOnPgmValidate() As Boolean
On Error GoTo errHandler

    EfuseOnPgmValidate = False
    If GlbUtility Is Nothing Then Set GlbUtility = New Utility
    GlbUtility.currStage = vbNullString
    If Not GlbUtility.Initialize(Replace(TheExec.CurrentJob, "_105C", "")) Then Exit Function
    Call DecideEarlyFuseAlgorithm
    If Not ParseCfgTable Then Exit Function
    If Not ParseBdfTable Then Exit Function
    If Not ParseCmpFuseTable Then Exit Function
    If Not ParseBKMTable Then Exit Function
    If Not ParseDramTable Then Exit Function
    If Not ParseCPMTable Then Exit Function
    If Not ParseBinChkTable Then Exit Function
    
    If Not ParseBincutTable Then Exit Function

    '20220215,Add for Multi RV Pat
    If Flag_CFG_Multi_RV_Enable = True Then
        Call CFG_Multi_RV_PatCalc
    End If
    If BdfDataBase Is Nothing Then 'This means this project only contains Config_Main_X table
        Set BdfDataBase = New eFuseBdfParser: BdfDataBase.Initialize (GlbUtility.currStage)
        BdfDataBase.CFG_Only = True
        BdfDataBase.MergeCfg2BdfTable CfgDataBase
        BdfDataBase.DumpBdfData BankDumpPath
    Else
        If Not BdfDataBase.Banks.Exists(FuseRevisionBankName) Then
            GlbUtility.MessageBox "Please check global parameter FuseRevisionBankName(" & FuseRevisionBankName & ") didn't exist in BDF's bank! "
            Exit Function
        End If
    End If
    
    '20220614, Add for parsing sramsoc_fusing tables
    If Not Parsing_SEPVM_Fusing_Table_dynamic("SRAM_SOC_fusing") Then Exit Function
    
    '20230530 ECID Sorting, Add it to parse ECID check list table
    If EFUSE_ECID_SORTING_ENABLE = True Then
        Call ECID_Dict_Buildup_S
    End If

    EfuseOnPgmValidate = True
    
Exit Function
errHandler:
    Call Print_Error_Message(Error_Info, "VBT_ZeFuse_Glb", "eFuse_GetReadDecimal")
    If AbortTest Then Exit Function Else Resume Next
End Function

'Add this function, to comply the requirement from DC lib
Public Function auto_eFuse_GetIDSResolution(FuseType As String, fieldStr As String, Optional showPrint As Boolean = False) As Double
On Error GoTo errHandler
Dim opbank As eFuseBdfBank, field As eFuseBdfField, bankstr As String: bankstr = FuseType

    Set opbank = GetBdfBank(bankstr)
    Set field = opbank.Fields(Trim(fieldStr))
    auto_eFuse_GetIDSResolution = field.Resolution
    If (showPrint) Then TheExec.Datalog.WriteComment vbTab & FuseType + GlbUtility.TxtFmt(fieldStr, 32) + " = " + GlbUtility.TxtFmt(field.Resolution, 16)

Exit Function
errHandler:
    Call Print_Error_Message(Error_Info, "VBT_ZeFuse_Glb", "auto_eFuse_GetIDSResolution")
    If AbortTest Then Exit Function Else Resume Next
End Function

Public Function GetEfuseHipValue(FuseType As String, fieldStr As String) As SiteDouble
On Error GoTo errHandler
Dim opbank As eFuseBdfBank, bankstr As String: bankstr = FuseType

    Set opbank = GetBdfBank(bankstr)
    Set GetEfuseHipValue = opbank.GeteFuseValue(Trim(fieldStr))

Exit Function
errHandler:
    Call Print_Error_Message(Error_Info, "VBT_ZeFuse_Glb", "GetEfuseHipValue")
    If AbortTest Then Exit Function Else Resume Next
End Function

Public Sub SetEfuseHipValue(FuseType As String, fieldStr As String, ByVal invalue As SiteVariant, ByVal PatPF As SiteBoolean)
On Error GoTo errHandler
Dim opbank As eFuseBdfBank, bankstr As String: bankstr = FuseType

    Set opbank = GetBdfBank(bankstr)
    Call opbank.SetEfuse(Trim(fieldStr), invalue, PatPF)
        
Exit Sub
errHandler:
    Call Print_Error_Message(Error_Info, "VBT_ZeFuse_Glb", "SetEfuseHipValue")
    If AbortTest Then Exit Sub Else Resume Next
End Sub

Public Function GetBdfBank(bankstr As String) As eFuseBdfBank
On Error GoTo errHandler

    If LCase(bankstr) = "config" Then bankstr = CfgBankName
    If GlbUtility.IsStrMatch(bankstr, "UDRE") Then bankstr = "UDR_E"
    
      ''''20210607,Add for Rhod''''
    If GlbUtility.IsStrMatch(bankstr, "UDRP0") Then bankstr = "UDR_P0"
    If GlbUtility.IsStrMatch(bankstr, "UDRP1") Then bankstr = "UDR_P1"
    If GlbUtility.IsStrMatch(bankstr, "UDRP") Then bankstr = "UDR_P"
      ''''''''''''''''''''''''''''''
    If BdfDataBase.Banks.Exists(bankstr) Then
        Set GetBdfBank = BdfDataBase.Banks(bankstr)
        'For HIP category name 210126
        'Call GetBdfBank.UpdateLogFormat
    Else
        GlbUtility.MessageBox "The specific bank didn't exist! " & bankstr
    End If
    
Exit Function
errHandler:
    Call Print_Error_Message(Error_Info, "VBT_ZeFuse_Glb", "GetBdfBank")
    If AbortTest Then Exit Function Else Resume Next
End Function

Public Sub DicRemoved(target As Dictionary, removal As Dictionary)
On Error GoTo errHandler
Dim item As Variant

    For Each item In removal.Keys
        If target.Exists(item) Then target.Remove (item)
    Next
    
Exit Sub
errHandler:
    Call Print_Error_Message(Error_Info, "VBT_ZeFuse_Glb", "DicRemoved")
    If AbortTest Then Exit Sub Else Resume Next
End Sub

Public Sub DicSameKept(target As Dictionary, compareSrc As Dictionary)
On Error GoTo errHandler
Dim item As Variant

    For Each item In target.Keys
        If Not compareSrc.Exists(item) Then target.Remove (item)
    Next
    
Exit Sub
errHandler:
    Call Print_Error_Message(Error_Info, "VBT_ZeFuse_Glb", "DicSameKept")
    If AbortTest Then Exit Sub Else Resume Next
End Sub

Public Sub DicCloned(target As Dictionary, Src As Dictionary)
On Error GoTo errHandler
Dim item As Variant

    If target Is Nothing Then Set target = New Dictionary
    For Each item In Src.Keys
        target.Add item, Src(item)
    Next
    
Exit Sub
errHandler:
    Call Print_Error_Message(Error_Info, "VBT_ZeFuse_Glb", "DicCloned")
    If AbortTest Then Exit Sub Else Resume Next
End Sub

Private Function GetEnableStr() As String
On Error GoTo errHandler
Dim iEnable As Integer, cfg As String
Dim flags() As String, word As Variant

    Call TheExec.DataManager.GetEnableFlags(flags)
    For Each word In flags
        cfg = UCase(word)
        If cfg Like "CFG_*" And CfgDataBase.DicCfgList.Exists(Replace(cfg, "CFG_", "")) Then
            If TheExec.enableWord(cfg) Then GetEnableStr = GetEnableStr + Replace(cfg, "CFG_", "") + " "
        End If
    Next
 
Exit Function
errHandler:
    Call Print_Error_Message(Error_Info, "VBT_ZeFuse_Glb", "GetEnableStr")
    If AbortTest Then Exit Function Else Resume Next
End Function

Public Function CfgEnableCheck() As Boolean
On Error GoTo errHandler
Dim enableCnt As Integer, cfgAct As String

    SelectedCfg = GetEnableStr()
    SelectedCfg = Trim(SelectedCfg)
    enableCnt = UBound(Split(SelectedCfg, " ")) + 1

    If enableCnt > 1 Or (enableCnt = 0 And Not bAllowNoCfgSel) Then
        CfgEnableCheck = False
        If enableCnt > 1 Then MsgBox "[Error] more than 1 configs selected! => " & SelectedCfg:   CfgEnableCheck = False: Exit Function
        If enableCnt = 0 Then
            If Not BdfDataBase.CFG_Only Then
                MsgBox "[Error] no config selected! "
                TheExec.Flow.TestLimit resultVal:=999, lowVal:=1, hiVal:=1, Tname:="CfgEnableCheck"
                Exit Function
            Else
                Set CfgDataBaseAct = CfgDataBase
                CfgEnableCheck = True: Exit Function 'in Tesla, it used default only
            End If
        End If
    End If

    If TheExec.enableWord("CFG_SVM") Then
        If CfgDataBaseSvm Is Nothing Then
            GlbUtility.MessageBox "No SVM config! Please check it again!"
            TheExec.Flow.TestLimit resultVal:=999, lowVal:=1, hiVal:=1, Tname:="CfgEnableCheck"
            CfgEnableCheck = False: Exit Function
        End If
    End If
    CfgEnableCheck = True
    
Exit Function
errHandler:
    Call Print_Error_Message(Error_Info, "VBT_ZeFuse_Glb", "CfgEnableCheck")
    If AbortTest Then Exit Function Else Resume Next
End Function

Public Function eFuse_DSSC_CapSetup(patt As String, bank As eFuseBdfBank, SignalName As String, capWave As DSPWave, Optional testType As ReadTestType = DefaultMode)
On Error GoTo errHandler
Dim SiteVarValue As SiteLong
Dim nBits As Long, SampleSize As Long

    nBits = IIf(bank.DoubleBits, 2, 1)
      
    If TestType = JTAG_APB Then
        SampleSize = bank.FullSize * nBits 'Size * double bits
    ElseIf TestType = JTAG_Or_Result Then
        SampleSize = bank.FullSize 'Size
    ElseIf bank.pgmMode = pgm_JTAG Then
        SampleSize = bank.FullSize * nBits 'Size
    Else
        SampleSize = IIf(bank.pgmMode = pgm_DAA, bank.FullSize / bank.CapturePinCnt * nBits, bank.FullSize * 1)  'Size * single or double bits`@
    End If
    thehdw.DSSC.Pins(bank.CapturePin).Pattern(patt).Capture.Signals.Add (SignalName)
    With thehdw.DSSC.Pins(bank.CapturePin).Pattern(patt).Capture.Signals.item(SignalName)
            .SampleSize = SampleSize
            .LoadSettings
    End With
    capWave = thehdw.DSSC.Pins(bank.CapturePin).Pattern(patt).Capture.Signals(SignalName).DSPWave
    
    thehdw.Digital.Patgen.HaltMode = tlHaltOnOpcode
    
Exit Function
errHandler:
    Call Print_Error_Message(Error_Info, "VBT_ZeFuse_Glb", "eFuse_DSSC_CapSetup")
    If AbortTest Then Exit Function Else Resume Next
End Function

Public Function eFuse_DSSC_SrcSetup(patt As String, bank As eFuseBdfBank, SignalName As String, srcWave As DSPWave, Optional BySite As Boolean = True)
On Error GoTo errHandler
Dim SiteVarValue As SiteLong
Dim WaveDef As String, SampleSize As Long

    WaveDef = "WaveDef"
    SampleSize = GlbUtility.GetSampleSize(srcWave)
    thehdw.DSSC.Pins(bank.SourcePin).Pattern(patt).Source.Signals.Add SignalName
    If BySite Then
        For Each site In TheExec.sites
            TheExec.WaveDefinitions.CreateWaveDefinition WaveDef & site, srcWave, True
            With thehdw.DSSC.Pins(bank.SourcePin).Pattern(patt).Source.Signals(SignalName)
                .WaveDefinitionName = WaveDef & site
            End With
        Next
    Else
        TheExec.WaveDefinitions.CreateWaveDefinition WaveDef, srcWave, True
        With thehdw.DSSC.Pins(bank.SourcePin).Pattern(patt).Source.Signals(SignalName)
            .WaveDefinitionName = WaveDef
        End With
    End If
    
    With thehdw.DSSC.Pins(bank.SourcePin).Pattern(patt).Source.Signals(SignalName)
        .SampleSize = SampleSize
        .Amplitude = 1
        '.LoadSamples
        .LoadSettings
    End With
    thehdw.DSSC.Pins(bank.SourcePin).Pattern(patt).Source.Signals.DefaultSignal = SignalName
    
Exit Function
errHandler:
    Call Print_Error_Message(Error_Info, "VBT_ZeFuse_Glb", "eFuse_DSSC_SrcSetup")
    If AbortTest Then Exit Function Else Resume Next
End Function

'20210705,Modify for CP series set PRR
Public Sub GenerateEcid(opbank As eFuseBdfBank, lotId As String, lotidpf As SiteBoolean, wfrid As Long, _
    wfridpf As SiteBoolean, xcoor As SiteLong, xcoorpf As SiteBoolean, ycoor As SiteLong, ycoorpf As SiteBoolean, dicLotInfo As Dictionary)
On Error GoTo errHandler
Dim ecid As New SiteVariant   ' LotID(36) +WfrID(5) +X(6) +Y(6)
Dim field As eFuseBdfField, hexecid As New SiteVariant, ttlbits As Long: ttlbits = 0
Dim hvalue As SiteVariant
Dim F_updateEarlyDSSC As Boolean
Dim realJob As String: realJob = vbNullString
Dim tmp As String

    F_updateEarlyDSSC = False
    '''LotID
    Set field = opbank.Fields(dicLotInfo("lotid"))
    ecid = GlbUtility.GetLotIdBin(lotId, field.size, True)
    If UCase(ecid) Like "*ERROR101*" Then
        TheExec.Flow.TestLimit resultVal:=0, lowVal:=1, hiVal:=1, Tname:="GetLotIdBin"
        Exit Sub
    End If
    realJob = BdfDataBase.GetRealStage(field.BlowLocation)
    If Replace(UCase(realJob), "_EARLY", "") = TheExec.CurrentJob Then
        If Not field.IsEarlyFuse Then
            If LCase(Replace(UCase(realJob), "_EARLY", "")) Like "wlft*" And field.Algorithm = alg_lotid Then
                tmp = StrReverse(GlbUtility.GetLotIdBin(lotId, field.size))
                ecid = GlbUtility.Bin2HexStr(tmp, (field.size - 1) \ 4 + 1)
            End If
            opbank.SetEfuse dicLotInfo("lotid"), ecid, lotidpf
        Else
            field.TrimAteValue = ecid
            field.TrimFlagBySite = True
            F_updateEarlyDSSC = True
        End If
    End If
    hexecid = GlbUtility.GetLotIdBin(lotid, field.size): ttlbits = ttlbits + field.size
    
    '''WaferID
    Set field = opbank.Fields(dicLotInfo("waferid"))
    If LCase(Replace(UCase(realJob), "_EARLY", "")) Like "wlft*" And field.Algorithm = alg_numeric Then
        ecid = GlbUtility.Bin2HexStr(StrReverse(GlbUtility.Dec2Bin(CDbl(wfrid), field.size)), (field.size - 1) \ 4 + 1)
    Else
        ecid = GlbUtility.Dec2HexStr(CDbl(wfrid), 2)
    End If
    If Replace(UCase(realJob), "_EARLY", "") = TheExec.CurrentJob Then
        If Not field.IsEarlyFuse Then
            opbank.SetEfuse dicLotInfo("waferid"), ecid, wfridpf
        Else
            field.TrimAteValue = ecid
            field.TrimFlagBySite = True
            F_updateEarlyDSSC = True
        End If
    End If
    hexecid = hexecid & GlbUtility.Dec2Bin(wfrid, field.size): ttlbits = ttlbits + field.size
    
    '''Coor X
    Set field = opbank.Fields(dicLotInfo("xcoordinate")): ttlbits = ttlbits + field.size
    For Each site In TheExec.sites
        If LCase(Replace(UCase(realJob), "_EARLY", "")) Like "wlft*" And field.Algorithm = alg_numeric Then
            ecid(site) = GlbUtility.Bin2HexStr(StrReverse(GlbUtility.Dec2Bin(CDbl(xcoor), field.size)), (field.size - 1) \ 4 + 1)
        Else
            ecid(site) = GlbUtility.Dec2HexStr(CDbl(xcoor), 2)
        End If
        hexecid = hexecid & GlbUtility.Dec2Bin(xcoor, field.size)
    Next
    
    If Replace(UCase(realJob), "_EARLY", "") = TheExec.CurrentJob Then
        If Not field.IsEarlyFuse Then
            opbank.SetEfuse dicLotInfo("xcoordinate"), ecid, xcoorpf
        Else
            field.TrimAteValue = ecid
            field.TrimFlagBySite = True
            F_updateEarlyDSSC = True
        End If
    End If
    
    '''Coor Y
    Set field = opbank.Fields(dicLotInfo("ycoordinate")): ttlbits = ttlbits + field.size
    For Each site In TheExec.sites
        If LCase(Replace(UCase(realJob), "_EARLY", "")) Like "wlft*" And field.Algorithm = alg_numeric Then
            ecid(site) = GlbUtility.Bin2HexStr(StrReverse(GlbUtility.Dec2Bin(CDbl(ycoor), field.size)), (field.size - 1) \ 4 + 1)
        Else
            ecid(site) = GlbUtility.Dec2HexStr(CDbl(ycoor), 2)
        End If
        hexecid = hexecid & GlbUtility.Dec2Bin(ycoor, field.size)
    Next
    
    If Replace(UCase(realJob), "_EARLY", "") = TheExec.CurrentJob Then
        If Not field.IsEarlyFuse Then
            opbank.SetEfuse dicLotInfo("ycoordinate"), ecid, xcoorpf
        Else
            field.TrimAteValue = ecid
            field.TrimFlagBySite = True
            F_updateEarlyDSSC = True
        End If
    End If

    Set opbank.EcidHexStr = New SiteVariant
    For Each site In TheExec.sites
        opbank.EcidHexStr = Replace(GlbUtility.Bin2HexStr(StrReverse(hexecid), 16), "0x", "")
    Next
    
    Call PutProd_YearMonthDay(opbank)

    If F_updateEarlyDSSC = True Then
        opbank.UpdateField2EarlyDSSC
    End If
    
Exit Sub
errHandler:
    Call Print_Error_Message(Error_Info, "VBT_ZeFuse_Glb", "GenerateEcid")
    If AbortTest Then Exit Sub Else Resume Next
End Sub

Public Function EfuseTrimData_ATETrimValue_BitWidth_Read(fieldStr As String) As SiteLong
On Error GoTo errHandler
Dim value As New SiteLong, field As eFuseBdfField

    If Not BdfDataBase.FieldBankMap.Exists(fieldStr) Then GoTo errHandler
    Set field = BdfDataBase.FieldBankMap(fieldStr)
    value = field.size
    Set EfuseTrimData_ATETrimValue_BitWidth_Read = value

Exit Function
errHandler:
    Call Print_Error_Message(Error_Info, "VBT_ZeFuse_Glb", "EfuseTrimData_ATETrimValue_BitWidth_Read")
    If AbortTest Then Exit Function Else Resume Next
End Function

Public Sub EfuseTrimData_ATETrimValue_Write(fieldStr As String, TrimValue As SiteVariant, Optional bankstr As String, Optional forceWrite As Boolean = False, Optional pf As SiteBoolean = True)
On Error GoTo errHandler
Dim opbank As eFuseBdfBank

    If AllFieldUnique Then
        If Not BdfDataBase.FieldBankMap.Exists(fieldStr) Then GoTo errHandler
        Set opbank = GetBdfBank(BdfDataBase.FieldBankMap(fieldStr).bankName)
        opbank.SetEfuse fieldStr, TrimValue, PatPF:=pf, forceWrite:=forceWrite
    Else
        Dim field As New eFuseBdfField
        bankstr = "CFG"
        Set opbank = GetBdfBank(bankstr)
        If Not opbank.Fields.Exists(fieldStr) Then GoTo errHandler
        Set field = opbank.Fields(fieldStr)
        opbank.SetEfuse fieldStr, TrimValue, PatPF:=pf, forceWrite:=forceWrite
    End If

Exit Sub
errHandler:
    Call Print_Error_Message(Error_Info, "VBT_ZeFuse_Glb", "EfuseTrimData_ATETrimValue_Write")
    If AbortTest Then Exit Sub Else Resume Next
End Sub

Public Function EfuseTrimData_ATETrimValue_Read(fieldStr As String) As SiteLong
On Error GoTo errHandler
Dim value As New SiteLong, opbank As eFuseBdfBank

    If Not BdfDataBase.FieldBankMap.Exists(fieldStr) Then GoTo errHandler
    Set opbank = GetBdfBank(BdfDataBase.FieldBankMap(fieldStr).bankName)
    If GlbUtility.testedStages.Count = 0 Or Not TheExec.enableWord("Dut2Db") Then
        value = opbank.GetTrimeFuseValue(fieldStr)
    Else
        value = opbank.GeteFuseValue(fieldStr)
    End If
    Set EfuseTrimData_ATETrimValue_Read = value

Exit Function
errHandler:
    Call Print_Error_Message(Error_Info, "VBT_ZeFuse_Glb", "EfuseTrimData_ATETrimValue_Read")
    If AbortTest Then Exit Function Else Resume Next
End Function

Private Sub PutProd_YearMonthDay(opbank As eFuseBdfBank)
On Error GoTo errHandler
Dim localDate As Date, ProdYear As New SiteLong, ProdMonth As New SiteLong, ProdDay As New SiteLong
Dim PatPF As New SiteBoolean: PatPF = True

    If Not opbank.Fields.Exists("prod_year") Then Exit Sub
    localDate = DateTime.Now
    ProdYear = CLng(Year(localDate) - 2000)
    ProdMonth = CLng(Month(localDate))
    ProdDay = CLng(Day(localDate))
    opbank.SetEfuse "prod_year", ProdYear, PatPF
    opbank.SetEfuse "prod_month", ProdMonth, PatPF
    opbank.SetEfuse "prod_day", ProdDay, PatPF
    
Exit Sub
errHandler:
    Call Print_Error_Message(Error_Info, "VBT_ZeFuse_Glb", "PutProd_YearMonthDay")
    If AbortTest Then Exit Sub Else Resume Next
End Sub

Public Function UpdateFullCfg2BdfEfuse()
On Error GoTo errHandler
Dim fieldStr As Variant, field As eFuseBdfField, cfgfield As eFuseCfgField

    For Each fieldStr In BdfDataBase.Bank_Cfg.Fields.Keys
        Set field = BdfDataBase.Bank_Cfg.Fields(CStr(fieldStr))
        Set cfgfield = CfgDataBaseAct.Fields(CStr(fieldStr))
        field.FieldDefault = cfgfield.FieldActive
        field.DefaultOrReal = cfgfield.DefaultOrReal
    Next

Exit Function
errHandler:
    Call Print_Error_Message(Error_Info, "VBT_ZeFuse_Glb", "UpdateFullCfg2BdfEfuse")
    If AbortTest Then Exit Function Else Resume Next
End Function

Public Function UpdateCfg2BdfEfuse(hvalue As String, data As DataRetrieve, category As String)
On Error GoTo errHandler
Dim field As eFuseBdfField ', category As String:category = CfgEarlyTable

    Set field = BdfDataBase.Bank_Cfg.Fields(category)
    Select Case data:
        Case DataRetrieve.DafaultValue
            field.FieldDefault = hvalue
            field.Default = hvalue
        Case DataRetrieve.AteTrimData
            field.TrimAteValue = hvalue
            field.TrimFlagBySite = True
           ' field.TrimTestInstance = category
        Case DataRetrieve.TrimmedData
            field.TrimmedData = hvalue
            field.TrimAteValue = hvalue
            field.TrimFlagBySite = True
    End Select

    Set BdfDataBase.Bank_Cfg.BlankMaskDspWave = CfgDataBaseAct.BlankMaskDspWave
    Set BdfDataBase.Bank_Cfg.BlankMaskDspWaveFutureStages = CfgDataBaseAct.BlankMaskDspWaveFutureStages
        
Exit Function
errHandler:
    Call Print_Error_Message(Error_Info, "VBT_ZeFuse_Glb", "UpdateCfg2BdfEfuse")
    If AbortTest Then Exit Function Else Resume Next
End Function

Public Function IsInstMatch(segstr As String) As Boolean
On Error GoTo errHandler
Dim instance As String: instance = TheExec.DataManager.instanceName

    IsInstMatch = GlbUtility.IsStrMatch(instance, segstr)

Exit Function
errHandler:
    Call Print_Error_Message(Error_Info, "VBT_ZeFuse_Glb", "IsInstMatch")
    If AbortTest Then Exit Function Else Resume Next
End Function

Public Function DoubleBitChecks(InWave As DSPWave, ByRef out1 As DSPWave, ByRef out2 As DSPWave) As SiteLong
On Error GoTo errHandler
Dim Result As New SiteLong

    Select Case DaaParaBits
        Case 32:
           Call Ze_SplitWave32(InWave, out1, out2)
        Case 8:
           Call Ze_SplitWave8(InWave, out1, out2)
        Case Else
            GlbUtility.MessageBox "No support parallel bits " & DaaParaBits
    End Select
    Result = 999
    Ze_SingleDoubleCal out1, out2, Result
    Set DoubleBitChecks = Result

Exit Function
errHandler:
    Call Print_Error_Message(Error_Info, "VBT_ZeFuse_Glb", "DoubleBitChecks")
    If AbortTest Then Exit Function Else Resume Next
End Function

Public Sub UpdateCmpFields(bankstr As String)
On Error GoTo errHandler
Dim cmpStr As String, cmpBank As eFuseBdfBank, udrBank As eFuseBdfBank
Dim fieldStr As Variant, cmpfield As eFuseBdfField, udrfield As eFuseBdfField

    cmpStr = Replace(LCase(bankstr), "udr", "cmp") 'udr_p => cmp_p
    Set udrBank = GetBdfBank(bankstr)
    Set cmpBank = GetBdfBank(cmpStr)
    For Each fieldStr In cmpBank.Fields.Keys
        Set cmpfield = cmpBank.Fields(fieldStr)
        Set udrfield = udrBank.Fields(fieldStr)
        cmpfield.TrimAteValue = udrfield.DsscValue
        If Not GlbUtility.OnlineMode Then
            cmpfield.TrimmedData = udrfield.TrimmedData
            cmpfield.DsscValue = udrfield.DsscValue
            cmpfield.DsscDecValue = udrfield.DsscDecValue
        End If
    Next

Exit Sub
errHandler:
    Call Print_Error_Message(Error_Info, "VBT_ZeFuse_Glb", "UpdateCmpFields")
    If AbortTest Then Exit Sub Else Resume Next
End Sub

Public Function AllBanksPatFailCnt() As SiteLong
On Error GoTo errHandler
Dim Result As New SiteLong
Dim bankstr As Variant, bank As eFuseBdfBank

    Result = 0
    For Each bankstr In BdfDataBase.Banks
        Set bank = BdfDataBase.Banks(bankstr)
        Set Result = Result.Add(PatFailCnt(bank))
    Next
    Set AllBanksPatFailCnt = Result

Exit Function
errHandler:
    Call Print_Error_Message(Error_Info, "VBT_ZeFuse_Glb", "AllBanksPatFailCnt")
    If AbortTest Then Exit Function Else Resume Next
End Function

Public Function PatFailCnt(bank As eFuseBdfBank) As SiteLong
On Error GoTo errHandler
Dim fieldStr As Variant, field As eFuseBdfField, Result As New SiteLong

    Result = 0
    For Each site In TheExec.sites
        For Each fieldStr In bank.Fields.Keys
           Set field = bank.Fields(fieldStr)
           If Not field.PatTestPass Then
               Result = Result + 1
           End If
        Next
    Next
    If EFUSE_BIN1_SETTING Then Result = 0
    Set PatFailCnt = Result

Exit Function
errHandler:
    Call Print_Error_Message(Error_Info, "VBT_ZeFuse_Glb", "PatFailCnt")
    If AbortTest Then Exit Function Else Resume Next
End Function

Public Sub PutData2RegKeySave(opbank As eFuseBdfBank, earlyfuse As Boolean)
On Error GoTo errHandler
Dim stage As String
Dim site As Variant
    
    If LCase(opbank.name) = "ecid" Then
        For Each site In TheExec.sites
            If TheExec.Datalog.Setup.WaferSetup.GetXCoord(site) = 0 And _
            TheExec.Datalog.Setup.WaferSetup.GetYCoord(site) = 0 Then
                Call Print_Error_Message(Error_Info, "VBT_ZeFuse_Glb", "PutData2RegKeySave", "Please check the x(" & TheExec.Datalog.Setup.WaferSetup.GetXCoord(site) & ")/y(" & TheExec.Datalog.Setup.WaferSetup.GetYCoord(site) & ") coordinates are correct or not")
                TheExec.Flow.TestLimit resultVal:=1, lowVal:=0, hiVal:=0, Tname:="Coordinates_Check"
                Exit Sub
            End If
        Next site
    End If
    If UCase(opbank.name) = UCase(FuseRevisionBankName) Then
        stage = TheExec.CurrentChanMap
        If UCase(stage) Like "CHANNELMAP_CP*" Or UCase(stage) Like "CHANNELMAP_WLFT*" Then
            Call GlbUtility.Put2IEDA("SVM_CFuse_288Bits", opbank.FuseRevisionStr, opbank.name)
        ElseIf UCase(stage) Like "CHANNELMAP_FT*" Then
            Call GlbUtility.Put2IEDA("SVM_CFuse_288Bits", opbank.DsscCFGCondStr, opbank.name)
        Else
            Call Print_Error_Message(Error_Info, "VBT_ZeFuse_Glb", "PutData2RegKeySave", "Can not judge channel map name(" + stage + ")")
        End If
    End If
    'If Not earlyFuse And opbank.Name = "CFG" Then opbank.PutField2IEDA "SVM_CFuse_288Bits", "CFG_condition"
    If LCase(opbank.name) = "ecid" Then
        Call GlbUtility.Put2IEDA("eFuseLotNumber", opbank.DsscLotStr, opbank.name)
        Call GlbUtility.Put2IEDA("eFuseWaferID", opbank.DsscWfrStr, opbank.name)
        Call GlbUtility.Put2IEDA("eFuseDieX", opbank.DsscXcoor, opbank.name)
        Call GlbUtility.Put2IEDA("eFuseDieY", opbank.DsscYcoor, opbank.name)
        Call GlbUtility.Put2IEDA("Hram_ECID_53bit", opbank.EcidBinStr, opbank.name)
    End If

Exit Sub
errHandler:
    Call Print_Error_Message(Error_Info, "VBT_ZeFuse_Glb", "PutData2RegKeySave")
    If AbortTest Then Exit Sub Else Resume Next
End Sub

Private Sub ClearRegKeys()
On Error GoTo errHandler
Dim regKey As String: regKey = vbNullString
Dim tmp As Variant, identifyItem As String

    identifyItem = "eFuseLotNumber,eFuseWaferID,eFuseDieX,eFuseDieY,Hram_ECID_53bit,SVM_CFuse_288Bits"
    For Each site In TheExec.sites.Existing
        If site <> TheExec.sites.Existing.Count - 1 Then regKey = regKey & ","
    Next
    For Each tmp In Split(identifyItem, ",")
        RegKeySave CStr(tmp), regKey
    Next

Exit Sub
errHandler:
    Call Print_Error_Message(Error_Info, "VBT_ZeFuse_Glb", "ClearRegKeys")
    If AbortTest Then Exit Sub Else Resume Next
End Sub

Private Function ObtainEcidCategories() As Boolean
On Error GoTo errHandler
Dim bankstr As Variant

    Set DicEcidCat = Nothing
    For Each bankstr In BdfDataBase.Banks.Keys
        If EcidBankSearch(BdfDataBase.Banks(bankstr)) Then: ObtainEcidCategories = True: If Not MultiEcidBanks Then Exit Function   ' if multiple ecid banks then not Exit Function
    Next
    If ObtainEcidCategories = True Then Exit Function
    If EcidCfgSearch() Then:  ObtainEcidCategories = True: Exit Function
    GlbUtility.MessageBox "Ecid information didn't find! Please confirm it!"
    TheExec.Flow.TestLimit resultVal:=999, lowVal:=1, hiVal:=1, Tname:="ObtainEcidCategories"

Exit Function
errHandler:
    Call Print_Error_Message(Error_Info, "VBT_ZeFuse_Glb", "ClearRegKeys")
    If AbortTest Then Exit Function Else Resume Next
End Function

Private Sub CollectPrintIdx()
On Error GoTo errHandler
Dim bankstr As Variant, bank As eFuseBdfBank

    For Each bankstr In BdfDataBase.Banks.Keys
        Set bank = BdfDataBase.Banks(bankstr)
        Call bank.CollectPrintIdx
    Next

Exit Sub
errHandler:
    Call Print_Error_Message(Error_Info, "VBT_ZeFuse_Glb", "CollectPrintIdx")
    If AbortTest Then Exit Sub Else Resume Next
End Sub

Private Function EcidBankSearch(bank As eFuseBdfBank) As Boolean
On Error GoTo errHandler
Dim item As Variant, field As eFuseBdfField

    EcidBankSearch = False
    If bank Is Nothing Then Exit Function
    Call IniDicEcidCat
    For Each item In bank.Fields.keys
        If DicEcidCat.Exists(CStr(item)) Then
             If DicEcidCat(CStr(item)) = "" Then
                DicEcidCat(CStr(item)) = bank.name
                Set field = bank.Fields(CStr(item))
                If GlbUtility.IsStrMatch(CStr(item), ".*lot.*id") Then LotIdBits = field.size:
                If GlbUtility.IsStrMatch(CStr(item), ".*wafer.*id") Then WfrIdBits = field.size:
                If GlbUtility.IsStrMatch(CStr(item), ".*x.*coor") Then XcoorBits = field.size:
                If GlbUtility.IsStrMatch(CStr(item), ".*y.*coor") Then YcoorBits = field.size
                
                EcidBankSearch = True
                bank.HadDeidFuse = True
             Else
                GlbUtility.MessageBox "Duplicate field name [" & CStr(item) & "]"
                GlbUtility.WriteDlg "[Fatal Error] Duplicate field name [" & CStr(item) & "]"
                EcidBankSearch = False
                Exit Function
             End If
        End If
    Next
    
Exit Function
errHandler:
    Call Print_Error_Message(Error_Info, "VBT_ZeFuse_Glb", "EcidBankSearch")
    If AbortTest Then Exit Function Else Resume Next
End Function

Private Function EcidCfgSearch() As Boolean
On Error GoTo errHandler
Dim item As Variant

    EcidCfgSearch = False
    If CfgDataBaseAct Is Nothing Then Exit Function
    Call IniDicEcidCat
    For Each item In CfgDataBaseAct.Fields.Keys
        If DicEcidCat.Exists(CStr(item)) Then
             If DicEcidCat(CStr(item)) = "" Then
                DicEcidCat(CStr(item)) = CfgDataBaseAct.sheetName
                EcidCfgSearch = True
             Else
                GlbUtility.MessageBox "Duplicate field name [" & CStr(item) & "]"
                Call Print_Error_Message(Error_Info, "VBT_ZeFuse_Glb", "EcidCfgSearch", "Duplicate field name [" & CStr(item) & "]")
                EcidCfgSearch = False
                Exit Function
             End If
        End If
    Next
    If EcidCfgSearch Then
        EcidCfgSearch = CheckEcidCatIntegrity
    End If

Exit Function
errHandler:
    Call Print_Error_Message(Error_Info, "VBT_ZeFuse_Glb", "EcidCfgSearch")
    If AbortTest Then Exit Function Else Resume Next
End Function

Private Sub IniDicEcidCat()
On Error GoTo errHandler
Dim tmp As Variant
Dim opbank As eFuseBdfBank
Dim field As eFuseBdfField
Dim bankstr As Variant
Dim fieldstr As Variant

    GlbUtility.IniDictionary DicEcidCat, MultiEcidBanks
    'If EcidCategories = "" Then Exit Sub     
    For Each bankstr In BdfDataBase.Banks
        If UCase(bankstr) Like "CFG" Or UCase(bankstr) Like "ECID" Then
            Set opbank = BdfDataBase.Banks(bankstr)
            For Each fieldstr In opbank.Fields
                Set field = opbank.Fields(fieldstr)
                
                If field.Algorithm = alg_lotid Or field.Algorithm = alg_numeric Then
                    If Not DicEcidCat.Exists(LCase(field.name)) Then DicEcidCat.Add LCase(field.name), ""
                End If
            Next fieldstr
        End If
    Next bankstr
        
Exit Sub
errHandler:
    Call Print_Error_Message(Error_Info, "VBT_ZeFuse_Glb", "EcidCfgSearch")
    If AbortTest Then Exit Sub Else Resume Next
End Sub

Private Function CheckEcidCatIntegrity() As Boolean
On Error GoTo errHandler
Dim item As Variant

    CheckEcidCatIntegrity = True
    For Each item In DicEcidCat.Keys
        If CStr(item) <> "" Then
            If DicEcidCat(CStr(item)) = "" Then
                CheckEcidCatIntegrity = False
                GlbUtility.MessageBox "Ecid category isn't found! """ & CStr(item) & """"
            End If
        End If
    Next
    If LotIdBits = 0 Or WfrIdBits = 0 Or XcoorBits = 0 Or YcoorBits = 0 Then CheckEcidCatIntegrity = False
        
Exit Function
errHandler:
    Call Print_Error_Message(Error_Info, "VBT_ZeFuse_Glb", "CheckEcidCatIntegrity")
    If AbortTest Then Exit Function Else Resume Next
End Function

Private Function ParseBdfTable() As Boolean
On Error GoTo errHandler

    ParseBdfTable = True
    If BdfSheetName = "" Then
        Exit Function
    ElseIf Not WorksheetExists(BdfSheetName) Then
        Call Print_Error_Message(Error_Info, "VBT_ZeFuse_Glb", "ParseBdfTable", "BDF sheet not found (" & BdfSheetName & "), please check it!!!")
        TheExec.Flow.TestLimit resultVal:=999, lowVal:=1, hiVal:=1, Tname:="BDF_Parser"
        ParseBdfTable = False
        Exit Function
    Else
        If Dir(BankDumpPath, vbDirectory) = "" And (bBankDump Or Not GlbUtility.OnlineMode) Then MkDir BankDumpPath
    
        Set BdfDataBase = New eFuseBdfParser: BdfDataBase.Initialize (GlbUtility.currStage)
        If BdfDataBase.Parse(BdfSheetName, CfgDataBaseAct, autoCfgAssign:=True) Then
            BdfDataBase.PrintCrcField
            If bBankDump Or Not GlbUtility.OnlineMode Then BdfDataBase.DumpBdfData BankDumpPath
        Else
            ParseBdfTable = False
        End If
    End If

Exit Function
errHandler:
    Call Print_Error_Message(Error_Info, "VBT_ZeFuse_Glb", "ParseBdfTable")
    If AbortTest Then Exit Function Else Resume Next
End Function

Private Function ParseCfgSheet(cfgDb As eFuseCfgParser, shName As String) As Boolean
On Error GoTo errHandler

    Set cfgDb = New eFuseCfgParser: cfgDb.Initialize
    ParseCfgSheet = cfgDb.Parse(shName, bLumpStages)

Exit Function
errHandler:
    Call Print_Error_Message(Error_Info, "VBT_ZeFuse_Glb", "ParseCfgSheet")
    If AbortTest Then Exit Function Else Resume Next
End Function

Private Function WorksheetExists(sName As String) As Boolean
On Error GoTo errHandler

    If sName = "" Then WorksheetExists = False: Exit Function
    WorksheetExists = Evaluate("ISREF('" & sName & "'!A1)")

Exit Function
errHandler:
    Call Print_Error_Message(Error_Info, "VBT_ZeFuse_Glb", "WorksheetExists")
    If AbortTest Then Exit Function Else Resume Next
End Function

Private Sub ObtainTrimSelectItems()
On Error GoTo errHandler
Dim filename As String: filename = "./eFuseCtrl/Efuse_Trim_Select.csv"

    If TheExec.enableWord("TrimFileSelect") Then
        Call ForceItemsClear
        ReadTrimSelectOverWrite filename
        FileTrimSelectEver = True
    ElseIf FileTrimSelectEver Then
        Call ForceItemsClear
        FileTrimSelectEver = False
    End If

Exit Sub
errHandler:
    Call Print_Error_Message(Error_Info, "VBT_ZeFuse_Glb", "ObtainTrimSelectItems")
    If AbortTest Then Exit Sub Else Resume Next
End Sub

Private Sub ForceItemsClear()
On Error GoTo errHandler
Dim bankstr As Variant, opbank As eFuseBdfBank

    For Each bankstr In BdfDataBase.Banks
        Set opbank = BdfDataBase.Banks(bankstr)
        opbank.HadForceFuse = False
        opbank.DicForceFields.RemoveAll
    Next

Exit Sub
errHandler:
    Call Print_Error_Message(Error_Info, "VBT_ZeFuse_Glb", "ForceItemsClear")
    If AbortTest Then Exit Sub Else Resume Next
End Sub

Private Sub ReadTrimSelectOverWrite(filename As String, Optional overwrite As Boolean = False)
On Error GoTo errHandler
Dim lineFromFile As String, lineNum As Integer: lineNum = 0
Dim opbank As eFuseBdfBank, field As eFuseBdfField
Dim tmp() As String, fmtIsFieldOnly As Boolean, bankstr As String, idxBank As Long
Dim value As New SiteVariant, pf As New SiteBoolean

    Open filename For Input As #1
    ''-----------------------------------
    ''Format1
    '' col_1     |   col_2   |   col_3
    '' BankName  |  FieldName|   Value
    ''-----------------------------------
    ''Format2
    '' col_1     |   col_2
    '' FieldName |   value
    ''-----------------------------------
    
    Do Until EOF(1)
        Line Input #1, lineFromFile
'            If lineNum = 0 Then
'                 If Not GlbUtility.IsStrMatch(lineFromFile, "\(Bank\),Field Name|" & "Bank,Field Name") Then GlbUtility.MessageBox "Error format on file - " & filename
'            ElseIf lineNum > 0 Then
        If lineNum > 0 Then
            idxBank = 0
            tmp = Split(Replace(Replace(lineFromFile, ",,", ""), " ", ""), ",")
            fmtIsFieldOnly = False
            If BdfDataBase.FieldBankMap.Exists(tmp(0)) Or (UBound(tmp) = 0 And Not overwrite) Or (UBound(tmp) = 1 And overwrite) Then
                fmtIsFieldOnly = True
            End If
            If fmtIsFieldOnly Then
                If Not BdfDataBase.FieldBankMap.Exists(tmp(0)) Then GlbUtility.MessageBox tmp(0) & " can't be found in TrimSelect or OverWrite csv!"
                bankstr = BdfDataBase.FieldBankMap(tmp(0)).bankName
                idxBank = -1
            Else
                bankstr = tmp(0)
            End If
            
            Set opbank = GetBdfBank(bankstr)
            If opbank.IsFieldExist(tmp(idxBank + 1)) Then
                If overwrite Then
                    pf = True:
                    value = GlbUtility.String2Dbl(tmp(idxBank + 2))
                    opbank.SetEfuse tmp(idxBank + 1), value, pf, True
                Else
                    opbank.HadForceFuse = True
                    If Not opbank.DicForceFields.Exists(tmp(idxBank + 1)) Then opbank.DicForceFields.Add tmp(idxBank + 1), bankstr
                End If
            End If
        End If
        lineNum = lineNum + 1
    Loop
    Close #1
    
Exit Sub
errHandler:
    Call Print_Error_Message(Error_Info, "VBT_ZeFuse_Glb", "ReadTrimSelectOverWrite")
    If AbortTest Then Exit Sub Else Resume Next
End Sub

'Public Function auto_eFuse_SwitchToEarlyJob(Optional condstr As String = "early") As Long
'On Error GoTo errHandler
'    Dim funcName As String:: funcName = "auto_eFuse_SwitchJobToEarly"
'
'    Dim mS_TestName As String:: mS_TestName = ""
'
'    mS_TestName = GlbUtility.currStage & "->" & GlbUtility.currStage & "_early"
'    GlbUtility.currStage = GlbUtility.currStage + "_" + condstr
'
'    TheExec.Flow.TestLimit resultVal:=1, LowVal:=0, HiVal:=1, Tname:=mS_TestName
'
'    Exit Function
'errHandler:
'    TheExec.Datalog.WriteComment "<Error> " + funcName + ":: please check it out."
'    If AbortTest Then Exit Function Else Resume Next
'End Function

'Public Function auto_eFuse_SwitchToCurrentJob()
'On Error GoTo errHandler
'    Dim funcName As String:: funcName = "auto_eFuse_SwitchJobToEarly"
'
'    Dim mS_TestName As String:: mS_TestName = ""
'
'    If (GlbUtility.currStage = "cp1_early") Then GlbUtility.currStage = "cp1"
'    mS_TestName = GlbUtility.currStage & "_early" & "->" & GlbUtility.currStage
'
'    TheExec.Flow.TestLimit resultVal:=0, LowVal:=0, HiVal:=1, Tname:=mS_TestName
'
'    Exit Function
'errHandler:
'    TheExec.Datalog.WriteComment "<Error> " + funcName + ":: please check it out."
'    If AbortTest Then Exit Function Else Resume Next
'
'End Function

Public Function BKM_Update(bank As String, cateName As String, Optional useDefautFusing As Boolean = False, Optional DefautFusingValue As Long = 0)
On Error GoTo errHandler
Dim m_LotID As String
Dim m_WaferID As String
Dim m_BKMPath As String
Dim m_BKMTmpName As String
Dim m_BKMFileName As String
Dim m_BKMDecode As String
Dim m_FileType As String
Dim m_DicData As String
Dim m_BKMMappingName() As String
Dim BKM_DECODE_Fuse As New SiteVariant
Dim m_ChangeWafer As Boolean
Dim field As eFuseBdfField
Dim opbank As eFuseBdfBank
Dim localFilePath As String
Dim OutputFile As String
Dim fso As New FileSystemObject
Dim msg As String
    '-------------------------------------------------------------------------
    'BKM folder and file format
    'Folder Path : X:\BKM\LotID_Name   (The LotID_Name depends on your device)
    '         ex : X:\BKM\1N9CG64
    '  File Name : LotID-WaferID_BKMx.x
    '         ex : 1N9CG64-01_BKM4.6
    '-------------------------------------------------------------------------
    m_LotID = TheExec.Datalog.Setup.LotSetup.lotid
    m_WaferID = Format(TheExec.Datalog.Setup.WaferSetup.ID, "00")
    
    m_BKMTmpName = m_LotID & "-" & m_WaferID
    If gS_BKMPreName <> m_BKMTmpName Then m_ChangeWafer = True '' Detect Lot ID Change
    
    gS_BKMPreName = m_BKMTmpName '' Keep the current Lot ID
    m_FileType = "txt"
    
    Set opbank = GetBdfBank(bank)
    If opbank.Fields.Exists(cateName) Then
        Set field = opbank.Fields(cateName)
    Else
        msg = "BKM Category didn't exist in BDF " + bank + " bank"
        GoTo skip
    End If
    
    If useDefautFusing = False Then
        If m_ChangeWafer Then '' Lot ID Change so read BKM again
            If TheExec.TesterMode = testModeOffline Then
                TheExec.Datalog.WriteComment "The TesterMode is Offline mode"
                m_BKMPath = "D:\TEMP_BKM_TRY\" & m_LotID & "\"
                If Not CheckFileExist(m_BKMPath, m_BKMTmpName, m_FileType, m_BKMFileName) Then
                    If (BdfDataBase.DicBKMMap.Count > 1) Then
                        OutputFile = m_BKMPath & m_BKMTmpName + "_" + BdfDataBase.DicBKMMap.Keys(1) & "." & m_FileType
                        fso.CreateTextFile (OutputFile)
                    Else
                        msg = "<Error> BKM info does not exist! "
                        GoTo skip
                    End If
                End If
            Else
                m_BKMPath = Local_BKM_path & m_LotID & "\"
            End If
            
            If Not CheckFileExist(m_BKMPath, m_BKMTmpName, m_FileType, m_BKMFileName) Then
                TheExec.Datalog.WriteComment "<Warning> File:: " + m_BKMPath + m_BKMTmpName + " does not exist! "
                m_BKMPath = BKM_Main_Path & m_LotID & "\"
                If Not (CheckFileExist(m_BKMPath, m_BKMTmpName, m_FileType, m_BKMFileName)) Then
                    TheExec.Datalog.WriteComment "<Error> File:: " + m_BKMPath + m_BKMTmpName + " does not exist! "
                    TheExec.Flow.TestLimit resultVal:=0, lowVal:=1, hiVal:=1, Tname:="BKM_File_Not_Exist"
                    gS_BKMPreName = Empty
                    Exit Function
                End If
            End If
            
            TheExec.Datalog.WriteComment "*** Lot ID Change read BKM ***"
            TheExec.Datalog.WriteComment "Use BKM file: " & m_BKMPath & m_BKMTmpName
            
            m_BKMMappingName = Split(m_BKMFileName, "_")
            m_BKMMappingName(1) = Replace(m_BKMMappingName(1), ".txt", "")
            
            If BdfDataBase.DicBKMMap.Exists(m_BKMMappingName(1)) Then
                m_DicData = BdfDataBase.DicBKMMap(m_BKMMappingName(1))
            Else
                TheExec.Flow.TestLimit resultVal:=0, lowVal:=1, hiVal:=1, Tname:="BKM_Number_Not_Match"
                TheExec.Datalog.WriteComment ""
                TheExec.Datalog.WriteComment "The related BKM Number does not match"
                TheExec.Datalog.WriteComment ""
                Exit Function
            End If
    
            gS_BKM_IEDA = CStr(GlbUtility.Bin2Dec(m_DicData, False))
    
            TheExec.Datalog.WriteComment ""
            TheExec.Datalog.WriteComment " BKM Number is    " & m_BKMMappingName(1) & "    Mapping to Fuse is    " & gS_BKM_IEDA
            TheExec.Datalog.WriteComment ""
            TheExec.Flow.TestLimit resultVal:=CInt(gS_BKM_IEDA), lowVal:=field.Llimit, hiVal:=field.Hlimit, Tname:="BKM_Value", ForceResults:=tlForceNone
            TheExec.Datalog.WriteComment ""
        End If  '' Lot ID Change so read BKM again
    Else
        gS_BKM_IEDA = CStr(DefautFusingValue)
        TheExec.Datalog.WriteComment " BKM Number use Default Value : " & gS_BKM_IEDA
    End If

    If LCase(field.BlowLocation) = LCase(GlbUtility.currStage) Then
        BKM_DECODE_Fuse = CInt(gS_BKM_IEDA)
        opbank.SetEfuse cateName, BKM_DECODE_Fuse
    End If

    BdfDataBase.ReadRealBKMdone = True

Exit Function
skip:
    GlbUtility.WriteDlg "<Error> " + msg
    TheExec.Flow.TestLimit resultVal:=999, lowVal:=1, hiVal:=1, Tname:="BKM_Update"
Exit Function
errHandler:
    Call Print_Error_Message(Error_Info, "VBT_ZeFuse_Glb", "BKM_Update")
    If AbortTest Then Exit Function Else Resume Next
End Function

Public Function CheckFileExist(folderPath As String, filename_tmp As String, fileType As String, ByRef m_file As String, Optional ByRef m_filepath As String, Optional ByRef FolderEmpty_Falg As Boolean) As Boolean
On Error GoTo errHandler
Dim Obj_GetFuseFile As New FileSystemObject
Dim Obj_FuseFolder As Folder
Dim Obj_FuseFile As File
Dim m_CheckFileName As String
Dim m_counter As Long: m_counter = 0
Dim ReTry_Count As Integer: ReTry_Count = 0
Dim ReTry_MaxLimit As Integer: ReTry_MaxLimit = 5

    ''Check the folder is exist
    If (Obj_GetFuseFile.FolderExists(folderPath)) Then
        FolderEmpty_Falg = False
    Else
    ''if the folder isn't exist, then create new one.
        Do While (ReTry_Count < ReTry_MaxLimit)
            Call Shell("cmd /k md " & folderPath, 0)
            TheExec.Flow.Wait 0.5
            TheHdw.Wait 0.5
            If (Obj_GetFuseFile.FolderExists(folderPath)) Then
                FolderEmpty_Falg = False
                Exit Do
            Else
                ReTry_Count = ReTry_Count + 1
            End If
            FolderEmpty_Falg = True
        Loop
    End If

    If (ReTry_Count >= ReTry_MaxLimit) Then Exit Function
    Set Obj_FuseFolder = Obj_GetFuseFile.GetFolder(folderPath)

    ''20210127
    m_CheckFileName = LCase("*" & filename_tmp & "*." & fileType)
    
    ''----------------------------------------------------------------------------------------------------------------
    ''Check the file of folder.
    ''counter  = 0 -> the file isn't exist.
    ''counter  = 1 -> find the correct file.
    ''counter >= 2 -> more than one file match the set condition. It means that we have no idea to use which one file.
    ''----------------------------------------------------------------------------------------------------------------
    
    For Each Obj_FuseFile In Obj_FuseFolder.Files
        If (LCase(Obj_FuseFile.name) Like m_CheckFileName) Then
            m_counter = m_counter + 1
            m_file = Obj_FuseFile.name
            m_filepath = Obj_FuseFile.path
            If (m_counter >= 2) Then
                Call Print_Error_Message(Error_Info, "VBT_ZeFuse_Glb", "CheckFileExist", "File:: " + filename_tmp + "is more than one , please check it out!!")
                CheckFileExist = False
                Exit Function
                'GoTo errHandler
            End If
        End If
    Next
        
    If (m_counter = 0) Then
        CheckFileExist = False
'        TheExec.Datalog.WriteComment "<Error> File:: " + filename_tmp + " doesn't exist, please check it!"
'        GoTo errHandler
        Exit Function
    End If

    CheckFileExist = True
    
Exit Function
errHandler:
    Call Print_Error_Message(Error_Info, "VBT_ZeFuse_Glb", "CheckFileExist")
    If AbortTest Then Exit Function Else Resume Next
End Function

'Parse BKM_Info Sheet to Dictionay
'BKM_Info sheet is a mapping table for BKM fused data.
'Get the mapping data from X:\\BKM\
Private Function ParseBKMTable() As Boolean
On Error GoTo errHandler

    ParseBKMTable = True
    If BKMSheetName <> "" Then
        If Not WorksheetExists(BKMSheetName) Then
            Call Print_Error_Message(Error_Info, "VBT_ZeFuse_Glb", "ParseBKMTable", "BKM sheet not found (" & BKMSheetName & ") in TP, please check it!!!")
            ParseBKMTable = False
        Else
            ParseBKMTable = BdfDataBase.ParseBKMInfo(BKMSheetName)
        End If
    End If
    '20210317 Add to get BKMInfo from file for backup
    'BdfDataBase.ParseBKMInfoFromFile
    
Exit Function
errHandler:
    Call Print_Error_Message(Error_Info, "VBT_ZeFuse_Glb", "ParseBKMTable")
    If AbortTest Then Exit Function Else Resume Next
End Function

'202004xx for ap
'Get fused BKMData from decode data with efuse obj.
Public Function GetFusedBKMData(FuseType As String, cateName As String)
On Error GoTo errHandler
Dim m_FusedData As New SiteDouble
Dim opbank As eFuseBdfBank
Dim m_site As Variant

    Set opbank = GetBdfBank(FuseType)
    m_FusedData = opbank.GeteFuseValue(cateName)
    
    For Each m_site In TheExec.sites
        opbank.DsscBKMStr = m_FusedData
    Next m_site
    
    If BdfDataBase.ReadRealBKMdone = True Then
        TheExec.Flow.TestLimit m_FusedData, CInt(gS_BKM_IEDA), CInt(gS_BKM_IEDA), Tname:="BKM_Group_Index"
    Else
        TheExec.Flow.TestLimit m_FusedData, 0, 15, Tname:="BKM_Group_Index"
    End If

Exit Function
errHandler:
    Call Print_Error_Message(Error_Info, "VBT_ZeFuse_Glb", "GetFusedBKMData")
    If AbortTest Then Exit Function Else Resume Next
End Function

Public Function WriteIEDARegistry(FuseType As String, RegistryName As String, cateName As String)
On Error GoTo errHandler
Dim opbank As eFuseBdfBank
Dim m_value As New SiteVariant
Dim m_valueStr As New SiteVariant
Dim m_site As Variant

    Set opbank = GetBdfBank(FuseType)
    '------------------------------------------------------------------
    'Follow T-Si rule.
    'Case "BKM"     - write the BKM programming data to IEDA register.
    'Case "BKMFuse" - write the BKM read back data to IEDA register.
    '------------------------------------------------------------------
    
    If RegistryName = "BKM" Then
        m_value = opbank.GetTrimeFuseValue(cateName)
    Else
        m_value = opbank.GeteFuseValue(cateName)
    End If
    
    For Each m_site In TheExec.sites
        m_valueStr = CStr(m_value)
    Next
    
    Call GlbUtility.Put2IEDA(RegistryName, m_valueStr)
    
Exit Function
errHandler:
    Call Print_Error_Message(Error_Info, "VBT_ZeFuse_Glb", "WriteIEDARegistry")
    If AbortTest Then Exit Function Else Resume Next
End Function

Public Sub Glb_Setup()
On Error GoTo errHandler

    gB_Dut2Db = False
    If TheExec.enableWord("Dut2Db") Then gB_Dut2Db = True
    
    gB_eFuse_Disable_ChkLMT_Flag = False
    gB_eFuse_Disable_ChkLMT_Flag = TheExec.Flow.enableWord("eFuse_Disable_ChkLMT")
    
    '20210812, Add for disable efuse datalog printing
    gB_eFuse_Disable_DSPwavePrint_Flag = False
    gB_eFuse_Disable_DSPwavePrint_Flag = TheExec.Flow.enableWord("eFuse_Disable_DSPwavePrint")
    
    gB_eFuse_Disable_DecodeDataPrint_Flag = False
    gB_eFuse_Disable_DecodeDataPrint_Flag = TheExec.Flow.enableWord("eFuse_Disable_DecodeDataPrint")
    
    ''20220524, modify for syntax check only check current job fields
    gB_eFuse_Disable_SyntaxCheckAll_Flag = False
    gB_eFuse_Disable_SyntaxCheckAll_Flag = TheExec.Flow.enableWord("eFuse_Disable_SyntaxCheckAll")
    
    '20220106, Add for Real value validation( reverse bit)
    gB_efuse_DicValue_Chk_Flag = False
    gB_efuse_DicValue_Chk_Flag = TheExec.Flow.enableWord("Efuse_DicValue_Chk")

    '20230530 ECID Sorting, Add ECID Sorting Enable word
    EFUSE_ECID_SORTING_ENABLE = False
    EFUSE_ECID_SORTING_ENABLE = TheExec.Flow.enableWord("ECIDSort_Enable")
    
    EFUSE_ECID_SORTING_2CMODE = False
    EFUSE_ECID_SORTING_2CMODE = TheExec.Flow.enableWord("Enable_ECID_Sorting_2CMode")
    
    gS_BKM_Unknown = vbNullString
    gB_Fuse_Skip = False

    MixPseudoFuseEnable = False
    ForceDecodeEnable = False   '20210812, Add for efuse TTR bank read do decode
    
    ''====20201230 add for efuse new code====
    If GetPseudoFuseFileOneTime = False Then
        LotID_ForPseudoFuse = TheExec.Datalog.Setup.LotSetup.lotid
        WaferID_ForPseudoFuse = TheExec.Datalog.Setup.WaferSetup.ID
        DateTime_ForPseudoFuse = Replace(Date, "/", "_") + "_" + Replace(mid(Format(Time, "hh:mm:ss"), 1, 2), ":", "")
        GetPseudoFuseFileOneTime = True
    Else
        If (LotID_ForPseudoFuse <> TheExec.Datalog.Setup.LotSetup.lotid) Or (WaferID_ForPseudoFuse <> TheExec.Datalog.Setup.WaferSetup.ID) Then
            LotID_ForPseudoFuse = TheExec.Datalog.Setup.LotSetup.lotid
            WaferID_ForPseudoFuse = TheExec.Datalog.Setup.WaferSetup.ID
            DateTime_ForPseudoFuse = Replace(Date, "/", "_") + "_" + Replace(mid(Format(Time, "hh:mm:ss"), 1, 2), ":", "")
            ParsePseudoFuseFile = False
        End If
    End If
    
    '20211210, Modify for FT pseudo fuse with real package
    If TheExec.Flow.enableWord("eFuse_FT_RealPackage_PseudoFuse") Then gB_Package_PsudoFuse = True
    '20230307, Modify for FT pseudo fuse with wafer
    If TheExec.Flow.enableWord("eFuse_FT_Wafer_PseudoFuse") Then gB_FT_Wafer_PsudoFuse = True
    
    '20220803, Add for single double bit debug print
    gB_efuse_DebugPrint_SingleDoubleBits_Flag = TheExec.Flow.enableWord("Efuse_DbgPrint_SingleDoubleBits")
    
Exit Sub
errHandler:
    Call Print_Error_Message(Error_Info, "VBT_ZeFuse_Glb", "Glb_Setup")
    If AbortTest Then Exit Sub Else Resume Next
End Sub

Private Function ParseDramTable() As Boolean
On Error GoTo errHandler

    ParseDramTable = True
    If DramSheetName = "" Or Not WorksheetExists(DramSheetName) Then Exit Function
    ParseDramTable = BdfDataBase.ParseDram(DramSheetName)

Exit Function
errHandler:
    Call Print_Error_Message(Error_Info, "VBT_ZeFuse_Glb", "ParseDramTable")
    If AbortTest Then Exit Function Else Resume Next
End Function

Private Function ParseCPMTable() As Boolean
On Error GoTo errHandler

    ParseCPMTable = True
    If CPMSheetName = "" Or Not WorksheetExists(CPMSheetName) Then Exit Function
    ParseCPMTable = BdfDataBase.ParseCPM(CPMSheetName)

Exit Function
errHandler:
    Call Print_Error_Message(Error_Info, "VBT_ZeFuse_Glb", "ParseCPMTable")
    If AbortTest Then Exit Function Else Resume Next
End Function

Private Function ParseBinChkTable() As Boolean
On Error GoTo errHandler

    ParseBinChkTable = True
    If BinChkSheetName = "" Or Not WorksheetExists(BinChkSheetName) Then Exit Function
    ParseBinChkTable = ParseBinChk(BinChkSheetName)

Exit Function
errHandler:
    Call Print_Error_Message(Error_Info, "VBT_ZeFuse_Glb", "ParseBinChkTable")
    If AbortTest Then Exit Function Else Resume Next
End Function

'20211110,Add for two type bincut table syntax check
Private Function ParseBincutTable() As Boolean
On Error GoTo errHandler
Dim i As Long, j As Long '20230526 redim

    ParseBincutTable = True
    ReDim EfuseBinCut(MaxPerformanceModeCount, Total_Bincut_Num)
    ReDim EfuseBinCutAddition(MaxPerformanceModeCount, Total_Bincut_Num) As EFUSE_BINCUT_TYPE
    For i = 0 To MaxPerformanceModeCount
        For j = 0 To Total_Bincut_Num
            ReDim EfuseBinCut(i, j).c(MaxEqnNum) As Double
            ReDim EfuseBinCut(i, j).m(MaxEqnNum) As Double
            ReDim EfuseBinCut(i, j).CP_Vmax(MaxEqnNum) As Double
            ReDim EfuseBinCut(i, j).CP_Vmin(MaxEqnNum) As Double
            ReDim EfuseBinCut(i, j).CP_GB(MaxEqnNum) As Double
            
            ReDim EfuseBinCutAddition(i, j).c(MaxEqnNum) As Double
            ReDim EfuseBinCutAddition(i, j).m(MaxEqnNum) As Double
            ReDim EfuseBinCutAddition(i, j).CP_Vmax(MaxEqnNum) As Double
            ReDim EfuseBinCutAddition(i, j).CP_Vmin(MaxEqnNum) As Double
            ReDim EfuseBinCutAddition(i, j).CP_GB(MaxEqnNum) As Double
        Next j
    Next i

    If ParseBincut(EfuseBinCut, BincutOriginalSheetName) Then
        If BincutAdditionalSheetName = "" Then Exit Function
        ParseBincutTable = ParseBincut(EfuseBinCutAddition, BincutAdditionalSheetName)
    Else
        ParseBincutTable = False
    End If
    
Exit Function
errHandler:
    Call Print_Error_Message(Error_Info, "VBT_ZeFuse_Glb", "ParseBincutTable")
    If AbortTest Then Exit Function Else Resume Next
End Function

'20210712,Add for cmp fusing
Private Function ParseCmpFuseTable() As Boolean
On Error GoTo errHandler

    gB_CmpFuseEnable = False
    ParseCmpFuseTable = True
    If CmpFuseSheetName = "" Or Not WorksheetExists(CmpFuseSheetName) Then Exit Function
    If BdfDataBase.ParseCmpFuse(CmpFuseSheetName) Then
        gB_CmpFuseEnable = True
    Else
        ParseCmpFuseTable = False
    End If

Exit Function
errHandler:
    Call Print_Error_Message(Error_Info, "VBT_ZeFuse_Glb", "ParseCmpFuseTable")
    If AbortTest Then Exit Function Else Resume Next
End Function

Private Function PrintEfuseConfigSetting()
On Error GoTo errHandler

    If Flag_Efuse_Config_Printed = False Then
        TheExec.Datalog.WriteComment "********************************"
        TheExec.Datalog.WriteComment "*print: " & "Efuse Config" & " start*"
        TheExec.Datalog.WriteComment "********************************"
        
        '''//Run Options/Enable Words
        'TheExec.Datalog.WriteComment "Dut2Db" & "=" & CStr(TheExec.Flow.EnableWord("Dut2Db"))
        
        '20210812, Add for disable efuse detail datalog printing
        TheExec.Datalog.WriteComment "eFuse_Disable_DSPwavePrint" & "=" & CStr(TheExec.Flow.enableWord("eFuse_Disable_DSPwavePrint"))
        TheExec.Datalog.WriteComment "eFuse_Disable_DecodeDataPrint" & "=" & CStr(TheExec.Flow.enableWord("eFuse_Disable_DecodeDataPrint"))
        
        ''20220524, modify for syntax check only check current job fields
        TheExec.Datalog.WriteComment "eFuse_Disable_SyntaxCheckAll" & "=" & CStr(TheExec.Flow.enableWord("eFuse_Disable_SyntaxCheckAll"))
        
        '20220106, Add for Real value validation( reverse bit)
        TheExec.Datalog.WriteComment "Efuse_DicValue_Chk" & "=" & CStr(TheExec.Flow.enableWord("Efuse_DicValue_Chk"))
        
        '20211210, Modify for FT pseudo fuse with real package
        TheExec.Datalog.WriteComment "eFuse_FT_RealPackage_PseudoFuse" & "=" & CStr(TheExec.Flow.enableWord("eFuse_FT_RealPackage_PseudoFuse"))
        '20230307, Modify for FT pseudo fuse with wafer
        TheExec.Datalog.WriteComment "eFuse_FT_Wafer_PseudoFuse" & "=" & CStr(TheExec.Flow.enableWord("eFuse_FT_Wafer_PseudoFuse"))
        
        TheExec.Datalog.WriteComment "eFuse_Disable_ChkLMT" & "=" & CStr(TheExec.Flow.enableWord("eFuse_Disable_ChkLMT"))
        TheExec.Datalog.WriteComment "Pgm2File" & "=" & CStr(TheExec.Flow.enableWord("Pgm2File"))

        '''//Flags of efuse global variable
        TheExec.Datalog.WriteComment "EFUSE_REFUSE_FOR_PTE" & "=" & CStr(EFUSE_REFUSE_FOR_PTE)
        TheExec.Datalog.WriteComment "EFUSE_BIN1_SETTING" & "=" & CStr(EFUSE_BIN1_SETTING)
        TheExec.Datalog.WriteComment "EFUSE_POWER_OFF_SETTING" & "=" & CStr(EFUSE_POWER_OFF_SETTING)
        
        TheExec.Datalog.WriteComment "EFUSE_CFG_READ_BYJTAG" & "=" & CStr(EFUSE_CFG_READ_BYJTAG)
        TheExec.Datalog.WriteComment "EFUSE_ECID_READ_BYJTAG" & "=" & CStr(EFUSE_ECID_READ_BYJTAG)
        
        TheExec.Datalog.WriteComment "EFUSE_SETFUSE_SHOW_ERRMSX" & "=" & CStr(EFUSE_SETFUSE_SHOW_ERRMSX)
        TheExec.Datalog.WriteComment "EFUSE_ALWAYS_CHECK_CRC" & "=" & CStr(EFUSE_ALWAYS_CHECK_CRC)
        TheExec.Datalog.WriteComment "EFUSE_RV_PATTERN_FULL" & "=" & CStr(EFUSE_RV_PATTERN_FULL)
        TheExec.Datalog.WriteComment "EFUSE_PRINT_DOUBLE_HEXMAP" & "=" & CStr(EFUSE_PRINT_DOUBLE_HEXMAP)
        TheExec.Datalog.WriteComment "EFUSE_ECID_SORTING_ENABLE" & "=" & CStr(EFUSE_ECID_SORTING_ENABLE)
        TheExec.Datalog.WriteComment "EFUSE_ECID_SORTING_2CMODE" & "=" & CStr(EFUSE_ECID_SORTING_2CMODE)
        'TheExec.Datalog.WriteComment "gB_Fuse_Skip" & "=" & CStr(gB_Fuse_Skip)
         
        TheExec.Datalog.WriteComment "******************************"
        TheExec.Datalog.WriteComment "*print: " & "Efuse Config" & " end*"
        TheExec.Datalog.WriteComment "******************************"
        
        '''//Use the flag to control printing the config once for all touchdowns.
        'Flag_Efuse_Config_Printed = True
    End If
    
Exit Function
errHandler:
    Call Print_Error_Message(Error_Info, "VBT_ZeFuse_Glb", "PrintEfuseConfigSetting")
    If AbortTest Then Exit Function Else Resume Next
End Function

Public Function DebugPrintSingleDoubleBitCheck(out1 As DSPWave, out2 As DSPWave, site As Variant)
On Error GoTo errHandler
Dim idx As Integer
Dim waferInfo As String
Dim failBits As String

    If (UCase(TheExec.CurrentChanMap) Like "CHANNELMAP_FT*") Then
        waferInfo = HramLotId(site) & "-" & HramWaferId(site) & "_" & HramXCoord(site) & "_" & HramYCoord(site)
    Else
        waferInfo = lotid & "-" & WaferID & "_" & XCoord(site) & "_" & YCoord(site)
    End If
    For idx = 0 To (out1(site).sampleSize - 1)
        If out1(site).data(idx) <> out2(site).data(idx) Then
            failBits = GetCompareFailBit(GlbUtility.Dec2Bin(out1(site).data(idx), 16, False), GlbUtility.Dec2Bin(out2(site).data(idx), 16, False))
            Call Print_Error_Message(Error_Info, "VBT_ZeFuse_Glb", "DebugPrintSingleDoubleBitCheck", "Site" + CStr(site) + "/" + waferInfo + "/Row:" + CStr(idx) + "/" + failBits)
        End If
    Next

Exit Function
errHandler:
    Call Print_Error_Message(Error_Info, "VBT_ZeFuse_Glb", "DebugPrintSingleDoubleBitCheck")
    If AbortTest Then Exit Function Else Resume Next
End Function

Public Function GetCompareFailBit(str1 As String, str2 As String) As String
On Error GoTo errHandler
Dim idx As Integer, strLenM As Integer, strLenL As Integer
Dim compareStr As String
Dim value1 As String
Dim value2 As String
    
    strLenL = Len(str1)
    strLenM = strLenL * 2
    For idx = 1 To (Len(str1))
        value1 = mid(str1, idx, 1)
        value2 = mid(str2, idx, 1)
        If value1 <> value2 Then
            If compareStr = "" Then
                compareStr = "bit" & (strLenM - idx) & "(" & value1 & "):bit" & (strLenL - idx) & "(" & value2 & ")"
            Else
                compareStr = compareStr & ",bit" & (strLenM - idx) & "(" & value1 & "):bit" & (strLenL - idx) & "(" & value2 & ")"
            End If
        End If
    Next
    
    GetCompareFailBit = compareStr
Exit Function
errHandler:
    Call Print_Error_Message(Error_Info, "VBT_ZeFuse_Glb", "GetCompareFailBit")
    If AbortTest Then Exit Function Else Resume Next
End Function

'20220215,Add for Multi RV Pat
Private Function CFG_Multi_RV_PatCalc()
On Error GoTo errHandler
Dim rtnPatNames() As String, patindex As Variant
Dim nBits As Long, sampleSize As Long, onelinesize As Long, onelinecnt As Long
Dim CheckSmapleSizeStart As Long, CheckSmapleSizeStop As Long
Dim i As Long, j As Long
Dim field As eFuseBdfField, item As Variant
Dim EachLineChkReal() As Boolean
    

    nBits = IIf(BdfDataBase.Bank_Cfg.DoubleBits, 2, 1)
    onelinesize = IIf(BdfDataBase.Bank_Cfg.DoubleBits, 16, 32)
    
    sampleSize = BdfDataBase.Bank_Cfg.FullSize
    
    CFG_Multi_RV_PatSetName = "CFG_" + UCase(TheExec.CurrentJob) & "RV_Write"
    rtnPatNames = TheExec.DataManager.Raw.GetPatternsInSet(CFG_Multi_RV_PatSetName, CFG_Multi_RV_PatCnt)
    
    ReDim EachLineChkReal(sampleSize / onelinesize - 1)
    ReDim CFG_Multi_RV_PatDsscCnt(CFG_Multi_RV_PatCnt - 1)
    CFG_Multi_RV_PatDsscTotalCnt = 0
    
    For Each item In BdfDataBase.Bank_Cfg.Fields.Keys
        Set field = BdfDataBase.Bank_Cfg.Fields(item)
        
        If field.DefaultOrReal = dr_real And field.Algorithm <> alg_cond And (Replace(LCase(BdfDataBase.GetRealStage(field.BlowLocation)), "_early", "") = LCase(TheExec.CurrentJob)) Then
            If field.LSB Mod onelinesize = 0 Then
                CheckSmapleSizeStart = field.LSB / onelinesize
            Else
                CheckSmapleSizeStart = Ceiling(field.LSB / (onelinesize)) - 1
            End If
            
            If field.msb Mod onelinesize = 0 Then
                CheckSmapleSizeStop = field.msb / onelinesize
            Else
                CheckSmapleSizeStop = Ceiling(field.msb / (onelinesize)) - 1
            End If
            
            For i = CheckSmapleSizeStart To CheckSmapleSizeStop
                If EachLineChkReal(i) = False Then
                    EachLineChkReal(i) = True
                End If
            Next i
        End If
    Next
    
    For i = 0 To UBound(rtnPatNames)
        onelinecnt = 0
        patindex = Split(rtnPatNames(i), "FLDSSC")
        CheckSmapleSizeStart = sampleSize / CFG_TotalWritePatCnt * CLng(patindex(UBound(patindex))) / onelinesize
        CheckSmapleSizeStop = sampleSize / CFG_TotalWritePatCnt * (CLng(patindex(UBound(patindex))) + 1) / onelinesize - 1
        For j = CheckSmapleSizeStart To CheckSmapleSizeStop
            If EachLineChkReal(j) = True Then
                onelinecnt = onelinecnt + 1
            End If
        Next j
        CFG_Multi_RV_PatDsscCnt(i) = onelinecnt * nBits * onelinesize
        CFG_Multi_RV_PatDsscTotalCnt = CFG_Multi_RV_PatDsscTotalCnt + CFG_Multi_RV_PatDsscCnt(i)
    Next i

Exit Function
errHandler:
    Call Print_Error_Message(Error_Info, "VBT_ZeFuse_Glb", "CFG_Multi_RV_PatCalc")
    If AbortTest Then Exit Function Else Resume Next
End Function

Public Function ShadowAndNonShadowSameStage()
On Error GoTo errHandler
    Dim VddField As Variant
    Dim t_Field As String
    
    If DicVddPmodeShadow.Count <> 0 Then
        For Each VddField In DicVddPmodeShadow.Keys
            t_Field = Replace(LCase(VddField), "_shadow", "")
            If DicVddPmode.Exists(t_Field) Then
                If DicVddPmode(t_Field) = DicVddPmodeShadow(VddField) Then
                    VddPmodeHadSameStage = True
                    Call Print_Error_Message(warning_info, "VBT_ZeFuse_Glb", "ShadowAndNonShadowSameStage", _
                    t_Field & " and " & VddField & " is same stage(" & DicVddPmode(t_Field) & ")!!")
                    TheExec.AddOutput "<Warning> " & t_Field & " and " & VddField & " is same stage(" & DicVddPmode(t_Field) & ")!!", vbBlue, True
                End If
            End If
        Next VddField
    Else
        Call Print_Error_Message(warning_info, "VBT_ZeFuse_Glb", "ShadowAndNonShadowSameStage", "Doesn't have any shadow vdd fields!!")
    End If

Exit Function
errHandler:
    Call Print_Error_Message(Error_Info, "VBT_ZeFuse_Glb", "ShadowAndNonShadowSameStage")
        If AbortTest Then Exit Function Else Resume Next
End Function
