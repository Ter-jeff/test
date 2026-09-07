VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "eFuseBdfBank"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Public name As String
Public sheetName As String
Public DoubleBits As Boolean
Public JtagMsbFirst As Boolean
Public FullSize As Long
Public pgmMode As ProgramMode
Public Fields As Dictionary   ' Database Return
Public Crcs As Dictionary
Public Crcs_Comp As Dictionary
Public CapturePin As PinList
Public CapturePinCnt As Long
Public SourcePin As PinList
Public myUtility As Utility
Public CateNameMaxLen As Long
Public TouchDownRun As Boolean 'for offline simulate use
Public AnyWriteRun As Boolean 'for offline simulate use

Public DaaCapturedWave As DSPWave 'Plain data from pattern capture, in parallel, this could be double bits or single bit
Public DaaCapWaveSerial As DSPWave 'The serial data of capture of entire bank
Public DaaCapWaveSerialFL As DSPWave 'Convert parallel wave to serial wave(has duplicate(2-bit) serial wave)

Public JtagCapturedWave As DSPWave 'Plain data from pattern capture, in serial, this mostly is  single bit
Public JtagCapturedSerial As DSPWave 'The serial data of capture of entire bank

'20221117 Support multi-type program stage
Public StageMaskMultiTyp As Dictionary
Public DicMultiTypFields As Dictionary
Public DicAllTypefields As Dictionary
Public MultiTypeJobList As Dictionary
Public MultiTyepRVsegment As Dictionary
Public MultiTyepRVIndexWave As Dictionary
Public MultiTyepRVReadWave As Dictionary
Public tempMaskMultiTyp As Dictionary

''202004xx
Public StageMask As DSPWave
Public EarlyTestedstagemask As DSPWave
Public PostTestedstagemask As DSPWave
Public EarlyTestedStageMaskCheck As Boolean
Public PostTestedStageMaskCheck As Boolean
Public EarlyStageMask As DSPWave
Public RVsegment As DSPWave
Public DecimalWave As DSPWave
Public DefaultWave As DSPWave
Public RVIndexWave As DSPWave
Public RVReadWave As DSPWave
Private EarlyMaskArray() As Long
Private MaskArray() As Long
Private EarlyTestedMaskArray() As Long
Private PostTestedMaskArray() As Long
Private DefaultArray() As Long

Public BlankMaskDspWave As DSPWave ' Config main table, ex 288 bits wave.
Public BlankMaskDspWaveFutureStages As DSPWave ' Config main table, ex 288 bits wave.

Public PseudoFuseDspWave As DSPWave '202004xx for ap 'Store PseudoFusing Dspwave for exporting to file.
Public PseudoFuseBankHeader As String
Public PseudoFuseBankValueStr As String

Public SingleBitWave As DSPWave 'The serial data of source of entire bank (AteTrimValue)
Public DoubleBitWave As DSPWave 'The serial data of source of entire bank (AteTrimValue)
Public DsscSourcedWave As DSPWave 'Complete data for pattern source, already include the vector repeat
Public DsscWave_Eff As DSPWave 'effective wave
Public EarlyDSSCWave_Eff As DSPWave 'effective wave
Public EarlyDsscSourcedWave As DSPWave  'Source wavet to pattern - effective wave * (2-bit, 1-bit)
Public EarlySingleBitWave As DSPWave 'The serial data of source of entire bank (AteTrimValue)


Private defaultOthersSingleBitWave As DSPWave
Private defaultOthersData() As Double
Private earlyData() As Double
Private earlySingleBitData() As Double
Public HadEarlyFuse As Boolean
Public HadDeidFuse As Boolean
Public HadCrcFuse As Boolean
Public HadCrcCompFuse As Boolean
Public HadVddBinFuse As Boolean
Public HadRandomFuse As Boolean
Public HadIdsFuse As Boolean
Public HadForceFuse As Boolean 'Get highest priority
Public SeparateEarly As Boolean
Public NonFirstfuse As Boolean
'For easily operation - Start
Public DicReadPrintIdx As Dictionary
Public DicWritePrintIdx As Dictionary
Private FieldIdxStData() As Double
Private FieldIdxSpData() As Double
Private FieldIdxMsbData() As Double
Public FieldIdxSt As DSPWave
Public FieldIdxSp As DSPWave
Public FieldIdxMsb As DSPWave
'For easily operation - End
Public DicBinCut As Dictionary
Public DicIds As Dictionary 'This is stored ids at current stage only
Public DicAllIdsList As Dictionary 'This is stored all fields with algorithm is ids
Public DicEarlyFused As Dictionary
Public DicRandomFused As Dictionary
Public DicOthers As Dictionary
Public DicNonDefaultNonEarly As Dictionary
Public DicDefaultIncludeEarly As Dictionary
Public DicSetDefaultVauleToReal As Dictionary
Public DicForceFields As Dictionary
Public SumIdsCodes As New SiteLong
Public SumVddbinCodes As New SiteLong
Public DicCondTables As Dictionary
Public CfgCmpValue As Dictionary '<<= This will pass from config table
Public CfgCmpValueLumpStages As Dictionary '<<= This will pass from config table
Public Base As Double
Public GetEarlyFuseOnly As Boolean
Public IsBlank As New SiteLong '1= Blank, 0 =non-Blank
Public DsscLotStr As SiteVariant, DsscWfrStr As SiteVariant, DsscXcoor As SiteVariant, DsscYcoor As SiteVariant, EcidBinStr As SiteVariant, EcidHexStr As New SiteVariant
Public DsscCFGCondStr As New SiteVariant    ''202004xx for ap
Public FuseRevisionStr As New SiteVariant   ''for CFG 288bit IEDA
Public DsscBKMStr As New SiteVariant    ''202004xx for ap
Public JobExistInBank As Boolean  ''202004xx for ap
Public cmpFlag As Boolean
Public PatName As String
Public IsSameCheckVar As New SiteLong   ''20221004, Add for multi blank check, 1 = same, 0 = different
Public BlankCheckCount As Integer       ''20221004, Add for multi blank check
Public HVLV_MultiJob As String          'Add for record multitype of BlkChk HV/LV
Public SplitReadDspWaveFlag As Boolean  'Add for need to split bank of read pattern
Public NeedJTAGRead As Boolean          'If DAA mode is locked, can trun on this flag for JTAG decode
Public EcidFieldSize As Long
Public CreateFakeValueTime As Long    'Recorded how many times tested fake values of bank have been created.

Public DicCmpFuse As Dictionary   ' 20210713,Add for cmp fuse
Public ReadPatResut As New SiteBoolean  '20240918, Add for store read pattern results to determine whether IEDA is set

Private Sub Class_Initialize()
On Error GoTo errHandler
    Set myUtility = New Utility
    FullSize = 0
    JtagMsbFirst = False
    CateNameMaxLen = 0
    HadEarlyFuse = False
    HadDeidFuse = False
    HadCrcFuse = False
    HadCrcCompFuse = False
    SeparateEarly = True
    NonFirstfuse = False
    SplitReadDspWaveFlag = False
    NeedJTAGRead = False
    Base = -1
    myUtility.IniDictionary Fields
    myUtility.IniDictionary Crcs
    myUtility.IniDictionary Crcs_Comp
    myUtility.IniDictionary DicOthers
    myUtility.IniDictionary DicEarlyFused
    myUtility.IniDictionary DicRandomFused
    myUtility.IniDictionary DicIds
    myUtility.IniDictionary DicAllIdsList
    myUtility.IniDictionary DicBinCut
    myUtility.IniDictionary DicCondTables 'Algorithm equals "cond"
    myUtility.IniDictionary DicNonDefaultNonEarly
    myUtility.IniDictionary DicDefaultIncludeEarly
    myUtility.IniDictionary DicReadPrintIdx
    myUtility.IniDictionary DicWritePrintIdx
    myUtility.IniDictionary DicSetDefaultVauleToReal
    myUtility.IniDictionary DicForceFields
    myUtility.IniDictionary DicCmpFuse   ' 20210713,Add for cmp fuse
    JobExistInBank = False ''202004xx for ap
    cmpFlag = False ''202005xx for ap
    IsSameCheckVar = 1    ''20221004, Add for multi blank check
    BlankCheckCount = 0   ''20221004, Add for multi blank check
    '20221117 Support multi-type program stage
    myUtility.IniDictionary StageMaskMultiTyp
    myUtility.IniDictionary MultiTypeJobList
    myUtility.IniDictionary DicMultiTypFields
    myUtility.IniDictionary DicAllTypefields
    myUtility.IniDictionary MultiTyepRVsegment
    myUtility.IniDictionary MultiTyepRVIndexWave
    myUtility.IniDictionary MultiTyepRVReadWave
    myUtility.IniDictionary tempMaskMultiTyp
    ReadPatResut = True  '20240918, Add for store read pattern result
Exit Sub
errHandler:
    Call Print_Error_Message(Error_Info, "eFuseBdfBank", "Class_Initialize")
    If AbortTest Then Exit Sub Else Resume Next
End Sub

Public Function IsEarlyFused() As Boolean
On Error GoTo errHandler
Dim fieldStr As Variant, field As eFuseBdfField

    IsEarlyFused = False
    If Not DicEarlyFused Is Nothing Then
        For Each fieldStr In DicEarlyFused.Keys
            Set field = Fields(fieldStr)
            If field.Trimmed Then
                IsEarlyFused = True
                Exit Function
            End If
        Next
    End If

Exit Function
errHandler:
    Call Print_Error_Message(Error_Info, "eFuseBdfBank", "IsEarlyFused")
    If AbortTest Then Exit Function Else Resume Next
End Function

Public Sub AddField(field As eFuseBdfField)
On Error GoTo errHandler
Dim lbankMsb As Long

    If Len(field.name) > CateNameMaxLen Then CateNameMaxLen = Len(field.name)
    
    lbankMsb = field.msb
    If field.msb < field.LSB Then
        lbankMsb = field.LSB
    End If
    If lbankMsb > FullSize Then
        FullSize = lbankMsb + 1
    End If
    ''202011XXX
    If field.LSB = field.msb Then FullSize = lbankMsb + 1
    '' check CRC
    CollectCrc field
    
    If Fields.Exists(field.name) Then
        myUtility.MessageBox "Dulpicate category in " & name & " bank"
        Call Print_Error_Message(Error_Info, "eFuseBdfBank", "AddField", "Dulpicate category in " & name & " bank")
    Else
        Fields.Add field.name, field
    End If

Exit Sub
errHandler:
    Call Print_Error_Message(Error_Info, "eFuseBdfBank", "AddField")
    If AbortTest Then Exit Sub Else Resume Next
End Sub

Public Sub AteTrimReset(currStage As String, testedStages As Dictionary, Optional CFG_Only As Boolean = False)
On Error GoTo errHandler
Dim myfield As eFuseBdfField, fieldStr As Variant
Dim realStage As String: realStage = vbNullString
Dim jobType As Variant, site As Variant
Dim Segment As DSPWave
Dim jobMask As New DSPWave

    IsBlank = 1
    SumIdsCodes = 0: SumVddbinCodes = 0
    Set DsscLotStr = New SiteVariant: Set DsscWfrStr = New SiteVariant: Set DsscXcoor = New SiteVariant: Set DsscYcoor = New SiteVariant: Set EcidBinStr = New SiteVariant
    Set DsscCFGCondStr = New SiteVariant    ''202004xx for ap
    Set DsscBKMStr = New SiteVariant    ''202004xx for ap
    Set FuseRevisionStr = New SiteVariant   ''for CFG 288bit IEDA

    If Fields.Exists(Empty) Then Fields.Remove (Empty)
    For Each fieldStr In Fields.Keys
        Set myfield = Fields(fieldStr)
        realStage = BdfDataBase.GetRealStage(myfield.BlowLocation)
        If Not CFG_Only Then myfield.FieldDefault = myfield.Default ' if CFG only, then it asigned already by previous function
        Call myfield.InitVariables

        If myfield.DefaultSetByUser Then
            myfield.DefaultOrReal = dr_default
            myfield.DefaultSetByUser = False 'clear the changes by user
        End If

        If myfield.Algorithm = alg_cond Or _
            DicEarlyFused.Exists(fieldStr) Then
            If UCase(Replace(LCase(realStage), "_early", "")) = UCase(TheExec.CurrentJob) Then
                myfield.TrimAteValue = myfield.FieldDefault
                If myfield.size < 32 Then
                    myfield.TrimAteDecValue = CLng(myUtility.Hex2Dec(myfield.FieldDefault))
                End If
                myfield.TrimFlagBySite = True
            Else
                myfield.TrimmedData = myfield.FieldDefault
                myfield.Trimmed = True
            End If
        ElseIf (DicEarlyFused.Exists(fieldStr)) Then
            If (myfield.DefaultOrReal = dr_default) Then
                myfield.TrimAteValue = myfield.FieldDefault
                If myfield.size < 32 Then
                    myfield.TrimAteDecValue = CLng(myUtility.Hex2Dec(myfield.FieldDefault))
                End If
                myfield.TrimFlagBySite = True
            End If
        ElseIf (testedStages.Exists(realStage)) Or (testedStages.Exists(Replace(LCase(realStage), "_early", ""))) Then
            myfield.TrimmedData = myfield.FieldDefault
            myfield.Trimmed = True
        ElseIf LCase(currStage) = LCase(realStage) Or bLumpStages Then
            If myfield.ReadOnly Then
                myfield.TrimAteValue = myfield.FieldDefault
                If myfield.size < 32 Then
                    myfield.TrimAteDecValue = CLng(myUtility.Hex2Dec(myfield.FieldDefault))
                End If
                myfield.TrimFlagBySite = True
            End If
        Else 'At the future stages, all limits should be "0"
            myfield.Llimit = 0
            myfield.Hlimit = 0
            myfield.BVLLimit = 0
            myfield.BVHLimit = 0
            myfield.LimitIsZero = True 'e.g. BinCut limits will get from VddBin sheets, but will set to 0 if it's at future stages
        End If
        If fieldStr = "fuse_revision" And BdfDataBase.Fuse_Revision = 0 Then
            BdfDataBase.Fuse_Revision = myUtility.String2Dbl(myfield.FieldDefault, -1)
        End If
        If fieldStr = "chip_rev_major" And BdfDataBase.Chip_rev_major = 0 Then
            BdfDataBase.Chip_rev_major = myUtility.String2Dbl(myfield.FieldDefault, -1)
        End If
        If fieldStr = "chip_rev_minor" And BdfDataBase.Chip_rev_minor = 0 Then
            BdfDataBase.Chip_rev_minor = myUtility.String2Dbl(myfield.FieldDefault, -1)
        End If
    Next fieldStr

    Set DaaCapturedWave = New DSPWave
    Set DaaCapWaveSerial = New DSPWave
    Set JtagCapturedWave = New DSPWave
    Set JtagCapturedSerial = New DSPWave
    Set SingleBitWave = New DSPWave
    Set DoubleBitWave = New DSPWave
    Set DsscSourcedWave = New DSPWave
    Set DaaCapWaveSerialFL = New DSPWave
    Set DsscWave_Eff = New DSPWave
    Set RVIndexWave = New DSPWave
    Set RVReadWave = New DSPWave

    'Set EarlyDSSCWave_Eff = New DSPWave
    If (Not defaultOthersData) = -1 Then
        For Each site In TheExec.sites
            defaultOthersData = GetDspWave(AteTrimData, dicCategory:=DicOthers).data
            If (Not myUtility.testedStages.Exists(myUtility.currStage) And DicEarlyFused.Count <> 0) Then
                DsscWave_Eff = GetDspWave(AteTrimData, dicCategory:=DicEarlyFused)
                'EarlyDSSCWave_Eff = GetDspWave(AteTrimData, dicCategory:=DicEarlyFused)
                earlySingleBitData = DsscWave_Eff.data
                'earlySingleBitData = EarlyDSSCWave_Eff.Data
                earlyData = GenDsscSrcDspWave(True, True).data
            Else
                DsscWave_Eff.CreateConstant 0, FullSize, DspLong
                'EarlyDSSCWave_Eff.CreateConstant 0, FullSize
            End If
            EarlyMaskArray = EarlyStageMask.data
            MaskArray = StageMask.data
            EarlyTestedMaskArray = EarlyTestedstagemask.data
            PostTestedMaskArray = PostTestedstagemask.data
            DefaultArray = DefaultWave.data
            For Each jobType In StageMaskMultiTyp.Keys
                tempMaskMultiTyp.Add jobType, StageMaskMultiTyp(jobType).data
            Next
            Exit For
        Next site
    End If
    If (Not FieldIdxStData) = -1 Then
        For Each site In TheExec.sites
            FieldIdxStData = FieldIdxSt.data
            FieldIdxSpData = FieldIdxSp.data
            FieldIdxMsbData = FieldIdxMsb.data
            Exit For
        Next
    End If

    RVIndexWave.CreateConstant 0, FullSize, DspLong
    RVReadWave.CreateConstant 0, UBound(segFlagArr) + 1, DspLong
    MultiTyepRVsegment.RemoveAll
    MultiTyepRVIndexWave.RemoveAll
    MultiTyepRVReadWave.RemoveAll
    If TheExec.enableWord("CFG_Partial") And name = "CFG" Then
        For Each jobType In multiTypeRVseg.Keys
            Set Segment = New DSPWave
            Segment.CreateConstant 0, 1, DspLong
            Segment.data = multiTypeRVseg(jobType)
            MultiTyepRVsegment.Add jobType, Segment
            MultiTyepRVIndexWave.Add jobType, RVIndexWave
            MultiTyepRVReadWave.Add jobType, RVReadWave
        Next jobType
    End If
    
    Set RVsegment = New DSPWave
    RVsegment.CreateConstant 0, 1, DspLong
    For Each site In TheExec.sites
       DsscWave_Eff.data = earlySingleBitData
       If TheExec.enableWord("CFG_Partial") And name = "CFG" Then
           RVsegment.data = segFlagArr
       End If
       'EarlyDSSCWave_Eff.Data = earlySingleBitData
    Next site
    
    Set defaultOthersSingleBitWave = New DSPWave
    defaultOthersSingleBitWave.data = defaultOthersData
    Set EarlyDsscSourcedWave = New DSPWave
    EarlyDsscSourcedWave.data = earlyData
    Set EarlySingleBitWave = New DSPWave
    EarlySingleBitWave.data = earlySingleBitData
     
    'mask wave
    Set EarlyStageMask = New DSPWave: EarlyStageMask.data = EarlyMaskArray
    Set StageMask = New DSPWave: StageMask.data = MaskArray
    Set EarlyTestedstagemask = New DSPWave: EarlyTestedstagemask.data = EarlyTestedMaskArray
    Set PostTestedstagemask = New DSPWave: PostTestedstagemask.data = PostTestedMaskArray
    Set DefaultWave = New DSPWave: DefaultWave.data = DefaultArray
    StageMaskMultiTyp.RemoveAll
    For Each jobType In tempMaskMultiTyp.Keys
        Set jobMask = New DSPWave
        jobMask.data = tempMaskMultiTyp(jobType)
        StageMaskMultiTyp.Add jobType, jobMask.Copy
    Next jobType
    Set FieldIdxSt = New DSPWave: FieldIdxSt.data = FieldIdxStData
    Set FieldIdxSp = New DSPWave: FieldIdxSp.data = FieldIdxSpData
    Set FieldIdxMsb = New DSPWave: FieldIdxMsb.data = FieldIdxMsbData
    TouchDownRun = True
    AnyWriteRun = False
    'PseudoFuseDspWave.CreateConstant 0, FullSize, DspLong
    Set DecimalWave = New DSPWave
    'Set DefaultWave = New DSPWave                      '20210202
    'DefaultWave.CreateConstant 0, FullSize, DspLong    '20210202
    HramLotId = vbNullString
    HramWaferId = 0
    HramXCoord = -32768
    HramYCoord = -32768
    g_Rvenable = False
    XCoord = -32768
    YCoord = -32768

Exit Sub
errHandler:
    Call Print_Error_Message(Error_Info, "eFuseBdfBank", "AteTrimReset")
    If AbortTest Then Exit Sub Else Resume Next
End Sub

Public Function GeteFuseValue(category As String) As SiteDouble
On Error GoTo errHandler
Dim result As New SiteDouble
Dim site As Variant

    If Not Fields.Exists(category) Then
        myUtility.MessageBox "[" + name + "] - Category is not exist! in EFUSE_BitDef_Table (" & category + ")"
        Exit Function
    End If
    For Each site In TheExec.sites
        result = GeteFuseValueBySite(category)
    Next
    Set GeteFuseValue = result

Exit Function
errHandler:
    Call Print_Error_Message(Error_Info, "eFuseBdfBank", "GeteFuseValue")
    If AbortTest Then Exit Function Else Resume Next
End Function

Public Function GetTrimeFuseValue(category As String) As SiteDouble
On Error GoTo errHandler
Dim result As New SiteDouble
Dim site As Variant 'Carter, 20240304

    If Not Fields.Exists(category) Then
        myUtility.MessageBox "[" + name + "] - Category is not exist! in EFUSE_BitDef_Table (" & category + ")"
        Exit Function
    End If
    For Each site In TheExec.sites
        result = GeteFuseValueBySite(category, False)
    Next
    Set GetTrimeFuseValue = result

Exit Function
errHandler:
    Call Print_Error_Message(Error_Info, "eFuseBdfBank", "GetTrimeFuseValue")
    If AbortTest Then Exit Function Else Resume Next
End Function

Public Function GeteFuseValueBySite(category As String, Optional DSSC As Boolean = True) As Double
On Error GoTo errHandler
Dim field As eFuseBdfField, result As Double

    Set field = Fields(category)
    result = 0
    If DSSC Then
        GeteFuseValueBySite = myUtility.Hex2Dbl(field.DsscValue)
    Else
        GeteFuseValueBySite = myUtility.Hex2Dbl(field.TrimAteValue)
    End If

Exit Function
errHandler:
    Call Print_Error_Message(Error_Info, "eFuseBdfBank", "GeteFuseValueBySite")
    If AbortTest Then Exit Function Else Resume Next
End Function

Public Sub UpdatedDefaultVauleToReal(itemList As String)
On Error GoTo errHandler
Dim tmp() As String, fieldStr As Variant

    tmp = Split(Replace(itemList, " ", ""), ",")
    For Each fieldStr In tmp
        DicSetDefaultVauleToReal.Add fieldStr, name
    Next

Exit Sub
errHandler:
    Call Print_Error_Message(Error_Info, "eFuseBdfBank", "UpdatedDefaultVauleToReal")
    If AbortTest Then Exit Sub Else Resume Next
End Sub

Public Sub PutDefaultVauleToReal(Optional userSpecified As Boolean = True)
On Error GoTo errHandler
Dim fieldStr As Variant, dicKeys As Dictionary

    Set dicKeys = IIf(userSpecified, DicSetDefaultVauleToReal, Fields)
    For Each fieldStr In dicKeys.Keys
        SetDefaultVauleToRealField CStr(fieldStr)
    Next

Exit Sub
errHandler:
    Call Print_Error_Message(Error_Info, "eFuseBdfBank", "PutDefaultVauleToReal")
    If AbortTest Then Exit Sub Else Resume Next
End Sub

Public Sub SetDefaultVauleToRealField(fieldStr As String)
On Error GoTo errHandler
Dim pf As New SiteBoolean: pf = True
Dim values As New SiteVariant: values = Fields(fieldStr).Default

    If Fields.Exists(fieldStr) Then
        If Fields(fieldStr).size < 32 Then values = myUtility.Hex2Dec(values)
        If Fields(fieldStr).ReadOnly Or Fields(fieldStr).Algorithm = alg_ids Then Exit Sub
        Call Fields(fieldStr).IsOverflow(values, True)
        Call Print_Error_Message(Warning_Info, "eFuseBdfBank", "SetDefaultVauleToRealField", "User use field DEFAULT value to REAL field! Bank = " & name & ", Field = " & fieldStr)
        If LCase(BdfDataBase.GetRealStage(Fields(fieldStr).BlowLocation)) = LCase(myUtility.currStage) Or bLumpStages Then
            SetEfuse fieldStr, values, pf, SetByDefault:=True
            Fields(fieldStr).TrimFlagBySite = True
        End If
    Else
        myUtility.MessageBox "[" + name + "] - [Error] Bank didn't contain the field " & fieldStr
    End If

Exit Sub
errHandler:
    Call Print_Error_Message(Error_Info, "eFuseBdfBank", "SetDefaultVauleToRealField")
    If AbortTest Then Exit Sub Else Resume Next
End Sub

'20220106, Modify for Real value validation( reverse bit)
Public Sub SetEfuse(category As String, ByVal invalue As SiteVariant, _
                    Optional ByVal PatPF As SiteBoolean = True, _
                    Optional forceWrite As Boolean = False, _
                    Optional skipStageChk As Boolean = False, _
                    Optional SetByDefault As Boolean = False, _
                    Optional PatFlagPrintOut As Boolean = False, _
                    Optional Dict_Store_Code_Name As String = vbNullString, _
                    Optional re_value As SiteVariant = 0)
On Error GoTo errHandler
Dim binstr As New SiteVariant
Dim bintmpstr As String, site As Variant
Dim field As eFuseBdfField, value As New SiteVariant
Dim hvalue As SiteVariant
    
    If (PatPF Is Nothing) Then Set PatPF = New SiteBoolean: PatPF = True
    If Not IsFieldExist(category) Then Exit Sub

    Set field = Fields(category)
    value = invalue
    ''202004xx for ap
    If LCase(BdfDataBase.GetRealStage(field.BlowLocation)) <> LCase(myUtility.currStage) And Not bLumpStages Then
        If Not skipStageChk Then
            Call Print_Error_Message(Error_Info, "eFuseBdfBank", "SetEfuse", "[" + name + "] - Category can't be set, different stages! in EFUSE_BitDef_Table (" & category + ")>>Only allow at " & field.BlowLocation)
            If EFUSE_SETFUSE_SHOW_ERRMSX = True Then
                myUtility.MessageBox "[" + name + "] - Category can't be set, different stages! in EFUSE_BitDef_Table (" & category + ")>>Only allow at " & field.BlowLocation
            End If
        End If
    Exit Sub
    End If
        
    If field.Algorithm = alg_ids Then
        value = IDS_Conversion(field, value)
    End If

    Call field.IsOverflow(value, bMsgBox:=False) ''20201112 add for obj

    Set hvalue = myUtility.oDec2HexStr(value, field.hBytes)
    If field.Trimmed Then
        Call Print_Error_Message(Error_Info, "eFuseBdfBank", "SetEfuse", "[" + name + "] - [Ignore Set eFuse]  " & category & " ,Value had already been Trimmed in previous stage! >> " & field.BlowLocation)
        If EFUSE_SETFUSE_SHOW_ERRMSX = True Then
            myUtility.MessageBox "Site(" & site & ") [" + name + "] - [Ignore Set eFuse]  " & category & " ,Value had already been Trimmed in previous stage! >> " & field.BlowLocation
        End If
        Exit Sub
    End If

    If Not field.ReadOnly Then
        For Each site In TheExec.sites
            If Not field.TrimFlagBySite(site) Then
                field.PatTestPass(site) = PatPF(site)
                ''202004xx for ap
                field.TrimAteDecValue(site) = value(site)
                If myUtility.xIsHex(value(site)) And field.size < 32 Then
                    bintmpstr = myUtility.xHex2BinStr(value(site), field.size)
                    field.TrimAteDecValue(site) = myUtility.Bin2Dec(bintmpstr)
                End If
                
                field.TrimAteValue(site) = hvalue(site)
                
                If field.Algorithm = alg_ids Then
                    myUtility.WriteDlg "Site(" & site & ") [" + name + "] - [Set eFuse]  " & category & " = " & field.TrimAteDecValue & " ( " & Format(invalue * 1000, "####0.000000") & " mA / " & field.Resolution & " mA ) at row#" & field.DataRowNum & " in sheet " & sheetName
                Else
                    '20220106, Add for Real value validation( reverse bit)
                    If Dict_Store_Code_Name <> "" Then ''Efuse_DicValue_Chk ==>
                        If re_value <> 0 Then
                            myUtility.WriteDlg "Site(" & site & ") [" + name + "] - [Set eFuse]  " & category & " = " & field.TrimAteDecValue & " at row#" & field.DataRowNum & " in sheet " & sheetName & " @ " & Dict_Store_Code_Name & " ReverseValue Bit : " & CStr(re_value)
                        Else
                            myUtility.WriteDlg "Site(" & site & ") [" + name + "] - [Set eFuse]  " & category & " = " & field.TrimAteDecValue & " at row#" & field.DataRowNum & " in sheet " & sheetName & " @ " & Dict_Store_Code_Name
                        End If
                    Else
                        If UCase(field.bankName) Like "CFG" And field.Algorithm = alg_lotid Then
                            myUtility.WriteDlg "Site(" & site & ") [" + name + "] - [Set eFuse]  " & category & " = " _
                            & myUtility.Bin2HexStr(StrReverse(myUtility.Hex2BinStr(field.TrimAteDecValue, field.size)), (field.size - 1) \ 4 + 1) & " (Reverse) at row#" & field.DataRowNum & " in sheet " & sheetName
                        ElseIf UCase(field.bankName) Like "CFG" And field.Algorithm = alg_numeric Then
                            myUtility.WriteDlg "Site(" & site & ") [" + name + "] - [Set eFuse]  " & category & " = " _
                            & myUtility.Bin2Dec(StrReverse(myUtility.Dec2Bin(field.TrimAteDecValue, field.size)), (field.size - 1) \ 4 + 1) & " (Reverse) at row#" & field.DataRowNum & " in sheet " & sheetName
                        Else
                            myUtility.WriteDlg "Site(" & site & ") [" + name + "] - [Set eFuse]  " & category & " = " & field.TrimAteDecValue & " at row#" & field.DataRowNum & " in sheet " & sheetName
                        End If
                    End If
                    '----------------------------------------------------------
                    'myUtility.WriteDlg "Site(" & site & ") [" + name + "] - [Set eFuse]  " & category & " = " & field.TrimAteDecValue & " at row#" & field.DataRowNum & " in sheet " & sheetName
                End If
'               If Not FieldDsscIs0(field) Then
'               myUtility.WriteDlg "Site(" & site & ") [" + name + "] - [Warning] [Set eFuse]  " & category & " ,it had DSSC result, field is not blank!"
'               End If
                field.TrimFlagBySite(site) = True
            ElseIf forceWrite Or field.RealSetByDefault Or SetByDefault Then
                ''202004xx for ap
                field.PatTestPass(site) = PatPF(site)
                field.TrimAteDecValue(site) = value(site)
                If myUtility.xIsHex(value(site)) And field.size < 32 Then
                    bintmpstr = myUtility.xHex2BinStr(value(site), field.size)
                    field.TrimAteDecValue(site) = myUtility.Bin2Dec(bintmpstr)
                End If
                field.TrimAteValue(site) = hvalue(site)
                myUtility.WriteDlg "Site(" & site & ") [" + name + "] - [Overwrite eFuse]  " & category & " , Value had been set! Set value again with " & field.TrimAteDecValue(site) & " at row#" & field.DataRowNum & " in sheet " & sheetName
                field.TrimFlagBySite(site) = True
            Else
                Call Print_Error_Message(Error_Info, "eFuseBdfBank", "SetEfuse", "Site(" & site & ") [" + name + "] - [Error, Ignore Set eFuse]  " & category & " ,Value had been set before")
                If EFUSE_SETFUSE_SHOW_ERRMSX = True Then
                    myUtility.MessageBox "Site(" & site & ") [" + name + "] - [Error, Ignore Set eFuse]  " & category & " ,Value had been set before, Error!"
                End If
                TheExec.Flow.TestLimit resultVal:=0, lowVal:=1, hiVal:=1, Tname:="SetEfuse Error"
            End If
        Next site
    Else
        If forceWrite Then
            For Each site In TheExec.sites
                field.PatTestPass(site) = PatPF(site)
                field.TrimAteValue(site) = hvalue(site)
            
                Call Print_Error_Message(Warning_Info, "eFuseBdfBank", "Site(" & site & ") [" + name + "] - [Set eFuse] User force overwrite  " & category & " with value " & field.TrimAteDecValue(site) & " at row#" & field.DataRowNum & " in sheet " & sheetName)
            Next
            field.DefaultSetByUser = True
            field.DefaultOrReal = dr_real
        Else
            Call Print_Error_Message(Error_Info, "eFuseBdfBank", "SetEfuse", "[" + name + "] - This is a ReadOnly category!  " & category & " at row#" & field.DataRowNum & " in sheet " & sheetName)
            If EFUSE_SETFUSE_SHOW_ERRMSX = True Then
                myUtility.MessageBox "[" + name + "] - [Error] This is a ReadOnly category!  " & category & " at row#" & field.DataRowNum & " in sheet " & sheetName
            End If
            TheExec.Flow.TestLimit resultVal:=0, lowVal:=1, hiVal:=1, Tname:="SetEfuse Error"
        End If
    End If
    
    field.TrimTestInstance = TheExec.DataManager.instancename
    field.RealSetByDefault = SetByDefault
        
    If PatFlagPrintOut Then
        For Each site In TheExec.sites
            myUtility.WriteDlg "Site(" & site & ") [" + name + "] - [Set PatTestPass_Flag]  " & category & " = " & PatPF
        Next
    End If

Exit Sub
errHandler:
    Call Print_Error_Message(Error_Info, "eFuseBdfBank", "SetEfuse")
    If AbortTest Then Exit Sub Else Resume Next
End Sub

Private Function FieldDsscIs0(field As eFuseBdfField) As Boolean
On Error GoTo errHandler
    FieldDsscIs0 = IIf(Hex(Replace(field.DsscValue, "0x", "&H")) = 0, True, False)
Exit Function
errHandler:
    Call Print_Error_Message(Error_Info, "eFuseBdfBank", "FieldDsscIs0")
    If AbortTest Then Exit Function Else Resume Next
End Function

Private Function IDS_Conversion(field As eFuseBdfField, ByVal values As SiteVariant) As SiteVariant
On Error GoTo errHandler
Dim funcName As String:: funcName = "IDS_Conversion"
Dim m_dbl As New SiteDouble
Dim m_dbl_round As New SiteDouble
Dim site As Variant 'Carter, 20240304

    values = values.Multiply(1000) ' A to mA, to do calculation
    For Each site In TheExec.sites.Selected
        m_dbl = (values / field.Resolution)
        m_dbl_round = Round(m_dbl)
        If ((CDec(m_dbl_round) - CDec(m_dbl)) >= 0) Then ''''MUST have
            m_dbl_round = m_dbl_round
        Else
            m_dbl_round = m_dbl_round + 1
        End If
        field.MeasureValue = m_dbl_round * field.Resolution '/ 1000 'mA
    Next
    Set IDS_Conversion = m_dbl_round
Exit Function
errHandler:
    Call Print_Error_Message(Error_Info, "eFuseBdfBank", "IDS_Conversion")
    If AbortTest Then Exit Function Else Resume Next
End Function

Public Function CalculateCRC(Optional specField As String = vbNullString, Optional specDspWave As DSPWave = Nothing, Optional getValOnly As Boolean = False) As String
On Error GoTo errHandler
Dim CrcBitsStr As String:
Dim item As Variant, crc As Dictionary, field As eFuseBdfField, bitem As Variant, bit As Long
Dim mywave As New DSPWave, CRCarray() As Byte, CrcBits As Long
Dim crcStr As String, site As Variant
    If Crcs.Count = 0 Then
        CalculateCRC = vbNullString
        If HadCrcFuse Then myUtility.MessageBox "CRC field didn't put the calculate bits info. in this bank!" & vbCrLf & "e.g. CRC_calcBits=[63:0],REAL_CODE_FROM_TEST @Bank=""" + name + """": CalculateCRC = "": Exit Function
        myUtility.WriteDlg "No CRC field in this bank! Bank=""" + name + """": CalculateCRC = "": Exit Function
    End If

    CalculateCRC = vbNullString
    For Each item In Crcs.Keys  ' field name => dictionary (bit, true)
        If specField = "" Or LCase(item) = LCase(specField) Then
            Set field = Fields(item)
            If ((LCase(myUtility.currStage) = LCase(BdfDataBase.GetRealStage(field.BlowLocation)) And Not field.Trimmed) Or bLumpStages) Or getValOnly Then
                If mywave.sampleSize = 0 Then
                    If specDspWave Is Nothing Then
                        Set mywave = GetDspWave(AteTrimData, forceGet:=True, calCrc:=True) 'Set true means it needs a MSB first Dspwave
                    Else
                        Set mywave = specDspWave
                    End If
                End If
                                        
                If field.DefaultOrReal = dr_real Then
                    CrcBits = field.size
                    ReDim CRCarray(CrcBits - 1)
                    Set crc = Crcs(item)
                    For Each site In TheExec.sites
                        CrcBitsStr = vbNullString
                        Call myUtility.CRC_0Array(CRCarray) '<< Should do reset in siteloop
                        For Each bitem In crc.Keys
                            Select Case CrcBits
                            Case 16:
                                Call myUtility.CalculateCRC16(CRCarray, CByte(mywave.Element(bitem)))
                            Case 32:
                                Call myUtility.CalculateCRC32(CRCarray, CByte(mywave.Element(bitem)))
                            End Select
                            If eFusePrinted Then CrcBitsStr = CrcBitsStr & mywave.Element(bitem)
                        Next

                        crcStr = vbNullString
                        For bit = 0 To UBound(CRCarray)
                            crcStr = CRCarray(bit) & crcStr     '[MSB......LSB]
                        Next
            
                        crcStr = myUtility.Bin2HexStr(crcStr, field.hBytes)
                        If Replace(crcStr, "0", "") = "x" Then 'crcStr = "0xC5EDF340" Or
                            Call Print_Error_Message(Error_Info, "eFuseBdfBank", "CalculateCRC", "The efuse fields are 0, CRC calculation is ignored! >>" & field.name)
                        Else
                            CalculateCRC = crcStr
                            If Not getValOnly Then
                                 field.TrimAteValue = crcStr
                                 field.TrimFlagBySite = True
                                 field.TrimTestInstance = TheExec.DataManager.instancename
                                 field.RealSetByDefault = False
                            End If
                        End If
                        If eFusePrinted Then
                            myUtility.WriteDlg "Site" & site & " bnak(" & item & ") CRC=" & crcStr & " ,Input Bits = " & CrcBitsStr
                        End If
                    Next
                Else
                    Call Print_Error_Message(Warning_Info, "eFuseBdfBank", "CalculateCRC", "bnak(" & item & ") skip set, it's CRC is not REAL!")
                End If
            Else
                Call Print_Error_Message(Warning_Info, "eFuseBdfBank", "bnak(" & item & ") skip set, it's CRC  at stage >> " & field.BlowLocation)
            End If
        End If
        Set mywave = Nothing
    Next
Exit Function
errHandler:
    Call Print_Error_Message(Error_Info, "eFuseBdfBank", "CalculateCRC")
    If AbortTest Then Exit Function Else Resume Next
End Function

Private Function GetSizeByFields(ByRef dicCategory As Dictionary) As Long
On Error GoTo errHandler
Dim fieldStr As Variant, field As eFuseBdfField, size As Long

    size = 0
    For Each fieldStr In dicCategory.Keys
        Set field = Fields(fieldStr)
        size = size + field.size
    Next
    GetSizeByFields = size
Exit Function
errHandler:
    Call Print_Error_Message(Error_Info, "eFuseBdfBank", "GetSizeByFields")
    If AbortTest Then Exit Function Else Resume Next
End Function

Public Function GetDsscWave(ByRef dicCategory As Dictionary, Optional useOthersDefault As Boolean = True, Optional RvOnly As Boolean, Optional ByRef dicRowNum As DSPWave, Optional MultiType As String = vbNullString) As DSPWave
On Error GoTo errHandler
Dim eFuseWave As New DSPWave, fieldStr As Variant, field As eFuseBdfField, iBit As Long
Dim stBit As Long, spBit As Long, bitStep As Long, binA() As New SiteLong, bitDealt As Long
Dim rvWaveIdx As New DSPWave
Dim site As Variant
Dim oWave As New DSPWave

    If useOthersDefault And Not RvOnly And DicForceFields.Count = 0 And MultiType = "" Then
        eFuseWave = defaultOthersSingleBitWave.Copy
    Else
        eFuseWave.CreateConstant 0, FullSize
        If RvOnly Then
            rvWaveIdx.CreateConstant 0, FullSize
        End If
    End If
    
    For Each fieldStr In dicCategory.Keys
        Set field = Fields(fieldStr)
                If (field.DefaultOrReal = dr_real And useOthersDefault) Or (field.Algorithm = alg_cond) Or Not useOthersDefault Then
                            bitDealt = field.size - 1
                            stBit = field.LSB: spBit = field.msb: bitStep = 1: ' MSB 1st
                            If field.Msb1st Then
                                 GetBitStSp field, stBit, spBit
                            ElseIf stBit > spBit Then
                                 bitStep = -1: 'LSB first, if doing CRC, it needs to use the MSB first
                            Else
                                 myUtility.MessageBox "Something wrong in dealing BDF GetDspWave"
                                 Call Print_Error_Message(Error_Info, "eFuseBdfBank", "GetDsscWave", "Something wrong in dealing BDF GetDspWave")
                            End If
                            
                            bitDealt = 0
                            binA = GetBinA(DataRetrieve.AteTrimData, field)
                            For iBit = stBit To spBit Step bitStep
                                        For Each site In TheExec.sites.Selected
                                                If RvOnly Then
                                                rvWaveIdx.ElementLite(iBit) = 1
                                                End If
                                                eFuseWave.ElementLite(iBit) = binA(bitDealt)
                                        Next
                                        bitDealt = bitDealt + 1
                            Next
                End If
    Next

    If RvOnly And name = "CFG" Then
        If Len(MultiType) <> 0 And MultiTyepRVsegment.Exists(MultiType) Then
            Ze_CombineShortWave eFuseWave, rvWaveIdx, MultiTyepRVsegment(MultiType), oWave, dicRowNum, DoubleBits
            Set MultiTyepRVIndexWave(MultiType) = rvWaveIdx
        Else
            Ze_CombineShortWave eFuseWave, rvWaveIdx, RVsegment, oWave, dicRowNum, DoubleBits
            Set RVIndexWave = rvWaveIdx
        End If
        Set eFuseWave = oWave
    End If

    For Each site In TheExec.sites
         SingleBitWave = eFuseWave
    Next
    Set GetDsscWave = eFuseWave
Exit Function
errHandler:
    Call Print_Error_Message(Error_Info, "eFuseBdfBank", "GetDsscWave")
    If AbortTest Then Exit Function Else Resume Next
End Function

Public Function GetOfflineDspWave(v As DataRetrieve, Optional ByVal stage As Boolean = True, Optional ByRef dicCategory As Dictionary, _
Optional forceGet As Boolean = False, Optional includeEarlyFuse As Boolean = False, Optional calCrc As Boolean = False, Optional earlyfuse As Boolean = False) As DSPWave
On Error GoTo errHandler
    Dim eFuseWave As New DSPWave, fieldStr As Variant, field As eFuseBdfField, iBit As Long
    Dim stBit As Long, spBit As Long, bitStep As Long, binA() As New SiteLong, bitDealt As Long
    Dim tmpFieldStage As String
    Dim site As Variant
    If v = TrimmedData Then stage = False
    If forceGet Then
        eFuseWave = DefaultWave.Copy
    Else
        eFuseWave.CreateConstant 0, FullSize
    End If
    If Not myUtility.OnlineMode Then eFuseWave.CreateConstant 0, FullSize
    Dim getLump As Boolean: getLump = False
    If Not CfgDataBaseAct Is Nothing Then
        If Not earlyfuse And myUtility.testedStages.Count = 0 And CfgDataBaseAct.EarlyStage Then
            getLump = True
        End If
    End If
    For Each fieldStr In Fields.Keys
        Set field = Fields(fieldStr)
        tmpFieldStage = Replace(UCase(BdfDataBase.GetRealStage(field.BlowLocation)), "_EARLY", "")
        If (myUtility.testedStages.Exists(tmpFieldStage)) Then
            If (Not field.Trimmed And v = AteTrimData) Or _
            (Not field.Trimmed And v = DafaultValue) Or _
            v = TrimmedData Or v = DsscValue Or (bLumpStages And myUtility.testedStages.Count > 0) Or _
            (forceGet And field.DefaultOrReal = dr_real) Or _
            (forceGet And field.DefaultOrReal = dr_default And field.Algorithm = alg_cond) Or _
            (Not myUtility.OnlineMode And field.Trimmed = True) Then
                If GetEarlyFuseOnly And Not DicEarlyFused.Exists(fieldStr) Then GoTo skip
                If Not dicCategory Is Nothing Then
                    If Not dicCategory.Exists(fieldStr) Then GoTo skip
                End If
                bitDealt = field.size - 1
                stBit = field.LSB: spBit = field.msb: bitStep = 1: ' MSB 1st
                If field.Msb1st Then
                    GetBitStSp field, stBit, spBit
                ElseIf stBit > spBit Then
                    bitStep = -1: 'LSB first, if doing CRC, it needs to use the MSB first
                Else
                    myUtility.MessageBox "Something wrong in dealing BDF GetDspWave"
                End If
                bitDealt = 0
                binA = GetBinA(v, field, calCrc:=calCrc, getLump:=Not earlyfuse)
                For iBit = stBit To spBit Step bitStep
                    For Each site In TheExec.sites.Selected
                        eFuseWave.ElementLite(iBit) = binA(bitDealt)
                    Next
                    bitDealt = bitDealt + 1
                Next
skip:
            ElseIf Not dicCategory Is Nothing Then
                If dicCategory.Exists(fieldStr) And v = AteTrimData And Not DicCondTables.Exists(fieldStr) Then   'fieldstr <> CfgEarlyTable Then
                    myUtility.MessageBox "[" & fieldStr & "] The category had been fused already! Please confirmed!"
                End If
            Else
            End If 'field.Trimmed End
        End If
    Next
    If Not forceGet And (stage Or Not dicCategory Is Nothing) Then
        Set eFuseWave = DspWaveByConditions(eFuseWave, stage, dicCategory, includeEarlyFuse)
    End If
    For Each site In TheExec.sites
        SingleBitWave = eFuseWave
    Next
    Set GetOfflineDspWave = eFuseWave
Exit Function
errHandler:
    Call Print_Error_Message(Error_Info, "eFuseBdfBank", "GetDspWave")
        If AbortTest Then Exit Function Else Resume Next
End Function

Public Function GetDspWave(v As DataRetrieve, Optional ByVal stage As Boolean = True, Optional ByRef dicCategory As Dictionary, _
Optional forceGet As Boolean = False, Optional includeEarlyFuse As Boolean = False, Optional calCrc As Boolean = False, Optional earlyfuse As Boolean = False) As DSPWave
On Error GoTo errHandler
Dim eFuseWave As New DSPWave, fieldStr As Variant, field As eFuseBdfField, iBit As Long
Dim stBit As Long, spBit As Long, bitStep As Long, binA() As New SiteLong, bitDealt As Long
Dim getLump As Boolean: getLump = False
Dim tmp As New DSPWave
Dim site As Variant
    
    If v = TrimmedData Then stage = False
    
    If forceGet Then
        eFuseWave = DefaultWave.Copy
    Else
        eFuseWave.CreateConstant 0, FullSize
    End If
    
    If Not myUtility.OnlineMode Then eFuseWave.CreateConstant 0, FullSize

    '20200120
    If Not CfgDataBaseAct Is Nothing Then
        If Not earlyfuse And myUtility.testedStages.Count = 0 And CfgDataBaseAct.EarlyStage Then
            getLump = True
        End If
    End If
    '-----------------
    '1. Update the value
    '-----------------
    For Each fieldStr In Fields.Keys
        Set field = Fields(fieldStr)
        If (Not field.Trimmed And v = AteTrimData) Or _
         (Not field.Trimmed And v = DafaultValue) Or _
         v = TrimmedData Or v = DsscValue Or (bLumpStages And myUtility.testedStages.Count > 0) Or _
         (forceGet And field.DefaultOrReal = dr_real) Or _
         (forceGet And field.DefaultOrReal = dr_default And field.Algorithm = alg_cond) Or _
         (Not myUtility.OnlineMode And field.Trimmed = True) Then
         '(Not GlbUtility.IsOnline And field.Trimmed = True And field.Algorithm <> alg_cond) Then
         'v = TrimmedData Or v = DsscValue Or forceGet Or (bLumpStages And myUtility.testedStages.Count > 0) Then
            If GetEarlyFuseOnly And Not DicEarlyFused.Exists(fieldStr) Then GoTo skip
            If Not dicCategory Is Nothing Then
                If Not dicCategory.Exists(fieldStr) Then GoTo skip
            End If
            bitDealt = field.size - 1
            stBit = field.LSB: spBit = field.msb: bitStep = 1: ' MSB 1st
            If field.Msb1st Then
                GetBitStSp field, stBit, spBit
            ElseIf stBit > spBit Then
                bitStep = -1: 'LSB first, if doing CRC, it needs to use the MSB first
            Else
                myUtility.MessageBox "Something wrong in dealing BDF GetDspWave"
                Call Print_Error_Message(Error_Info, "eFuseBdfBank", "GetDspWave", "Something wrong in dealing BDF GetDspWave")
            End If

            bitDealt = 0
            binA = GetBinA(v, field, calCrc:=calCrc, getLump:=Not earlyfuse)
            For iBit = stBit To spBit Step bitStep
                For Each site In TheExec.sites.Selected
                    eFuseWave.ElementLite(iBit) = binA(bitDealt)
                Next
                bitDealt = bitDealt + 1
            Next
skip:
        ElseIf Not dicCategory Is Nothing Then
            If dicCategory.Exists(fieldStr) And v = AteTrimData And Not DicCondTables.Exists(fieldStr) Then   'fieldstr <> CfgEarlyTable Then
                myUtility.MessageBox "[" & fieldStr & "] The category had been fused already! Please confirmed!"
                Call Print_Error_Message(Error_Info, "eFuseBdfBank", "GetDspWave", "[" & fieldStr & "] The category had been fused already! Please confirmed!")
            End If
        Else
            'myUtility.MessageBox "Something wrong in GetDspWave!"
        End If 'field.Trimmed End
    Next
    
    '----------------------------------
    '2. Use stage mask to filter the wave
    '----------------------------------
   If Not forceGet And (stage Or Not dicCategory Is Nothing) Then
       Set eFuseWave = DspWaveByConditions(eFuseWave, stage, dicCategory, includeEarlyFuse)
    End If
    
    For Each site In TheExec.sites
        tmp = SingleBitWave.ConvertDataTypeTo(DspLong).LogicalCompare(NotEqualTo, eFuseWave.ConvertDataTypeTo(DspLong))
        SingleBitWave = eFuseWave
    Next
    Set GetDspWave = eFuseWave
Exit Function
errHandler:
    Call Print_Error_Message(Error_Info, "eFuseBdfBank", "GetDspWave")
    If AbortTest Then Exit Function Else Resume Next
End Function

Public Function GetBinA(v As DataRetrieve, field As eFuseBdfField, Optional ByRef SetTrimmed As Dictionary = Nothing, Optional calCrc As Boolean = False, Optional getLump As Boolean = False) As SiteLong()
On Error GoTo errHandler
Dim dealValue As New SiteVariant
               Select Case v ' binA is LSB first already' Dec9 5bits => result = binA[0..4] = [10010], lSB=>[0] MSB=>[nbit-1]
                    Case DataRetrieve.TrimmedData:
                                Set dealValue = field.TrimmedData
                    Case DataRetrieve.AteTrimData:
                                If field.Algorithm = alg_cond And (calCrc Or getLump) Then
                                      dealValue = CfgCmpValueLumpStages(field.name)
                                ElseIf field.Trimmed Then 'dptx_lpdp0_pll_fcal
                                      Set dealValue = field.TrimmedData
                                Else
                                      Set dealValue = field.TrimAteValue
                                End If
'                                If field.Trimmed Then 'dptx_lpdp0_pll_fcal
                                If Not SetTrimmed Is Nothing Then SetTrimmed.Add field.name, True
                    Case DataRetrieve.DafaultValue:
                                Set dealValue = field.FieldDefault
                    Case DataRetrieve.DsscValue:
                                Set dealValue = field.DsscValue
                End Select
                GetBinA = myUtility.Hex2BinArray(dealValue, field.size)
Exit Function
errHandler:
    Call Print_Error_Message(Error_Info, "eFuseBdfBank", "GetBinA")
    If AbortTest Then Exit Function Else Resume Next
End Function
 
Public Function GenDsscSrcDspWave(Optional dealDoubleBit As Boolean = False, Optional dealVecRpt As Boolean = False, Optional rptVerSpec As Long = -1) As DSPWave
On Error GoTo errHandler
Dim eFuseWave As New DSPWave, rptVer As Long
Dim idxWave As New DSPWave, waveSize As Long
Dim site As Variant 'Carter, 20240304

    rptVer = IIf(rptVerSpec = -1, PatRepeatCnt, rptVerSpec)
    If pgmMode = pgm_JTAG Then rptVer = 1

    Set eFuseWave = SingleBitWave
    If DoubleBits And dealDoubleBit Then
        Set DoubleBitWave = New DSPWave
        Ze_XExpandDblBit SingleBitWave, DoubleBitWave
        Set eFuseWave = DoubleBitWave
    End If
    If JtagMsbFirst Then
        waveSize = GlbUtility.GetSampleSize(eFuseWave)
        idxWave.CreateRamp waveSize - 1, -1, waveSize, DspLong
        For Each site In TheExec.sites
            eFuseWave = eFuseWave.Lookup(idxWave)
        Next
    End If
    'myUtility.PersonalTimerLog Name & ">> XExpandDblBit", True
    If dealVecRpt And rptVer > 1 Then
        Set DsscSourcedWave = New DSPWave
        Ze_XExpandBitVector eFuseWave, rptVer, DsscSourcedWave
        Set eFuseWave = DsscSourcedWave
    End If
    'myUtility.PersonalTimerLog Name & ">> XExpandBitVector", True
    Set GenDsscSrcDspWave = eFuseWave

Exit Function
errHandler:
    Call Print_Error_Message(Error_Info, "eFuseBdfBank", "GenDsscSrcDspWave")
    If AbortTest Then Exit Function Else Resume Next
End Function

Public Function DspWaveByConditions(inwave As DSPWave, Optional ByVal stage As Boolean = True, _
                                                                       Optional ByVal dicCategory As Dictionary, Optional includeEarlyFuse As Boolean = False, Optional chkNonEarly As Boolean = False) As DSPWave
On Error GoTo errHandler
Dim dicCared As New DSPWave
Dim stBit As Long, spBit As Long, iBit As Long
Dim fieldStr As Variant, field As eFuseBdfField
Dim dicCatCheck As Dictionary
Dim outwave As New DSPWave

    If inwave Is Nothing Then
        myUtility.MessageBox "The input DspWave is empty! Please check it."
        Call Print_Error_Message(Error_Info, "eFuseBdfBank", "DspWaveByConditions", "The input DspWave is empty! Please check it.")
    Else
        If Not dicCategory Is Nothing Then
            Set dicCatCheck = New Dictionary
            For Each fieldStr In dicCategory.Keys
                dicCatCheck.Add fieldStr, False
            Next
        End If
        dicCared.CreateConstant 0, FullSize
        For Each fieldStr In Fields.Keys
            Set field = Fields(fieldStr)
            GetBitStSp field, stBit, spBit
            
            If (stage And (UCase(myUtility.currStage) = UCase(BdfDataBase.GetRealStage(field.BlowLocation)) Or bLumpStages)) Or (includeEarlyFuse And DicEarlyFused.Exists(fieldStr)) Then
                If dicCategory Is Nothing Then
                    For iBit = stBit To spBit
                        dicCared.Element(iBit) = 1
                    Next
                ElseIf dicCategory.Exists(fieldStr) Then
                    If dicCategory(fieldStr) = name Then
                        dicCatCheck(fieldStr) = True
                        For iBit = stBit To spBit
                            dicCared.Element(iBit) = 1
                        Next
                    End If
                End If
            ElseIf Not dicCategory Is Nothing Then
                If dicCategory.Exists(fieldStr) Then
                    If dicCategory(fieldStr) = name Then
                        dicCatCheck(fieldStr) = True
                        For iBit = stBit To spBit
                            dicCared.Element(iBit) = 1
                        Next
                    End If
                End If
            Else
                myUtility.MessageBox "Conditons is ambigious... Please confirm!"
                Call Print_Error_Message(Error_Info, "eFuseBdfBank", "DspWaveByConditions", "Conditons is ambigious... Please confirm!")
                GoTo errHandler
            End If
        Next
        'inWave(0).Plot
        If Not dicCatCheck Is Nothing Then
            For Each fieldStr In dicCategory.Keys
                If Not dicCatCheck.Exists(fieldStr) And dicCategory(fieldStr) = name And fieldStr <> "" Then
                    myUtility.MessageBox "Category specified didn't found! " & fieldStr
                    Call Print_Error_Message(Error_Info, "eFuseBdfBank", "DspWaveByConditions", "Category specified didn't found! " & fieldStr)
                End If
            Next
        End If
        
        Ze_GetDspWaveExMask inwave, outwave, dicCared
        Set DspWaveByConditions = outwave
    End If

Exit Function
errHandler:
    Call Print_Error_Message(Error_Info, "eFuseBdfBank", "DspWaveByConditions")
    If AbortTest Then Exit Function Else Resume Next
End Function

Public Function UpdateTrimmed(dicTrimmed As Dictionary, Optional chkNonEarly As Boolean = False)
On Error GoTo errHandler
Dim fieldStr As Variant, field As eFuseBdfField

    For Each fieldStr In dicTrimmed.Keys
        Set field = Fields(fieldStr)
        If field.Algorithm = alg_cond Then
            If chkNonEarly And bLumpStages Then
                field.TrimmedData = CfgCmpValueLumpStages(field.name)
            Else
                field.TrimmedData = CfgCmpValue(field.name)
            End If
        Else
            field.TrimmedData = field.TrimAteValue
        End If
        field.Trimmed = True
    Next

Exit Function
errHandler:
    Call Print_Error_Message(Error_Info, "eFuseBdfBank", "UpdateTrimmed")
    If AbortTest Then Exit Function Else Resume Next
End Function

Public Function GetTrimBitCount(dicTrimmed As Dictionary) As Long
On Error GoTo errHandler
Dim fieldStr As Variant, field As eFuseBdfField, Count As Long: Count = 0

    For Each fieldStr In dicTrimmed.Keys
        Set field = Fields(fieldStr)
        Count = Count + field.size
    Next
    GetTrimBitCount = Count

Exit Function
errHandler:
    Call Print_Error_Message(Error_Info, "eFuseBdfBank", "GetTrimBitCount")
    If AbortTest Then Exit Function Else Resume Next
End Function

Public Function TrimedMissingCnt(ByVal dicCategory As Dictionary, Optional TrimMissingFilter_Flag As Boolean = False) As SiteLong
On Error GoTo errHandler
Dim fieldStr As Variant, field As eFuseBdfField
Dim m_TrimedMissingCnt As New SiteLong
Dim site As Variant

    For Each fieldStr In dicCategory.Keys
        Set field = Fields(fieldStr)
        If (Not field.Trimmed) And (UCase(BdfDataBase.GetRealStage(field.BlowLocation)) = UCase(myUtility.currStage) Or bLumpStages) Then
            If Not (TrimMissingFilter_Flag And ((LCase(field.name) Like "efuse_database_revision*") Or (field.Algorithm = alg_crc))) Then
                For Each site In TheExec.sites.Active
                    If Not field.TrimFlagBySite Then
                        If eFusePrinted Then Call Print_Error_Message(Error_Info, "eFuseBdfBank", "TrimedMissingCnt", "[Trim missing] Didn't set trim value for field >> [" & field.bankName & "]" & field.name)
                        m_TrimedMissingCnt = m_TrimedMissingCnt + 1
                        If EFUSE_BIN1_SETTING Then
                            field.TrimFlagBySite = True
                            field.TrimAteValue = "0x0"
                            field.TrimAteDecValue = 0
                            field.RealSetByDefault = False
                            m_TrimedMissingCnt = m_TrimedMissingCnt - 1
                        End If
                    End If
                Next site
            End If
        End If
    Next
    Set TrimedMissingCnt = m_TrimedMissingCnt

Exit Function
errHandler:
    Call Print_Error_Message(Error_Info, "eFuseBdfBank", "TrimedMissingCnt")
    If AbortTest Then Exit Function Else Resume Next
End Function

Public Sub CollectPrintIdx()
On Error GoTo errHandler
Dim fieldStr As Variant, field As eFuseBdfField, ST As String, SP As String, LSB As Long, msb As Long, prtWidth As Long
Dim decodeStr As New SiteVariant, DecodeBin As New SiteVariant, tmp As String, belowLoc As String, iPrintWr As Long, iPrintRd As Long

    iPrintWr = 0: iPrintRd = 0: prtWidth = IIf(FullSize > 9999, 5, 4)
    DicReadPrintIdx.RemoveAll
    DicWritePrintIdx.RemoveAll
    For Each fieldStr In Fields.Keys
        Set field = Fields(fieldStr)
        ST = IIf(field.Msb1st, "MSB", "LSB")
        SP = IIf(field.Msb1st, "LSB", "MSB")
        LSB = IIf(field.msb > field.LSB, field.LSB, field.msb)
        msb = IIf(field.msb > field.LSB, field.msb, field.LSB)
        DicReadPrintIdx.Add iPrintRd, ") Read from DSSC :" & myUtility.TxtFmt(field.name, 48, padding:="_") _
                & " [(" & ST & ")" & myUtility.TxtFmt(msb, prtWidth, "0") & ":" & myUtility.TxtFmt(LSB, prtWidth, "0") _
                & "(" & SP & ")]"
        iPrintRd = iPrintRd + 1
        DicWritePrintIdx.Add iPrintWr, ") Programming    :" & myUtility.TxtFmt(field.name, 48, "_") _
                & " [(" & ST & ")" & myUtility.TxtFmt(msb, prtWidth, "0") & ":" & myUtility.TxtFmt(LSB, prtWidth, "0") _
                & "(" & SP & ")]"
        iPrintWr = iPrintWr + 1
    Next

Exit Sub
errHandler:
    Call Print_Error_Message(Error_Info, "eFuseBdfBank", "CollectPrintIdx")
    If AbortTest Then Exit Sub Else Resume Next
End Sub

Public Sub PrintReadFuseDatalog(Optional PrintSingleBitWave As Boolean = False, Optional PrintDoubleBitWave As Boolean = False, _
Optional limitFieldCnt As Long = -1, Optional dicItem4Print As Dictionary, Optional printdecode As Boolean = True, Optional isEarly As Boolean = False)
On Error GoTo errHandler
Dim fieldStr As Variant, field As eFuseBdfField, ST As String, SP As String, LSB As Long, msb As Long
Dim decodeStr As New SiteVariant, DecodeBin As New SiteVariant, belowLoc As String, iPrint As Long
Dim instance As String: instance = TheExec.DataManager.instancename
Dim Serial2Print As New DSPWave, dicPrint As New Dictionary, dicKeys As Dictionary
Dim isSpecificField As Boolean: isSpecificField = False
Dim startRowNum As Long
Dim printText As String
Dim myStr As Variant
Dim tmpWave As New DSPWave
Dim site As Variant 'Carter, 20240304

    If Not eFusePrinted Then Exit Sub
    If limitFieldCnt = -1 Then limitFieldCnt = eFusePrintedCnt
    If PrintSingleBitWave Then
        If pgmMode = pgm_DAA Then Ze_DspwaveToParallelMsb DaaCapWaveSerial, 32, Serial2Print
        If pgmMode = pgm_JTAG Then Ze_DspwaveToParallelMsb JtagCapturedSerial, 32, Serial2Print
    End If
    If dicItem4Print Is Nothing Then
        Set dicKeys = Fields
    Else
        Set dicKeys = dicItem4Print
        For Each fieldStr In Fields.Keys
            Set field = Fields(fieldStr)
            startRowNum = field.DataRowNum
            Exit For
        Next
        isSpecificField = True
    End If
    Set dicKeys = IIf(dicItem4Print Is Nothing, Fields, dicItem4Print)
    For Each site In TheExec.sites
        iPrint = 0: dicPrint.RemoveAll:
        If printdecode Then
            If cmpFlag = True Then myUtility.WriteDlg "Site(" & site & "), " & UCase(name) & " pat: " & PatName
                myUtility.WriteDlg "====== Print Out EFuse ReadOut Bits ( " & instance & ") (Site" & site & ")============"
                For Each fieldStr In dicKeys.Keys
                    If limitFieldCnt = -1 Or iPrint < limitFieldCnt Then
                        Set field = Fields(fieldStr)
                        If isSpecificField Then iPrint = field.DataRowNum - startRowNum
                        
                        '20230703, Support partial fusing (middle\early) when "Pgm2File" is enabled and "SwitchFlagForPseudoFuse" is executed, just decode it's fuse item
                        If PseudoFuseEnable And MixPseudoFuseEnable Then
                            If isEarly Then
                                If Not DicEarlyFused.Exists(field.name) Then GoTo SkipPrintField
                            Else
                                'do nothing
                            End If
                        End If

                        decodeStr = "0": DecodeBin = vbNullString: belowLoc = "   "
                        'Hex Infomation Appended
                        If DicEcidCat.Exists(fieldStr) Then  ' Ecid decode
                             decodeStr = Decode(field, True)
                        Else
                            Select Case field.Algorithm  ' IDS/VDDBin/Base decode
                                Case alg_ids:
                                    decodeStr = myUtility.TxtFmt(field.IdsValue_A * 1000 & " mA", 10) 'to mA
                                Case alg_vddbin:
                                    decodeStr = myUtility.TxtFmt(field.VddBinValue_V & " mV", 10)  ' 'to mV
                                Case alg_base:
                                    decodeStr = myUtility.TxtFmt(field.BaseValue_V * 1000 & " mV", 10) ' 'to mV
                                Case Else:
                                    If field.size < 32 And Not field.Algorithm = alg_crc Then
                                        decodeStr = myUtility.TxtFmt(field.DsscDecValue, field.size)
                                    Else
                                        decodeStr = myUtility.TxtFmt(field.DsscValue, 10)
                                    End If
                            End Select
                        End If
                          
                        'Binary Infomation Appended
                        If field.size < 32 And field.Algorithm <> alg_crc Then '20210202
                                DecodeBin = myUtility.Dec2Bin(field.DsscDecValue, field.size)
                        Else
                                DecodeBin = myUtility.xHex2BinStr(field.DsscValue, field.size)
                        End If
                        If Not field.Msb1st Then DecodeBin = StrReverse(DecodeBin)
                        printText = decodeStr
                        dicPrint.Add iPrint, "Site(" & site & DicReadPrintIdx(iPrint) & "=" & myUtility.TxtFmt(printText, 10) & " [" & DecodeBin & "] "
SkipPrintField:
                        iPrint = iPrint + 1
                    End If
                Next fieldStr
                ' Print the read out result string
                For Each myStr In dicPrint.Keys
                    myUtility.WriteDlg dicPrint(myStr)
                Next
        End If
        ' Print the read out dspwave result
        If PrintSingleBitWave Then
            PrintDspWave site, Serial2Print, instance, "SingleBit"
        End If
        
        If PrintDoubleBitWave Then
            If pgmMode = pgm_DAA Then
                PrintDspWave site, DaaCapturedWave, instance, "DoubleBit"
            ElseIf pgmMode = pgm_JTAG Then
                tmpWave.CreateConstant 0, FullSize * 2, DspLong
                tmpWave = JtagCapturedWave.ConvertDataTypeTo(DspLong).ConvertStreamTo(tldspParallel, 32, 0, Bit0IsMsb)
                PrintDspWave site, tmpWave, instance, "DoubleBit"
            End If
        End If
    Next

Exit Sub
errHandler:
    Call Print_Error_Message(Error_Info, "eFuseBdfBank", "PrintReadFuseDatalog")
    If AbortTest Then Exit Sub Else Resume Next
End Sub

Public Sub DumpDspWave(ByRef inwave As DSPWave, Optional idname As String = "PrimeDspWave", Optional conv2ParaBit As Long = -1, Optional dicRowNum As DSPWave = Nothing)
On Error GoTo errHandler
Dim instance As String: instance = TheExec.DataManager.instancename
Dim myParaWave As New DSPWave
Dim site As Variant 'Carter, 20240304

    If conv2ParaBit <> -1 Then
        Ze_DspwaveToParallelMsb inwave, conv2ParaBit, myParaWave
    Else
        Set myParaWave = inwave
    End If
    For Each site In TheExec.sites
        PrintDspWave site, myParaWave, instance, idname, conv2ParaBit, dicRowNum
    Next

Exit Sub
errHandler:
    Call Print_Error_Message(Error_Info, "eFuseBdfBank", "DumpDspWave")
    If AbortTest Then Exit Sub Else Resume Next
End Sub

Public Sub DumpDspWave4Stdf(ByRef inwave As DSPWave)
On Error GoTo errHandler
Dim instance As String: instance = TheExec.DataManager.instancename
Dim myParaWave As New DSPWave
Dim site As Variant 'Carter, 20240304

    Ze_DspwaveToParallelMsb inwave, 32, myParaWave
    For Each site In TheExec.sites
        PrintDspWaveForStdf myParaWave
    Next

Exit Sub
errHandler:
    Call Print_Error_Message(Error_Info, "eFuseBdfBank", "DumpDspWave4Stdf")
    If AbortTest Then Exit Sub Else Resume Next
End Sub

Public Sub PrintWriteFuseDatalog(ByVal dicCategory As Dictionary, Optional PrintSingleBitWave As Boolean = False, Optional PrintDoubleBitWave As Boolean = False, _
Optional PrintDsscSourcedWave As Boolean = False, Optional limitFieldCnt As Long = -1, Optional printdecode As Boolean = True, Optional CompactPrint As Boolean = False, _
Optional dicRowNum As Dictionary, Optional Early As Boolean = False)
On Error GoTo errHandler
Dim fieldStr As Variant, field As eFuseBdfField, ST As String, SP As String, LSB As Long, msb As Long
Dim decodeStr As New SiteVariant, DecodeBin As New SiteVariant, belowLoc As String, iPrint As Long
Dim instance As String: instance = TheExec.DataManager.instancename
Dim DaaSerial2Print As New DSPWave, DaaSerialDb2Print As New DSPWave, DaaSerialDssc2Print As New DSPWave, dicPrint As New Dictionary
Dim myStr As Variant
Dim site As Variant 'Carter, 20240304

    If Not eFusePrinted Then Exit Sub
    If limitFieldCnt = -1 Then limitFieldCnt = eFusePrintedCnt

    If Not SingleBitWave Is Nothing And PrintSingleBitWave Then
        Ze_DspwaveToParallelMsb SingleBitWave, 32, DaaSerial2Print
    End If
    If Not DoubleBitWave Is Nothing And PrintDoubleBitWave And DoubleBits Then
        If Early And name <> "ecid" Then
            Ze_DspwaveToParallelMsb EarlyDsscSourcedWave, 32, DaaSerialDb2Print
        Else
            Ze_DspwaveToParallelMsb DoubleBitWave, 32, DaaSerialDb2Print
        End If
    End If
    If Not DsscSourcedWave Is Nothing And PrintDsscSourcedWave Then
        Ze_DspwaveToParallelMsb DsscSourcedWave, 32, DaaSerialDssc2Print
    End If
    For Each site In TheExec.sites
        iPrint = 0: dicPrint.RemoveAll
        If printdecode Then
            myUtility.WriteDlg "====== Print Out EFuse Program Bits ( " & instance & ") (Site" & site & ")============"
            For Each fieldStr In Fields.Keys
                If limitFieldCnt = -1 Or iPrint < limitFieldCnt Then
                    Set field = Fields(fieldStr)
                    decodeStr = "0":  DecodeBin = vbNullString: belowLoc = "   "
                    If field.TrimFlagBySite And dicCategory.Exists(fieldStr) Then
                        If DicEcidCat.Exists(fieldStr) Then
                                      decodeStr = Decode(field)
                        Else
                            If field.size < 32 And Not field.Algorithm = alg_crc Then
                               decodeStr = myUtility.TxtFmt(field.TrimAteDecValue, 10)
                            Else
                               decodeStr = field.TrimAteValue
                            End If
                        End If
                        'Binary Infomation Appended
                        DecodeBin = myUtility.xHex2BinStr(field.TrimAteValue, field.size)
                        If Not field.Msb1st Then DecodeBin = StrReverse(DecodeBin)
                    End If

                    If dicCategory.Exists(fieldStr) Or Not CompactPrint Then
                        dicPrint.Add iPrint, "Site(" & site & DicWritePrintIdx.item(iPrint) & belowLoc & " = " & myUtility.TxtFmt(decodeStr, 10) & " [" & DecodeBin & "]"
                    End If
                    iPrint = iPrint + 1
                End If
            Next

            For Each myStr In dicPrint.Keys
                   myUtility.WriteDlg dicPrint(myStr)
            Next
        End If
        ' Print the dssc source dspwave
        If Not SingleBitWave Is Nothing And PrintSingleBitWave Then
            PrintDspWave site, DaaSerial2Print, instance, "SingleBit"
        End If
        If Not DoubleBitWave Is Nothing And PrintDoubleBitWave And DoubleBits Then
            PrintDspWave site, DaaSerialDb2Print, instance, "DoubleBit"
        End If
        If Not DsscSourcedWave Is Nothing And PrintDsscSourcedWave Then
            PrintDspWave site, DaaSerialDssc2Print, instance, "DsscSrc"
        End If
    Next

Exit Sub
errHandler:
    Call Print_Error_Message(Error_Info, "eFuseBdfBank", "PrintWriteFuseDatalog")
    If AbortTest Then Exit Sub Else Resume Next
End Sub

Public Function Decode(field As eFuseBdfField, Optional dsscRt As Boolean = False) As String
On Error GoTo errHandler
Dim ecid As String
Dim tmp As String
Dim fieldStage As String
    
        fieldStage = BdfDataBase.GetRealStage(field.BlowLocation)
        fieldStage = Replace(UCase(fieldStage), "_EARLY", "")
        
        If field.Algorithm = alg_lotid Then
            If dsscRt Then
                ecid = myUtility.xHex2BinStr(field.DsscValue, LotIdBits)
                If (LCase(fieldStage) Like "wlft*" And field.Algorithm = alg_lotid) Then
                    Decode = myUtility.DecodeLotId(StrReverse(ecid), False)
                Else
                    Decode = myUtility.DecodeLotId(ecid, False)
                End If
                DsscLotStr = Decode
                If Len(EcidBinStr) < EcidFieldSize Then EcidBinStr = EcidBinStr + ecid
            Else
                ecid = myUtility.xHex2BinStr(field.TrimAteValue, LotIdBits)
                If (LCase(fieldStage) Like "wlft*" And field.Algorithm = alg_lotid) Then
                    Decode = myUtility.DecodeLotId(StrReverse(ecid), False)
                Else
                    Decode = myUtility.DecodeLotId(ecid, False)
                End If
            End If
            Exit Function
        Else
            If field.Algorithm = alg_numeric Then
                If dsscRt Then
                    If (LCase(fieldStage) Like "wlft*" And field.Algorithm = alg_numeric) Then
                        Decode = myUtility.Bin2Dec(StrReverse(myUtility.xHex2BinStr(field.DsscValue, field.size)))
                    Else
                        Decode = myUtility.Hex2Dec(field.DsscValue)
                    End If
                    ecid = myUtility.xHex2BinStr(field.DsscValue, field.size)
                    If Len(EcidBinStr) < EcidFieldSize Then EcidBinStr = EcidBinStr + ecid
                    If LCase(field.name) Like "*wafer_id*" Then DsscWfrStr = Decode
                    If LCase(field.name) Like "*x_coordinate*" Then DsscXcoor = Decode
                    If LCase(field.name) Like "*y_coordinate*" Then DsscYcoor = Decode
                Else
                    If (LCase(fieldStage) Like "wlft*" And field.Algorithm = alg_numeric) Then
                        Decode = myUtility.Bin2Dec(StrReverse(myUtility.xHex2BinStr(field.TrimAteValue, field.size)))
                    Else
                        Decode = myUtility.Hex2Dec(field.TrimAteValue)
                    End If
                End If
                Exit Function
            End If
        End If
Exit Function
errHandler:
    Call Print_Error_Message(Error_Info, "eFuseBdfBank", "Decode")
    If AbortTest Then Exit Function Else Resume Next
End Function

Private Function GetCfgIdxStr(msb As Long, LSB As Long) As String
On Error GoTo errHandler
Dim field As eFuseCfgField, fieldStr As Variant, lsbin As String, msbin As String, tmp() As String

    For Each fieldStr In CfgDataBaseAct.Fields.Keys
        Set field = CfgDataBaseAct.Fields(fieldStr)
        If field.LSB <= LSB And field.msb >= LSB Then
            tmp = Split(field.name, "_")
            lsbin = tmp(UBound(tmp))
        End If
        If field.LSB <= msb And field.msb >= msb Then
             tmp = Split(field.name, "_")
             msbin = tmp(UBound(tmp))
        End If
    Next
    GetCfgIdxStr = msbin & "_" & lsbin

Exit Function
errHandler:
    Call Print_Error_Message(Error_Info, "eFuseBdfBank", "GetCfgIdxStr")
    If AbortTest Then Exit Function Else Resume Next
End Function

Private Sub PrintDspWave(site As Variant, inwave As DSPWave, instance As String, idname As String, Optional bitsPata As Long = -1, Optional dicRowNum As DSPWave = Nothing)
On Error GoTo errHandler
Dim decodeStr As String, stcon As String, bitwidth As Long, rowStr As String, iBit As Long
Dim nBits As Long

    If Not eFusePrinted Then Exit Sub
    bitwidth = IIf(bitsPata = -1, 32, bitsPata)
    stcon = vbCrLf & "====== Print Out EFuse Bits Content(" & instance & ") (Site" & site & ") **" & idname & "** ============" & vbCrLf
    stcon = stcon & headerStr(bitwidth)
    If Not dicRowNum Is Nothing Then stcon = stcon & myUtility.TxtFmt("[Bits in Bdf]", 13) & myUtility.TxtFmt("[Bits in PAT]", 13)
    stcon = stcon & vbCrLf
    myUtility.WriteDlg stcon
    For iBit = 0 To inwave.sampleSize - 1
        stcon = vbNullString
        If dicRowNum Is Nothing Then
            rowStr = myUtility.TxtFmt(CStr(iBit), 3, "0")
        Else
            rowStr = myUtility.TxtFmt(dicRowNum.ElementLite(iBit), 3, "0")
        End If
        decodeStr = myUtility.Dec2Bin(inwave.Element(iBit), bitwidth)
        stcon = stcon & "Row = " & rowStr & "[" & myUtility.TxtFmt(CStr(bitwidth * (iBit + 1) - 1), 5, "0") & ":" & myUtility.TxtFmt(CStr(bitwidth * iBit), 5, "0") & "]: " & decodeStr
        If Not dicRowNum Is Nothing Then
            nBits = IIf(DoubleBits, 2, 1)
            stcon = stcon & "[" & myUtility.TxtFmt(CStr(bitwidth * (dicRowNum.ElementLite(iBit) + 1) - 1), 5, "0") & ":" & myUtility.TxtFmt(CStr(bitwidth * (dicRowNum.ElementLite(iBit))), 5, "0") & "]"
            stcon = stcon & "[" & myUtility.TxtFmt(CStr(bitwidth * (dicRowNum.ElementLite(iBit) + 1) * nBits - 1), 5, "0") & ":" & myUtility.TxtFmt(CStr(bitwidth * (dicRowNum.ElementLite(iBit))) * nBits, 5, "0") & "]"
        End If
    ''20210616
    ''change printing row by row.(previously print whole data)
    ''bank CFG size is over then string data type limitation.
        myUtility.WriteDlg stcon
    Next
    stcon = vbCrLf
    stcon = stcon & "====== End of printing out all Program Bit ============" & vbCrLf
    myUtility.WriteDlg stcon

Exit Sub
errHandler:
    Call Print_Error_Message(Error_Info, "eFuseBdfBank", "PrintDspWave")
    If AbortTest Then Exit Sub Else Resume Next
End Sub

Private Function headerStr(bitwidth As Long) As String
On Error GoTo errHandler
Dim stcon  As String, iBit As Long

    stcon = vbNullString
    For iBit = bitwidth - 1 To 0 Step -1
        stcon = stcon & (iBit \ 10)
    Next
    headerStr = Space(24) & stcon & vbCrLf
    
    stcon = vbNullString
    For iBit = bitwidth - 1 To 0 Step -1
        stcon = stcon & (iBit Mod 10)
    Next
    headerStr = headerStr & Space(24) & stcon & vbCrLf
    If bitwidth = 16 Then
        stcon = myUtility.TxtFmt(".+....V....+....", bitwidth)
    Else
        stcon = myUtility.TxtFmt("..V....+....V....+....V....+....", 32)
    End If
    headerStr = headerStr & Space(24) & stcon & vbCrLf

Exit Function
errHandler:
    Call Print_Error_Message(Error_Info, "eFuseBdfBank", "headerStr")
    If AbortTest Then Exit Function Else Resume Next
End Function

Private Sub PrintDspWaveForStdf(inwave As DSPWave)
On Error GoTo errHandler
Dim decodeStr As String, stcon As String, printName As String, m_THStr As String
Dim iBit As Long, Seg512 As Long, iSplit As Long, MaxBit As Long

    printName = UCase(Replace(name, CfgBankName, "CONFIG"))
    m_THStr = "1"

    myUtility.WriteDlg ""
    If EFUSE_PRINT_DOUBLE_HEXMAP Then myUtility.WriteDlg printName + "," + CStr(TheExec.sites.siteNumber) + ",Double-Bits = " + CStr(DoubleBits)
    
    Seg512 = (inwave.sampleSize - 1) \ 16 + 1 ' 512/32 =16
    MaxBit = inwave.sampleSize * 32 - 1
    For iSplit = 0 To Seg512 - 1
        If iSplit = Seg512 - 1 Then
            stcon = "FUSE," & printName & "," & m_THStr & "," & TheExec.sites.siteNumber & "," & iSplit * 512 & "," & MaxBit & ","
        Else
            stcon = "FUSE," & printName & "," & m_THStr & "," & TheExec.sites.siteNumber & "," & iSplit * 512 & "," & (iSplit + 1) * 512 - 1 & ","
        End If
        decodeStr = vbNullString
        For iBit = iSplit * 16 To (iSplit + 1) * 16 - 1
            If iBit < inwave.sampleSize Then decodeStr = myUtility.Dec2Bin(inwave.Element(iBit), 32) + decodeStr
        Next
        stcon = stcon & Replace(myUtility.Bin2HexStr(decodeStr, (Len(decodeStr) - 1) \ 4 + 1), "0x", "")
        myUtility.WriteDlg stcon
    Next

Exit Sub
errHandler:
    Call Print_Error_Message(Error_Info, "eFuseBdfBank", "PrintDspWaveForStdf")
    If AbortTest Then Exit Sub Else Resume Next
End Sub

Private Sub CollectCrc(myfield As eFuseBdfField)
On Error GoTo errHandler
Dim CrcBits As New Dictionary
Dim reParse As New RegExp, reParseRng As New RegExp, TypeParseRng As New RegExp, v As String, reCheckParseRng As New RegExp
Dim ibitSt As Integer, ibitSp As Integer, itmp As Long
Dim objmatches As Variant, obj As Variant, tmp As Variant
Dim xCrcbits As New Dictionary

    If LCase(myfield.Algorithm) = Algorithm.alg_crc Then
    '   CRC_IgnoreBits=[287:0],[381:381],[389:388],[1375:1344],[2585:1536],[6143:6128],REAL_CODE_FROM_TEST
    '   Compute CRC-32 over MON_FUSE[639:0]. UID_MODE Fuse Programming spec is a good reference.,N/A
    '   PgmFlow=nonDEID, CRC_calcBits=[63:0],REAL_CODE_FROM_TEST
        myfield.Msb1st = True ' Force it to true, no matter it's ecid or not
        v = myfield.Description

        TypeParseRng.IgnoreCase = True
        TypeParseRng.Pattern = "one_complement_target"
        TypeParseRng.Global = True
        
        If TypeParseRng.test(v) Then    '20220602, modify for one complement
            reCheckParseRng.IgnoreCase = True
            reCheckParseRng.Pattern = "\[\w+]"
            reCheckParseRng.Global = True
            myfield.crc_type = crc_onecomp
            Set objmatches = reCheckParseRng.Execute(v)
            If objmatches.Count <> 1 Then
                myUtility.MessageBox myfield.name & " has error format about completment field!" & "(" & v & ")"
                myUtility.WriteDlg "[Error] " & myfield.name & " has error format about completment field!" & "(" & v & ")"
            End If
            
            For Each obj In objmatches
                tmp = Replace(Replace(obj.value, "[", ""), "]", "")
            Next
            Crcs_Comp.Add myfield.name, tmp
        Else
            myfield.crc_type = crc_normal
            reParseRng.IgnoreCase = True
            reParseRng.Pattern = "\[\d+:\d+\]"
            reParseRng.Global = True
            
            reCheckParseRng.IgnoreCase = True
            reCheckParseRng.Pattern = "\[\d+\]"
            reCheckParseRng.Global = True
    
            If reParseRng.test(v) And Not reCheckParseRng.test(v) Then
                Set objmatches = reParseRng.Execute(v)
                For Each obj In objmatches
                    tmp = Split(Replace(Replace(obj.value, "[", ""), "]", ""), ":")
                    ibitSt = CInt(tmp(0)): ibitSp = CInt(tmp(1))
                    If ibitSt > ibitSp Then
                       itmp = ibitSt
                       ibitSt = ibitSp
                       ibitSp = itmp
                    End If
                    For itmp = ibitSt To ibitSp
                        CrcBits.Add itmp, True
                    Next
                Next
            Else
                myUtility.MessageBox myfield.name & " has error format about ignore bit!" & "(" & v & ")"
                Call Print_Error_Message(Error_Info, "eFuseBdfBank", "CollectCrc", myfield.name & " has error format about ignore bit!" & "(" & v & ")")
            End If
            
            reParse.IgnoreCase = True
            reParse.Pattern = "CRC_IgnoreBits"
            If reParse.test(v) Then
                For itmp = 0 To FullSize - 1
                   If Not CrcBits.Exists(itmp) Then xCrcbits.Add itmp, True
                Next
                Set xCrcbits = myUtility.DicKeysSort(xCrcbits)
                Crcs.Add myfield.name, xCrcbits
            ElseIf CrcBits.Count > 0 Then
                Set CrcBits = myUtility.DicKeysSort(CrcBits)
                Crcs.Add myfield.name, CrcBits
            Else
                Call Print_Error_Message(Error_Info, "eFuseBdfBank", "CollectCrc", "[" & myfield.name & "] CRC calculation bits didn't declare in BDF file, please check it!")
                If myfield.name = "ecid_crc" Then
                    For itmp = 0 To 63
                        CrcBits.Add itmp, True
                    Next
                    Set CrcBits = myUtility.DicKeysSort(CrcBits)
                    Crcs.Add myfield.name, CrcBits
                End If
            End If
         End If
    Else
        Exit Sub
    End If

Exit Sub
errHandler:
    Call Print_Error_Message(Error_Info, "eFuseBdfBank", "CollectCrc")
    If AbortTest Then Exit Sub Else Resume Next
End Sub

'Public Sub PutDaaCapToDb(ByVal InWave As DSPWave)
'    Dim nBits As Long, opWave As New DSPWave, sWave As New DSPWave
'    Dim flWave As New DSPWave ''202004xx for ap
'    nBits = IIf(DoubleBits, 2, 1)
'    For Each site In TheExec.sites
'        If InWave.SampleSize <> FullSize * nBits / DaaParaBits + IIf(FullSize * nBits / DaaParaBits Mod 4 > 0, (4 - FullSize * nBits / DaaParaBits Mod 4), 0) Then '(FullSize + (2 - myDspwave.SampleSize Mod 2)) * nBits Then ' FullSize * nBits / DaaParaBits Then
'            myUtility.MessageBox "The sample size should be " & FullSize * nBits / DaaParaBits
'            Exit Sub
'        End If
'    Next
'    Select Case DaaParaBits
'        Case 32:
'            Set opWave = InWave
'        Case 8:
'            Ze_DspwaveParallel8ToParallel32 InWave, opWave
'        Case Else
'           myUtility.MessageBox "No support parallel bits " & DaaParaBits
'    End Select
'
'    If DoubleBits Then
'        Ze_GetSingleBitWave opWave, sWave, flWave ''202004xx for ap
'        'Ze_GetSingleBitWave opWave, sWave
'        Ze_DspwaveToSerialMsb sWave, 16
'    Else
'        Set sWave = opWave
'        Ze_DspwaveToSerialMsb sWave, 32
'    End If
'    For Each site In TheExec.sites
'         DaaCapturedWave = InWave
'         DaaCapWaveSerial = sWave
'         DaaCapWaveSerialFL = flWave
'    Next
'End Sub

Public Sub PutJtagCapToDb(ByVal inwave As DSPWave, Optional ForceDoubleBit As Boolean = False)
On Error GoTo errHandler
Dim nBits As Long, sWave As New DSPWave, waveSize As Long
Dim idxWave As New DSPWave
Dim site As Variant

    nBits = IIf(DoubleBits And ForceDoubleBit, 2, 1) 'Typically the JTAG read is the result of OR, APB is designed for read double bits through JTAG
    For Each site In TheExec.sites.Selected
        waveSize = inwave.sampleSize
        If waveSize <> FullSize * nBits Then
            myUtility.MessageBox "The sample size should be " & FullSize * nBits
            Call Print_Error_Message(Error_Info, "eFuseBdfBank", "PutJtagCapToDb", "The sample size should be " & FullSize * nBits)
            Exit Sub
        End If
    Next

    If nBits = 2 Then 'Typically the JTAG read is the result of OR, APB is designed for read double bits through JTAG
        Ze_GetSingleBitWave_Jtag inwave, sWave
    Else
        Set sWave = inwave
    End If

    If JtagMsbFirst Then idxWave.CreateRamp waveSize - 1, -1, waveSize, DspLong
    For Each site In TheExec.sites
        If JtagMsbFirst Then sWave = sWave.Lookup(idxWave)
        JtagCapturedWave = inwave
        JtagCapturedSerial = sWave
    Next

Exit Sub
errHandler:
    Call Print_Error_Message(Error_Info, "eFuseBdfBank", "PutJtagCapToDb")
    If AbortTest Then Exit Sub Else Resume Next
End Sub

Public Sub GetBitStSp(field As eFuseBdfField, ByRef ST As Long, ByRef SP As Long)
On Error GoTo errHandler
Dim stBit As Long, spBit As Long, iBit As Long
Dim sttmp As Long

    stBit = field.LSB: spBit = field.msb:
    If stBit > spBit Then
        sttmp = spBit:    spBit = stBit:    stBit = sttmp:
    End If
    ST = stBit: SP = spBit

Exit Sub
errHandler:
    Call Print_Error_Message(Error_Info, "eFuseBdfBank", "GetBitStSp")
    If AbortTest Then Exit Sub Else Resume Next
End Sub

Public Function RestoreFromFileToDb()
On Error GoTo errHandler
Dim oWave As New DSPWave, sitev As New SiteBoolean

    sitev = myUtility.GetStoreWaveform(name, ".\offlinewave\", oWave)
    PutDsscRtToDb oWave, sitev

Exit Function
errHandler:
    Call Print_Error_Message(Error_Info, "eFuseBdfBank", "RestoreFromFileToDb")
    If AbortTest Then Exit Function Else Resume Next
End Function

Public Function PutDsscRtToDb(ByRef inwave As DSPWave, Optional put2Db As SiteBoolean, Optional dicItem4Print As Dictionary, Optional isEarly As Boolean = False) As Long
On Error GoTo errHandler
Dim funcName As String: funcName = "PutDsscRtToDb"
Dim field As eFuseBdfField, item As Variant, site As Variant
Dim stBit As Long, spBit As Long, iBit As Long, iField As Long
Dim fieldValue As New SiteVariant, binstr As String ', DecimalWave As DSPWave
Dim lotfield As String, wfrfield As String, xcoorfield As String, ycoorfield As String
Dim cfgStr As String
Dim hexecid As String, ttlbits As Long

    If HadDeidFuse Then Set EcidHexStr = New SiteVariant
    DsscCFGCondStr = vbNullString
    FuseRevisionStr = vbNullString

    For Each site In TheExec.sites
        iField = 0
        If put2Db Then
            For Each item In Fields.Keys
                Set field = Fields(item)
                '20230703, Support partial fusing (middle\early) when "Pgm2File" is enabled and execute "SwitchFlagForPseudoFuse" item, just decode it's fuse items
                If PseudoFuseEnable And MixPseudoFuseEnable Then
                    If isEarly Then
                       If Not DicEarlyFused.Exists(field.name) Then GoTo SkipPutReadBuf
                    ElseIf Not dicItem4Print Is Nothing Then
                        If Not dicItem4Print.Exists(field.name) Then GoTo SkipPutReadBuf
                   Else
                        'do nothing
                   End If
                End If
                If field.size < 32 And Not field.Algorithm = alg_crc Then
                    If (Not PseudoFuseEnable) Or MixPseudoFuseEnable Then
                        fieldValue = myUtility.Dec2HexStr(DecimalWave.ElementLite(iField), field.hBytes)
                        field.DsscDecValue = DecimalWave.ElementLite(iField)
                    Else
                        If myUtility.testedStages.Exists(BdfDataBase.GetRealStage(field.BlowLocation)) Or myUtility.testedStages.Exists(Replace(UCase(field.BlowLocation), "_EARLY", "")) Then
                            '20210716,Add for FT1 pseudo fuse read from real device
                            '20211210, Modify for FT pseudo fuse with real package
                            If gB_Package_PsudoFuse = True And (UCase(TheExec.CurrentJob) Like "*FT*" And ParsePseudoFuseFile = False) Then
                                fieldValue = myUtility.Dec2HexStr(DecimalWave.ElementLite(iField), field.hBytes)
                                field.DsscDecValue = DecimalWave.ElementLite(iField)
                            Else
                                fieldValue = field.DsscValue
                                If myUtility.IsHex(fieldValue) Then
                                    field.DsscDecValue = myUtility.Hex2Dec(fieldValue)
                                Else
                                    field.DsscDecValue = fieldValue
                                End If
                            End If
                        Else
                            fieldValue = field.TrimAteValue
                            If myUtility.IsHex(fieldValue) Then
                                field.DsscDecValue = myUtility.Hex2Dec(fieldValue)
                            Else
                                field.DsscDecValue = fieldValue
                            End If
                        End If
                    End If

                Else
                    If (Not PseudoFuseEnable) Or MixPseudoFuseEnable Then
                        GetBitStSp field, stBit, spBit
                        fieldValue = vbNullString:  binstr = vbNullString
                        For iBit = stBit To spBit
                               binstr = inwave.ElementLite(iBit) & binstr
                        Next
                        If Not field.Msb1st Then binstr = StrReverse(binstr)
                        fieldValue = myUtility.Bin2HexStr(binstr, field.hBytes)
                   Else
                       If myUtility.testedStages.Exists(BdfDataBase.GetRealStage(field.BlowLocation)) Or myUtility.testedStages.Exists(Replace(UCase(field.BlowLocation), "_EARLY", "")) Then
                           '20210716,Add for FT1 pseudo fuse read from real device
                           '20211210, Modify for FT pseudo fuse with real package
                           If gB_Package_PsudoFuse = True And (UCase(TheExec.CurrentJob) Like "*FT*" And ParsePseudoFuseFile = False) Then
                               GetBitStSp field, stBit, spBit
                               fieldValue = vbNullString:  binstr = vbNullString
                               For iBit = stBit To spBit
                                      binstr = inwave.ElementLite(iBit) & binstr
                               Next
                               If Not field.Msb1st Then binstr = StrReverse(binstr)
                               fieldValue = myUtility.Bin2HexStr(binstr, field.hBytes)
                           Else
                               fieldValue = field.DsscValue
                           End If
                       Else
                           fieldValue = field.TrimAteValue
                       End If
                   End If
                End If
                field.DsscValue = fieldValue
                If field.Trimmed Then field.TrimmedData = field.DsscValue
SkipPutReadBuf:
                iField = iField + 1
                If HadDeidFuse And DicEcidCat.Exists(CStr(item)) Then
                     If CStr(item) Like "*lot*id" Then lotfield = CStr(item)
                     If CStr(item) Like "*wafer*id" Then wfrfield = CStr(item)
                     If CStr(item) Like "*x*coor*" Then xcoorfield = CStr(item)
                     If CStr(item) Like "*y*coor*" Then ycoorfield = CStr(item)
                End If
                         
                ''202004xx for ap
                If field.Algorithm = alg_cond Then
                    cfgStr = myUtility.xHex2BinStr(field.DsscValue, field.size)
                    '20210722,Modify IEDA should be LSB->MSB, EX:DsscValue ="0011" to "1100"
                    cfgStr = StrReverse(cfgStr)
                    DsscCFGCondStr = DsscCFGCondStr + cfgStr
                End If
                
                If LCase(field.name) = "fuse_revision" Then
                    FuseRevisionStr = myUtility.xHex2BinStr(field.DsscValue, field.size)
                End If
            Next
            If HadDeidFuse Then
                ttlbits = 0
                Set field = Fields(lotfield):   hexecid = myUtility.Hex2BinStr(field.DsscValue, field.size): ttlbits = ttlbits + field.size
                HramLotId = Decode(field, True)
                Set field = Fields(wfrfield):   hexecid = hexecid & myUtility.Hex2BinStr(field.DsscValue, field.size): ttlbits = ttlbits + field.size
                HramWaferId = Decode(field, True)
                Set field = Fields(xcoorfield): hexecid = hexecid & myUtility.Hex2BinStr(field.DsscValue, field.size): ttlbits = ttlbits + field.size
                HramXCoord = Decode(field, True)
'                If (myUtility.currStage = "FT1" Or myUtility.currStage = "FT2" Or myUtility.currStage = "FT3") And (UCase(name) Like "ECID") Then
                If UCase(TheExec.CurrentChanMap) Like "CHANNELMAP_FT*" And (UCase(name) Like "ECID") Then 'JackChou 202404 To allow Stage_FT with ChannelMap_CP case [Requester-TSMC-KoPei]
                    Call TheExec.Datalog.Setup.WaferSetup.SetXCoord(site, myUtility.Hex2Dec(field.DsscValue))
                    XCoord = HramXCoord
                End If
                Set field = Fields(ycoorfield): hexecid = hexecid & myUtility.Hex2BinStr(field.DsscValue, field.size): ttlbits = ttlbits + field.size
                HramYCoord = Decode(field, True)
'                If (myUtility.currStage = "FT1" Or myUtility.currStage = "FT2" Or myUtility.currStage = "FT3") And (UCase(name) Like "ECID") Then
                If UCase(TheExec.CurrentChanMap) Like "CHANNELMAP_FT*" And (UCase(name) Like "ECID") Then 'JackChou 202404 To allow Stage_FT with ChannelMap_CP case [Requester-TSMC-KoPei]
                    Call TheExec.Datalog.Setup.WaferSetup.SetYCoord(site, myUtility.Hex2Dec(field.DsscValue))
                    YCoord = HramYCoord
                End If
                EcidHexStr = Replace(myUtility.Bin2HexStr(StrReverse(hexecid), 16), "0x", "") '(ttlbits - 1) \ 4 + 1
                       
             End If
        End If
    Next

Exit Function
errHandler:
    Call Print_Error_Message(Error_Info, "eFuseBdfBank", "PutDsscRtToDb")
    If AbortTest Then Exit Function Else Resume Next
End Function

Public Sub PutRandomCodes()
On Error GoTo errHandler
Dim funcName As String:: funcName = "PutRandomCodes"
Dim fieldStr As Variant, UTIL_BYTE_SIZE As Long: UTIL_BYTE_SIZE = 8
Dim field As eFuseBdfField, Output As String, i As Long
Dim site As Variant

    For Each site In TheExec.sites
        For Each fieldStr In DicRandomFused.Keys
           Set field = Fields(fieldStr)
            If field.DefaultOrReal Then
                Output = vbNullString
                For i = 0 To field.size \ UTIL_BYTE_SIZE - 1
                    Output = Output + myUtility.Dec2HexStr(myUtility.RNG_cryptoRandomByte, UTIL_BYTE_SIZE \ 4)
                Next
                field.TrimAteValue = "0x" & Replace(Output, "0x", "")
                field.TrimFlagBySite = True
            End If
        Next
    Next site

Exit Sub
errHandler:
    Call Print_Error_Message(Error_Info, "eFuseBdfBank", "PutRandomCodes")
    If AbortTest Then Exit Sub Else Resume Next
End Sub

Public Sub PutBinCutCodes()
On Error GoTo errHandler
Dim fieldStr As Variant, field As eFuseBdfField
Dim m_Pmode As Long, values As New SiteVariant, PatPF As New SiteBoolean

    PatPF = True
    If Not HadVddBinFuse Then Exit Sub
    For Each fieldStr In DicBinCut 'This dictionary contains items for current test stage only
        Set field = Fields(fieldStr)
        If field.DefaultOrReal = dr_real Then
            m_Pmode = VddBinStr2Enum(Replace(LCase(field.name), "_shadow", ""))
            Set values = GenVddBinFuseCode(field.name, m_Pmode, field)
            SetEfuse field.name, values, PatPF
'       Else
'           myUtility.WriteErrDlg "Missing VddBin Value for """ & field.Name & """"
        End If
    Next

Exit Sub
errHandler:
    Call Print_Error_Message(Error_Info, "eFuseBdfBank", "PutBinCutCodes")
    If AbortTest Then Exit Sub Else Resume Next
End Sub

Public Sub PutIdsCodes()
On Error GoTo errHandler
Dim fieldStr As Variant, field As eFuseBdfField
Dim m_Pmode As Long, values As New SiteVariant, PatPF As New SiteBoolean

    PatPF = True
    If Not HadIdsFuse Then Exit Sub
    For Each fieldStr In DicIds ' This dictionary contains items for current test stage only
        Set field = Fields(fieldStr)
        If field.DefaultOrReal = dr_real Then
            Set values = ObtainIdsValues(field)
            SetEfuse field.name, values, PatPF, True, , , True
        End If
    Next

Exit Sub
errHandler:
    Call Print_Error_Message(Error_Info, "eFuseBdfBank", "PutIdsCodes")
    If AbortTest Then Exit Sub Else Resume Next
End Sub

Public Sub OfflinePutTestedStageEcidFields()
On Error GoTo errHandler
Dim fieldStr As Variant, field As eFuseBdfField
Dim ecid As String
Dim site As Variant
    For Each site In TheExec.sites
        For Each fieldStr In Fields
            Set field = Fields(fieldStr)
            Select Case fieldStr
                Case "lot_id":
                    ecid = GlbUtility.GetLotIdBin(TheExec.Datalog.Setup.LotSetup.lotId, field.size, True)
                    field.DsscValue = ecid
                    field.DsscDecValue = GlbUtility.Hex2Dec(ecid)
                    field.TrimmedData = field.DsscValue
                    field.TrimAteValue = field.DsscValue
                Case "wafer_id":
                    field.DsscValue = GlbUtility.Dec2HexStr(val(TheExec.Datalog.Setup.WaferSetup.ID))
                    field.DsscDecValue = val(TheExec.Datalog.Setup.WaferSetup.ID)
                    field.TrimmedData = field.DsscValue
                    field.TrimAteValue = field.DsscValue
                Case "x_coordinate"
                    field.DsscValue = GlbUtility.Dec2HexStr(TheExec.Datalog.Setup.WaferSetup.GetXCoord(site))
                    field.DsscDecValue = val(TheExec.Datalog.Setup.WaferSetup.GetXCoord(site))
                    field.TrimmedData = field.DsscValue
                    field.TrimAteValue = field.DsscValue
                Case "y_coordinate"
                    field.DsscValue = GlbUtility.Dec2HexStr(TheExec.Datalog.Setup.WaferSetup.GetYCoord(site))
                    field.DsscDecValue = val(TheExec.Datalog.Setup.WaferSetup.GetYCoord(site))
                    field.TrimmedData = field.DsscValue
                    field.TrimAteValue = field.DsscValue
                Case "ecid_crc"
                    CalculateCRC "ECID_CRC"
            End Select
        Next
    Next site
Exit Sub
errHandler:
    Call Print_Error_Message(Error_Info, "eFuseBdfBank", "PutTestedStageEcidFilds")
        If AbortTest Then Exit Sub Else Resume Next
End Sub

Public Sub OfflinePutTestedStageValues()
On Error GoTo errHandler
Dim fieldStr As Variant, field As eFuseBdfField
Dim m_Pmode As Long, values As New SiteVariant, PatPF As New SiteBoolean
Dim bintmpstr As String
Dim hvalue As SiteVariant
Dim site As Variant
PatPF = True
    For Each fieldStr In Fields
        Set field = Fields(fieldStr)
        If GlbUtility.testedStages.Exists(field.BlowLocation) Then
            If field.DefaultOrReal = dr_real Then
                Select Case field.Algorithm
                Case alg_ids:
                    Set values = ObtainIdsValues(field)
                    values = IDS_Conversion(field, values)
                    Set hvalue = myUtility.oDec2HexStr(values, field.hBytes)
                    For Each site In TheExec.sites
                        If Not field.TrimFlagBySite(site) Then
                            field.PatTestPass(site) = PatPF(site)
                            field.TrimFlagBySite(site) = True
                            field.Trimmed = True
                            If myUtility.xIsHex(values(site)) And field.size < 32 Then
                                bintmpstr = myUtility.xHex2BinStr(values(site), field.size)
                                field.TrimAteDecValue(site) = myUtility.Bin2Dec(bintmpstr)
                            End If
                            field.TrimAteValue(site) = hvalue(site)
                            field.TrimmedData = hvalue(site)
                            field.DsscValue = hvalue(site)
                            field.DsscDecValue = values(site)
                        End If
                    Next
                Case alg_vddbin:
                    m_Pmode = VddBinStr2Enum(Replace(LCase(field.name), "_shadow", ""))
                    Set values = GenVddBinFuseCode(field.name, m_Pmode, field)
                    For Each site In TheExec.sites
                        If Not field.TrimFlagBySite(site) Then
                            field.PatTestPass(site) = PatPF(site)
                            field.TrimFlagBySite(site) = True
                            field.Trimmed = True
                            If myUtility.xIsHex(values(site)) And field.size < 32 Then
                                bintmpstr = myUtility.xHex2BinStr(values(site), field.size)
                                field.TrimAteDecValue(site) = myUtility.Bin2Dec(bintmpstr)
                            End If
                            field.TrimmedData = myUtility.Dec2HexStr(values(site))
                            field.DsscValue = myUtility.Dec2HexStr(values(site))
                            field.DsscDecValue = values(site)
                        End If
                    Next
                Case alg_app:
                    If field.size < 32 Then
                        If field.name = "product_identifier" Then
                            field.DsscValue = "0x0"
                            field.DsscDecValue = 0
                        ElseIf field.name = "bkm_process" Or field.name = "bkm_package" Then
                            field.DsscValue = myUtility.Bin2HexStr(BdfDataBase.DicBKMMap("BKM2.3"), field.size)
                            field.DsscDecValue = myUtility.Bin2Dec(BdfDataBase.DicBKMMap("BKM2.3"))
                            field.TrimmedData = field.DsscValue
                        Else
                            field.DsscDecValue = Round(CDbl((field.Hlimit - field.Llimit) * Rnd(Now) + field.Llimit), 0)
                            field.DsscValue = myUtility.Dec2HexStr(field.DsscDecValue)
                            field.TrimmedData = field.DsscValue
                        End If
                    Else
                         field.DsscValue = "0x0" ' if size >=32, needs to set it to hex format
                    End If
                Case alg_crc:
                End Select
            ElseIf field.Algorithm = alg_na Then
                Dim cmpStr As String, udrStr As String
                Dim udrBank As New eFuseBdfBank, cmpbank As New eFuseBdfBank
                Dim cmpfield As New eFuseBdfField, udrfield As New eFuseBdfField
                udrStr = Replace(LCase(name), "cmp", "udr")
                Set udrBank = GetBdfBank(udrStr)
                Set cmpbank = GetBdfBank(name)
                For Each site In TheExec.sites
                    Set cmpfield = cmpbank.Fields(fieldStr)
                    Set udrfield = udrBank.Fields(fieldStr)
                    cmpfield.TrimAteValue = udrfield.DsscValue
                    cmpfield.TrimmedData = udrfield.DsscValue
                    cmpfield.DsscValue = udrfield.DsscValue
                    cmpfield.DsscDecValue = udrfield.DsscDecValue
                    cmpfield.PatTestPass(site) = PatPF(site)
                    cmpfield.TrimFlagBySite(site) = True
                Next site
            End If
        End If
    Next fieldStr
Exit Sub
errHandler:
    Call Print_Error_Message(Error_Info, "eFuseBdfBank", "OfflinePutTestedStageValues")
    If AbortTest Then Exit Sub Else Resume Next
End Sub

Public Sub PseudoFusedFillup(dic4Fused As Dictionary) 'use for offline only
On Error GoTo errHandler
Dim fieldStr As Variant, field As eFuseBdfField, value As New SiteVariant, pf As New SiteBoolean
Dim needSetFuse As Boolean
Dim site As Variant

    If myUtility.OnlineMode Then Exit Sub
    For Each fieldStr In dic4Fused
        Set field = Fields(fieldStr)
        needSetFuse = False
        If Not field.Trimmed And Not field.ReadOnly And Not field.Algorithm = alg_crc Then
            For Each site In TheExec.sites
                If Not field.TrimFlagBySite(site) Then
                    needSetFuse = True
                    If field.size < 32 Then
                        If field.name = "product_identifier" Then
                            value = 0
                        Else
                            value = Round(CDbl((field.Hlimit - field.Llimit) * Rnd(Now) + field.Llimit), 0)
                        End If
                    Else
                         value = "0x0" ' if size >=32, needs to set it to hex format
                    End If
                End If
            Next
            If needSetFuse Then
                pf = True
                SetEfuse field.name, value, pf
            End If
        End If
    Next

Exit Sub
errHandler:
    Call Print_Error_Message(Error_Info, "eFuseBdfBank", "PseudoFusedFillup")
    If AbortTest Then Exit Sub Else Resume Next
End Sub

Public Sub UpdateLogFormat()
On Error GoTo errHandler
    myUtility.UpdateDLogColumns (CateNameMaxLen + 25)
Exit Sub
errHandler:
    Call Print_Error_Message(Error_Info, "eFuseBdfBank", "UpdateLogFormat")
    If AbortTest Then Exit Sub Else Resume Next
End Sub

Public Sub PutField2IEDA(identifyItem As String, fieldStr As String)
On Error GoTo errHandler
Dim regKey As String: regKey = vbNullString
Dim field As eFuseBdfField: Set field = Fields(fieldStr)
Dim binstr As New SiteVariant

    binstr = myUtility.Hex2BinStr(field.DsscValue, field.size)
    myUtility.Put2IEDA identifyItem, binstr

Exit Sub
errHandler:
    Call Print_Error_Message(Error_Info, "eFuseBdfBank", "PutField2IEDA")
    If AbortTest Then Exit Sub Else Resume Next
End Sub

Public Function IsFieldExist(category As String) As Boolean
On Error GoTo errHandler

    If Not Fields.Exists(category) Then
        myUtility.MessageBox "[" + name + "] - Category is not exist! in EFUSE_BitDef_Table (" & category + ")"
        Call Print_Error_Message(Error_Info, "eFuseBdfBank", "IsFieldExist", "[" + name + "] - Category is not exist! in EFUSE_BitDef_Table (" & category + ")")
        IsFieldExist = False: Exit Function
    End If
    IsFieldExist = True

Exit Function
errHandler:
    Call Print_Error_Message(Error_Info, "eFuseBdfBank", "IsFieldExist")
    If AbortTest Then Exit Function Else Resume Next
End Function

''the default value set in efuse initialization.
''update the real code for cp1_early.
Public Sub UpdateField2EarlyDSSC()
On Error GoTo errHandler
Dim myfield As eFuseBdfField, fieldStr As Variant, site As Variant

    Set DsscWave_Eff = New DSPWave
    Set EarlyDsscSourcedWave = New DSPWave
    If myUtility.testedStages.Exists(myUtility.currStage) Then
        DsscWave_Eff.CreateConstant 0, FullSize
    End If
    For Each site In TheExec.sites
        If Not myUtility.testedStages.Exists(myUtility.currStage) Then
            DsscWave_Eff = GetDspWave(0, dicCategory:=DicEarlyFused)
        Else
            
        End If
        earlySingleBitData = DsscWave_Eff.data
        earlyData = GenDsscSrcDspWave(True, True).data
        
        EarlyDsscSourcedWave.data = earlyData
        EarlySingleBitWave.data = earlySingleBitData
    Next

Exit Sub
errHandler:
    Call Print_Error_Message(Error_Info, "eFuseBdfBank", "UpdateField2EarlyDSSC")
    If AbortTest Then Exit Sub Else Resume Next
End Sub

Public Function StoreDataToBank(key As String, data() As String, header() As String, ByRef offsetValue As Long)
On Error GoTo errHandler
Dim fieldStr As Variant
Dim field As eFuseBdfField
Dim LoopIndex As Long
Dim binarr() As SiteLong
Dim iBit As Long, stBit As Long, spBit As Long, bitStep As Long, bitDealt As Long
Dim ss As Variant

    ss = TheExec.sites.siteNumber
    LoopIndex = 0
    For Each fieldStr In Fields.Keys
        Set field = Fields(fieldStr)
        If myUtility.testedStages.Exists(BdfDataBase.GetRealStage(field.BlowLocation)) Then
            field.Trimmed = True
        End If

        If myUtility.IsStrMatch(header(offsetValue + LoopIndex), field.name) Then
            If field.Algorithm = alg_ids Then
                field.DsscValue = field.IdsValue_Hex(CDbl(data(offsetValue + LoopIndex)) * 0.001)
                field.DsscDecValue = myUtility.Hex2Dec(field.DsscValue)
                field.FuseMeasureValue = CDbl(data(offsetValue + LoopIndex)) * 0.001
            ElseIf field.Algorithm = alg_vddbin Then
                field.DsscValue = field.VddBinValue_Hex(CDbl(data(offsetValue + LoopIndex)))
                field.DsscDecValue = myUtility.Hex2Dec(field.DsscValue)
                field.FuseMeasureValue = CDbl(data(offsetValue + LoopIndex))
            ElseIf field.Algorithm = alg_base Then
                field.DsscValue = field.BaseValue_Hex(CDbl(data(offsetValue + LoopIndex)))
                field.DsscDecValue = (CDbl(data(offsetValue + LoopIndex)) / field.Resolution) - 1
            ElseIf field.Algorithm = alg_lotid Then
                field.DsscValue = myUtility.GetLotIdBin(data(offsetValue + LoopIndex), field.size, True)
            ElseIf field.size >= 32 Or field.Algorithm = alg_crc Then
                field.DsscValue = data(offsetValue + LoopIndex)
            Else
                field.DsscValue = myUtility.Dec2HexStr(CDbl(data(offsetValue + LoopIndex)), field.size / 4 + 1)
                field.DsscDecValue = CDbl(data(offsetValue + LoopIndex))
            End If
            If field.Trimmed Then field.TrimmedData = field.DsscValue
            stBit = field.LSB: spBit = field.msb: bitStep = 1: ' MSB 1st
            If field.Msb1st Then
                GetBitStSp field, stBit, spBit
            ElseIf stBit > spBit Then
                bitStep = -1: 'LSB first, if doing CRC, it needs to use the MSB first
            End If
            binarr = myUtility.Hex2BinArray(field.DsscValue, field.size)
            
            bitDealt = 0
            For iBit = stBit To spBit Step bitStep
                PseudoFuseDspWave.ElementLite(iBit) = binarr(bitDealt)
                bitDealt = bitDealt + 1
            Next
        End If
        TheExec.Datalog.WriteComment "site(" & ss & ")  " & "[set efuse]" & field.name & " = " & CStr(field.DsscValue(ss))
        LoopIndex = LoopIndex + 1
    Next
    offsetValue = offsetValue + Fields.Count

Exit Function
errHandler:
    Call Print_Error_Message(Error_Info, "eFuseBdfBank", "StoreDataToBank")
    If AbortTest Then Exit Function Else Resume Next
End Function

'202006xx dsp modify
Public Function DigCapWaveProcess(blankChk As Boolean, _
                                   ByRef capWave As DSPWave, _
                                   result As SiteLong, _
                                   Optional ByVal stage As Boolean = True, _
                                   Optional ByVal dicCategory As Dictionary, _
                                   Optional ForceDoubleBit As Boolean = False, _
                                   Optional includeEarlyFuse As Boolean = False, _
                                   Optional chkNonEarly As Boolean = False, _
                                   Optional Rvenable As Boolean = False, Optional MultitypJob As String = vbNullString)
On Error GoTo errHandler
Dim nBits As Long
Dim dicCared As New DSPWave
Dim stBit As Long, spBit As Long, iBit As Long
Dim fieldStr As Variant, jobType As Variant, field As eFuseBdfField
Dim dicCatCheck As Dictionary
Dim paraWave As New DSPWave
Dim seriWave As New DSPWave
Dim seriWave_FL As New DSPWave
Dim RVOut As New DSPWave
Dim decWave As New DSPWave
Dim capwave_count As New SiteLong: capwave_count = 0
Dim TestedWaveCheck As Boolean
Dim Testedstagemask As New DSPWave
Dim RVsegmentFinal As New DSPWave
Dim RVIndexWaveFinal As New DSPWave
Dim site As Variant 'Carter, 20240304

    nBits = IIf(DoubleBits And ForceDoubleBit, 2, 1)
    dicCared.CreateConstant 0, FullSize

    If blankChk Or g_Rvenable Then
        If Not dicCategory Is Nothing Then
            Set dicCatCheck = New Dictionary
                For Each fieldStr In dicCategory.Keys
                      dicCatCheck.Add fieldStr, False
                Next
        End If
        For Each fieldStr In Fields.Keys
            Set field = Fields(fieldStr)
            GetBitStSp field, stBit, spBit
            ''for non-ECID case
            If (stage And (UCase(GlbUtility.currStage) = UCase(BdfDataBase.GetRealStage(field.BlowLocation)) Or bLumpStages)) Or (includeEarlyFuse And DicEarlyFused.Exists(fieldStr)) Then
                If dicCategory Is Nothing Then
                            For iBit = stBit To spBit
                                dicCared.Element(iBit) = 1
                            Next
                ElseIf dicCategory.Exists(fieldStr) Then
                       If dicCategory(fieldStr) = name Then
                            dicCatCheck(fieldStr) = True
                            '20210714, Modify for Complement Fuse
                            If (gB_CmpFuseEnable) = True And (DicCmpFuse.Exists(fieldStr)) Then
                                GlbUtility.WriteDlg "Do '" + fieldStr + "' complement fuse mask change."
                                If field.DefaultOrReal = dr_default Then
                                    For iBit = stBit To spBit
                                        dicCared.Element(iBit) = DicCmpFuse(fieldStr)(iBit - stBit)
                                        If dicCared.Element(iBit) = 0 Then
                                            GlbUtility.WriteDlg "Bit: " + CStr(iBit) + ",Set to don't check."
                                        End If
                                    Next
                                Else
                                    For iBit = stBit To spBit
                                        dicCared.Element(iBit) = 0
                                        GlbUtility.WriteDlg "Bit: " + CStr(iBit) + ",Set to don't check."
                                    Next
                                End If
                            Else
                                For iBit = stBit To spBit
                                    dicCared.Element(iBit) = 1
                                Next
                            End If
                      End If
                End If
            ElseIf Not dicCategory Is Nothing Then
                If dicCategory.Exists(fieldStr) Then
                    If dicCategory(fieldStr) = name Then
                        dicCatCheck(fieldStr) = True
                        For iBit = stBit To spBit
                            dicCared.Element(iBit) = 1
                        Next
                    End If
                End If
            Else
                GlbUtility.MessageBox "Conditons is ambigious... Please confirm!"
                Call Print_Error_Message(Error_Info, "eFuseBdfBank", "DigCapWaveProcess", "Conditons is ambigious... Please confirm!")
                GoTo errHandler
            End If
        Next
        'inWave(0).Plot
        If Not dicCatCheck Is Nothing Then
            For Each fieldStr In dicCategory.Keys
                If Not dicCatCheck.Exists(fieldStr) And dicCategory(fieldStr) = name And fieldStr <> "" Then
                    GlbUtility.MessageBox "Category specified didn't found! " & fieldStr
                    Call Print_Error_Message(Error_Info, "eFuseBdfBank", "DigCapWaveProcess", "Category specified didn't found! " & fieldStr)
                End If
            Next
        End If
    End If
    
    ''------------------------------------------------------------------------------------
    ''Input before the process
    ''------------------------------------------------------------------------------------
    ''- capWave     => Capture Wave
    ''- dicCared    => StageMaskWave
    ''- DaaParaBits => Parallel bit size of Direct access mode
    ''- pgmMode     => Direct access mode / JTag mod
    ''- nBits       => 2-bit / 1-bit
    ''- FieldIdxSt  => the field's start index
    ''- FieldIdxSp  => the field's end index
    ''- FieldIdxMsb => the field's MSB bit
    ''- g_Rvenable  => DVRV is enable
    ''- RVIndexWave => If the field corresponds with the RV condition,
    ''                 the RVindex set to "1" from the field's lsb bit to msb bit.
    ''                 The size of RVIndexWave is total memory size of the bank.
    ''- RVsegment   => The segment is "1" which corresponds with the RV condition.
    ''                 The size of RVsegment is total field size of the bank.
    ''-------------------------------------------------------------------------------------
    ''Output after the process
    ''-------------------------------------------------------------------------------------
    ''
    ''
    ''-------------------------------------------------------------------------------------
    Set RVsegmentFinal = RVsegment
    Set RVIndexWaveFinal = RVIndexWave
    If g_Rvenable Then
        If MultiTyepRVIndexWave.Exists(MultitypJob) Then
            For Each site In TheExec.sites.Starting
                If MultiTyepRVIndexWave(MultitypJob).CalcSum = 0 Then
                    MultiTyepRVIndexWave(MultitypJob).CreateConstant 0, FullSize, DspLong
                End If
            Next site
            Set RVIndexWaveFinal = MultiTyepRVIndexWave(MultitypJob)
        Else
            For Each site In TheExec.sites.Starting
                If RVIndexWave.CalcSum = 0 Then
                    RVIndexWave.CreateConstant 0, FullSize, DspLong
                End If
            Next site
            Set RVIndexWaveFinal = RVIndexWave
        End If
        
        If MultiTyepRVsegment.Exists(MultitypJob) Then Set RVsegmentFinal = MultiTyepRVsegment(MultitypJob)
    End If

    Testedstagemask.CreateConstant 0, FullSize, DspDouble
    If blankChk Then
        If EarlyTestedStageMaskCheck And Not chkNonEarly Then
            EarlyTestedstagemask = EarlyTestedstagemask.ConvertDataTypeTo(DspDouble)
            EarlyTestedstagemask.NumFormat = FormatUnknown
            Set Testedstagemask = EarlyTestedstagemask.Copy
            TestedWaveCheck = True
            NonFirstfuse = True
        ElseIf PostTestedStageMaskCheck And chkNonEarly Then
            PostTestedstagemask = PostTestedstagemask.ConvertDataTypeTo(DspDouble)
            PostTestedstagemask.NumFormat = FormatUnknown
            Set Testedstagemask = PostTestedstagemask.Copy
            TestedWaveCheck = True
            NonFirstfuse = True
        End If
    End If

    Call rundsp.ProcessCapWaveAndCheck(capWave, dicCared, DaaParaBits, pgmMode, nBits, JtagMsbFirst, _
                                       paraWave, seriWave, seriWave_FL, result, FieldIdxSt, FieldIdxSp, FieldIdxMsb, decWave, _
                                g_Rvenable, RVIndexWaveFinal, RVsegmentFinal, RVOut, TestedWaveCheck, Testedstagemask, capwave_count, CapturePinCnt)
    If blankChk Then
        If Not NonFirstfuse Then
            ' If first fuse, it will check other fileds(exculde current fields) is all 0.
            Call Bank_OtherMaskCheck(seriWave, MultitypJob, chkNonEarly, result)
        Else
            If TestedWaveCheck Then TheExec.Flow.TestLimit resultVal:=capwave_count, lowVal:=1, hiVal:=FullSize, Tname:=name + "_CapWave_Data_Count"
        End If
    End If
    If pgmMode = pgm_DAA Then
        Set DaaCapturedWave = paraWave
        Set DaaCapWaveSerial = seriWave
        Set DaaCapWaveSerialFL = seriWave_FL
    Else
        Set JtagCapturedWave = capWave
        Set JtagCapturedSerial = seriWave
    End If
    Set DecimalWave = decWave
    
    If g_Rvenable Then
        If MultiTyepRVReadWave.Exists(MultitypJob) Then
            Set MultiTyepRVReadWave(MultitypJob) = RVOut
        Else
            Set RVReadWave = RVOut
        End If
    End If

Exit Function
errHandler:
    Call Print_Error_Message(Error_Info, "eFuseBdfBank", "DigCapWaveProcess")
    If AbortTest Then Exit Function Else Resume Next
End Function

Public Sub PrintTapReadFuseDatalog(Optional printdecode As Boolean = False)
On Error GoTo errHandler
Dim fieldStr As Variant, site As Variant, field As eFuseBdfField
Dim stBit As Long, spBit As Long, binstr As String, iBit As Long, printText As String
Dim ST As String, SP As String, LSB As Long, msb As Long
Dim decodeStr As New SiteVariant, DecodeBin As New SiteVariant, belowLoc As String, iPrint As Long
Dim instance As String: instance = TheExec.DataManager.instancename
Dim Serial2Print As New DSPWave, dicPrint As New Dictionary
Dim lotfield As String, wfrfield As String, xcoorfield As String, ycoorfield As String
Dim myStr As Variant, hexecid As String, ttlbits As Long
Dim cfgStr As String

    If Not eFusePrinted Then Exit Sub
    Ze_DspwaveToParallelMsb JtagCapturedSerial, 32, Serial2Print

    DsscCFGCondStr = vbNullString
    FuseRevisionStr = vbNullString

    For Each site In TheExec.sites
        PrintDspWave site, Serial2Print, instance, "SingleBit"
        iPrint = 0: dicPrint.RemoveAll:
                
        If printdecode Then
            myUtility.WriteDlg "====== Print Out EFuse ReadOut Bits ( " & instance & ") (Site" & site & ")============"
            For Each fieldStr In Fields.Keys
                Set field = Fields(fieldStr)
                
                GetBitStSp field, stBit, spBit
                binstr = vbNullString
                For iBit = stBit To spBit
                    binstr = JtagCapturedSerial.ElementLite(iBit) & binstr
                Next
                If Not field.Msb1st Then binstr = StrReverse(binstr)  ''20211123 add for Rhodes FT2 ECID JTAG read ,EFUSE_READ_ECID_BYJTAG = True

                field.DsscValue = myUtility.Bin2HexStr("b" & binstr, field.hBytes)
                If field.size < 32 And Not field.Algorithm = alg_crc Then
                    field.DsscDecValue = myUtility.Hex2Dec(field.DsscValue)
                End If
                                        
                decodeStr = "0": DecodeBin = vbNullString: belowLoc = "   "
                'Hex Infomation Appended
                If DicEcidCat.Exists(fieldStr) Then  ' Ecid decode
                    decodeStr = Decode(field, True)
                Else
                    Select Case field.Algorithm  ' IDS/VDDBin/Base decode
                        Case alg_ids:
                            decodeStr = myUtility.TxtFmt(field.IdsValue_A * 1000 & " mA", 10) 'to mA
                        Case alg_vddbin:
                            decodeStr = myUtility.TxtFmt(field.VddBinValue_V & " mV", 10)  ' 'to mV
                        Case alg_base:
                            decodeStr = myUtility.TxtFmt(field.BaseValue_V * 1000 & " mV", 10) ' 'to mV
                        Case alg_cond:
                            cfgStr = StrReverse(binstr)
                            DsscCFGCondStr = DsscCFGCondStr + cfgStr
                    Case Else:
                            If field.size < 32 And Not field.Algorithm = alg_crc Then
                                decodeStr = myUtility.TxtFmt(field.DsscDecValue, field.size)
                            Else
                                decodeStr = myUtility.TxtFmt(field.DsscValue, 10)
                            End If
                     End Select
                End If

                If LCase(field.name) = "fuse_revision" Then
                    FuseRevisionStr = binstr
                End If
                                          
                'Binary Infomation Appended
                If field.size < 32 And field.Algorithm <> alg_crc Then '20210202
                    DecodeBin = myUtility.Dec2Bin(field.DsscDecValue, field.size)
                Else
                    DecodeBin = myUtility.xHex2BinStr(field.DsscValue, field.size)
                End If
                
                If Not field.Msb1st Then DecodeBin = StrReverse(DecodeBin)
                printText = decodeStr
                dicPrint.Add iPrint, "Site(" & site & DicReadPrintIdx(iPrint) & "=" & myUtility.TxtFmt(printText, 10) & " [" & DecodeBin & "] "
                iPrint = iPrint + 1
                
                If HadDeidFuse Then
                    If myUtility.IsStrMatch(CStr(fieldStr), "lot.*id") Then lotfield = CStr(fieldStr):
                    If myUtility.IsStrMatch(CStr(fieldStr), "wafer.*id") Then wfrfield = CStr(fieldStr):
                    If myUtility.IsStrMatch(CStr(fieldStr), "x.*coor") Then xcoorfield = CStr(fieldStr):
                    If myUtility.IsStrMatch(CStr(fieldStr), "y.*coor") Then ycoorfield = CStr(fieldStr)
                End If
            Next
            
            ' Print the read out result string
            For Each myStr In dicPrint.Keys
                myUtility.WriteDlg dicPrint(myStr)
            Next
            
            If HadDeidFuse Then
                ttlbits = 0
                Set field = Fields(lotfield):      hexecid = myUtility.Hex2BinStr(field.DsscValue, field.size): ttlbits = ttlbits + field.size
                HramLotId = Decode(field, True)
                Set field = Fields(wfrfield):     hexecid = hexecid & myUtility.Hex2BinStr(field.DsscValue, field.size): ttlbits = ttlbits + field.size
                HramWaferId = field.DsscDecValue
                Set field = Fields(xcoorfield): hexecid = hexecid & myUtility.Hex2BinStr(field.DsscValue, field.size): ttlbits = ttlbits + field.size
                HramXCoord = field.DsscDecValue
'                If (myUtility.currStage = "FT1" Or myUtility.currStage = "FT2" Or myUtility.currStage = "FT3") And (UCase(name) Like "ECID") Then
                If UCase(TheExec.CurrentChanMap) Like "CHANNELMAP_FT*" And (UCase(name) Like "ECID") Then 'JackChou 202404 To allow Stage_FT with ChannelMap_CP case [Requester-TSMC-KoPei]
                    Call TheExec.Datalog.Setup.WaferSetup.SetXCoord(site, myUtility.Hex2Dec(field.DsscValue))
                    XCoord = HramXCoord
                End If
                Set field = Fields(ycoorfield): hexecid = hexecid & myUtility.Hex2BinStr(field.DsscValue, field.size): ttlbits = ttlbits + field.size
                HramYCoord = field.DsscDecValue
'                If (myUtility.currStage = "FT1" Or myUtility.currStage = "FT2" Or myUtility.currStage = "FT3") And (UCase(name) Like "ECID") Then
                If UCase(TheExec.CurrentChanMap) Like "CHANNELMAP_FT*" And (UCase(name) Like "ECID") Then 'JackChou 202404 To allow Stage_FT with ChannelMap_CP case [Requester-TSMC-KoPei]
                    Call TheExec.Datalog.Setup.WaferSetup.SetYCoord(site, myUtility.Hex2Dec(field.DsscValue))
                    YCoord = HramYCoord
                End If
                EcidHexStr = Replace(myUtility.Bin2HexStr(StrReverse(hexecid), 16), "0x", "") '(ttlbits - 1) \ 4 + 1
            End If
        End If
   Next

Exit Sub
errHandler:
    Call Print_Error_Message(Error_Info, "eFuseBdfBank", "PrintTapReadFuseDatalog")
    If AbortTest Then Exit Sub Else Resume Next
End Sub

Public Function CalculateCRCComp(Optional specField As String = vbNullString, Optional getValOnly As Boolean = False) As String
On Error GoTo errHandler
Dim item As Variant, field As eFuseBdfField, targetfield As eFuseBdfField
Dim CrcCompBitsStr As String
Dim iChr As Long, BinBit As String, crccompStr As String
Dim binstr As New SiteVariant
Dim site As Variant 'Carter, 20240304

    If Crcs_Comp.Count = 0 Then
        CalculateCRCComp = vbNullString
        If HadCrcCompFuse Then myUtility.MessageBox "CRC_comp field didn't put the target field info. in this bank!": CalculateCRCComp = "": Exit Function
        myUtility.WriteDlg "No CRC_comp field in this bank! Bank=""" + name + """": CalculateCRCComp = False: Exit Function
    End If
                                                                                                                                         'ProfileMarkLeave eFuseBdfBank_ProfileMark_513    ' Profile Mark
    CalculateCRCComp = vbNullString
    For Each item In Crcs_Comp.Keys  ' field name => dictionary (bit, true)
        If specField = vbNullString Or LCase(item) = LCase(specField) Then
            Set field = Fields(item)
            Set targetfield = Fields(Crcs_Comp(item))
            
            If ((LCase(myUtility.currStage) = LCase(BdfDataBase.GetRealStage(field.BlowLocation)) And Not field.Trimmed) Or bLumpStages) Or getValOnly Then
                If field.DefaultOrReal = dr_real Then
                    If Not getValOnly Then
                        binstr = myUtility.Hex2BinStr(targetfield.TrimAteValue, targetfield.size)
                    Else
                        binstr = myUtility.Hex2BinStr(targetfield.DsscValue, targetfield.size)
                    End If
                    
                    For Each site In TheExec.sites
                        CrcCompBitsStr = vbNullString
                        For iChr = 0 To Len(binstr) - 1
                            BinBit = mid(binstr, iChr + 1, 1)
                            If BinBit = "1" Then
                                BinBit = "0"
                            Else
                                BinBit = "1"
                            End If
                            CrcCompBitsStr = CrcCompBitsStr + BinBit
                        Next iChr
                        crccompStr = myUtility.Bin2HexStr(CrcCompBitsStr, field.hBytes)
                        CalculateCRCComp = crccompStr
                        
                        If Not getValOnly Then
                             field.TrimAteValue = crccompStr
                             field.TrimFlagBySite = True
                             field.TrimTestInstance = TheExec.DataManager.instancename
                             field.RealSetByDefault = False
                        End If
                        If eFusePrinted Then
                            myUtility.WriteDlg "Site" & site & " bnak(" & item & ") CRC One Complement=" & crccompStr & " ,Input Bits = " & binstr
                        End If
                    Next site
                Else
                    Call Print_Error_Message(Warning_Info, "eFuseBdfBank", "CalculateCRCComp", "bnak(" & item & ") skip set, it's CRC One Complement is not REAL! ")
                End If 'If Not field.DefaultOrReal Then
            Else
                Call Print_Error_Message(Warning_Info, "eFuseBdfBank", "CalculateCRCComp", "bnak(" & item & ") skip set, it's CRC One Complement  at stage >> " & field.BlowLocation)
            End If
        End If ' If specField = "" Or item = specField Then
    Next

Exit Function
errHandler:
    Call Print_Error_Message(Error_Info, "eFuseBdfBank", "CalculateCRCComp")
    If AbortTest Then Exit Function Else Resume Next
End Function

Public Function IsExistMultiTypeProgramStage(inputStage As String) As Boolean
On Error GoTo errHandler

    If inputStage = "" Then Exit Function
    If MultiTypeJobList.Count <> 0 And MultiTypeJobList.Exists(inputStage) Then
        IsExistMultiTypeProgramStage = True
    Else
        IsExistMultiTypeProgramStage = False
        Call Print_Error_Message(Error_Info, "eFuseBdfBank", "IsExistMultiTypeProgramStage", "BDF table not exist multi-type program stage but instance argument(MultitypJob) is not empty!!!")
    End If

Exit Function
errHandler:
    Call Print_Error_Message(Error_Info, "eFuseBdfBank", "IsExistMultiTypeProgramStage")
    If AbortTest Then Exit Function Else Resume Next
End Function

Public Function IsMultiTypInst(instNameTyp As String, ByRef argVal As String) As Boolean
On Error GoTo errHandler

    IsMultiTypInst = False
    If instNameTyp <> "" And argVal <> "" Then
        If instNameTyp = argVal Then
            argVal = UCase(GlbUtility.currStage) + "_" + instNameTyp
            IsMultiTypInst = IsExistMultiTypeProgramStage(argVal)
        End If
    End If

    If Not IsMultiTypInst Then
        Call Print_Error_Message(Error_Info, "eFuseBdfBank", "IsMultiTypInst", "For Multi-type programming function, instance name and argument mismatch")
    End If

Exit Function
errHandler:
    Call Print_Error_Message(Error_Info, "eFuseBdfBank", "IsMultiTypInst")
    If AbortTest Then Exit Function Else Resume Next
End Function

Public Function IsExistMultiTypeKey(InstName As String, ByRef key As String) As Boolean
On Error GoTo errHandler
Dim keyword As String: keyword = vbNullString
Dim InstNameSegs() As String

    IsExistMultiTypeKey = False

    InstNameSegs = Split(InstName, "_")
    If MultiTypeJobList.Exists(UCase(GlbUtility.currStage) + "_" + InstNameSegs(UBound(InstNameSegs))) Then
        key = InstNameSegs(UBound(InstNameSegs))
    ElseIf MultiTypeJobList.Exists(UCase(GlbUtility.currStage) + "_" + InstNameSegs(UBound(InstNameSegs) - 1)) Then
        key = InstNameSegs(UBound(InstNameSegs) - 1)
    End If
    If key <> "" Then IsExistMultiTypeKey = True

Exit Function
errHandler:
    Call Print_Error_Message(Error_Info, "eFuseBdfBank", "IsExistMultiTypeKey")
    If AbortTest Then Exit Function Else Resume Next
End Function

Public Function Bank_OtherMaskCheck(seriWave As DSPWave, MultitypJob As String, chkNonEarly As Boolean, result As SiteLong)
On Error GoTo errHandler
Dim OtherMask As New DSPWave
Dim failBits As New SiteLong
Dim site As Variant

    OtherMask.CreateConstant 1, FullSize, DspLong
    ' Select which othermask
    If Not chkNonEarly Then
        OtherMask = EarlyStageMask.BitwiseXor(OtherMask)
    Else
        If MultitypJob = Empty Then
            ' Check BlkChk_LV/HV of multijobs is same test condition
            If HVLV_MultiJob = Empty Then
                OtherMask = StageMask.BitwiseXor(OtherMask)
            Else
                NonFirstfuse = True
                Exit Function
            End If
        Else
            If HVLV_MultiJob = Empty Then HVLV_MultiJob = MultitypJob
            
            OtherMask = StageMaskMultiTyp(MultitypJob).BitwiseXor(OtherMask)
            
            If HVLV_MultiJob <> MultitypJob Then
                NonFirstfuse = True
                Exit Function
            End If
        End If
    End If
    
    Call rundsp.GetDspWaveFailbits(seriWave, OtherMask, failBits)
    
    For Each site In TheExec.sites
        If result(site) > 0 Then failBits(site) = 0
    Next site
    TheExec.Flow.TestLimit resultVal:=failBits, lowVal:=0, hiVal:=0, Tname:="Failbits_Count"

Exit Function
errHandler:
    Call Print_Error_Message(Error_Info, "eFuseBdfBank", "Bank_OtherMaskCheck")
    If AbortTest Then Exit Function Else Resume Next
End Function

Public Function GetBinCheckDsscWave() As DSPWave
On Error GoTo errHandler
Dim eFuseWave As New DSPWave, iBit As Long
Dim fieldStr As Variant, field As eFuseBdfField
Dim stBit As Long, spBit As Long, bitStep As Long, binA() As New SiteLong, bitDealt As Long
Dim site As Variant 'Carter, 20240304

    eFuseWave.CreateConstant 0, FullSize
    
    For Each fieldStr In Fields
        Set field = Fields(fieldStr)
        If (field.DefaultOrReal = dr_real) Then
            bitDealt = field.size - 1
            stBit = field.LSB: spBit = field.msb: bitStep = 1: ' MSB 1st
            If field.Msb1st Then
                 GetBitStSp field, stBit, spBit
            ElseIf stBit > spBit Then
                 bitStep = -1: 'LSB first, if doing CRC, it needs to use the MSB first
            Else
                 myUtility.MessageBox "Something wrong in dealing BDF GetDspWave"
                 Call Print_Error_Message(Error_Info, "eFuseBdfBank", "GetBinchkDsscWave", "Something wrong in dealing BDF GetDspWave")
            End If
            
            bitDealt = 0
            binA = GetBinA(DataRetrieve.AteTrimData, field)
            For iBit = stBit To spBit Step bitStep
                For Each site In TheExec.sites.Selected
                    eFuseWave.ElementLite(iBit) = binA(bitDealt)
                Next
                bitDealt = bitDealt + 1
            Next
        End If
    Next
    
    Set GetBinCheckDsscWave = eFuseWave
    
Exit Function
errHandler:
    Call Print_Error_Message(Error_Info, "eFuseBdfBank", "GetBinCheckDsscWave")
    If AbortTest Then Exit Function Else Resume Next
End Function

Public Sub InitialMultiCheckVarVariable()
On Error GoTo errHandler

    IsSameCheckVar = 1
    BlankCheckCount = 0

Exit Sub
errHandler:
    Call Print_Error_Message(Error_Info, "eFuseBdfBank", "InitialMultiCheckVarVariable")
    If AbortTest Then Exit Sub Else Resume Next
End Sub
