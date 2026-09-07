Attribute VB_Name = "VBT_LIB_SRM"
Option Explicit

Private Const TL_E_AT_PATSET_BREAKPT = &HC0000014
''==============================================================================================
Public Type PinsInfo
    PinsName As String
    ChanMapType As String
    Init_step As Long
    MergedN As Long
    CheckListYIndex As Long
    Range_List() As Double
    Accuracy_List() As Double
    WaitTime_List() As Double
    Init_CurrentRange As Double
    Scenario As Double
    VMain As String
    VAlt As String
    SramPair As String
    LogicPair() As String
End Type

Public Type SrmPatterns
    BlockType As String
    PatStr As String
    Patterns() As String
End Type

Public Type SelSrmPinsInfo
    CurrentSRMPat As String
    CurrentSRMData As String
    CurrentSRMReadPat As String
    SELSRM_BLOCK As String
    CapBitSize() As Long
    SrcBitSize() As Long
    SrcBits() As String
    TestCase() As String
    ReadLimit() As String
    HighLowLimit As Double
    FIND_AllHEADER As Boolean
    SramPins() As PinsInfo
    LogicPins() As PinsInfo
    AllPins() As PinsInfo
    SrmPattern() As SrmPatterns
    SelSrmDic As New dictionary
    TempDicSRM As New dictionary
    Check() As Integer
    CheckCount As Integer
End Type

Public Type AutoRangeInfo
    MergedN As Long
    Init_step As Long
    PinName As String
    MergeType As String
    PinMapType As String
    ChanMapType As String
    hiLimit As Double
    PowerSeq As Double
    PowerDownSeq As Double
    Range_List() As Double
    Accuracy_List() As Double
    WaitTime_List() As Double
    Init_CurrentRange As Double
End Type

Public Enum SelSrmCnd
    LHSL_SRAMPAT0_DSRAMPAT1 = 1
    LHSL_SRAMPAT1_DSRAMPAT0 = 2
    LLSH_SRAMPAT0_DSRAMPAT1 = 3
    LLSH_SRAMPAT1_DSRAMPAT0 = 4
End Enum

Public Enum SelSrmType
    S_SC_HC             'default 0
    S_SC_DSSC           '1
    S_BI_H              '2
    S_BI_DSSC           '3
    C_SC_HC             '4
    C_SC_DSSC           '5
    C_BI_HC             '6
    C_BI_DSSC           '7
    L_SC_HC             '8
    L_SC_DSSC           '9
    L_BI_HC             '10
    L_BI_DSSC           '11
    S_SSBBIST_DSSC      '12
    LGIndex             '13
    SRMPins             '14
    Write_data          '15
    Read_data           '16
    Src_Cap_bits        '17
End Enum

Public Enum ChnMapHeader
    Enum_PinName = 1
    Enum_PinType = 2
End Enum
''==============================================================================================
Public Const gS_SELSRM_SHEET = "SelSram_ChkList"
Public SRMCheckListInfo As SelSrmPinsInfo
Public glbArr_AllPinInfo() As AutoRangeInfo
Public glbArr_PowerPinInfo() As AutoRangeInfo
Public glbInstanceName As String

Public Const glbLng_PatCnt = 14
Public Const glbConstVar_PatLoopCnt = "SelSrm_LP_Var"
Public Const glbConstVar_PatLoopCntEnd = "LPCount_End"
Public Const glbConstVar_TestCaseCnt = "SEL_CASE_CNT"
Public Const glbConstVar_TestCaseCntEnd = "SEL_CASE_End"

Private Const glbConst_LOGICPIN = "LOGIC_PIN"
Private Const glbConst_SRAMPINS = "SRAM_PINS"
Private Const glbConst_WRITE = "WRITE"
Private Const glbConst_READ = "READ"
Private Const glbConst_SRCCAPBITS = "SRC_CAP_BITS"
Private Const glbConst_CHECK = "CHECK"
Private Const glbConst_COMMENT = "COMMENT"

Private Const glbConstKeyWD_NoConnected = "N/C"
Private Const glbConstKeyWD_Empty = ""
Private Const glbConstKeyWD_POWER = "*POWER*"
Private Const glbConstKeyWD_DCVS = "*DCVS*"
Private Const glbConstKeyWD_DCVI = "*DCVI*"
Private Const glbConstKeyWD_DCVSMERGED = "*DCVSMERGED*"
Private Const glbConstKeyWD_DCVIMERGED = "*DCVIMERGED*"

''==============================================================================================
''Public Methods
''Perform a digital functional test.
''Return TL_SUCCESS if the test executes without problems, else TL_ERROR.
''==============================================================================================
Public Function Functional_T_SRM(Optional PatternTimeout As String = "30", _
                                    Optional Step_ As SubType, _
                                    Optional Init_Pat1 As Pattern, Optional Init_Pat2 As Pattern, _
                                    Optional Init_Pat3 As Pattern, Optional Init_Pat4 As Pattern, _
                                    Optional Init_Pat5 As Pattern, Optional Init_Pat6 As Pattern, _
                                    Optional Init_Pat7 As Pattern, Optional Init_Pat8 As Pattern, _
                                    Optional Init_Pat9 As Pattern, Optional Init_Pat10 As Pattern, _
                                    Optional Payload_Pat1 As Pattern, _
                                    Optional Payload_Pat2 As Pattern, _
                                    Optional Payload_Pat3 As Pattern, _
                                    Optional Payload_Pat4 As Pattern, _
                                    Optional Payload_Pat5 As Pattern, _
                                    Optional DSSC_BitSize As String, _
                                    Optional DSSC_Seg As String, _
                                    Optional DSSC_SrcPin As String, _
                                    Optional DSSC_EQ As String, _
                                    Optional CGVoltage As String, _
                                    Optional Wait As String, _
                                    Optional SRMType As String, _
                                    Optional Validating_ As Boolean = True, _
                                    Optional Apply_Flag As Boolean = False) As Long
On Error GoTo errHandler 'Add ErrHandler 2023/08/18
''==============================================================================================
''Pre-Load Patterns when validation
''==============================================================================================
    Functional_T_SRM = TL_SUCCESS   ' be optimistic
    If Not TheExec.Flow.IsRunning Then Exit Function

    Call SRM_CoustomSetting
    If Validating_ Then
        Dim PatStr As String:: PatStr = ""
        Dim patArr() As String
        Dim i As Long
        
        PatStr = SRM_ArgsToStr(Init_Pat1, Init_Pat2, Init_Pat3, Init_Pat4, Init_Pat5, Init_Pat6, Init_Pat7, Init_Pat8, Init_Pat9, Init_Pat10, Payload_Pat1, Payload_Pat2, Payload_Pat3, Payload_Pat4, Payload_Pat5)
        patArr = Split(PatStr, ",")
        For i = 0 To UBound(patArr)
            If patArr(i) <> "" Then Call PrLoadPattern(patArr(i))
        Next
        Exit Function
    End If
''==============================================================================================
''PreBody, Apply default values to parameters whose values were not specified.
''==============================================================================================
    glbInstanceName = TheExec.DataManager.instanceName

    If Step_ = subAllBody Or Step_ = subPrebody Then
        If Not Apply_Flag Then TheHdw.Digital.ApplyLevelsTiming True, True, True, tlPowered
    End If
''==============================================================================================
''Body
''==============================================================================================
    If Step_ = subAllBody Or Step_ = subBody Then
        If TheExec.sites.ActiveCount > 0 Then
            Dim SRMLib As New Class_SRMLib
            Call SRMLib.Initialize(Init_Pat1, Init_Pat2, Init_Pat3, Init_Pat4, Init_Pat5, Init_Pat6, Init_Pat7, Init_Pat8, Init_Pat9, Init_Pat10, Payload_Pat1, Payload_Pat2, Payload_Pat3, Payload_Pat4, Payload_Pat5, _
                                Wait, DSSC_BitSize, DSSC_Seg, DSSC_SrcPin, DSSC_EQ, CGVoltage, SRMType, PatternTimeout)

            Select Case UCase(SRMType)
                Case "DSSC":::::: SRMLib.SRMViaDSSC
                Case "STATIC":::: SRMLib.SRMViaPAT
                Case "CUSTOM":::: SRMLib.CUSTOM
                Case "SRMEXP":::: SRMLib.SRMEXPLH
                Case "SRMREAD"::: SRMLib.SRMREAD
                Case "HARDIP":::: SRMLib.HARDIP
                Case Else: GoTo errHandler
            End Select
        End If
    End If
''==============================================================================================
''PostBody
''==============================================================================================
    If Step_ = subAllBody Or Step_ = subPostbody Then
        'Add your VBT code here
    End If
''==============================================================================================
Exit Function
errHandler:
    Call Print_Error_Message(Error_Info, "VBT_LIB_SRM", "Functional_T_SRM") 'Add ErrHandler 2023/08/18
    Call TheExec.ErrorReport
    Functional_T_SRM = TL_ERROR
    If AbortTest Then Exit Function Else Resume Next
End Function
''==============================================================================================
''If Coustom wants to add something, add VBT here.
''==============================================================================================
Public Function SRM_CoustomSetting()
On Error GoTo errHandler
    Dim funcName As String:: funcName = "SRM_CoustomSetting"

Exit Function
errHandler:
    Call Print_Error_Message(Error_Info, "VBT_LIB_SRM", "SRM_CoustomSetting") 'Add ErrHandler 2023/08/18
    If AbortTest Then Exit Function Else Resume Next
End Function
''==============================================================================================
''Parsing Sheet data.
''==============================================================================================
'[20231106][T-Tah][Oliver] add multiple SRAM pin compares with one logic pin method
Public Function SRM_ParseChkList(ShtName As String)
On Error GoTo errHandler
    Dim funcName As String: funcName = "SRM_ParseChkList"
''==============================================================================================
''Declare variables
''==============================================================================================
    Dim MySheet As Worksheet:: Set MySheet = Nothing
    Dim MaxColumn As Long, MaxRow As Long
    Dim CurXIndex As Long, CurYIndex As Long
    Dim CurColumn As Long, CurRow As Long, i As Long
    Dim PatCount As Long, TempCnt As Long
    Dim SRMIndex As Long, SRMPIndex As Long, LGPIndex As Long, WDIndex As Long, RDIndex As Long, SCBIndex As Long, CMTIndex As Long, CheckIndex As Long
    Dim PreLoadPats As String: PreLoadPats = ""
    Dim PreLoadAllPats As String: PreLoadAllPats = ""
    Dim SrmPatNotExistStr As String, SrmPatNotExistArr() As String
    Dim TempArr() As String, TestCaseArr() As String, PattArray() As String, LimitArr() As String
    Dim SrmChkListArr As Variant, SrmChkListArrTemp() As Variant
    
    SRMCheckListInfo.FIND_AllHEADER = False
''==============================================================================================
''Parsing Sheet data to SrmChkListArr.
''==============================================================================================
'    ShtName = gS_SELSRM_SHEET
    If SRM_FindSheet(ShtName) Then
        Set MySheet = ActiveWorkbook.Sheets(ShtName)
        MySheet.Select
        MaxRow = MySheet.UsedRange.Rows.Count
        MaxColumn = MySheet.UsedRange.Columns.Count
        SrmChkListArr = MySheet.range(Cells(1, 1), Cells(MaxRow, MaxColumn)).value
        ReDim SrmChkListArrTemp(MaxColumn - 1, MaxRow - 1)
''==============================================================================================
''Transfer SrmChkListArr data to SrmChkListArrTemp
''==============================================================================================
        For CurColumn = 1 To MaxColumn
            For CurRow = 1 To MaxRow
                SrmChkListArrTemp(CurColumn - 1, CurRow - 1) = SrmChkListArr(CurRow, CurColumn)
            Next CurRow
        Next CurColumn
''==============================================================================================
''Check FIND_AllHEADER state and Create SRMCheckListInfo.SelSrmDic
''==============================================================================================
        'SRMCheckListInfo.FIND_AllHEADER = False
        If SRMCheckListInfo.FIND_AllHEADER = False Then
            SRM_CreatePatSetAllDic
            SRMCheckListInfo.SelSrmDic.RemoveAll
            SRMCheckListInfo.TempDicSRM.RemoveAll
            For CurColumn = 1 To MaxColumn
                If SrmChkListArr(1, CurColumn) = "" Then
                    Exit For
                Else
                    If Not SRMCheckListInfo.SelSrmDic.Exists(UCase(Trim(SrmChkListArr(1, CurColumn)))) Then
                        SRMCheckListInfo.SelSrmDic.Add UCase(Trim(SrmChkListArr(1, CurColumn))), CurColumn
                    End If
                End If
            Next
            If SRMCheckListInfo.SelSrmDic.Exists(glbConst_COMMENT) Then
                SRMCheckListInfo.FIND_AllHEADER = True
            Else
                TheExec.Datalog.WriteComment "Please check sheet 'SelSram_ChkList' Header"
                Stop
            End If
''==============================================================================================
''Assign Index for each Header
''==============================================================================================
            SRMPIndex = SRMCheckListInfo.SelSrmDic(glbConst_SRAMPINS) - 1
            LGPIndex = SRMCheckListInfo.SelSrmDic(glbConst_LOGICPIN) - 1
            WDIndex = SRMCheckListInfo.SelSrmDic(glbConst_WRITE) - 1
            RDIndex = SRMCheckListInfo.SelSrmDic(glbConst_READ) - 1
            CheckIndex = SRMCheckListInfo.SelSrmDic(glbConst_CHECK) - 1
            SCBIndex = SRMCheckListInfo.SelSrmDic(glbConst_SRCCAPBITS) - 1
            CMTIndex = SRMCheckListInfo.SelSrmDic(glbConst_COMMENT) - 1
''==============================================================================================
''Pasing and preload All Patterns from SelSrm_ChkList sheet.
''==============================================================================================
            ReDim SRMCheckListInfo.SrmPattern(LGPIndex)
            For CurXIndex = 0 To LGPIndex - 1
                If Not SrmChkListArrTemp(CurXIndex, 1) = "" Then
                    ReDim Preserve SRMCheckListInfo.SrmPattern(CurXIndex)
                    SRMCheckListInfo.SrmPattern(CurXIndex).BlockType = UCase(Trim(SrmChkListArrTemp(CurXIndex, 0)))
                    For CurYIndex = 1 To MaxRow - 1
                        If SrmChkListArrTemp(CurXIndex, CurYIndex) = "" Then
                            Exit For
                        Else
                            If Not SRM_PatPreCheck(UCase(Trim(SrmChkListArrTemp(CurXIndex, CurYIndex)))) Then
                                SrmPatNotExistStr = SrmPatNotExistStr & UCase(Trim(SrmChkListArrTemp(CurXIndex, CurYIndex))) & ","
                            Else
                                ReDim Preserve SRMCheckListInfo.SrmPattern(CurXIndex).Patterns(CurYIndex - 1)
                                If SRMCheckListInfo.SrmPattern(CurXIndex).BlockType Like "*DSSC*" Then
                                    Call PATT_GetPatListFromPatternSet(UCase(Trim(SrmChkListArrTemp(CurXIndex, CurYIndex))), PattArray, PatCount)
                                    PreLoadPats = PreLoadPats & PattArray(0) & ","
                                Else
                                    PreLoadPats = PreLoadPats & UCase(Trim(SrmChkListArrTemp(CurXIndex, CurYIndex))) & ","
                                End If
                                SRMCheckListInfo.SrmPattern(CurXIndex).Patterns(CurYIndex - 1) = UCase(Trim(SrmChkListArrTemp(CurXIndex, CurYIndex)))
                            End If
                        End If
                    Next
                    SRMCheckListInfo.SrmPattern(CurXIndex).PatStr = PreLoadPats
                    PreLoadAllPats = PreLoadAllPats & PreLoadPats:: PreLoadPats = ""
                Else
                    ReDim Preserve SRMCheckListInfo.SrmPattern(CurXIndex)
                    ReDim Preserve SRMCheckListInfo.SrmPattern(CurXIndex).Patterns(0)
                    SRMCheckListInfo.SrmPattern(CurXIndex).BlockType = UCase(Trim(SrmChkListArrTemp(CurXIndex, 0)))
                End If
            Next

            If Not SrmPatNotExistStr = "" Then
                SrmPatNotExistStr = Replace(left(SrmPatNotExistStr, Len(SrmPatNotExistStr) - 1), ",", vbCrLf)
                MsgBox "Did not find Patterns in the PatSets_All sheet." & vbCrLf & SrmPatNotExistStr, vbOKOnly + vbCritical, "Error: Selsrm Patterns PreCheck!!!":: SrmPatNotExistStr = ""
                Stop
                Exit Function
            Else
                If PreLoadAllPats <> "" Then
                    TheExec.AddOutput "Please Wait, PreLoad ALL SelSrm Patterns. ", vbRed, True
                    TheExec.Datalog.WriteComment "*******************************************************************"
                    TheExec.Datalog.WriteComment "*            PLEASE WAIT, PRELOAD ALL SELSRAM PATTERNS.           *"
                    TheExec.Datalog.WriteComment "*******************************************************************"
                    If TheExec.TesterMode = testModeOnline Then
                        TheHdw.Patterns(PreLoadAllPats).Load:: PreLoadAllPats = ""
                    End If
                End If
            End If
''==============================================================================================
''Pasing "Logic_Pins" from SelSrm_ChkList sheet.
''==============================================================================================
            Dim LastCount As Integer
            LastCount = 0
            For CurYIndex = 1 To MaxRow - 1
                If SrmChkListArrTemp(LGPIndex, CurYIndex) = "" Then
                    Exit For
                Else
                    Dim PinTemp() As String
                    Dim PinIndex As Integer
                    Dim PinSize As Integer
                    
                    PinTemp = Split(SrmChkListArrTemp(LGPIndex, CurYIndex), ",")
                    PinSize = UBound(PinTemp) + LastCount + 1
                    ReDim Preserve SRMCheckListInfo.LogicPins(PinSize - 1)
                    ReDim Preserve SRMCheckListInfo.AllPins(PinSize - 1)
                    
                    For PinIndex = LastCount To PinSize - 1
                        SRMCheckListInfo.LogicPins(PinIndex).PinsName = UCase(Trim(PinTemp(PinIndex - LastCount)))
                        SRMCheckListInfo.AllPins(PinIndex).PinsName = UCase(Trim(PinTemp(PinIndex - LastCount)))
                        SRMCheckListInfo.LogicPins(PinIndex).SramPair = UCase(Trim(SrmChkListArrTemp(SRMPIndex, CurYIndex)))
                        SRMCheckListInfo.AllPins(PinIndex).Scenario = 0
                        SRMCheckListInfo.LogicPins(PinIndex).CheckListYIndex = CurYIndex
                    Next PinIndex
                    LastCount = PinSize
                End If
            Next
''==============================================================================================
''Pasing "SRAM_Pins" from SelSrm_ChkList sheet.
''==============================================================================================
            For CurYIndex = 1 To MaxRow - 1
                If SrmChkListArrTemp(SRMPIndex, CurYIndex) = "" Then
                    Exit For
                Else
                    '20230601 split sram pins by "," for Tahiti
                    Dim SrmPinTemp() As String
                    Dim SrmTempIndex As Integer
                    Dim SrmTempSize As Integer

                    SrmPinTemp = Split(SrmChkListArrTemp(SRMPIndex, CurYIndex), ",")
                    SrmTempSize = UBound(SrmPinTemp)
                    For SrmTempIndex = 0 To SrmTempSize
                        If Not SRMCheckListInfo.TempDicSRM.Exists(UCase(Trim(SrmPinTemp(SrmTempIndex)))) Then
                            SRMCheckListInfo.TempDicSRM.Add UCase(Trim(SrmPinTemp(SrmTempIndex))), SRMCheckListInfo.TempDicSRM.Count + 1
                            ReDim Preserve SRMCheckListInfo.SramPins(SRMCheckListInfo.TempDicSRM.Count - 1)
                            ReDim SRMCheckListInfo.SramPins(SRMCheckListInfo.TempDicSRM.Count - 1).LogicPair(0)
                            SRMCheckListInfo.SramPins(SRMCheckListInfo.TempDicSRM.Count - 1).PinsName = UCase(Trim(SrmPinTemp(SrmTempIndex)))
                            SRMCheckListInfo.SramPins(SRMCheckListInfo.TempDicSRM.Count - 1).LogicPair(0) = UCase(Trim(SrmChkListArrTemp(LGPIndex, CurYIndex)))
                            SRMCheckListInfo.SramPins(SRMCheckListInfo.TempDicSRM.Count - 1).Scenario = 0
                            
                            ReDim Preserve SRMCheckListInfo.AllPins(UBound(SRMCheckListInfo.AllPins) + 1)
                            SRMCheckListInfo.AllPins(UBound(SRMCheckListInfo.AllPins)).PinsName = UCase(Trim(SrmPinTemp(SrmTempIndex)))
                            SRMCheckListInfo.AllPins(UBound(SRMCheckListInfo.AllPins)).Scenario = 0
                        Else
                            SRMIndex = SRMCheckListInfo.TempDicSRM(UCase(Trim(SrmPinTemp(SrmTempIndex)))) - 1
                            TempCnt = UBound(SRMCheckListInfo.SramPins(SRMIndex).LogicPair) + 1
                            ReDim Preserve SRMCheckListInfo.SramPins(SRMIndex).LogicPair(TempCnt)
                            SRMCheckListInfo.SramPins(SRMIndex).LogicPair(TempCnt) = UCase(Trim(SrmChkListArrTemp(LGPIndex, CurYIndex)))
                        End If
                    Next SrmTempIndex

                    
                    'SRMCheckListInfo.LogicPins(CurYIndex - 1).SramPair = UCase(Trim(SrmChkListArrTemp(SRMPIndex, CurYIndex)))
'                    If Not SRMCheckListInfo.TempDicSRM.Exists(UCase(Trim(SrmChkListArrTemp(SRMPIndex, CurYIndex)))) Then 'pin
'                        SRMCheckListInfo.TempDicSRM.Add UCase(Trim(SrmChkListArrTemp(SRMPIndex, CurYIndex))), SRMCheckListInfo.TempDicSRM.Count + 1  'pin
'                        ReDim Preserve SRMCheckListInfo.SramPins(SRMCheckListInfo.TempDicSRM.Count - 1)
'                        ReDim SRMCheckListInfo.SramPins(SRMCheckListInfo.TempDicSRM.Count - 1).LogicPair(0)
'                        SRMCheckListInfo.SramPins(SRMCheckListInfo.TempDicSRM.Count - 1).PinsName = UCase(Trim(SrmChkListArrTemp(SRMPIndex, CurYIndex)))
'                        SRMCheckListInfo.SramPins(SRMCheckListInfo.TempDicSRM.Count - 1).LogicPair(0) = UCase(Trim(SrmChkListArrTemp(LGPIndex, CurYIndex)))
'                        SRMCheckListInfo.SramPins(SRMCheckListInfo.TempDicSRM.Count - 1).Scenario = 0
'
'                        ReDim Preserve SRMCheckListInfo.AllPins(UBound(SRMCheckListInfo.AllPins) + 1)
'                        SRMCheckListInfo.AllPins(UBound(SRMCheckListInfo.AllPins)).PinsName = UCase(Trim(SrmChkListArrTemp(SRMPIndex, CurYIndex)))
'                        SRMCheckListInfo.AllPins(UBound(SRMCheckListInfo.AllPins)).Scenario = 0
'                    Else
'                        SRMIndex = SRMCheckListInfo.TempDicSRM(SRMCheckListInfo.LogicPins(CurYIndex - 1).SramPair) - 1
'                        TempCnt = UBound(SRMCheckListInfo.SramPins(SRMIndex).LogicPair) + 1
'                        ReDim Preserve SRMCheckListInfo.SramPins(SRMIndex).LogicPair(TempCnt)
'                        SRMCheckListInfo.SramPins(SRMIndex).LogicPair(TempCnt) = UCase(Trim(SrmChkListArrTemp(LGPIndex, CurYIndex)))
'                    End If
                End If
            Next
''==============================================================================================
''Pasing "Write" data from SelSrm_ChkList sheet.
''==============================================================================================
            For CurYIndex = 1 To MaxRow - 1
                If SrmChkListArrTemp(WDIndex, CurYIndex) = "" Then
                    Exit For
                Else
                    ReDim Preserve SRMCheckListInfo.TestCase(CurYIndex - 1)
                    TestCaseArr = Split(UCase(Trim(SrmChkListArrTemp(WDIndex, CurYIndex))), "SRC")
                    SRMCheckListInfo.TestCase(CurYIndex - 1) = TestCaseArr(1) 'Get digsrc only
                End If
            Next
''==============================================================================================
''Pasing "Read" data from SelSrm_ChkList sheet.
''==============================================================================================
            For CurYIndex = 1 To MaxRow - 1
                If SrmChkListArrTemp(RDIndex, CurYIndex) = "" Then
                    Exit For
                Else
                    ReDim Preserve SRMCheckListInfo.ReadLimit(CurYIndex - 1)
                    LimitArr = Split(UCase(Trim(SrmChkListArrTemp(RDIndex, CurYIndex))), "CAP")
                    SRMCheckListInfo.ReadLimit(CurYIndex - 1) = Bin2Dec(LimitArr(1))
                End If
            Next
''==============================================================================================
''Pasing "Src_Cap_bits" from SelSrm_ChkList sheet.
''==============================================================================================
            For CurYIndex = 1 To MaxRow - 1
                If SrmChkListArrTemp(SCBIndex, CurYIndex) = "" Then
                    Exit For
                Else
                    ReDim Preserve SRMCheckListInfo.CapBitSize(CurYIndex - 1)
                    ReDim Preserve SRMCheckListInfo.SrcBitSize(CurYIndex - 1)
                    TempArr = Split(UCase(Trim(SrmChkListArrTemp(SCBIndex, CurYIndex))), ",")
                    SRMCheckListInfo.CapBitSize(CurYIndex - 1) = TempArr(1)
                    SRMCheckListInfo.SrcBitSize(CurYIndex - 1) = TempArr(0)
                End If
            Next
''==============================================================================================
''Pasing "Check" from SelSrm_ChkList sheet.
''==============================================================================================
            ReDim Preserve SRMCheckListInfo.Check(UBound(SRMCheckListInfo.TestCase))
            SRMCheckListInfo.CheckCount = 0
            For CurYIndex = 1 To MaxRow - 1
                If SrmChkListArrTemp(CheckIndex, CurYIndex) = "v" Then
                    SRMCheckListInfo.Check(SRMCheckListInfo.CheckCount) = CurYIndex - 1
                    SRMCheckListInfo.CheckCount = SRMCheckListInfo.CheckCount + 1
                End If
            Next
            ReDim Preserve SRMCheckListInfo.Check(SRMCheckListInfo.CheckCount - 1)
        End If
''==============================================================================================
''Initialize SRMCheckListInfo
''==============================================================================================
    ReDim SRMCheckListInfo.SrcBits(UBound(SRMCheckListInfo.LogicPins))
    Set MySheet = Nothing
''==============================================================================================
    End If
Exit Function
errHandler:
    Call Print_Error_Message(Error_Info, "VBT_LIB_SRM", "SRM_ParseChkList") 'Add ErrHandler 2023/08/18
    If AbortTest Then Exit Function Else Resume Next
End Function
''==============================================================================================
''SRM_FindSheet
''==============================================================================================
Private Function SRM_FindSheet(sheetName As String) As Boolean
On Error GoTo errHandler
    Dim funcName As String:: funcName = "SRM_FindSheet"

    Dim ShtCnt As Integer
    Dim ShtTypeArr() As String
    Dim SheetNameType As Variant

    If UCase(sheetName) Like "*JOBLIST*" Then
        SheetNameType = DMGR_SHEET_TYPE_JOBLISTSHEET
    ElseIf UCase(sheetName) Like "DC*SPECS*" Then
        SheetNameType = DMGR_SHEET_TYPE_DCSPECSHEET
    ElseIf UCase(sheetName) Like "*CHANNEL*" Then
        SheetNameType = DMGR_SHEET_TYPE_CHANMAP
    ElseIf UCase(sheetName) Like "PATSETS_*" Then
        SheetNameType = DMGR_SHEET_TYPE_PATTERNSETSHEET
    Else
        SheetNameType = DMGR_SHEET_TYPE_USER
    End If

    SRM_FindSheet = False
    ShtTypeArr = TheExec.job.GetSheetNamesOfType(SheetNameType)

    For ShtCnt = 0 To UBound(ShtTypeArr)
        If LCase(ShtTypeArr(ShtCnt)) Like LCase(sheetName) Then
            SRM_FindSheet = True
            Exit For
        End If
    Next ShtCnt

    If SRM_FindSheet = False Then
        TheExec.Datalog.WriteComment sheetName & " doesn't exist in current program which was specified in VBT, Please check the input for SRM_FindSheet. Error!!!"
        TheExec.ErrorLogMessage sheetName & " doesn't exist in current program which was specified in VBT, Please check the input for SRM_FindSheet. Error!!!"
    End If

Exit Function
errHandler:
    Call Print_Error_Message(Error_Info, "VBT_LIB_SRM", "SRM_FindSheet") 'Add ErrHandler 2023/08/18
    If AbortTest Then Exit Function Else Resume Next
End Function

''==============================================================================================
''SRM_ArgsToStr
''==============================================================================================
Public Function SRM_ArgsToStr(Optional Init_Pat1 As Pattern, Optional Init_Pat2 As Pattern, _
                                Optional Init_Pat3 As Pattern, Optional Init_Pat4 As Pattern, _
                                Optional Init_Pat5 As Pattern, Optional Init_Pat6 As Pattern, _
                                Optional Init_Pat7 As Pattern, Optional Init_Pat8 As Pattern, _
                                Optional Init_Pat9 As Pattern, Optional Init_Pat10 As Pattern, _
                                Optional Payload_Pat1 As Pattern, _
                                Optional Payload_Pat2 As Pattern, _
                                Optional Payload_Pat3 As Pattern, _
                                Optional Payload_Pat4 As Pattern, _
                                Optional Payload_Pat5 As Pattern) As String
On Error GoTo errHandler
    Dim funcName As String:: funcName = "SRM_ArgsToArr"

    Dim PatStr As String: PatStr = ""
    If Init_Pat1 <> "" Then PatStr = Init_Pat1.value
    If Init_Pat2 <> "" Then PatStr = PatStr & "," & Init_Pat2.value
    If Init_Pat3 <> "" Then PatStr = PatStr & "," & Init_Pat3.value
    If Init_Pat4 <> "" Then PatStr = PatStr & "," & Init_Pat4.value
    If Init_Pat5 <> "" Then PatStr = PatStr & "," & Init_Pat5.value
    If Init_Pat6 <> "" Then PatStr = PatStr & "," & Init_Pat6.value
    If Init_Pat7 <> "" Then PatStr = PatStr & "," & Init_Pat7.value
    If Init_Pat8 <> "" Then PatStr = PatStr & "," & Init_Pat8.value
    If Init_Pat9 <> "" Then PatStr = PatStr & "," & Init_Pat9.value
    If Init_Pat10 <> "" Then PatStr = PatStr & "," & Init_Pat10.value

    If Payload_Pat1 <> "" Then PatStr = PatStr & "," & Payload_Pat1.value
    If Payload_Pat2 <> "" Then PatStr = PatStr & "," & Payload_Pat2.value
    If Payload_Pat3 <> "" Then PatStr = PatStr & "," & Payload_Pat3.value
    If Payload_Pat4 <> "" Then PatStr = PatStr & "," & Payload_Pat4.value
    If Payload_Pat5 <> "" Then PatStr = PatStr & "," & Payload_Pat5.value

    SRM_ArgsToStr = PatStr

Exit Function
errHandler:
    Call Print_Error_Message(Error_Info, "VBT_LIB_SRM", "SRM_ArgsToStr") 'Add ErrHandler 2023/08/18
    If AbortTest Then Exit Function Else Resume Next
End Function
''==============================================================================================
''SRM_CreatePatSetAllDic
''==============================================================================================
Public Function SRM_CreatePatSetAllDic()
On Error GoTo errHandler
    Dim funcName As String:: funcName = "SRM_CreatePatSetAllDic"

    Dim MySheet As Worksheet:: Set MySheet = Nothing
    Dim sheetName As String
    Dim MaxColumn As Long, MaxRow As Long
    Dim CurColumn As Long, CurRow As Long
    Dim PatSetAllArr As Variant

    Dim startRow As Integer: startRow = 4
    Dim PatternSetColumn As Integer: PatternSetColumn = 2
    Dim FileGroupNameColumn As Integer: FileGroupNameColumn = 6

    PatSetAllDic.RemoveAll
'==============================================================================
'Parsing PatSets_All Sheet data.
'==============================================================================
    sheetName = "PatSets_All"
    If SRM_FindSheet(sheetName) Then
        Set MySheet = ActiveWorkbook.Sheets(sheetName)
        MySheet.Select
        MaxRow = MySheet.UsedRange.Rows.Count
        MaxColumn = MySheet.UsedRange.Columns.Count
        PatSetAllArr = MySheet.range(Cells(startRow, PatternSetColumn), Cells(MaxRow, FileGroupNameColumn)).value
        For CurRow = 1 To UBound(PatSetAllArr)
            If PatSetAllArr(CurRow, PatternSetColumn - 1) = "" Then
                Exit For
            Else
                If Not PatSetAllDic.Exists(UCase(Trim(PatSetAllArr(CurRow, PatternSetColumn - 1)))) Then
                    PatSetAllDic.Add UCase(Trim(PatSetAllArr(CurRow, PatternSetColumn - 1))), UCase(Trim(PatSetAllArr(CurRow, FileGroupNameColumn - 1)))
                End If
            End If
        Next CurRow
        Set MySheet = Nothing
    Else
        MsgBox "Sheet PatSets_All is NOT exist!!!", vbOKOnly + vbCritical, "Error"
    End If

Exit Function
errHandler:
    Call Print_Error_Message(Error_Info, "VBT_LIB_SRM", "SRM_CreatePatSetAllDic") 'Add ErrHandler 2023/08/18
    If AbortTest Then Exit Function Else Resume Next
End Function
''==============================================================================================
''SRM_PatPreCheck
''==============================================================================================
Public Function SRM_PatPreCheck(PatName As String) As Boolean
On Error GoTo errHandler
    Dim funcName As String:: funcName = "SRM_PatPreCheck"
'==============================================================================
'Check If PatName exist or not.
'==============================================================================
    SRM_PatPreCheck = True

    If Not PatSetAllDic.Exists(UCase(PatName)) Then
        SRM_PatPreCheck = False
    End If

Exit Function
errHandler:
    Call Print_Error_Message(Error_Info, "VBT_LIB_SRM", "SRM_PatPreCheck") 'Add ErrHandler 2023/08/18
    If AbortTest Then Exit Function Else Resume Next
End Function
''==============================================================================================
''SRM_InitDatalogSetup
''==============================================================================================
Public Function SRM_InitDatalogSetup()
On Error GoTo errHandler
    Dim funcName As String:: funcName = "SRM_InitDatalogSetup"

    TheExec.Datalog.Setup.Shared.ascii.Columns.EnableCustomWidths = True
    TheExec.Datalog.Setup.Shared.ascii.Columns.Parametric.TestName.width = 130
    TheExec.Datalog.Setup.Shared.ascii.Columns.Functional.TestName.width = 130
    TheExec.Datalog.Setup.Shared.ascii.Columns.Functional.Pattern.width = 130
    TheExec.Datalog.Setup.Shared.ascii.Columns.Parametric.Force.Enable = True
    TheExec.Datalog.ApplySetup  'must need to Apply after datalog setup

Exit Function
errHandler:
    Call Print_Error_Message(Error_Info, "VBT_LIB_SRM", "SRM_InitDatalogSetup") 'Add ErrHandler 2023/08/18
    If AbortTest Then Exit Function Else Resume Next
End Function
''==============================================================================================
''SRM_InitArrComma
''==============================================================================================
Public Function SRM_InitArrComma(InputArr() As String, ArrCnt As Long)
On Error GoTo errHandler
    Dim funcName As String:: funcName = "SRM_InitArrComma"

    Dim i As Integer
    For i = 0 To ArrCnt
        InputArr(i) = ","
    Next

Exit Function
errHandler:
    Call Print_Error_Message(Error_Info, "VBT_LIB_SRM", "SRM_InitArrComma") 'Add ErrHandler 2023/08/18
    If AbortTest Then Exit Function Else Resume Next
End Function
''==============================================================================================
''SRM_ReDimArrSize
''==============================================================================================
Public Function SRM_ReDimArrSize(InputArr() As String, ArrCnt As Long)
On Error GoTo errHandler
    Dim funcName As String:: funcName = "SRM_ReDimArrSize"

    Dim i As Integer
    For i = 0 To ArrCnt
        If InputArr(i) = "" Then
            Exit For
        End If
    Next

    If Not i = 0 Then ReDim Preserve InputArr(i - 1)
    ArrCnt = UBound(InputArr)

Exit Function
errHandler:
    Call Print_Error_Message(Error_Info, "VBT_LIB_SRM", "SRM_ReDimArrSize") 'Add ErrHandler 2023/08/18
    If AbortTest Then Exit Function Else Resume Next
End Function
''==============================================================================================
''SRM_InitPayloadArr
''==============================================================================================
Public Function SRM_InitPayloadArr(OutPutArr() As String, OutputArrCnt As Long, IintArr() As String, InitCnt As Long, PayloadArr() As String, PayloadCnt As Long)
On Error GoTo errHandler
    Dim funcName As String:: funcName = "SRM_SetInitPayloadArr"

    Dim i As Long, InitCntTemp As Long, PayloadCntTemp As Long

    PayloadCntTemp = PayloadCnt + 1
    InitCntTemp = InitCnt + 1

    For i = 0 To PayloadCnt
        If PayloadArr(i) = "" Then
            Exit For
        Else
            ReDim Preserve IintArr(InitCntTemp + i)
            IintArr(InitCntTemp + i) = PayloadArr(i)
        End If
    Next
    OutPutArr = IintArr
    OutputArrCnt = UBound(OutPutArr)
Exit Function
errHandler:
    Call Print_Error_Message(Error_Info, "VBT_LIB_SRM", "SRM_InitPayloadArr") 'Add ErrHandler 2023/08/18
    If AbortTest Then Exit Function Else Resume Next
End Function
''==============================================================================================
''SRM_FindSrmPatIndex
''==============================================================================================
Public Function SRM_FindSrmPatIndex(InputPutArr() As String, index As Long)
On Error GoTo errHandler
    Dim funcName As String:: funcName = "SRM_FindSrmPatIndex"

    Dim i As Long
    For i = 0 To UBound(InputPutArr)
        If UCase(InputPutArr(i)) Like "*SRM*" Then
            index = i
            Exit For
        End If
    Next

Exit Function
errHandler:
    Call Print_Error_Message(Error_Info, "VBT_LIB_SRM", "SRM_FindSrmPatIndex") 'Add ErrHandler 2023/08/18
    If AbortTest Then Exit Function Else Resume Next
End Function
''==============================================================================================
''SRM_DecideType
''==============================================================================================
Public Function SRM_DecideType(InputPutStr As String)
On Error GoTo errHandler
    Dim funcName As String:: funcName = "SRM_DecideType"

    If InputPutStr Like "*DSSC*" Then
       InputPutStr = "DSSC"
    ElseIf InputPutStr Like "*SRMEXP*" Then
        InputPutStr = "SRMEXP"
    ElseIf InputPutStr Like "*SRMREAD" Then
        InputPutStr = "SRMREAD"
    ElseIf InputPutStr Like "*HC*" Then
        InputPutStr = "STATIC"
    ElseIf InputPutStr Like "*CUSTOM*" Then
        InputPutStr = "CUSTOM"
    End If
Exit Function
errHandler:
    Call Print_Error_Message(Error_Info, "VBT_LIB_SRM", "SRM_DecideType") 'Add ErrHandler 2023/08/18
    If AbortTest Then Exit Function Else Resume Next
End Function
''==============================================================================================
''SRM_GetInfo
''==============================================================================================
Public Function SRM_GetInfo(InputPutStr As String)
On Error GoTo errHandler
    Dim funcName As String:: funcName = "SRM_GetInfo"

Exit Function
errHandler:
    Call Print_Error_Message(Error_Info, "VBT_LIB_SRM", "SRM_GetInfo") 'Add ErrHandler 2023/08/18
    If AbortTest Then Exit Function Else Resume Next
End Function

''==============================================================================================
''SRM_InitLoopCnt
''==============================================================================================
Public Function SRM_InitLoopCnt(SRMBlock As String)
On Error GoTo errHandler
    Dim funcName As String:: funcName = "SRM_InitLoopCnt"
''==============================================================================================
''Decide loop counter.
''==============================================================================================
    If Not SRMCheckListInfo.SrmPattern(SRMCheckListInfo.SelSrmDic(UCase(SRMBlock)) - 1).Patterns(0) = "" Then
        If (UCase(SRMBlock) Like "*HC*") Then
            Call SRM_AutoSetLoopCnt(UBound(SRMCheckListInfo.SrmPattern(SRMCheckListInfo.SelSrmDic(UCase(SRMBlock)) - 1).Patterns))
        ElseIf (UCase(SRMBlock) Like "*DSSC*") Then
            Call SRM_AutoSetLoopCnt(UBound(SRMCheckListInfo.SrmPattern(SRMCheckListInfo.SelSrmDic(UCase(SRMBlock)) - 1).Patterns), SRMCheckListInfo.CheckCount - 1)
        End If
    Else
        Call SRM_AutoSetLoopCnt(-1, -1)
    End If

Exit Function
errHandler:
    Call Print_Error_Message(Error_Info, "VBT_LIB_SRM", "SRM_InitLoopCnt") 'Add ErrHandler 2023/08/18
    If AbortTest Then Exit Function Else Resume Next
End Function
''==============================================================================================
''SRM_AutoSetLoopCnt
''==============================================================================================
Public Function SRM_AutoSetLoopCnt(LP_Count As Long, Optional SEL_Case_Cnt As Long) As Long
On Error GoTo errHandler
    Dim funcName As String:: funcName = "Auto_SelSrm_SetLoopCNT"
    Dim site As Variant
''==============================================================================================
''Set Flow site-var value.
''==============================================================================================
    For Each site In TheExec.sites.Active
        TheExec.sites(site).SiteVariableValue("SelSrm_LP_Var") = 0
        TheExec.sites(site).SiteVariableValue(glbConstVar_PatLoopCntEnd) = LP_Count + 1
        TheExec.sites(site).SiteVariableValue(glbConstVar_TestCaseCnt) = 0
        TheExec.sites(site).SiteVariableValue(glbConstVar_TestCaseCntEnd) = SEL_Case_Cnt + 1
    Next site

Exit Function
errHandler:
    Call Print_Error_Message(Error_Info, "VBT_LIB_SRM", "SRM_AutoSetLoopCnt") 'Add ErrHandler 2023/08/18
    If AbortTest Then Exit Function Else Resume Next
End Function
''==============================================================================================
''SRM_GetChnType
''==============================================================================================
Public Function SRM_GetChnType()
On Error GoTo errHandler
    Dim funcName As String:: funcName = "SRM_GetChnType"
    
    Dim i As Long
    Dim j As Long
    Dim lclLng_TypeCol As Long
    Dim lclLng_HeaderRow As Long
    Dim lclLng_PinNameCol As Long
    Dim lclLng_PowerPinCnt As Long:: lclLng_PowerPinCnt = 0
    Dim lclLng_MaxColumn As Long, lclLng_MaxRow As Long
    Dim lclLng_FindAllHeader As Long:: lclLng_FindAllHeader = 0
    
    Dim lclStr_CurChnMap As String
    Dim lclStrArr_ChnMapSheet() As String
    Dim lclStr_ChnMapType As String:: lclStr_ChnMapType = "TYPE"
    Dim lclStr_ChnMapPinName As String:: lclStr_ChnMapPinName = "PIN NAME"

    Dim lclVar_ChnMapInfo As Variant

    Dim lclWS_MySheet As Worksheet:: Set lclWS_MySheet = Nothing
''==============================================================================================
''Define the HexVS's Rangelist. It cannot get directly, if we use the merge pin.
''==============================================================================================
    Dim lclDbl_HexVSRangeListSize As Double:: lclDbl_HexVSRangeListSize = 3
    Dim lclDblArr_HexVSRangeList() As Double:: ReDim lclDblArr_HexVSRangeList(lclDbl_HexVSRangeListSize)
    lclDblArr_HexVSRangeList(0) = 0.01
    lclDblArr_HexVSRangeList(1) = 0.1
    lclDblArr_HexVSRangeList(2) = 1
    lclDblArr_HexVSRangeList(3) = 15
''==============================================================================================
''Choose the current Channel Map
''==============================================================================================
    lclStr_CurChnMap = LCase(TheExec.CurrentChanMap)
    If SRM_FindSheet(lclStr_CurChnMap) Then
        Set lclWS_MySheet = ActiveWorkbook.Sheets(lclStr_CurChnMap)
        lclWS_MySheet.Select
        lclLng_MaxRow = lclWS_MySheet.UsedRange.Rows.Count
        lclLng_MaxColumn = lclWS_MySheet.UsedRange.Columns.Count
        lclVar_ChnMapInfo = lclWS_MySheet.range(Cells(1, 1), Cells(lclLng_MaxRow, lclLng_MaxColumn)).value
        lclLng_HeaderRow = lclLng_MaxRow
''==============================================================================================
''Get the Header column "Pin Name" and "Type" from current channel map
''==============================================================================================
        For i = 1 To lclLng_MaxColumn - 1
            For j = 1 To lclLng_HeaderRow
                If lclVar_ChnMapInfo(j, i) <> "" Then
                    If UCase(lclVar_ChnMapInfo(j, i)) Like lclStr_ChnMapPinName Then
                        lclLng_HeaderRow = j
                        lclLng_PinNameCol = i
                        lclLng_FindAllHeader = lclLng_FindAllHeader + Enum_PinName
                        Exit For
                    ElseIf UCase(lclVar_ChnMapInfo(j, i)) Like lclStr_ChnMapType Then
                        lclLng_TypeCol = i
                        lclLng_FindAllHeader = lclLng_FindAllHeader + Enum_PinType
                        Exit For
                    End If
                End If
            Next j
            If lclLng_FindAllHeader > Enum_PinType Then
                lclLng_FindAllHeader = 0
                Exit For
            End If
        Next i
''==============================================================================================
''Get Pin Name and Pin Type from Channel Map
''==============================================================================================
        ReDim glbArr_AllPinInfo(lclLng_MaxRow - lclLng_HeaderRow - 1)
        For i = lclLng_HeaderRow + 1 To lclLng_MaxRow
            glbArr_AllPinInfo(i - (lclLng_HeaderRow + 1)).PinName = lclVar_ChnMapInfo(i, lclLng_PinNameCol)
            glbArr_AllPinInfo(i - (lclLng_HeaderRow + 1)).MergeType = lclVar_ChnMapInfo(i, lclLng_TypeCol)
        Next
''==============================================================================================
''Do if PinName is not emtpy and MergeType is not N/C pin in channel map sheet
''==============================================================================================
        For i = 0 To UBound(glbArr_AllPinInfo)
            If glbArr_AllPinInfo(i).PinName <> glbConstKeyWD_Empty And glbArr_AllPinInfo(i).MergeType <> glbConstKeyWD_NoConnected Then
                glbArr_AllPinInfo(i).PinMapType = UCase(TheExec.DataManager.pinType(glbArr_AllPinInfo(i).PinName))
                '''Do if PinMapType is Power in PinMap sheet
                If glbArr_AllPinInfo(i).PinMapType Like glbConstKeyWD_POWER Then
                    ReDim Preserve glbArr_PowerPinInfo(lclLng_PowerPinCnt)
                    
                    glbArr_PowerPinInfo(lclLng_PowerPinCnt).PinName = UCase(glbArr_AllPinInfo(i).PinName)
                    glbArr_PowerPinInfo(lclLng_PowerPinCnt).MergeType = UCase(glbArr_AllPinInfo(i).MergeType)
                    glbArr_PowerPinInfo(lclLng_PowerPinCnt).PinMapType = UCase(glbArr_AllPinInfo(i).PinMapType)
''==============================================================================================
''Get the instrument type, i.g., "DCVS", "DCVSMergedN" and "DCVI"...
''==============================================================================================
                    glbArr_PowerPinInfo(lclLng_PowerPinCnt).ChanMapType = UCase(GetInstrument(glbArr_PowerPinInfo(lclLng_PowerPinCnt).PinName, 0))
''==============================================================================================
''Get merged value from the the instrument type, i.g., "DCVS" => 1, "DCVSMergedN" => N
 ''==============================================================================================
                    glbArr_PowerPinInfo(lclLng_PowerPinCnt).MergedN = SRM_FindChnMergeCase(glbArr_PowerPinInfo(lclLng_PowerPinCnt).ChanMapType)
''==============================================================================================
''Get the initial current range value from HW
''==============================================================================================
                    If glbArr_PowerPinInfo(lclLng_PowerPinCnt).MergeType Like glbConstKeyWD_DCVS Then
                        glbArr_PowerPinInfo(lclLng_PowerPinCnt).Init_CurrentRange = TheHdw.DCVS.Pins(glbArr_PowerPinInfo(lclLng_PowerPinCnt).PinName).CurrentRange.value
                    ElseIf glbArr_PowerPinInfo(lclLng_PowerPinCnt).MergeType Like glbConstKeyWD_DCVI Then
                        glbArr_PowerPinInfo(lclLng_PowerPinCnt).Init_CurrentRange = TheHdw.DCVI.Pins(glbArr_PowerPinInfo(lclLng_PowerPinCnt).PinName).CurrentRange.value
                    End If
''==============================================================================================
''Get the current range list from HW
''==============================================================================================
                    If glbArr_PowerPinInfo(lclLng_PowerPinCnt).ChanMapType Like glbConstIns_VHDVS _
                        Or glbArr_PowerPinInfo(lclLng_PowerPinCnt).ChanMapType Like glbConstIns_VSM _
                        Or glbArr_PowerPinInfo(lclLng_PowerPinCnt).ChanMapType Like glbConstIns_VS5A _
                        Or glbArr_PowerPinInfo(lclLng_PowerPinCnt).ChanMapType Like glbConstIns_VS800MA Then
                        
                        glbArr_PowerPinInfo(lclLng_PowerPinCnt).Range_List = TheHdw.DCVS.Pins(glbArr_PowerPinInfo(lclLng_PowerPinCnt).PinName).CurrentRange.List
                    ElseIf glbArr_PowerPinInfo(lclLng_PowerPinCnt).ChanMapType Like glbConstIns_HEXVS Then
                        glbArr_PowerPinInfo(lclLng_PowerPinCnt).Range_List = lclDblArr_HexVSRangeList
                        If glbArr_PowerPinInfo(lclLng_PowerPinCnt).MergeType Like glbConstKeyWD_DCVSMERGED Then
                            ReDim Preserve glbArr_PowerPinInfo(lclLng_PowerPinCnt).Range_List(4)
                            glbArr_PowerPinInfo(lclLng_PowerPinCnt).Range_List(4) = TheHdw.DCVS.Pins(glbArr_PowerPinInfo(lclLng_PowerPinCnt).PinName).Meter.CurrentRange
                        End If
                    ElseIf glbArr_PowerPinInfo(lclLng_PowerPinCnt).ChanMapType Like glbConstIns_DC07 Then
                        glbArr_PowerPinInfo(lclLng_PowerPinCnt).Range_List = TheHdw.DCVI.Pins(glbArr_PowerPinInfo(lclLng_PowerPinCnt).PinName).CurrentRange.List
                    End If
''==============================================================================================
''Get the waittime list for each current range list, respectively
''==============================================================================================
                    glbArr_PowerPinInfo(lclLng_PowerPinCnt).WaitTime_List = SRM_FindRangeWaitTime(glbArr_PowerPinInfo(lclLng_PowerPinCnt).Range_List, glbArr_PowerPinInfo(lclLng_PowerPinCnt).ChanMapType, glbArr_PowerPinInfo(lclLng_PowerPinCnt).MergedN)
''==============================================================================================
''Get the Accuracy list for each current range list, respectively
''==============================================================================================
                    glbArr_PowerPinInfo(lclLng_PowerPinCnt).Accuracy_List = SRM_FindRangeAccuracy(glbArr_PowerPinInfo(lclLng_PowerPinCnt).Range_List, glbArr_PowerPinInfo(lclLng_PowerPinCnt).ChanMapType, glbArr_PowerPinInfo(lclLng_PowerPinCnt).MergedN)
''==============================================================================================
''Update PowerPin Count
''==============================================================================================
                    lclLng_PowerPinCnt = lclLng_PowerPinCnt + 1
                End If
            End If
        Next
''==============================================================================================
''Reset WorkSheet to Nothing
''==============================================================================================
        Set lclWS_MySheet = Nothing
    End If
    Exit Function
errHandler:
    Call Print_Error_Message(Error_Info, "VBT_LIB_SRM", "SRM_GetChnType") 'Add ErrHandler 2023/08/18
    If AbortTest Then Exit Function Else Resume Next
End Function
''==============================================================================================
''SRM_FindRangeWaitTime
''==============================================================================================
Public Function SRM_FindRangeWaitTime(Input_RangeList() As Double, Input_InstrucmentType As String, Optional Input_MergedN As Long = 1) As Variant
On Error GoTo errHandler
    Dim funcName As String:: funcName = "SRM_FindRangeWaitTime"
    
    Dim var As Variant
    Dim lclDbl_SattleTime() As Double:: ReDim lclDbl_SattleTime(UBound(Input_RangeList))

    If Input_InstrucmentType Like glbConstIns_HEXVS Then
        For var = 0 To UBound(Input_RangeList)
            If Input_RangeList(var) = 0.01 Then
                lclDbl_SattleTime(var) = 100 * ms
            ElseIf Input_RangeList(var) = 0.1 Then
                lclDbl_SattleTime(var) = 10 * ms
            ElseIf Input_RangeList(var) = 1 Then
                lclDbl_SattleTime(var) = 1 * ms
            ElseIf Input_RangeList(var) >= 15 Then
                lclDbl_SattleTime(var) = 100 * us
            End If
        Next var
    ElseIf Input_InstrucmentType Like glbConstIns_VHDVS Then
        For var = 0 To UBound(Input_RangeList)
            If Input_RangeList(var) = 0.000004 Then
                lclDbl_SattleTime(var) = 18 * ms
            ElseIf Input_RangeList(var) = 0.00002 Then
                lclDbl_SattleTime(var) = 4 * ms
            ElseIf Input_RangeList(var) = 0.0002 Then
                lclDbl_SattleTime(var) = 4 * ms
            ElseIf Input_RangeList(var) = 0.002 Then
                lclDbl_SattleTime(var) = 3.5 * ms
            ElseIf Input_RangeList(var) = 0.02 Then
                lclDbl_SattleTime(var) = 540 * us
            ElseIf Input_RangeList(var) = 0.04 Then
                lclDbl_SattleTime(var) = 260 * us
            ElseIf Input_RangeList(var) = 0.2 Then
                lclDbl_SattleTime(var) = 210 * us
            ElseIf Input_RangeList(var) = 0.4 Then
                lclDbl_SattleTime(var) = 90 * us
            ElseIf Input_RangeList(var) = 0.7 Then
                lclDbl_SattleTime(var) = 100 * us
            ElseIf Input_RangeList(var) = 0.8 Then
                lclDbl_SattleTime(var) = 100 * us
            ElseIf Input_RangeList(var) = 1.4 Then
                lclDbl_SattleTime(var) = 50 * us
            ElseIf Input_RangeList(var) = 2.8 Then
                lclDbl_SattleTime(var) = 45 * us
            ElseIf Input_RangeList(var) = 5.6 Then
                lclDbl_SattleTime(var) = 30 * us
            End If
        Next var
    ElseIf Input_InstrucmentType Like glbConstIns_DC07 Then
        For var = 0 To UBound(Input_RangeList)
            If Input_RangeList(var) = 0.000002 * Input_MergedN Then
                lclDbl_SattleTime(var) = 6 * ms
            ElseIf Input_RangeList(var) = 0.00002 * Input_MergedN Then
                lclDbl_SattleTime(var) = 1.5 * ms
            ElseIf Input_RangeList(var) = 0.0002 * Input_MergedN Then
                lclDbl_SattleTime(var) = 1.4 * ms
            ElseIf Input_RangeList(var) = 0.002 * Input_MergedN Then
                lclDbl_SattleTime(var) = 11 * ms
            ElseIf Input_RangeList(var) = 0.02 * Input_MergedN Then
                lclDbl_SattleTime(var) = 1.5 * ms
            ElseIf Input_RangeList(var) = 0.2 * Input_MergedN Then
                lclDbl_SattleTime(var) = 260 * us
            ElseIf Input_RangeList(var) >= 1 Then
                lclDbl_SattleTime(var) = 1.6 * ms
            End If
        Next var
    ElseIf Input_InstrucmentType Like glbConstIns_VSM Then
        For var = 0 To UBound(Input_RangeList)
            If Input_RangeList(var) <= 10 Then
                lclDbl_SattleTime(var) = 6 * ms
            Else
                lclDbl_SattleTime(var) = 1.5 * ms
            End If
        Next var
    Else
        Call Print_Error_Message(Error_Warning_Info.Warning_Info, "VBT_LIB_SRM", "SRM_FindRangeWaitTime", "Instrument not define !!")
    End If
    SRM_FindRangeWaitTime = lclDbl_SattleTime

Exit Function
errHandler:
    Call Print_Error_Message(Error_Info, "VBT_LIB_SRM", "SRM_FindRangeWaitTime") 'Add ErrHandler 2023/08/18
    If AbortTest Then Exit Function Else Resume Next
End Function
''==============================================================================================
''SRM_FindRangeAccuracy
''==============================================================================================
Public Function SRM_FindRangeAccuracy(Input_RangeList() As Double, Input_InstrucmentType As String, Optional Input_MergedN As Long = 1) As Variant
On Error GoTo errHandler
    Dim funcName As String:: funcName = "SRM_FindRangeAccuracy"

    Dim var As Variant
    Dim lclDbl_Accuracy() As Double:: ReDim lclDbl_Accuracy(UBound(Input_RangeList))

    If Input_InstrucmentType Like glbConstIns_HEXVS Then
        For var = 0 To UBound(Input_RangeList)
            If Input_RangeList(var) = 0.01 Then
                lclDbl_Accuracy(var) = Input_RangeList(var) * 0.01 + 0.00005
            ElseIf Input_RangeList(var) = 0.1 Then
                lclDbl_Accuracy(var) = Input_RangeList(var) * 0.01 + 0.0005
            ElseIf Input_RangeList(var) = 1 Then
                lclDbl_Accuracy(var) = Input_RangeList(var) * 0.01 + 0.005
            ElseIf Input_RangeList(var) >= 15 Then
                lclDbl_Accuracy(var) = Input_RangeList(var) * 0.01 + (Input_RangeList(var) \ 15) * 0.05
            End If
        Next var
    ElseIf Input_InstrucmentType Like glbConstIns_VHDVS Then
        For var = 0 To UBound(Input_RangeList)
            If Input_RangeList(var) = 0.000004 Then
                lclDbl_Accuracy(var) = Input_RangeList(var) * 0.005 + 0.000000026
            ElseIf Input_RangeList(var) = 0.00002 Then
                lclDbl_Accuracy(var) = Input_RangeList(var) * 0.005 + 0.00000012
            ElseIf Input_RangeList(var) = 0.0002 Then
                lclDbl_Accuracy(var) = Input_RangeList(var) * 0.005 + 0.0000012
            ElseIf Input_RangeList(var) = 0.002 Then
                lclDbl_Accuracy(var) = Input_RangeList(var) * 0.005 + 0.000012
            ElseIf Input_RangeList(var) = 0.02 Then
                lclDbl_Accuracy(var) = Input_RangeList(var) * 0.005 + 0.00012
            ElseIf Input_RangeList(var) = 0.04 Then
                lclDbl_Accuracy(var) = Input_RangeList(var) * 0.007 + 0.00024
            ElseIf Input_RangeList(var) = 0.2 Then
                lclDbl_Accuracy(var) = Input_RangeList(var) * 0.005 + 0.0012
            ElseIf Input_RangeList(var) = 0.4 Then
                lclDbl_Accuracy(var) = Input_RangeList(var) * 0.007 + 0.0024
            ElseIf Input_RangeList(var) = 0.7 Then
                lclDbl_Accuracy(var) = Input_RangeList(var) * 0.005 + 0.006
            ElseIf Input_RangeList(var) = 0.8 Then
                lclDbl_Accuracy(var) = Input_RangeList(var) * 0.005 + 0.006
            ElseIf Input_RangeList(var) = 1.4 Then
                lclDbl_Accuracy(var) = Input_RangeList(var) * 0.007 + 0.01
            ElseIf Input_RangeList(var) = 2.8 Then
                lclDbl_Accuracy(var) = Input_RangeList(var) * 0.007 + 0.02
            ElseIf Input_RangeList(var) = 5.6 Then
                lclDbl_Accuracy(var) = Input_RangeList(var) * 0.007 + 0.04
            End If
        Next var
    ElseIf Input_InstrucmentType Like glbConstIns_DC07 Then
        For var = 0 To UBound(Input_RangeList)
            If Input_RangeList(var) = 0.00002 * Input_MergedN Then
                lclDbl_Accuracy(var) = Input_RangeList(var) * 0.002 + 0.000000075
            ElseIf Input_RangeList(var) = 0.0002 * Input_MergedN Then
                lclDbl_Accuracy(var) = Input_RangeList(var) * 0.002 + 0.0000004
            ElseIf Input_RangeList(var) = 0.002 * Input_MergedN Then
                lclDbl_Accuracy(var) = Input_RangeList(var) * 0.002 + 0.000004
            ElseIf Input_RangeList(var) = 0.02 * Input_MergedN Then
                lclDbl_Accuracy(var) = Input_RangeList(var) * 0.002 + 0.00004
            ElseIf Input_RangeList(var) = 0.2 * Input_MergedN Then
                lclDbl_Accuracy(var) = Input_RangeList(var) * 0.002 + 0.0004
            ElseIf Input_RangeList(var) >= 1 Then
                lclDbl_Accuracy(var) = Input_RangeList(var) * 0.002 + 0.008
            End If
        Next var
   ElseIf Input_InstrucmentType Like glbConstIns_VSM Then
        For var = 0 To UBound(Input_RangeList)
            If Input_RangeList(var) <= 10 Then
                lclDbl_Accuracy(var) = Input_RangeList(var) * 0.0025 + Input_RangeList(var) * 0.007
            ElseIf (Input_RangeList(var) Mod 11) = 0 Then
                lclDbl_Accuracy(var) = Input_RangeList(var) * 0.0025 + (Input_RangeList(var) \ 11) * 0.04
            ElseIf (Input_RangeList(var) Mod 21) = 0 Then
                lclDbl_Accuracy(var) = Input_RangeList(var) * 0.0025 + (Input_RangeList(var) \ 21) * 0.08
            ElseIf (Input_RangeList(var) Mod 51) = 0 Then
                lclDbl_Accuracy(var) = Input_RangeList(var) * 0.0025 + (Input_RangeList(var) \ 51) * 0.2
            ElseIf (Input_RangeList(var) Mod 81) = 0 Then
                lclDbl_Accuracy(var) = Input_RangeList(var) * 0.0025 + (Input_RangeList(var) \ 81) * 0.32
            End If
        Next var
    Else
        Call Print_Error_Message(Error_Warning_Info.Warning_Info, "VBT_LIB_SRM", "SRM_FindRangeAccuracy", "Instrument not define !!")
    End If
    SRM_FindRangeAccuracy = lclDbl_Accuracy
    
Exit Function
errHandler:
    Call Print_Error_Message(Error_Info, "VBT_LIB_SRM", "SRM_FindRangeAccuracy") 'Add ErrHandler 2023/08/18
    If AbortTest Then Exit Function Else Resume Next
End Function
''==============================================================================================
''SRM_FindChnMergeCase
''==============================================================================================
Public Function SRM_FindChnMergeCase(Input_MergeType As String) As Variant
On Error GoTo errHandler
    Dim funcName As String:: funcName = "SRM_FindChnMergeCase"

    SRM_FindChnMergeCase = 1
    
    If Input_MergeType Like glbConstKeyWD_DCVSMERGED Then
        SRM_FindChnMergeCase = Replace(Input_MergeType, Replace(glbConstKeyWD_DCVSMERGED, "*", ""), "")
    ElseIf Input_MergeType Like glbConstKeyWD_DCVIMERGED Then
        SRM_FindChnMergeCase = 2
    End If

Exit Function
errHandler:
    Call Print_Error_Message(Error_Info, "VBT_LIB_SRM", "SRM_FindChnMergeCase") 'Add ErrHandler 2023/08/18
    If AbortTest Then Exit Function Else Resume Next
End Function


