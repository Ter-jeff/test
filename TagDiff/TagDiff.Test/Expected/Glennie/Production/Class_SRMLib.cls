VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Class_SRMLib"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private glbStr_WaitTime As String
Private glbStr_DigSrcBitSize As String
Private glbStr_DigSrcSeg As String
Private glbStr_DigSrcPin As String
Private glbStr_DigSrcEQ As String
Private glbStr_CGVol As String

Private atrStr_ForceVoltage As String

Private SrmInfo As SrmInstArgs

Private Const glbConst_DigSrcPin = "JTAG_TDI"
Private Const glbConst_DigCapPin = "JTAG_TDO"
Private Const glbConst_SrcSignalName = "FUNC_SRC"
Private Const glbConst_CapSignalName = "FUNC_CAP"
Private Const glbConst_DigCapName = "DigCap"

Private Const glbConst_SRMREAD = "SRMREAD "
Private Const glbConst_SRMEXPL = "SRMEXPL "
Private Const glbConst_SRMEXPH = "SRMEXPH "
Private Const glbConst_SRMCurrentLHSL = "SRMCurrent LHSL "
Private Const glbConst_SRMCurrentLLSH = "SRMCurrent LLSH "

Private Const glbConstKeyWD_LOGICH_SRAML = "*LOGICH_SRAML*"
Private Const glbConstKeyWD_LOGICL_SRAMH = "*LOGICL_SRAMH*"
Private Const glbConstKeyWD_SRM = "*_SRM*"
Private Const glbConstKeyWD_SELSRM = "*_SELSRM*"
Private Const glbConstKeyWD_DSRM = "*_DSRM*"
Private Const glbConstKeyWD_PRESERVE = "PRESERVE"
Private Const glbConstKeyWD_EXPH = "*EXPH"
Private Const glbConstKeyWD_EXPL = "*EXPL"
Private Const glbConstKeyWD_SRAM = "*SRAM*"

Private glbPinList_DigSrcPin As New PinList
Private glbPinList_DigCapPin As New PinList

Private Type SrmInstArgs
    InitPatsArr() As String
    InitPatsCnt As Long
    PayloadPatsArr() As String
    PayloadPatsCnt As Long
    InitPatSRMIndex As Long
    AllPatsArr() As String
    AllPatsCnt As Long
    SRMType As String
End Type
''==============================================================================================
''Initialize
''==============================================================================================
Public Function Initialize(ByVal Init_Pat1 As Pattern, ByVal Init_Pat2 As Pattern, _
                             ByVal Init_Pat3 As Pattern, ByVal Init_Pat4 As Pattern, _
                             ByVal Init_Pat5 As Pattern, ByVal Init_Pat6 As Pattern, _
                             ByVal Init_Pat7 As Pattern, ByVal Init_Pat8 As Pattern, _
                             ByVal Init_Pat9 As Pattern, ByVal Init_Pat10 As Pattern, _
                             ByVal Payload_Pat1 As Pattern, _
                             ByVal Payload_Pat2 As Pattern, _
                             ByVal Payload_Pat3 As Pattern, _
                             ByVal Payload_Pat4 As Pattern, _
                             ByVal Payload_Pat5 As Pattern, _
                             ByVal WaitTime As String, _
                             ByVal DigSrcBitSize As String, _
                             ByVal DigSrcSeg As String, _
                             ByVal DigSrcPin As String, _
                             ByVal DigSrcEQ As String, _
                             ByVal CGVoltage As String, _
                             SRMType As String, PatternTimeout As String) As Long

    On Error GoTo errHandler:: Dim funcName As String:: funcName = "Class_SRMLib.Initialize"
''==============================================================================================
''Initialize and get value
''==============================================================================================
    Dim lclStr_InitPat As String
    Dim lclStr_PayloadPat As String
    Dim lclStrArr_DigSrcBitSize() As String
    Dim lclStrArr_DigSrcSeg() As String
    Dim lclStrArr_DigSrcPin() As String
    Dim lclStrArr_DigSrcEQ() As String
    Dim lclStrArr_WaitTime() As String
    Dim lclStrArr_Pat() As String
'    Dim lclStr_DigCapBitSize As String:: lclStr_DigCapBitSize =
'    Dim lclStr_DigCapPin As String:: lclStr_DigCapPin =
    Dim i As Integer
''==============================================================================================
''Enable the pattern timeout counter
''==============================================================================================
    TheHdw.Digital.Patgen.TimeoutEnable = True
    TheHdw.Digital.Patgen.TimeOut = PatternTimeout
''==============================================================================================
''Init class Property
''==============================================================================================
    atrStr_ForceVoltage = CGVoltage
''==============================================================================================
''Init and payload
''==============================================================================================
    lclStr_InitPat = Init_Pat1 & ";" & Init_Pat2 & ";" & Init_Pat3 & ";" & Init_Pat4 & ";" & Init_Pat5 & ";" & Init_Pat6 & ";" & Init_Pat7 & ";" & Init_Pat8 & ";" & Init_Pat9 & ";" & Init_Pat10
    SrmInfo.InitPatsArr = Split(lclStr_InitPat, ";")
    SrmInfo.InitPatsCnt = UBound(SrmInfo.InitPatsArr)
    Call SRM_ReDimArrSize(SrmInfo.InitPatsArr, SrmInfo.InitPatsCnt)

    lclStr_PayloadPat = Payload_Pat1 & ";" & Payload_Pat2 & ";" & Payload_Pat3 & ";" & Payload_Pat4 & ";" & Payload_Pat5
    SrmInfo.PayloadPatsArr = Split(lclStr_PayloadPat, ";")
    SrmInfo.PayloadPatsCnt = UBound(SrmInfo.PayloadPatsArr)
    Call SRM_ReDimArrSize(SrmInfo.PayloadPatsArr, SrmInfo.PayloadPatsCnt)
    
    SrmInfo.AllPatsArr = SrmInfo.InitPatsArr
    Call SRM_InitPayloadArr(SrmInfo.AllPatsArr, SrmInfo.AllPatsCnt, SrmInfo.InitPatsArr, SrmInfo.InitPatsCnt, SrmInfo.PayloadPatsArr, SrmInfo.PayloadPatsCnt)
    Call SRM_FindSrmPatIndex(SrmInfo.AllPatsArr, SrmInfo.InitPatSRMIndex)
''==============================================================================================
''Check SRM Type
''==============================================================================================
    SrmInfo.SRMType = SRMType
    Call SRM_DecideType(SRMType)
''==============================================================================================
''WaitTime
''==============================================================================================
    If InStr(WaitTime, ",") > 0 Then
       lclStrArr_WaitTime = Split(Replace(WaitTime, "'", ""), ",")
    Else
       ReDim lclStrArr_WaitTime(glbLng_PatCnt) As String
       Call SRM_InitArrComma(lclStrArr_WaitTime, glbLng_PatCnt)
       glbStr_WaitTime = Join(lclStrArr_WaitTime, "")
    End If
''==============================================================================================
''lclStr_DigSrcBitSize
''==============================================================================================
    
    If InStr(DigSrcBitSize, ",") > 0 Then
        lclStrArr_DigSrcBitSize = Split(Replace(DigSrcBitSize, "'", ""), ",")
    Else
        ReDim lclStrArr_DigSrcBitSize(glbLng_PatCnt) As String
        Call SRM_InitArrComma(lclStrArr_DigSrcBitSize, glbLng_PatCnt)
        glbStr_DigSrcBitSize = Join(lclStrArr_DigSrcBitSize, "")
    End If
''==============================================================================================
''lclStr_DigSrcSeg
''==============================================================================================
    If InStr(DigSrcSeg, ",") > 0 Then
        lclStrArr_DigSrcSeg = Split(Replace(DigSrcSeg, "'", ""), ",")
    Else
        ReDim lclStrArr_DigSrcSeg(glbLng_PatCnt) As String
        Call SRM_InitArrComma(lclStrArr_DigSrcSeg, glbLng_PatCnt)
        glbStr_DigSrcSeg = Join(lclStrArr_DigSrcSeg, "")
    End If
''==============================================================================================
''lclStr_DigSrcPin
''==============================================================================================
    glbPinList_DigSrcPin.value = IIf(DigSrcPin = "", glbConst_DigSrcPin, DigSrcPin)
    glbPinList_DigCapPin.value = glbConst_DigCapPin
    If InStr(DigSrcPin, ",") > 0 Then
        lclStrArr_DigSrcPin = Split(Replace(DigSrcPin, "'", ""), ",")
    Else
        ReDim lclStrArr_DigSrcPin(glbLng_PatCnt) As String
        Call SRM_InitArrComma(lclStrArr_DigSrcPin, glbLng_PatCnt)
        glbStr_DigSrcPin = Join(lclStrArr_DigSrcPin, "")
    End If
''==============================================================================================
''lclStr_DigSrcEQ
''==============================================================================================
    If InStr(DigSrcEQ, ",") > 0 Then
        lclStrArr_DigSrcEQ = Split(Replace(DigSrcEQ, "'", ""), ",")
    Else
        ReDim lclStrArr_DigSrcEQ(glbLng_PatCnt) As String
        Call SRM_InitArrComma(lclStrArr_DigSrcEQ, glbLng_PatCnt)
        glbStr_DigSrcEQ = Join(lclStrArr_DigSrcEQ, "")
    End If
    
    Exit Function
errHandler:
    If AbortTest Then Exit Function Else Resume Next
End Function
''==============================================================================================
''DSRMDSSC/SRMDSSC
''Auto Change DSSC pattern from SelSram_Chklist data structure
''==============================================================================================
Public Function SRMViaDSSC()
    On Error GoTo errHandler:: Dim funcName As String:: funcName = "Class_SRMLib.SRMViaDSSC"
    
    Dim i As Long, TestNum As Long
    Dim lclSLng_PatLoopCnt As New SiteLong
    Dim lclSLng_TestCaseCnt As New SiteLong
    Dim lclLng_SRMType As Long
    Dim lclLng_InitPatSRMIndex As Long
    Dim lclStr_TestCase As String
    Dim lclVar_Site As Variant
    
    lclLng_SRMType = SRMCheckListInfo.SelSrmDic(UCase(SrmInfo.SRMType))
    lclLng_InitPatSRMIndex = SrmInfo.InitPatSRMIndex
    'TestNum = TheExec.sites.item(lclVar_Site).TestNumber
''==============================================================================================
''Set Local variables value
''==============================================================================================
    For Each lclVar_Site In TheExec.sites.Active
        lclSLng_PatLoopCnt(lclVar_Site) = TheExec.sites(lclVar_Site).SiteVariableValue(glbConstVar_PatLoopCnt)
        lclSLng_TestCaseCnt(lclVar_Site) = TheExec.sites(lclVar_Site).SiteVariableValue(glbConstVar_TestCaseCnt)
        lclStr_TestCase = SRMCheckListInfo.TestCase(SRMCheckListInfo.Check(lclSLng_TestCaseCnt))
        SrmInfo.AllPatsArr(lclLng_InitPatSRMIndex) = SRMCheckListInfo.SrmPattern(lclLng_SRMType - 1).patterns(lclSLng_PatLoopCnt(lclVar_Site))
''==============================================================================================
''Set HighLowLimit by difference DSSC pattern and test case
''==============================================================================================
        If SrmInfo.AllPatsArr(SrmInfo.InitPatSRMIndex) Like glbConstKeyWD_DSRM Then
            SRMCheckListInfo.HighLowLimit = AutoBinStr2DecSRM(lclStr_TestCase)
        ElseIf SrmInfo.AllPatsArr(SrmInfo.InitPatSRMIndex) Like glbConstKeyWD_SRM Then
            SRMCheckListInfo.HighLowLimit = AutoBinStr2DecSRM(lclStr_TestCase)
        Else
            TheExec.Datalog.WriteComment "Not DSRM or SRM pattern naming, please check pattern."
            Exit Function
        End If
        SRMCheckListInfo.CurrentSRMData = lclStr_TestCase
        SRMCheckListInfo.CurrentSRMPat = SrmInfo.AllPatsArr(lclLng_InitPatSRMIndex)
        Exit For
    Next lclVar_Site

    Call TheExec.Datalog.SetDynamicTestName(SrmInfo.SRMType & " " & SRMCheckListInfo.CurrentSRMPat, False)
''==============================================================================================
'' Run pattern, HRAM setup capture on first fail, SetupDigSrc
''==============================================================================================
    Call InitHRAM
    For i = 0 To UBound(SrmInfo.AllPatsArr)
        If SrmInfo.AllPatsArr(i) Like glbConstKeyWD_DSRM Or SrmInfo.AllPatsArr(i) Like glbConstKeyWD_SRM Then
            Call SetupDigSrc(SrmInfo.AllPatsArr(i), Len(lclStr_TestCase), lclStr_TestCase, glbPinList_DigSrcPin)
        End If
        SRMTestPattern SrmInfo.AllPatsArr(i)
    Next i
''==============================================================================================
    Exit Function
errHandler:
    If AbortTest Then Exit Function Else Resume Next
End Function
''==============================================================================================
''DSRMHC/SRMHC
''==============================================================================================
Public Function SRMViaPAT()
    On Error GoTo errHandler:: Dim funcName As String:: funcName = "Class_SRMLib.SRMViaPAT"
    
    Dim i As Long, TestNum As Long
    Dim lclSLng_PatLoopCnt As New SiteLong
    Dim lclLng_SRMType As Long
    Dim lclLng_InitPatSRMIndex As Long
    Dim lclVar_Site As Variant
    Dim lclStrArr_Temp() As String:: ReDim lclStrArr_Temp(glbLng_PatCnt)
    Dim lclStr_StaticData As String
    
    lclLng_SRMType = SRMCheckListInfo.SelSrmDic(UCase(SrmInfo.SRMType))
    lclLng_InitPatSRMIndex = SrmInfo.InitPatSRMIndex
    'TestNum = TheExec.sites.item(lclVar_Site).TestNumber
''==============================================================================================
''Set Local variables value
''==============================================================================================
    For Each lclVar_Site In TheExec.sites.Active
        lclSLng_PatLoopCnt(lclVar_Site) = TheExec.sites(lclVar_Site).SiteVariableValue(glbConstVar_PatLoopCnt)
        SrmInfo.AllPatsArr(lclLng_InitPatSRMIndex) = SRMCheckListInfo.SrmPattern(lclLng_SRMType - 1).patterns(lclSLng_PatLoopCnt(lclVar_Site))
        SRMCheckListInfo.CurrentSRMPat = SrmInfo.AllPatsArr(lclLng_InitPatSRMIndex)
        If SrmInfo.AllPatsArr(SrmInfo.InitPatSRMIndex) Like glbConstKeyWD_DSRM Then
            lclStrArr_Temp = Split(Replace(SRMCheckListInfo.CurrentSRMPat, Replace(glbConstKeyWD_DSRM, "*", ""), ","), ",")
            lclStr_StaticData = auto_Hex2BinStr(mid(lclStrArr_Temp(1), 1, 2), SRMCheckListInfo.SrcBitSize(0))
            SRMCheckListInfo.HighLowLimit = AutoBinStr2DecSRM(lclStr_StaticData)
        ElseIf SrmInfo.AllPatsArr(SrmInfo.InitPatSRMIndex) Like glbConstKeyWD_SRM Then
            lclStrArr_Temp = Split(Replace(SRMCheckListInfo.CurrentSRMPat, Replace(glbConstKeyWD_SRM, "*", ""), ","), ",")
            lclStr_StaticData = auto_Hex2BinStr(mid(lclStrArr_Temp(1), 1, 2), SRMCheckListInfo.SrcBitSize(0))
            SRMCheckListInfo.HighLowLimit = AutoBinStr2DecSRM(lclStr_StaticData)
        ElseIf SrmInfo.AllPatsArr(SrmInfo.InitPatSRMIndex) Like glbConstKeyWD_SELSRM Then
            lclStrArr_Temp = Split(Replace(SRMCheckListInfo.CurrentSRMPat, Replace(glbConstKeyWD_SELSRM, "*", ""), ","), ",")
            lclStr_StaticData = lclStrArr_Temp(1)
            SRMCheckListInfo.HighLowLimit = CInt(lclStr_StaticData)
        Else
            TheExec.Datalog.WriteComment "Not DSRM or SRM pattern naming, please check pattern."
            Exit Function
        End If
        SRMCheckListInfo.CurrentSRMData = lclStr_StaticData
        Exit For
    Next lclVar_Site
    
    Call TheExec.Datalog.SetDynamicTestName(SrmInfo.SRMType & " " & SRMCheckListInfo.CurrentSRMPat, False)
''==============================================================================================
'' Run pattern, HRAM setup capture on first fail, SetupDigSrc
''==============================================================================================
    Call InitHRAM
    For i = 0 To UBound(SrmInfo.AllPatsArr)
        SRMTestPattern SrmInfo.AllPatsArr(i)
    Next i
''==============================================================================================
    Exit Function
errHandler:
    If AbortTest Then Exit Function Else Resume Next
End Function
''==============================================================================================
''SRMEXPayloadH
''==============================================================================================
Public Function SRMEXPLH()
    On Error GoTo errHandler:: Dim funcName As String:: funcName = "Class_SRMLib.SRMEXPLH"
    
    Dim i As Long
    Dim lclStr_Prefix As String
    Dim lclStr_Tname As String
    
    SRMCheckListInfo.CurrentSRMReadPat = SrmInfo.AllPatsArr(0)
    
    If SrmInfo.AllPatsArr(0) Like glbConstKeyWD_EXPL Then
        lclStr_Prefix = glbConst_SRMEXPL
    ElseIf SrmInfo.AllPatsArr(0) Like glbConstKeyWD_EXPH Then
        lclStr_Prefix = glbConst_SRMEXPH
    Else
        TheExec.Datalog.WriteComment "Not EXPL or EXPH pattern naming, please check pattern."
        Exit Function
    End If
    
    If GetForceVoltage <> "" Then
        'Call SRMChangeVMeasI(GetForceVoltage, glbInstanceName, SRMCheckListInfo.CurrentSRMData, SRMCheckListInfo.CurrentSRMPat)
        Call SRMChangeVMeasIUFP(GetForceVoltage, glbInstanceName, SRMCheckListInfo.CurrentSRMData, SRMCheckListInfo.CurrentSRMPat)
    End If
    
    lclStr_Tname = lclStr_Prefix & SRMCheckListInfo.CurrentSRMPat & " " & SRMCheckListInfo.CurrentSRMData & " " & SRMCheckListInfo.CurrentSRMReadPat
    Call TheExec.Datalog.SetDynamicTestName(lclStr_Tname, False)
    
    For i = 0 To UBound(SrmInfo.AllPatsArr)
        SRMTestPattern SrmInfo.AllPatsArr(i)
    Next i

    Exit Function
errHandler:
    If AbortTest Then Exit Function Else Resume Next
End Function
''==============================================================================================
''SRMREAD, Setup digcap and change voltage value
''==============================================================================================
Public Function SRMREAD()
    On Error GoTo errHandler:: Dim funcName As String:: funcName = "Class_SRMLib.SRMREAD"
    
    Dim i As Long
    Dim lclSDSP_OutDspWave As New DSPWave:: Set lclSDSP_OutDspWave = Nothing
    Dim lclStr_Tname As String
    Dim lclStrArr_DigCapName() As String:: ReDim lclStrArr_DigCapName(0):: lclStrArr_DigCapName(0) = glbConst_DigCapName
    Dim lclDic_DigCapInfo As New Dictionary:: lclDic_DigCapInfo.RemoveAll:: lclDic_DigCapInfo.Add lclStrArr_DigCapName(0), SRMCheckListInfo.CapBitSize(0)

    SRMCheckListInfo.CurrentSRMReadPat = SrmInfo.AllPatsArr(0)
    lclStr_Tname = Replace(glbConst_SRMREAD, " ", "") & "_" & SRMCheckListInfo.CurrentSRMReadPat
    
    If GetForceVoltage <> "" Then
        'Call SRMChangeVMeasI(GetForceVoltage, glbInstanceName, SRMCheckListInfo.CurrentSRMData, SRMCheckListInfo.CurrentSRMPat)
        Call SRMChangeVMeasIUFP(GetForceVoltage, glbInstanceName, SRMCheckListInfo.CurrentSRMData, SRMCheckListInfo.CurrentSRMPat)
    End If

    Call TheExec.Datalog.SetDynamicTestName(lclStr_Tname, False)
    
    For i = 0 To UBound(SrmInfo.AllPatsArr)
        Call SetupDigCap(SrmInfo.AllPatsArr(i), glbPinList_DigCapPin, lclSDSP_OutDspWave)
        SRMTestPattern SrmInfo.AllPatsArr(i)
    Next i
    
    Call CreateSimulateDSPWaveData(lclSDSP_OutDspWave, SRMCheckListInfo.CapBitSize(0))
    Call ProcessDSPCap(lclStrArr_DigCapName, lclSDSP_OutDspWave, lclDic_DigCapInfo, glbPinList_DigCapPin.value)
    
    Exit Function
errHandler:
    If AbortTest Then Exit Function Else Resume Next
End Function
''==============================================================================================
''CUSTOM
''==============================================================================================
Public Function CUSTOM()
    On Error GoTo errHandler:: Dim funcName As String:: funcName = "Class_SRMLib.CUSTOM"

    Dim i As Long
    Dim lclStr_StaticData As String
    Dim lclStrArr_Temp() As String
    
    For i = 0 To UBound(SrmInfo.AllPatsArr)
        If SrmInfo.AllPatsArr(i) Like glbConstKeyWD_DSRM Then
            lclStrArr_Temp = Split(Replace(SrmInfo.AllPatsArr(i), Replace(glbConstKeyWD_DSRM, "*", ""), ","), ",")
            lclStr_StaticData = auto_Hex2BinStr(mid(lclStrArr_Temp(1), 1, 2), SRMCheckListInfo.SrcBitSize(0))
            SRMCheckListInfo.CurrentSRMPat = SrmInfo.AllPatsArr(i)
            SRMCheckListInfo.CurrentSRMData = lclStr_StaticData
            Call TheExec.Datalog.SetDynamicTestName(SrmInfo.SRMType & " " & SrmInfo.AllPatsArr(i), False)
        ElseIf SrmInfo.AllPatsArr(i) Like glbConstKeyWD_SRM Then
            lclStrArr_Temp = Split(Replace(SrmInfo.AllPatsArr(i), Replace(glbConstKeyWD_SRM, "*", ""), ","), ",")
            lclStr_StaticData = auto_Hex2BinStr(mid(lclStrArr_Temp(1), 1, 2), SRMCheckListInfo.SrcBitSize(0))
            SRMCheckListInfo.CurrentSRMPat = SrmInfo.AllPatsArr(i)
            SRMCheckListInfo.CurrentSRMData = lclStr_StaticData
            Call TheExec.Datalog.SetDynamicTestName(SrmInfo.SRMType & " " & SrmInfo.AllPatsArr(i), False)
        End If
    Next i
''==============================================================================================
'' Run pattern, HRAM setup capture on first fail, SetupDigSrc
''==============================================================================================
    Call InitHRAM
    For i = 0 To UBound(SrmInfo.AllPatsArr)
        SRMTestPattern SrmInfo.AllPatsArr(i)
    Next i
''==============================================================================================

    Exit Function
errHandler:
    If AbortTest Then Exit Function Else Resume Next
End Function
''==============================================================================================
''HARDIP
''==============================================================================================
Public Function HARDIP()
    On Error GoTo errHandler:: Dim funcName As String:: funcName = "Class_SRMLib.HARDIP"

    Exit Function
errHandler:
    If AbortTest Then Exit Function Else Resume Next
End Function
''==============================================================================================
''SRMTestPattern
''==============================================================================================
Public Function SRMTestPattern(ByVal InputPat As String)
    On Error GoTo errHandler:: Dim funcName As String:: funcName = "Class_SRMLib.SRMTestPattern"
'==============================================================================
'Loading pattern and test
'==============================================================================
    Call TheHdw.patterns(InputPat).Load
    Call TheHdw.patterns(InputPat).test(pfAlways, 0)
'==============================================================================
    Exit Function
errHandler:
    TheExec.Datalog.WriteComment "<Error> " + funcName + ":: Please check it out."
    If AbortTest Then Exit Function Else Resume Next
End Function
''==============================================================================================
''WriteSrmResult
''==============================================================================================
Public Property Get GetForceVoltage() As String
    GetForceVoltage = atrStr_ForceVoltage
End Property

''==============================================================================================
''SRMWriteResult
'===============================================================================================
'                                             VDD_Logic             VDD_SRAM*
'Logic/Sram           SELSRAM  DSELSRAM  Low_Limit/High_limit  Low_Limit/High_limit   Condition
'LogicH_SramL(LHSL)      0         1        0mA   /   N/A         0mA   /   N/A         valid
'LogicH_SramL(LHSL)      1         0        0mA   /   N/A         N/A   /   0mA        illegal
'LogicL_SramH(LLSH)      0         1        N/A   /   0mA         0mA   /   N/A        illegal
'LogicL_SramH(LLSH)      1         0        0mA   /   N/A         0mA   /   N/A         valid
'
'Summary each LogicPin's Scenario
'LHSL_SRAMPAT0_DSRAMPAT1   '1
'LHSL_SRAMPAT1_DSRAMPAT0   '2
'LLSH_SRAMPAT0_DSRAMPAT1   '3
'LLSH_SRAMPAT1_DSRAMPAT0   '4
'===============================================================================================
Public Function SRMWriteResult(InputInstName As String, InputMeasureCurrAllPins As PinListData, InputTempPinsVAlt() As String, InputTestName As String, Optional InputSrmPatName As String, Optional InputSramData As String, Optional InputPinCnt As Long)
    On Error GoTo errHandler:: Dim funcName As String:: funcName = "SRMWriteResult"

    Dim j As Long
    Dim lclLng_TestNum As Long
    Dim lclVar_Site As Variant

    For Each lclVar_Site In TheExec.sites
        lclLng_TestNum = TheExec.sites.item(lclVar_Site).TestNumber
        Exit For
    Next
'==============================================================================
'Print LogicPins Hi/Low test limit
'==============================================================================
    For j = 0 To UBound(SRMCheckListInfo.LogicPins)
        If Not SRMCheckListInfo.LogicPins(j).PinsName = glbConstKeyWD_PRESERVE Then
            SRMCheckListInfo.LogicPins(j).VMain = Format(TheHdw.DCVS.pins(SRMCheckListInfo.LogicPins(j).PinsName).Voltage.Main.value, "0.000")
            SRMCheckListInfo.LogicPins(j).VAlt = Format(TheHdw.DCVS.pins(SRMCheckListInfo.LogicPins(j).PinsName).Voltage.Alt.value, "0.000")
            If SRMCheckListInfo.LogicPins(j).Scenario = LLSH_SRAMPAT0_DSRAMPAT1 Then
                TheExec.flow.TestLimit resultVal:=InputMeasureCurrAllPins.pins(SRMCheckListInfo.LogicPins(j).PinsName), hiVal:=0, scaletype:=scaleNone, unit:=unitAmp, formatStr:="%.3f", Tname:=InputTestName & SRMCheckListInfo.CurrentSRMPat & " " & SRMCheckListInfo.CurrentSRMData, ForceVal:=CDbl(SRMCheckListInfo.LogicPins(j).VAlt), ForceUnit:=unitVolt, ForceResults:=tlForceNone, tNum:=lclLng_TestNum
            Else
                TheExec.flow.TestLimit resultVal:=InputMeasureCurrAllPins.pins(SRMCheckListInfo.LogicPins(j).PinsName), lowVal:=0, scaletype:=scaleNone, unit:=unitAmp, formatStr:="%.3f", Tname:=InputTestName & SRMCheckListInfo.CurrentSRMPat & " " & SRMCheckListInfo.CurrentSRMData, ForceVal:=CDbl(SRMCheckListInfo.LogicPins(j).VAlt), ForceUnit:=unitVolt, ForceResults:=tlForceNone, tNum:=lclLng_TestNum
            End If
        End If
    Next j
'==============================================================================
'Print SramPins Hi/Low test limit
'==============================================================================
    For j = 0 To UBound(SRMCheckListInfo.SramPins)
        If Not SRMCheckListInfo.SramPins(j).PinsName = glbConstKeyWD_PRESERVE Then
            SRMCheckListInfo.SramPins(j).VMain = Format(TheHdw.DCVS.pins(SRMCheckListInfo.SramPins(j).PinsName).Voltage.Main.value, "0.000")
            SRMCheckListInfo.SramPins(j).VAlt = Format(TheHdw.DCVS.pins(SRMCheckListInfo.SramPins(j).PinsName).Voltage.Alt.value, "0.000")
            If SRMCheckListInfo.SramPins(j).Scenario = LHSL_SRAMPAT1_DSRAMPAT0 Then
                TheExec.flow.TestLimit resultVal:=InputMeasureCurrAllPins.pins(SRMCheckListInfo.SramPins(j).PinsName), hiVal:=0, scaletype:=scaleNone, unit:=unitAmp, formatStr:="%.3f", Tname:=InputTestName & SRMCheckListInfo.CurrentSRMPat & " " & SRMCheckListInfo.CurrentSRMData, ForceVal:=CDbl(SRMCheckListInfo.SramPins(j).VAlt), ForceUnit:=unitVolt, ForceResults:=tlForceNone, tNum:=lclLng_TestNum
            Else
                TheExec.flow.TestLimit resultVal:=InputMeasureCurrAllPins.pins(SRMCheckListInfo.SramPins(j).PinsName), lowVal:=0, scaletype:=scaleNone, unit:=unitAmp, formatStr:="%.3f", Tname:=InputTestName & SRMCheckListInfo.CurrentSRMPat & " " & SRMCheckListInfo.CurrentSRMData, ForceVal:=CDbl(SRMCheckListInfo.SramPins(j).VAlt), ForceUnit:=unitVolt, ForceResults:=tlForceNone, tNum:=lclLng_TestNum
            End If
        End If
    Next j
    'lclLng_TestNum = lclLng_TestNum + 1
'==============================================================================
    Exit Function
errHandler:
    TheExec.Datalog.WriteComment "<Error> " + funcName + ":: please check it out."
    If AbortTest Then Exit Function Else Resume Next
End Function
''==============================================================================================
''PrintDigSrcInfo
''==============================================================================================
Public Function PrintDigSrcInfo(InputPat As String, ByVal InputDigSrcBitsSize As Long, ByVal InputDigSrcBits As String, ByVal InputDigSrcPin As String)
    On Error GoTo errHandler:: Dim funcName As String:: funcName = "Class_SRMLib.PrintDigSrcInfo"

    TheExec.Datalog.WriteComment ("")
    TheExec.Datalog.WriteComment "[DigSrcInfo] pattern: " & InputPat & ", Pin: " & InputDigSrcPin & ", Size: " & InputDigSrcBitsSize & ", Bits; LSB(L) ==> MSB(R): " & InputDigSrcBits

    Exit Function
errHandler:
    TheExec.Datalog.WriteComment "<Error> " + funcName + ":: Please check it out."
    If AbortTest Then Exit Function Else Resume Next
End Function
''==============================================================================================
''PrintDigCapInfo
''==============================================================================================
Public Function PrintDigCapInfo(InputPat As String, ByVal InputDigCapBitsSize As Long, ByVal InputDigCapPin As String)
    On Error GoTo errHandler:: Dim funcName As String:: funcName = "Class_SRMLib.PrintDigCapInfo"

    TheExec.Datalog.WriteComment ("")
    TheExec.Datalog.WriteComment "[DigCapInfo] pattern: " & InputPat & ", Pin: " & InputDigCapPin & ", Size: " & InputDigCapBitsSize

    Exit Function
errHandler:
    TheExec.Datalog.WriteComment "<Error> " + funcName + ":: Please check it out."
    If AbortTest Then Exit Function Else Resume Next
End Function
''==============================================================================================
''SetupDigSrc
''==============================================================================================
Public Function SetupDigSrc(InputPat As String, ByVal InputDigSrcBitsSize As Long, ByVal InputDigSrcBits As String, ByVal InputDigSrcPin As String)
    On Error GoTo errHandler:: Dim funcName As String:: funcName = "Class_SRMLib.SetupDigSrc"
    
    Dim i As Long
    Dim lclVar_Site As Variant
    Dim lclDSP_DigSrcWave As New DSPWave
    Dim lclStrArr_Pat() As String
    Dim lclLng_PatCnt As Long
    Dim lclStr_WaveDef As String
    Dim tempVarArray As Variant
    Dim lclStr_SrcSignalName As String:: lclStr_SrcSignalName = glbConst_SrcSignalName
    
    Call PrintDigSrcInfo(InputPat, InputDigSrcBitsSize, InputDigSrcBits, InputDigSrcPin)
    Call PATT_GetPatListFromPatternSet(InputPat, lclStrArr_Pat, lclLng_PatCnt)
    InputPat = lclStrArr_Pat(0)
    
    tempVarArray = TheHdw.DSSC.pins(InputDigSrcPin).Pattern(InputPat).Source.Labels.list
    If tempVarArray(0) <> "" Then lclStr_SrcSignalName = tempVarArray(0)
    
    Set lclDSP_DigSrcWave = Nothing
    lclDSP_DigSrcWave.CreateConstant 0, InputDigSrcBitsSize
    
    For i = 0 To InputDigSrcBitsSize - 1
       For Each lclVar_Site In TheExec.sites.Active
           lclDSP_DigSrcWave.Element(i) = CDbl(mid(InputDigSrcBits, i + 1, 1))
       Next lclVar_Site
    Next i

    lclStr_WaveDef = "WaveDef" & lclStr_SrcSignalName
    
    TheHdw.patterns(InputPat).Load
    TheHdw.DSSC.pins(InputDigSrcPin).Pattern(InputPat).Source.Signals.Add lclStr_SrcSignalName
    
    For Each lclVar_Site In TheExec.sites
        TheExec.WaveDefinitions.CreateWaveDefinition lclStr_WaveDef & lclVar_Site, lclDSP_DigSrcWave, True
        With TheHdw.DSSC.pins(InputDigSrcPin).Pattern(InputPat).Source.Signals(lclStr_SrcSignalName)
            .WaveDefinitionName = lclStr_WaveDef & lclVar_Site
'            .SampleSize = SegmentSize
'            .Amplitude = 1
'            .Amplitude = pc_Def_DSSC_Amplitude
'            .LoadSamples
        End With
    Next

    TheHdw.DSSC.pins(InputDigSrcPin).Pattern(InputPat).Source.Signals(lclStr_SrcSignalName).SampleSize = InputDigSrcBitsSize
    TheHdw.DSSC.pins(InputDigSrcPin).Pattern(InputPat).Source.Signals(lclStr_SrcSignalName).Amplitude = pc_Def_DSSC_Amplitude
    TheHdw.DSSC.pins(InputDigSrcPin).Pattern(InputPat).Source.Signals(lclStr_SrcSignalName).LoadSettings
    TheHdw.DSSC.pins(InputDigSrcPin).Pattern(InputPat).Source.Signals.DefaultSignal = lclStr_SrcSignalName
    
    Exit Function
errHandler:
    TheExec.Datalog.WriteComment "<Error> " + funcName + ":: Please check it out."
    If AbortTest Then Exit Function Else Resume Next
End Function
''==============================================================================================
''SetupDigCap
''==============================================================================================
Public Function SetupDigCap(InputPat As String, Input_DigCapPin As PinList, Output_DspWave As DSPWave) As Long
    On Error GoTo errHandler:: Dim funcName As String:: funcName = "Class_SRMLib.SetupDigCap"

    Dim i As Long
    Dim lclLng_PatCnt As Long
    Dim lclStrArr_Pat() As String
    Dim lclStr_PatName As String:: lclStr_PatName = ""
    Dim lclStr_CapSignalName As String:: lclStr_CapSignalName = glbConst_CapSignalName
    Dim lclVar_Site As Variant
    Dim lclLng_CapBitSize As Long:: lclLng_CapBitSize = SRMCheckListInfo.CapBitSize(0)

    Call AnalyzePatName(InputPat, lclStr_PatName)
    Call PrintDigCapInfo(InputPat, lclLng_CapBitSize, Input_DigCapPin)

    Output_DspWave.CreateConstant 0, lclLng_CapBitSize

    Call PATT_GetPatListFromPatternSet(InputPat, lclStrArr_Pat, lclLng_PatCnt)
    InputPat = lclStrArr_Pat(0)
    TheHdw.patterns(InputPat).Load
    TheHdw.DSSC.pins(Input_DigCapPin).Pattern(InputPat).Capture.Signals.Add lclStr_CapSignalName

    With TheHdw.DSSC.pins(Input_DigCapPin).Pattern(InputPat).Capture.Signals(lclStr_CapSignalName)
        .SampleSize = lclLng_CapBitSize
        .LoadSettings
    End With

    Output_DspWave = TheHdw.DSSC.pins(Input_DigCapPin).Pattern(InputPat).Capture.Signals(lclStr_CapSignalName).DSPWave

    For Each lclVar_Site In TheExec.sites
        Output_DspWave(lclVar_Site).info.WaveName = Input_DigCapPin
    Next lclVar_Site


    Exit Function
errHandler:
    TheExec.Datalog.WriteComment "<Error> " + funcName + ":: Please check it out."
    If AbortTest Then Exit Function Else Resume Next
End Function
'==============================================================================================
'ProcessDSPCap
'==============================================================================================
Public Function ProcessDSPCap(InputDigCapName() As String, InputDSPWave As DSPWave, InputDigCapInfoDic As Dictionary, InputDigCapPin As String) As Long
    On Error GoTo errHandler:: Dim funcName As String:: funcName = "Class_SRMLib.ProcessDSPCap"
        
        Dim i As Long
        Dim lclDbl_Bin2Dec As Double
        Dim lclLng_StartNum As Long:: lclLng_StartNum = 0
        Dim lclLng_OutPutLen As Long
        Dim lclLng_CaptBits As Long
        Dim lclStr_DigCapOut As String
        Dim lclStr_Prefix As String:: lclStr_Prefix = glbConst_SRMREAD
        Dim lclStr_Tname As String:: lclStr_Tname = lclStr_Prefix & SRMCheckListInfo.CurrentSRMPat & " " & SRMCheckListInfo.CurrentSRMData & " " & SRMCheckListInfo.CurrentSRMReadPat
        Dim lclStr_DigCapSgmtName As String
        Dim lclStr_OutPutSgmtName As String
        Dim lclVar_Site As Variant
        Dim lclVar_CapName As Variant

        For Each lclVar_Site In TheExec.sites
            lclLng_StartNum = 0
            For i = 0 To InputDSPWave(lclVar_Site).SampleSize - 1
                lclStr_DigCapOut = lclStr_DigCapOut & CStr(InputDSPWave(lclVar_Site).Element(i))
            Next i
            
            For Each lclVar_CapName In InputDigCapName
                If InputDigCapInfoDic.Exists(lclVar_CapName) Then
                    lclStr_DigCapSgmtName = ""
                    lclDbl_Bin2Dec = 0
                    lclLng_CaptBits = CLng(InputDigCapInfoDic.item(lclVar_CapName))
                    For i = lclLng_StartNum To (lclLng_StartNum + lclLng_CaptBits - 1)
                        lclStr_DigCapSgmtName = lclStr_DigCapSgmtName & CStr(InputDSPWave(lclVar_Site).Element(i))
                    Next i
                    lclLng_StartNum = Len(lclStr_DigCapSgmtName) - 1
                    For i = 0 To lclLng_StartNum
                        lclDbl_Bin2Dec = CDbl(lclDbl_Bin2Dec + mid(lclStr_DigCapSgmtName, (Len(lclStr_DigCapSgmtName) - i), 1) * 2 ^ i)
                    Next i
                    lclStr_OutPutSgmtName = lclVar_CapName & "(OutPutString:" & lclStr_DigCapSgmtName & ")"
                    TheExec.flow.TestLimit resultVal:=lclDbl_Bin2Dec, ForceResults:=tlForceNone, Tname:=lclStr_Tname, _
                                        PinName:=InputDigCapPin, lowVal:=SRMCheckListInfo.HighLowLimit, hiVal:=SRMCheckListInfo.HighLowLimit
                    lclLng_StartNum = lclLng_StartNum + lclLng_CaptBits
                End If
            Next lclVar_CapName
        Next lclVar_Site
    Exit Function
errHandler:
    TheExec.Datalog.WriteComment "<Error> " + funcName + ":: please check it out."
    If AbortTest Then Exit Function Else Resume Next
End Function
''==============================================================================================
''CreateSimulateDSPWaveData
''==============================================================================================
Public Function CreateSimulateDSPWaveData(OutputDspWave As DSPWave, InputDigCapSize As Long)
    On Error GoTo errHandler:: Dim funcName As String:: funcName = "Class_SRMLib.CreateSimulateDSPWaveData"
    
    Dim i As Integer
    Dim lclVar_Site As Variant
    Dim lclSVar_TempDSPData As New SiteVariant
    If TheExec.TesterMode = testModeOffline Then
        If InputDigCapSize <> 0 Then
            For Each lclVar_Site In TheExec.sites
                For i = 0 To InputDigCapSize - 1
                    If site = 0 Then
                        If i Mod 3 = 0 Then
                            OutputDspWave(lclVar_Site).Element(i) = 1
                        Else
                            OutputDspWave(lclVar_Site).Element(i) = 0
                        End If
                    Else
                        If i Mod 2 = 0 Then
                            OutputDspWave(lclVar_Site).Element(i) = 0
                        Else
                            OutputDspWave(lclVar_Site).Element(i) = 1
                        End If
                    End If
                Next i
            Next lclVar_Site
        Else
            For Each lclVar_Site In TheExec.sites
                If lclVar_Site = 0 Then
                    OutputDspWave(lclVar_Site).Element(0) = 1
                    OutputDspWave(lclVar_Site).Element(1) = 1
                    OutputDspWave(lclVar_Site).Element(2) = 1
                    OutputDspWave(lclVar_Site).Element(3) = 1
                    OutputDspWave(lclVar_Site).Element(4) = 1
                    OutputDspWave(lclVar_Site).Element(5) = 0
                Else
                    OutputDspWave(lclVar_Site).Element(0) = 1
                    OutputDspWave(lclVar_Site).Element(1) = 0
                    OutputDspWave(lclVar_Site).Element(2) = 0
                    OutputDspWave(lclVar_Site).Element(3) = 0
                    OutputDspWave(lclVar_Site).Element(4) = 0
                    OutputDspWave(lclVar_Site).Element(5) = 1
                End If
            Next lclVar_Site
        End If
        
        For Each lclVar_Site In TheExec.sites
            For i = 0 To OutputDspWave(lclVar_Site).SampleSize - 1
                lclSVar_TempDSPData(lclVar_Site) = lclSVar_TempDSPData(lclVar_Site) & CStr(OutputDspWave(lclVar_Site).Element(i))
            Next i
            TheExec.Datalog.WriteComment ("Site_" & lclVar_Site & " simulate data = " & lclSVar_TempDSPData(lclVar_Site))
        Next lclVar_Site
    End If
    Exit Function
errHandler:
    TheExec.Datalog.WriteComment "<Error> " + funcName + ":: Please check it out."
    If AbortTest Then Exit Function Else Resume Next
End Function
''==============================================================================================
''InitHRAM
''==============================================================================================
Public Function InitHRAM()
    On Error GoTo errHandler:: Dim funcName As String:: funcName = "Class_SRMLib.InitHRAM"
    
    TheHdw.Digital.Patgen.HaltMode = tlHaltOnOpcode
    TheHdw.Digital.hram.size = 512
    TheHdw.Digital.hram.CaptureType = captFail
    TheHdw.Digital.hram.SetTrigger trigFirst, True, 0, True
    Exit Function
errHandler:
    TheExec.Datalog.WriteComment "<Error> " + funcName + ":: Please check it out."
    If AbortTest Then Exit Function Else Resume Next
End Function
''==============================================================================================
''SRMChangeVMeasI
''After pattern done, change logic/sram voltage, check Pin's state and Meas I.
''InputForceVoltage = VDD_SOC,VDD_GPU,VDD_DISP,VDD_DCS_DDR,VDD_SRAM_GPU,VDD_SRAM|logic:-0.2;sram:+0.2
''==============================================================================================
Public Sub SRMChangeVMeasI(InputForceVoltage As String, InputInstName As String, InputDigSrcEQTemp As String, InputInitPat As String)
    On Error GoTo errHandler:: Dim funcName As String:: funcName = "SRMChangeVMeasI"
''==============================================================================================
''Initialize
''==============================================================================================
    Dim var As Variant
    Dim i As Long
    Dim lclLng_PinCnt As Long
    Dim lclLng_SramPinsIndex As Long:: lclLng_SramPinsIndex = 1
    Dim lclLng_LogicPinsIndex As Long:: lclLng_LogicPinsIndex = 0
    Dim lclLng_SrmPairIndex As Long:: lclLng_SrmPairIndex = 0
    Dim lclLng_SrmPairForceIndex As Long:: lclLng_SrmPairForceIndex = 1
    Dim lclLng_VoltageIndex As Long:: lclLng_VoltageIndex = 1
    Dim lclDbl_IFactor As Double
    Dim lclDbl_SRAMVolt As Double
    Dim lclDbl_LogicVolt As Double
    Dim lclDbl_SelPlusVol As Double
    Dim lclDbl_PinFoldLimit As Double
    
    Dim lclBool_GateState As Boolean
    
    Dim lclStr_Temp As String
    Dim lclStr_GateState As String
    Dim lclStr_SlotType As String
    Dim lclStr_Testame As String
    
    Dim lclStr_DcCategory As String
    Dim lclStr_DcSelector As String
    Dim lclStr_AcCategory As String
    Dim lclStr_AcSelector As String
    Dim lclStr_TimeSetSheet As String
    Dim lclStr_EdgeSetSheet As String
    Dim lclStr_LevelsSheet As String
    Dim lclStr_Overlay As String
    
    Dim lclStr_ChannelType As String
    Dim lclStr_VHDVSPins As String: lclStr_VHDVSPins = ""
    Dim lclStr_HEXVSPins As String: lclStr_HEXVSPins = ""
    
    Dim lclStrArr_Pins() As String
    Dim lclStrArr_VoltInfo() As String
    Dim lclStrArr_TempVolt() As String
    Dim lclStrArr_ChangePinsInfo() As String
    
    Dim lclPinListData_ALLMeasIResult As New PinListData
    Dim lclPinListData_VHDVSMeasResult As New PinListData
    Dim lclPinListData_HEXVSMeasResult As New PinListData
''==============================================================================================
''Get current context for VRS
''==============================================================================================
    TheExec.DataManager.GetInstanceContext lclStr_DcCategory, lclStr_DcSelector, lclStr_AcCategory, lclStr_AcSelector, lclStr_TimeSetSheet, lclStr_EdgeSetSheet, lclStr_LevelsSheet, lclStr_Overlay
    lclStrArr_ChangePinsInfo = Split(InputForceVoltage, "|")
''==============================================================================================
''Get Change Value from argument.
''==============================================================================================
    lclStrArr_VoltInfo = Split(lclStrArr_ChangePinsInfo(lclLng_SrmPairForceIndex), ";")
    lclStrArr_TempVolt = Split(lclStrArr_VoltInfo(lclLng_LogicPinsIndex), ":")
    lclDbl_LogicVolt = CDbl(lclStrArr_TempVolt(lclLng_VoltageIndex))
    
    lclStrArr_TempVolt = Split(lclStrArr_VoltInfo(lclLng_SramPinsIndex), ":")
    lclDbl_SRAMVolt = CDbl(lclStrArr_TempVolt(lclLng_VoltageIndex))

    If lclDbl_LogicVolt > lclDbl_SRAMVolt Then
        lclStr_Testame = glbConst_SRMCurrentLHSL
        If Not InputInstName Like "*Logic*SRAM*" Then
            InputInstName = InputInstName & "_" & Replace(glbConstKeyWD_LOGICH_SRAML, "*", "")
        End If
    Else
        lclStr_Testame = glbConst_SRMCurrentLLSH
        If Not InputInstName Like "*Logic*SRAM*" Then
            InputInstName = InputInstName & "_" & Replace(glbConstKeyWD_LOGICL_SRAMH, "*", "")
        End If
    End If
''==============================================================================================
''Pre-setting: By instruments to sort pins.(HexVS/VHDVS)
''T_Sic
''VHDVS: VDD_DISP , VDD_AVE, VDD_DCS_DDR, VDD_SRAM_GPU, VDD_SRAM_SOC
''HexVS: VDD_GPU,VDD_SOC
''T_Ton
''VHDVS: VDD_DISP,VDD_GPU,VDD_SOC
''HexVS: VDD_DCS_DDR,VDD_SRAM_GPU,VDD_SRAM
''==============================================================================================
    TheExec.DataManager.DecomposePinList lclStrArr_ChangePinsInfo(lclLng_SrmPairIndex), lclStrArr_Pins, lclLng_PinCnt
    For i = 0 To lclLng_PinCnt - 1
        If TheExec.DataManager.ChannelType(lclStrArr_Pins(i)) <> "N/C" Then
            lclStr_SlotType = LCase(GetInstrument(lclStrArr_Pins(i), 0))
            lclPinListData_ALLMeasIResult.AddPin lclStrArr_Pins(i)
            Select Case UCase(lclStr_SlotType)
                Case glbConstIns_HEXVS
                    lclPinListData_HEXVSMeasResult.AddPin lclStrArr_Pins(i)
                    If lclStr_HEXVSPins = "" Then
                        lclStr_HEXVSPins = lclStrArr_Pins(i)
                    Else
                        lclStr_HEXVSPins = lclStr_HEXVSPins & "," & lclStrArr_Pins(i)
                    End If
                Case glbConstIns_VHDVS
                    lclPinListData_VHDVSMeasResult.AddPin lclStrArr_Pins(i)
                    If lclStr_VHDVSPins = "" Then
                        lclStr_VHDVSPins = lclStrArr_Pins(i)
                    Else
                        lclStr_VHDVSPins = lclStr_VHDVSPins & "," & lclStrArr_Pins(i)
                    End If
            End Select
        End If
    Next i
''==============================================================================================
''Check Pin's Gate state before measurement, Pins setup then apply voltage to VAlt,
''==============================================================================================
    TheExec.Datalog.WriteComment "********************** START **********************"
    TheExec.Datalog.WriteComment "[GATE STATE CHECK *BEFORE* MEASURE CURRENT]"
    For i = 0 To lclLng_PinCnt - 1
        If TheExec.DataManager.ChannelType(lclStrArr_Pins(i)) <> "N/C" Then
            lclStr_SlotType = GetInstrument(lclStrArr_Pins(i), 0)
            lclStr_ChannelType = TheExec.DataManager.ChannelType(lclStrArr_Pins(i))
            If IsNumeric(right(lclStr_ChannelType, 1)) Then
                lclDbl_IFactor = CDbl(right(lclStr_ChannelType, 1))
            Else
                lclDbl_IFactor = 1
            End If
           ' Debug.Print lclStrArr_Pins(i) & ":   " & lclStr_ChannelType
            'Debug.Print lclStrArr_Pins(i) & ":   " & lclStr_SlotType
            lclDbl_PinFoldLimit = TheHdw.DCVS.pins(lclStrArr_Pins(i)).CurrentLimit.Source.FoldLimit.level.value
            lclDbl_SelPlusVol = IIf(lclStrArr_Pins(i) Like "*SRAM*", lclDbl_SRAMVolt, lclDbl_LogicVolt)
            If UCase(lclStr_SlotType) = glbConstIns_VHDVS Then
                lclDbl_IFactor = 0.7 * lclDbl_IFactor
            ElseIf UCase(lclStr_SlotType) = glbConstIns_HEXVS Then
                lclDbl_IFactor = 15 * lclDbl_IFactor
            Else
                 TheExec.ErrorLogMessage "Instrument " & lclStr_SlotType & " for pin " & lclStrArr_Pins(i) & " is not supported in Shmoo_Restore_Power_per_site_Vbump"
                 Exit For
            End If
            lclDbl_IFactor = IIf(lclDbl_IFactor < lclDbl_SelPlusVol, lclDbl_IFactor, lclDbl_PinFoldLimit)
            TheHdw.DCVS.pins(lclStrArr_Pins(i)).Voltage.Alt.value = TheExec.Specs.DC.item(lclStrArr_Pins(i) & "_VAR").Categories(lclStr_DcCategory).Selectors.item(lclStr_DcSelector).ContextValue + lclDbl_SelPlusVol
            lclBool_GateState = TheHdw.DCVS.pins(lclStrArr_Pins(i)).Gate
            With TheHdw.DCVS.pins(lclStrArr_Pins(i))
                .Meter.mode = tlDCVSMeterCurrent
                .SetCurrentRanges lclDbl_IFactor, lclDbl_IFactor
            End With
'            Select Case UCase(lclStr_SlotType)
'                Case glbConstIns_VHDVS
'                    TheHdw.DCVS.Pins(lclStrArr_Pins(i)).Voltage.Alt.value = TheExec.Specs.DC.item(lclStrArr_Pins(i) & "_VAR").Categories(lclStr_DcCategory).Selectors.item(lclStr_DcSelector).ContextValue + lclDbl_SelPlusVol
'                    lclBool_GateState = TheHdw.DCVS.Pins(lclStrArr_Pins(i)).Gate
'                    With TheHdw.DCVS.Pins(lclStrArr_Pins(i))
'                        .Meter.mode = tlDCVSMeterCurrent
'                        .SetCurrentRanges 0.2 * lclDbl_IFactor, 0.2 * lclDbl_IFactor
'                    End With
'                Case glbConstIns_HEXVS
'                    TheHdw.DCVS.Pins(lclStrArr_Pins(i)).Voltage.Alt.value = TheExec.Specs.DC.item(lclStrArr_Pins(i) & "_VAR").Categories(lclStr_DcCategory).Selectors.item(lclStr_DcSelector).ContextValue + lclDbl_SelPlusVol
'                    lclBool_GateState = TheHdw.DCVS.Pins(lclStrArr_Pins(i)).Gate
'                    With TheHdw.DCVS.Pins(lclStrArr_Pins(i))
'                        .Meter.mode = tlDCVSMeterCurrent
'                        .SetCurrentRanges 15 * lclDbl_IFactor, 15 * lclDbl_IFactor
'                    End With
'                Case Else
'                    TheExec.ErrorLogMessage "Instrument " & lclStr_SlotType & " for pin " & lclStrArr_Pins(i) & " is not supported in Shmoo_Restore_Power_per_site_Vbump"
'            End Select
            Select Case lclBool_GateState
                 Case True: lclStr_GateState = "GATE ON"
                 Case False: lclStr_GateState = "GATE OFF"
            End Select
            lclStr_Temp = FormatNumericDatalog(lclStrArr_Pins(i), 15, False) & "= " & lclStr_GateState
            TheExec.Datalog.WriteComment lclStr_Temp
            lclStr_Temp = ""
        End If
    Next i
    TheExec.Datalog.WriteComment "[GATE STATE CHECK *BEFORE* MEASURE CURRENT]"
    TheExec.Datalog.WriteComment "**********************  END  **********************"
''==============================================================================================
''Print voltage setting
''==============================================================================================
    Dim lclStr_PinsVMain As String: lclStr_PinsVMain = ""
    Dim lclStr_PinsVAlt As String: lclStr_PinsVAlt = ""
    Dim lclDbl_GetPinsVolt As Double: lclDbl_GetPinsVolt = 0
    Dim lclStrArr_InputTempPinsVAlt() As String
    ReDim lclStrArr_InputTempPinsVAlt(lclLng_PinCnt - 1)
    
    For i = 0 To UBound(lclStrArr_Pins)
        lclDbl_GetPinsVolt = TheHdw.DCVS.pins(lclStrArr_Pins(i)).Voltage.Main.value
        lclStr_PinsVMain = lclStr_PinsVMain & lclStrArr_Pins(i) & ":" & CStr(Format(lclDbl_GetPinsVolt, "0.000")) & "V, "
        lclDbl_GetPinsVolt = TheHdw.DCVS.pins(lclStrArr_Pins(i)).Voltage.Alt.value
        lclStr_PinsVAlt = lclStr_PinsVAlt & lclStrArr_Pins(i) & ":" & CStr(Format(lclDbl_GetPinsVolt, "0.000")) & "V, "
        lclStrArr_InputTempPinsVAlt(i) = Format(lclDbl_GetPinsVolt, "0.000")
    Next i

    TheExec.Datalog.WriteComment ""
    TheExec.Datalog.WriteComment "ORIGINAL VOLTAGE(VMain):" & lclStr_PinsVMain
    TheExec.Datalog.WriteComment "FORCE    VOLTAGE(VAlt) :" & lclStr_PinsVAlt
    TheHdw.DCVS.pins(lclStrArr_ChangePinsInfo(lclLng_SrmPairIndex)).Voltage.Output = tlDCVSVoltageAlt
    TheHdw.Wait 0.01
''==============================================================================================
''Meas I by difference instruments
''==============================================================================================
    TheHdw.DCVS.pins(lclStrArr_ChangePinsInfo(lclLng_SrmPairIndex)).Alarm(tlDCVSAlarmOverrange) = tlAlarmOff
    TheHdw.Wait (1 * ms)
    If lclStr_HEXVSPins <> "" Then
        lclPinListData_HEXVSMeasResult = TheHdw.DCVS.pins(lclStr_HEXVSPins).Meter.Read(tlStrobe, 10)
    End If
    TheHdw.Wait (1 * ms)
    If lclStr_VHDVSPins <> "" Then
        lclPinListData_VHDVSMeasResult = TheHdw.DCVS.pins(lclStr_VHDVSPins).Meter.Read(tlStrobe, 10)
    End If
''==============================================================================================
''Print Change Voltage and Hi/Low limit.
''==============================================================================================
    For i = 0 To lclLng_PinCnt - 1
        If i < lclPinListData_HEXVSMeasResult.pins.Count Then
            lclPinListData_ALLMeasIResult.pins(lclPinListData_HEXVSMeasResult.pins(i).name) = lclPinListData_HEXVSMeasResult.pins(i)
        Else
            lclPinListData_ALLMeasIResult.pins(lclPinListData_VHDVSMeasResult.pins(i - lclPinListData_HEXVSMeasResult.pins.Count).name) = lclPinListData_VHDVSMeasResult.pins(i - lclPinListData_HEXVSMeasResult.pins.Count)
        End If
    Next i
    Call SRMInfoScenario(InputInstName, InputInitPat, InputDigSrcEQTemp)
    Call SRMWriteResult(InputInstName, lclPinListData_ALLMeasIResult, lclStrArr_InputTempPinsVAlt, lclStr_Testame, InputInitPat, InputDigSrcEQTemp, lclLng_PinCnt)

    TheHdw.DCVS.pins(lclStrArr_ChangePinsInfo(lclLng_SrmPairIndex)).Alarm(tlDCVSAlarmOverrange) = tlAlarmDefault
''==============================================================================================
''Setup Voltage to VMain(Re-store voltage)
''==============================================================================================
    TheHdw.DCVS.pins(lclStrArr_ChangePinsInfo(lclLng_SrmPairIndex)).Voltage.Output = tlDCVSVoltageMain
    TheHdw.Wait 0.001
''==============================================================================================
''Check Pin's Gate state after measurement.
''==============================================================================================
    lclStr_GateState = ""
    TheExec.Datalog.WriteComment "********************** START **********************"
    TheExec.Datalog.WriteComment "[GATE STATE CHECK *AFTER* MEASURE CURRENT]"
    For i = 0 To UBound(lclStrArr_Pins)
        If TheExec.DataManager.ChannelType(lclStrArr_Pins(i)) <> "N/C" Then
            lclBool_GateState = TheHdw.DCVS.pins(lclStrArr_Pins(i)).Gate
            Select Case lclBool_GateState
                Case True: lclStr_GateState = "GATE ON"
                Case False: lclStr_GateState = "GATE OFF"
            End Select
            lclStr_Temp = FormatNumericDatalog(lclStrArr_Pins(i), 15, False) & "= " & lclStr_GateState
            TheExec.Datalog.WriteComment lclStr_Temp
            lclStr_Temp = ""
        End If
    Next i
    TheExec.Datalog.WriteComment "[GATE STATE CHECK *AFTER* MEASURE CURRENT]"
    TheExec.Datalog.WriteComment "**********************  END  **********************"
''==============================================================================================
    Exit Sub
errHandler:
    TheExec.Datalog.WriteComment "<Error> " + funcName + ":: please check it out."
    If AbortTest Then Exit Sub Else Resume Next
End Sub
''==============================================================================================================================
''SRMChangeVMeasIUFP
''After pattern done, change logic/sram voltage, check Pin's state and Meas I.
''InputForceVoltage = VDD_SOC,VDD_GPU,VDD_DISP,VDD_DCS_DDR,VDD_SRAM_GPU,VDD_SRAM|logic:-0.2;sram:+0.2
''==============================================================================================================================
Public Sub SRMChangeVMeasIUFP(InputForceVoltage As String, InputInstName As String, InputDigSrcEQTemp As String, InputInitPat As String)
    On Error GoTo errHandler:: Dim funcName As String:: funcName = "SRMChangeVMeasIUFP"
''==============================================================================================================================
''Initialize
''==============================================================================================================================
    Dim i As Long, j As Long
    Dim lclStr_Temp As String
    Dim lclStr_Testame As String
    
    Dim lclStr_GateState As String
    Dim lclBool_GateState As Boolean
    
    Dim lclLng_SramPinsIndex As Long:: lclLng_SramPinsIndex = 1
    Dim lclLng_LogicPinsIndex As Long:: lclLng_LogicPinsIndex = 0
    Dim lclLng_SrmPairIndex As Long:: lclLng_SrmPairIndex = 0
    Dim lclLng_SrmPairForceIndex As Long:: lclLng_SrmPairForceIndex = 1
    Dim lclLng_VoltageIndex As Long:: lclLng_VoltageIndex = 1
    
    Dim lclDbl_SRAMVolt As Double
    Dim lclDbl_LogicVolt As Double

    Dim lclStr_Pin As String
    Dim lclLng_PinCnt As Long
    Dim lclStrArr_Pins() As String
    Dim lclStrArr_VoltInfo() As String
    Dim lclStrArr_TempVolt() As String
    Dim lclStrArr_ChangePinsInfo() As String
    
    Dim lclStr_PinsVMain As String: lclStr_PinsVMain = ""
    Dim lclStr_PinsVAlt As String: lclStr_PinsVAlt = ""
    Dim lclDbl_GetPinsVolt As Double: lclDbl_GetPinsVolt = 0
    Dim lclStrArr_InputTempPinsVAlt() As String
  
    Dim lclPinListData_ALLMeasIResult As New PinListData
''==============================================================================================================================
    Dim lclStr_HEXVSPins As String::: lclStr_HEXVSPins = ""
    Dim lclStr_UVSPins As String::::: lclStr_UVSPins = ""
    Dim lclStr_VSMPins As String::::: lclStr_VSMPins = ""
    Dim lclStr_VS800MAPins As String: lclStr_VS800MAPins = ""
    Dim lclStr_VS5APins As String:::: lclStr_VS5APins = ""
    Dim lclStr_UFPPins As String::::: lclStr_UFPPins = ""
    
    Dim lclStrArr_HEXVSPins() As String
    Dim lclStrArr_UVSPins() As String
    Dim lclStrArr_VSMPins() As String
    Dim lclStrArr_VS800MAPins() As String
    Dim lclStrArr_VS5APins() As String
    Dim lclStrArr_UFPPins() As String
    
    Dim lclPinListData_HEXVSMeasIResult As New PinListData
    Dim lclPinListData_UVSMeasIResult As New PinListData
    Dim lclPinListData_VSMMeasIResult As New PinListData
    Dim lclPinListData_VS800mAMeasIResult As New PinListData
    Dim lclPinListData_VS5AMeasIResult As New PinListData
    Dim lclPinListData_UFPMeasIResult As New PinListData
''==============================================================================================================
''Get current context for VRS
''==============================================================================================================
    lclStrArr_ChangePinsInfo = Split(InputForceVoltage, "|")
''==============================================================================================================
''Get Change Value from argument.
''==============================================================================================================
    lclStrArr_VoltInfo = Split(lclStrArr_ChangePinsInfo(lclLng_SrmPairForceIndex), ";")
    lclStrArr_TempVolt = Split(lclStrArr_VoltInfo(lclLng_LogicPinsIndex), ":")
    lclDbl_LogicVolt = CDbl(lclStrArr_TempVolt(lclLng_VoltageIndex))
    
    lclStrArr_TempVolt = Split(lclStrArr_VoltInfo(lclLng_SramPinsIndex), ":")
    lclDbl_SRAMVolt = CDbl(lclStrArr_TempVolt(lclLng_VoltageIndex))

    If lclDbl_LogicVolt > lclDbl_SRAMVolt Then
        lclStr_Testame = glbConst_SRMCurrentLHSL
        If Not InputInstName Like "*Logic*SRAM*" Then
            InputInstName = InputInstName & "_" & Replace(glbConstKeyWD_LOGICH_SRAML, "*", "")
        End If
    Else
        lclStr_Testame = glbConst_SRMCurrentLLSH
        If Not InputInstName Like "*Logic*SRAM*" Then
            InputInstName = InputInstName & "_" & Replace(glbConstKeyWD_LOGICL_SRAMH, "*", "")
        End If
    End If
''==============================================================================================================
'SRM_Info.logic/srm Scenario setup
''Pre-setting: By instruments to sort pins.(HexVS/VHDVS)
''T_Sic
''VHDVS: VDD_DISP , VDD_AVE, VDD_DCS_DDR, VDD_SRAM_GPU, VDD_SRAM_SOC
''HexVS: VDD_GPU,VDD_SOC
''T_Ton
''VHDVS: VDD_DISP,VDD_GPU,VDD_SOC
''HexVS: VDD_DCS_DDR,VDD_SRAM_GPU,VDD_SRAM
''==============================================================================================================
    Call SRMInfoScenario(InputInstName, InputInitPat, InputDigSrcEQTemp)
''==============================================================================================================
''Get selsrm pins's ChanMapType from glbArr_PowerPinInfo
''==============================================================================================================
    'Dim lclType_AllPinsInfo() As PinsInfo:: lclType_AllPinsInfo = SRMCheckListInfo.AllPins
    Call GetPinsInfo(SRMCheckListInfo.LogicPins)
    Call GetPinsInfo(SRMCheckListInfo.SramPins)
    Call GetPinsInfo(SRMCheckListInfo.AllPins)
    Call GetScenario(SRMCheckListInfo.AllPins, SRMCheckListInfo.LogicPins)
    Call GetScenario(SRMCheckListInfo.AllPins, SRMCheckListInfo.SramPins)
''==============================================================================================================
    TheExec.Datalog.WriteComment "********************** START **********************"
    TheExec.Datalog.WriteComment "[GATE STATE CHECK *BEFORE* MEASURE CURRENT]"
    SortByInstruments SRMCheckListInfo.AllPins, lclStr_HEXVSPins, lclStr_UVSPins, lclStr_VSMPins, lclStr_UFPPins, lclStr_Pin, lclDbl_LogicVolt, lclDbl_SRAMVolt
    RestoreIRange SRMCheckListInfo.AllPins
    RemoveFirstChar lclStr_Pin
    TheExec.DataManager.DecomposePinList lclStr_Pin, lclStrArr_Pins, lclLng_PinCnt
    ReDim lclStrArr_InputTempPinsVAlt(UBound(lclStrArr_Pins))

    For i = 0 To UBound(lclStrArr_Pins)
        lclDbl_GetPinsVolt = TheHdw.DCVS.pins(lclStrArr_Pins(i)).Voltage.Main.value
        lclStr_PinsVMain = lclStr_PinsVMain & lclStrArr_Pins(i) & ":" & CStr(Format(lclDbl_GetPinsVolt, "0.000")) & "V, "

        lclDbl_GetPinsVolt = TheHdw.DCVS.pins(lclStrArr_Pins(i)).Voltage.Alt.value
        lclStr_PinsVAlt = lclStr_PinsVAlt & lclStrArr_Pins(i) & ":" & CStr(Format(lclDbl_GetPinsVolt, "0.000")) & "V, "
        lclStrArr_InputTempPinsVAlt(i) = Format(lclDbl_GetPinsVolt, "0.000")
    Next i
''==============================================================================================================
    'Remove first ","
    RemoveFirstChar lclStr_HEXVSPins
    RemoveFirstChar lclStr_UVSPins
    RemoveFirstChar lclStr_VSMPins
    RemoveFirstChar lclStr_UFPPins
    
    lclStrArr_HEXVSPins = Split(lclStr_HEXVSPins, ",")
    lclStrArr_UVSPins = Split(lclStr_UVSPins, ",")
    lclStrArr_VSMPins = Split(lclStr_VSMPins, ",")
    lclStrArr_UFPPins = Split(lclStr_UFPPins, ",")
'    lclStrArr_VS5APins = Split(lclStr_VS5APins, ",")
'    lclStrArr_VS800MAPins = Split(lclStr_VS800MAPins, ",")
''==============================================================================================================
''Print voltage setting
''==============================================================================================================
    TheExec.Datalog.WriteComment "[GATE STATE CHECK *BEFORE* MEASURE CURRENT]"
    TheExec.Datalog.WriteComment "**********************  END  **********************"
    TheExec.Datalog.WriteComment ""
    TheExec.Datalog.WriteComment "ORIGINAL VOLTAGE(VMain):" & lclStr_PinsVMain
    TheExec.Datalog.WriteComment "FORCE    VOLTAGE(VAlt) :" & lclStr_PinsVAlt
    TheHdw.DCVS.pins(lclStr_Pin).Voltage.Output = tlDCVSVoltageAlt
    TheHdw.Wait 0.01
''==============================================================================================================
''Meas I by difference instruments
''==============================================================================================================
    TheHdw.DCVS.pins(lclStr_Pin).Alarm(tlDCVSAlarmOverrange) = tlAlarmOff
    
    GetMeasIResult lclStr_HEXVSPins, lclPinListData_HEXVSMeasIResult, , 100
    GetMeasIResult lclStr_UVSPins, lclPinListData_UVSMeasIResult, 1, 5
    GetMeasIResult lclStr_VSMPins, lclPinListData_VSMMeasIResult, , 100
    GetMeasIResult lclStr_UFPPins, lclPinListData_UFPMeasIResult, , 5
''===============================================================================================================
''AutoRange setup
''===============================================================================================================
    Call MeasIAutoRange(SRMCheckListInfo.AllPins, lclStr_HEXVSPins, lclPinListData_HEXVSMeasIResult, lclStr_UVSPins, lclPinListData_UVSMeasIResult, lclStr_VSMPins, lclPinListData_VSMMeasIResult, lclStr_UFPPins, lclPinListData_UFPMeasIResult)
''==============================================================================================================
''Print Change Voltage and Hi/Low limit.
''==============================================================================================================
    Call CreateAllPinListData(lclStrArr_HEXVSPins, lclPinListData_HEXVSMeasIResult, lclPinListData_ALLMeasIResult)
    Call CreateAllPinListData(lclStrArr_UVSPins, lclPinListData_UVSMeasIResult, lclPinListData_ALLMeasIResult)
    Call CreateAllPinListData(lclStrArr_VSMPins, lclPinListData_VSMMeasIResult, lclPinListData_ALLMeasIResult)
    Call CreateAllPinListData(lclStrArr_UFPPins, lclPinListData_UFPMeasIResult, lclPinListData_ALLMeasIResult)
    Call SRMWriteResult(InputInstName, lclPinListData_ALLMeasIResult, lclStrArr_InputTempPinsVAlt, lclStr_Testame, InputInitPat, InputDigSrcEQTemp, lclLng_PinCnt)
''==============================================================================================================
''Setup Voltage to VMain(Re-store voltage)
''==============================================================================================================
    TheHdw.DCVS.pins(lclStr_Pin).Alarm(tlDCVSAlarmOverrange) = tlAlarmDefault
    TheHdw.DCVS.pins(lclStr_Pin).Voltage.Output = tlDCVSVoltageMain
    TheHdw.Wait 0.001
''==============================================================================================================
''Check Pin's Gate state after measurement.
''==============================================================================================================
    lclStr_GateState = ""
    TheExec.Datalog.WriteComment "********************** START **********************"
    TheExec.Datalog.WriteComment "[GATE STATE CHECK *AFTER* MEASURE CURRENT]"
    For i = 0 To UBound(lclStrArr_Pins)
        If TheExec.DataManager.ChannelType(lclStrArr_Pins(i)) <> "N/C" Then
            lclBool_GateState = TheHdw.DCVS.pins(lclStrArr_Pins(i)).Gate
            Select Case lclBool_GateState
                Case True: lclStr_GateState = "GATE ON"
                Case False: lclStr_GateState = "GATE OFF"
            End Select
            lclStr_Temp = FormatNumericDatalog(lclStrArr_Pins(i), 15, False) & "= " & lclStr_GateState
            TheExec.Datalog.WriteComment lclStr_Temp
            lclStr_Temp = ""
        End If
    Next i
    TheExec.Datalog.WriteComment "[GATE STATE CHECK *AFTER* MEASURE CURRENT]"
    TheExec.Datalog.WriteComment "**********************  END  **********************"
    
    RestoreIRange SRMCheckListInfo.AllPins
''==============================================================================================================
    Exit Sub
errHandler:
    TheExec.Datalog.WriteComment "<Error> " + funcName + ":: please check it out."
    If AbortTest Then Exit Sub Else Resume Next
End Sub

''==============================================================================================================
''MeasIAutoRange
''Measure Current AutoRange setup
''==============================================================================================================
Public Sub MeasIAutoRange(InputAllPins() As PinsInfo, Optional ByVal InputHEXVSPins As String, Optional InputHEXVSMeasIResult As PinListData, _
                        Optional ByVal InputUVSPins As String, Optional InputUVSMeasIResult As PinListData, _
                        Optional ByVal InputVSMPins As String, Optional InputVSMMeasIResult As PinListData, _
                        Optional ByVal InputUFPPins As String, Optional InputUFPMeasIResult As PinListData)
    On Error GoTo errHandler:: Dim funcName As String:: funcName = "MeasIAutoRange"
''==============================================================================================================
''Initialize
''==============================================================================================================
    Dim i As Long, j  As Long
    Dim lclLng_StepNo As Long
    Dim lclLng_StopStep As Long
    Dim lclLng_SearchStep As Long
    
    Dim lclDbl_WaitTime As Double:: lclDbl_WaitTime = 260 * us
    Dim lclDbl_SattleTime As Double
    
    Dim lclPinData_PinVal As New PinData
    Dim lclSBool_DropRng As New SiteBoolean
    
    If lclLng_SearchStep = 0 Then
        lclLng_StopStep = 5
    Else
        lclLng_StopStep = lclLng_SearchStep
    End If
''==============================================================================================================
''Initialize
''==============================================================================================================
    For j = 1 To lclLng_StopStep
        For i = 0 To UBound(InputAllPins)
            If Not InputAllPins(i).PinsName Like glbConstKeyWD_PRESERVE Then
                If (UCase(InputAllPins(i).ChanMapType) = glbConstIns_HEXVS) And ((InputAllPins(i).Scenario = 3)) And Not InputAllPins(i).PinsName Like glbConstKeyWD_SRAM Then '''illegal condition LHSL_SRAMPAT1_DSRAMPAT0, LLSH_SRAMPAT0_DSRAMPAT1
'                    TheExec.Datalog.WriteComment ("NO IRANGE AUTORANGE!!!")
                ElseIf (UCase(InputAllPins(i).ChanMapType) = glbConstIns_HEXVS) And ((InputAllPins(i).Scenario = 2)) And InputAllPins(i).PinsName Like glbConstKeyWD_SRAM Then '''illegal condition LHSL_SRAMPAT1_DSRAMPAT0, LLSH_SRAMPAT0_DSRAMPAT1
'                    TheExec.Datalog.WriteComment ("NO IRANGE AUTORANGE!!!")
                Else
                    If UCase(InputAllPins(i).ChanMapType) = glbConstIns_VSM Then
                        lclPinData_PinVal = InputVSMMeasIResult.pins(InputAllPins(i).PinsName).Abs
                    ElseIf UCase(InputAllPins(i).ChanMapType) = glbConstIns_VHDVS Then
                        lclPinData_PinVal = InputUVSMeasIResult.pins(InputAllPins(i).PinsName).Abs
                    ElseIf UCase(InputAllPins(i).ChanMapType) = glbConstIns_HEXVS Then
                        lclPinData_PinVal = InputHEXVSMeasIResult.pins(InputAllPins(i).PinsName).Abs
                    ElseIf UCase(InputAllPins(i).ChanMapType) = glbConstIns_VS800MA Or UCase(InputAllPins(i).ChanMapType) = glbConstIns_VS5A Then
                        lclPinData_PinVal = InputUFPMeasIResult.pins(InputAllPins(i).PinsName).Abs
                    End If
                    lclLng_StepNo = InputAllPins(i).Init_step - j
                    If lclLng_StepNo >= 0 Then
                        If InputAllPins(i).Range_List(lclLng_StepNo) >= 0.00002 Then
                            lclSBool_DropRng = lclPinData_PinVal.compare(LessThan, InputAllPins(i).Range_List(lclLng_StepNo) - InputAllPins(i).Accuracy_List(lclLng_StepNo))
                            If lclSBool_DropRng.Any(True) Then
                                TheExec.sites.Selected = lclSBool_DropRng
                                TheHdw.DCVS.pins(InputAllPins(i).PinsName).SetCurrentRanges InputAllPins(i).Range_List(lclLng_StepNo), InputAllPins(i).Range_List(lclLng_StepNo)
                                TheExec.sites.Selected = True
                                lclDbl_SattleTime = InputAllPins(i).WaitTime_List(lclLng_StepNo)
                                If lclDbl_SattleTime > lclDbl_WaitTime Then lclDbl_WaitTime = lclDbl_SattleTime
                            End If
                        End If
                        'If gl_Disable_SelSRAM_debug_log = False Then
                            'TheExec.Datalog.WriteComment (TheExec.DataManager.instanceName & " =====> Curr_meas Meter I range setting, " & InputAllPins(i).PinsName & " =" & TheHdw.DCVS.Pins(InputAllPins(i).PinsName).Meter.CurrentRange.value)
                            ''TheExec.Datalog.WriteComment (TheExec.DataManager.instanceName & " =====> Curr_meas Meter I range Program setting, " & InputAllPins(i).PinsName & " =" & InputAllPins(i).Range_List(lclLng_StepNo))
                        'End If
                    Else
                        lclLng_StepNo = 0
                    End If
                End If
            End If
        Next i
        Set lclPinData_PinVal = Nothing
        If Not TheExec.TesterMode = testModeOffline Then
            TheHdw.Wait lclDbl_WaitTime
        End If
        GetMeasIResult InputHEXVSPins, InputHEXVSMeasIResult, , 100
        GetMeasIResult InputUVSPins, InputUVSMeasIResult, 1, 5
        GetMeasIResult InputVSMPins, InputVSMMeasIResult, , 100
        GetMeasIResult InputUFPPins, InputUFPMeasIResult, , 5
    Next j

    Exit Sub
errHandler:
    TheExec.Datalog.WriteComment "<Error> " + funcName + ":: please check it out."
    If AbortTest Then Exit Sub Else Resume Next
End Sub
''==============================================================================================
''AutoBinStr2DecDSRM
''==============================================================================================
Public Function AutoBinStr2DecDSRM(bin As String) As Integer
On Error GoTo errHandler
    Dim funcName As String:: funcName = "AutoBinStr2DecDSRM"

    Dim tempStr As String
    Dim mL_Loop As Long
'==============================================================================
'Capture bits (MSB -> LSB). Need confirm which bit need readback twice.
'==============================================================================
'    Dim BitsArr() As String
'    Dim i As Long, j As Long, ArrSize As Long
'    For i = 1 To Len(Bin)
'        For j = 0 To 2 'Times of bits repeat.Nedd add to SRMCheckListInfo.
'            ReDim Preserve BitsArr(ArrSize)
'            ArrSize = ArrSize + 1
'            BitsArr(ArrSize - 1) = Mid(Bin, i, 1)
'        Next
'    Next
'    TempStr = Join(BitsArr, "")
'==============================================================================
    'TempStr = Bin & Mid(Bin, Len(Bin)) 'For T_Ton, T_Sic amd T_Bor
    'TempStr = Bin & Mid(Bin, 2, 1)   'For T_Stn
    ''TempStr = Bin   'For T_Grackle
    'TempStr = Mid(Bin, 1, 1) & Mid(Bin, 1, 1) & Mid(Bin, 2, 4) 'For Ellis
    For mL_Loop = 0 To UBound(SRMCheckListInfo.ReadLimit)
        If SRMCheckListInfo.TestCase(mL_Loop) = bin Then
            tempStr = SRMCheckListInfo.ReadLimit(mL_Loop)
            Exit For
        End If
    Next mL_Loop
'    For mL_Loop = 0 To UBound(SRMCheckListInfo.ReadBackBits)
'        TempStr = TempStr & mid(Bin, SRMCheckListInfo.ReadBackBits(mL_Loop) + 1, 1)
'    Next
'==============================================================================
    AutoBinStr2DecDSRM = (2 ^ Len(tempStr)) - 1 - auto_bitStr2Dec(tempStr)
    'Debug.Print AutoBinStr2DecDSRM

Exit Function
errHandler:
    TheExec.Datalog.WriteComment "<Error> " + funcName + ":: please check it out."
    If AbortTest Then Exit Function Else Resume Next
End Function
''==============================================================================================
''AutoBinStr2DecSRM
''==============================================================================================
Public Function AutoBinStr2DecSRM(bin As String) As Integer
On Error GoTo errHandler
    Dim funcName As String:: funcName = "AutoBinStr2DecSRM"

    Dim tempStr As String
    Dim mL_Loop As Long
    
    For mL_Loop = 0 To UBound(SRMCheckListInfo.ReadLimit)
        If SRMCheckListInfo.TestCase(mL_Loop) = bin Then
            tempStr = SRMCheckListInfo.ReadLimit(mL_Loop)
            Exit For
        End If
    Next mL_Loop

    AutoBinStr2DecSRM = CInt(tempStr)

Exit Function
errHandler:
    TheExec.Datalog.WriteComment "<Error> " + funcName + ":: please check it out."
    If AbortTest Then Exit Function Else Resume Next
End Function

''==============================================================================================
''SRMInfoScenario
'===============================================================================================
'                                             VDD_Logic             VDD_SRAM*
'Logic/Sram           SELSRAM  DSELSRAM  Low_Limit/High_limit  Low_Limit/High_limit   Condition
'LogicH_SramL(LHSL)      0         1        0mA   /   N/A         0mA   /   N/A         valid
'LogicH_SramL(LHSL)      1         0        0mA   /   N/A         N/A   /   0mA        illegal
'LogicL_SramH(LLSH)      0         1        N/A   /   0mA         0mA   /   N/A        illegal
'LogicL_SramH(LLSH)      1         0        0mA   /   N/A         0mA   /   N/A         valid
'
'Summary each LogicPin's Scenario
'LHSL_SRAMPAT0_DSRAMPAT1   '1
'LHSL_SRAMPAT1_DSRAMPAT0   '2
'LLSH_SRAMPAT0_DSRAMPAT1   '3
'LLSH_SRAMPAT1_DSRAMPAT0   '4
'===============================================================================================
Public Function SRMInfoScenario(InputInstName As String, Optional InputSrmPatName As String, Optional InputSramData As String)
    On Error GoTo errHandler:: Dim funcName As String:: funcName = "SRMInfoScenario"

    Dim j As Long
    
    For j = 0 To UBound(SRMCheckListInfo.SramPins)
        SRMCheckListInfo.SramPins(j).Scenario = LHSL_SRAMPAT0_DSRAMPAT1
    Next
    
    For j = 0 To UBound(SRMCheckListInfo.LogicPins)
    
        '20230601 split sram pins by "," for Tahiti
        Dim SrmPinTemp() As String
        Dim SrmTempIndex As Integer
        Dim SrmTempSize As Integer
        SrmPinTemp = Split(SRMCheckListInfo.LogicPins(j).SramPair, ",")
        SrmTempSize = UBound(SrmPinTemp)
        
        SRMCheckListInfo.SrcBits(j) = mid(InputSramData, SRMCheckListInfo.LogicPins(j).CheckListYIndex, 1)
        If Not SRMCheckListInfo.LogicPins(j).PinsName = glbConstKeyWD_PRESERVE Then
            If SRMCheckListInfo.SrcBits(j) Like "0" Then
                If UCase(InputInstName) Like glbConstKeyWD_LOGICH_SRAML And UCase(InputSrmPatName) Like glbConstKeyWD_SRM Then
                    SRMCheckListInfo.LogicPins(j).Scenario = LHSL_SRAMPAT0_DSRAMPAT1
                ElseIf UCase(InputInstName) Like glbConstKeyWD_LOGICH_SRAML And UCase(InputSrmPatName) Like glbConstKeyWD_DSRM Then
                    SRMCheckListInfo.LogicPins(j).Scenario = LHSL_SRAMPAT1_DSRAMPAT0
                    '20230601 SRAM pins looping
                    For SrmTempIndex = 0 To SrmTempSize
                        SRMCheckListInfo.SramPins(SRMCheckListInfo.TempDicSRM(UCase(SrmPinTemp(SrmTempIndex))) - 1).Scenario = LHSL_SRAMPAT1_DSRAMPAT0
                    Next SrmTempIndex
                    'SRMCheckListInfo.SramPins(SRMCheckListInfo.TempDicSRM(SRMCheckListInfo.LogicPins(j).SramPair) - 1).Scenario = LHSL_SRAMPAT1_DSRAMPAT0
                ElseIf UCase(InputInstName) Like glbConstKeyWD_LOGICL_SRAMH And UCase(InputSrmPatName) Like glbConstKeyWD_SRM Then
                    SRMCheckListInfo.LogicPins(j).Scenario = LLSH_SRAMPAT0_DSRAMPAT1
                ElseIf UCase(InputInstName) Like glbConstKeyWD_LOGICL_SRAMH And UCase(InputSrmPatName) Like glbConstKeyWD_DSRM Then
                    SRMCheckListInfo.LogicPins(j).Scenario = LLSH_SRAMPAT1_DSRAMPAT0
                End If
            Else    '"1"
                If UCase(InputInstName) Like glbConstKeyWD_LOGICH_SRAML And UCase(InputSrmPatName) Like glbConstKeyWD_SRM Then
                    SRMCheckListInfo.LogicPins(j).Scenario = LHSL_SRAMPAT1_DSRAMPAT0
                    '20230601 SRAM pins looping
                    For SrmTempIndex = 0 To SrmTempSize
                        SRMCheckListInfo.SramPins(SRMCheckListInfo.TempDicSRM(UCase(SrmPinTemp(SrmTempIndex))) - 1).Scenario = LHSL_SRAMPAT1_DSRAMPAT0
                    Next SrmTempIndex
'                   SRMCheckListInfo.SramPins(SRMCheckListInfo.TempDicSRM(SRMCheckListInfo.LogicPins(j).SramPair) - 1).Scenario = LHSL_SRAMPAT1_DSRAMPAT0
                ElseIf UCase(InputInstName) Like glbConstKeyWD_LOGICH_SRAML And UCase(InputSrmPatName) Like glbConstKeyWD_DSRM Then
                    SRMCheckListInfo.LogicPins(j).Scenario = LHSL_SRAMPAT0_DSRAMPAT1
                ElseIf UCase(InputInstName) Like glbConstKeyWD_LOGICL_SRAMH And UCase(InputSrmPatName) Like glbConstKeyWD_SRM Then
                    SRMCheckListInfo.LogicPins(j).Scenario = LLSH_SRAMPAT1_DSRAMPAT0
                ElseIf UCase(InputInstName) Like glbConstKeyWD_LOGICL_SRAMH And UCase(InputSrmPatName) Like glbConstKeyWD_DSRM Then
                    SRMCheckListInfo.LogicPins(j).Scenario = LLSH_SRAMPAT0_DSRAMPAT1
                End If
            End If
        End If
    Next j

    Exit Function
errHandler:
    TheExec.Datalog.WriteComment "<Error> " + funcName + ":: please check it out."
    If AbortTest Then Exit Function Else Resume Next
End Function
''==============================================================================================================
''GetInputPins
''==============================================================================================================
Public Sub GetPinsInfo(InputPins() As PinsInfo)
    On Error GoTo errHandler:: Dim funcName As String:: funcName = "GetInputPins"
''==============================================================================================================
''Initialize
''==============================================================================================================
    Dim i As Long, j As Long

    For i = 0 To UBound(InputPins)
        If Not InputPins(i).PinsName Like glbConstKeyWD_PRESERVE Then
            For j = 0 To UBound(glbArr_PowerPinInfo)
                If InputPins(i).PinsName Like glbArr_PowerPinInfo(j).PinName Then
                    InputPins(i).ChanMapType = glbArr_PowerPinInfo(j).ChanMapType
                    InputPins(i).MergedN = glbArr_PowerPinInfo(j).MergedN
                    InputPins(i).Range_List = glbArr_PowerPinInfo(j).Range_List
                    InputPins(i).WaitTime_List = glbArr_PowerPinInfo(j).WaitTime_List
                    InputPins(i).Init_step = UBound(glbArr_PowerPinInfo(j).Range_List)
                    InputPins(i).Accuracy_List = glbArr_PowerPinInfo(j).Accuracy_List
                    InputPins(i).Init_CurrentRange = glbArr_PowerPinInfo(j).Init_CurrentRange
                End If
            Next j
        End If
    Next i

    Exit Sub
errHandler:
    TheExec.Datalog.WriteComment "<Error> " + funcName + ":: please check it out."
    If AbortTest Then Exit Sub Else Resume Next
End Sub
''==============================================================================================================
''GetScenario
''==============================================================================================================
Public Sub GetScenario(InputAllPins() As PinsInfo, InputPins() As PinsInfo)
    On Error GoTo errHandler:: Dim funcName As String:: funcName = "GetScenario"
''==============================================================================================================
''Initialize
''==============================================================================================================
    Dim i As Long, j As Long

    For i = 0 To UBound(InputAllPins)
        If Not InputAllPins(i).PinsName Like glbConstKeyWD_PRESERVE Then
            For j = 0 To UBound(InputPins)
                If InputAllPins(i).PinsName Like InputPins(j).PinsName Then
                    InputAllPins(i).Scenario = InputPins(j).Scenario
                    InputAllPins(i).LogicPair = InputPins(j).LogicPair
                    InputAllPins(i).SramPair = InputPins(j).SramPair
                    Exit For
                End If
            Next j
        End If
    Next i

    Exit Sub
errHandler:
    TheExec.Datalog.WriteComment "<Error> " + funcName + ":: please check it out."
    If AbortTest Then Exit Sub Else Resume Next
End Sub
''==============================================================================================================
''CreateAllPinListData
''==============================================================================================================
Public Sub CreateAllPinListData(InputInstrumentPins() As String, InputInstrumentMeasIResult As PinListData, InputALLMeasIResult As PinListData)
    On Error GoTo errHandler:: Dim funcName As String:: funcName = "CreateAllPinListData"
''==============================================================================================================
''Initialize
''==============================================================================================================
    Dim i As Long

    For i = 0 To UBound(InputInstrumentPins)
        InputALLMeasIResult.AddPin (InputInstrumentPins(i))
        InputALLMeasIResult.pins(InputInstrumentPins(i)) = InputInstrumentMeasIResult.pins(InputInstrumentPins(i))
    Next i

    Exit Sub
errHandler:
    TheExec.Datalog.WriteComment "<Error> " + funcName + ":: please check it out."
    If AbortTest Then Exit Sub Else Resume Next
End Sub

''==============================================================================================================
''GetMeasIResult
''==============================================================================================================
Public Sub GetMeasIResult(InputInstrumentPins As String, InputInstrumentMeasIResult As PinListData, Optional InputSampleSize As Long = 10, Optional InputWaitTime As Long = 1)
    On Error GoTo errHandler:: Dim funcName As String:: funcName = "GetMeasIResult"
    
    If Not TheExec.TesterMode = testModeOffline Then
       TheHdw.Wait InputWaitTime * ms
    End If
    If InputInstrumentPins <> "" Then InputInstrumentMeasIResult = TheHdw.DCVS.pins(InputInstrumentPins).Meter.Read(tlStrobe, 10, , tlDCVSMeterReadingFormatAverage)
    
    Exit Sub
errHandler:
    TheExec.Datalog.WriteComment "<Error> " + funcName + ":: please check it out."
    If AbortTest Then Exit Sub Else Resume Next
End Sub

''==============================================================================================================
''RemoveFirstChar
''==============================================================================================================
Public Sub RemoveFirstChar(InputStr As String)
    On Error GoTo errHandler:: Dim funcName As String:: funcName = "RemoveFirstChar"

    'Remove first character
    If InputStr <> "" Then InputStr = right(InputStr, Len(InputStr) - 1)

    Exit Sub
errHandler:
    TheExec.Datalog.WriteComment "<Error> " + funcName + ":: please check it out."
    If AbortTest Then Exit Sub Else Resume Next
End Sub
''==============================================================================================================
''SortByInstruments
''==============================================================================================================
Public Sub SortByInstruments(InputAllPins() As PinsInfo, InputHEXVSPins As String, InputUVSPins As String, InputVSMPins As String, InputUFPPins As String, InputPins As String, InputLogicVal As Double, InputSramVal As Double)
    On Error GoTo errHandler:: Dim funcName As String:: funcName = "SortByInstruments"
    
    Dim i As Long
    Dim j As Long
    Dim Voltage_Avg As Double
    Dim lclStr_Temp As String
    Dim lclStr_GateState As String
    Dim lclBool_GateState As Boolean
    
    Dim lclStr_DcCategory As String
    Dim lclStr_DcSelector As String
    Dim lclStr_AcCategory As String
    Dim lclStr_AcSelector As String
    Dim lclStr_TimeSetSheet As String
    Dim lclStr_EdgeSetSheet As String
    Dim lclStr_LevelsSheet As String
    Dim lclStr_Overlay As String
    
    TheExec.DataManager.GetInstanceContext lclStr_DcCategory, lclStr_DcSelector, lclStr_AcCategory, lclStr_AcSelector, lclStr_TimeSetSheet, lclStr_EdgeSetSheet, lclStr_LevelsSheet, lclStr_Overlay
    For j = 0 To UBound(InputAllPins)
        Voltage_Avg = Voltage_Avg + TheExec.Specs.DC.item(InputAllPins(j).PinsName & "_VAR").Categories(lclStr_DcCategory).Selectors.item(lclStr_DcSelector).ContextValue
    Next j
    Voltage_Avg = Voltage_Avg / (UBound(InputAllPins) + 1)
    
      
    For i = 0 To UBound(InputAllPins)
        If Not InputAllPins(i).PinsName = glbConstKeyWD_PRESERVE Then
            InputPins = InputPins & "," & InputAllPins(i).PinsName
            If Not InputAllPins(i).PinsName Like glbConstKeyWD_SRAM Then
                TheHdw.DCVS.pins(InputAllPins(i).PinsName).Voltage.Alt.value = Voltage_Avg + InputLogicVal
            Else
                TheHdw.DCVS.pins(InputAllPins(i).PinsName).Voltage.Alt.value = Voltage_Avg + InputSramVal
            End If
            lclBool_GateState = TheHdw.DCVS.pins(InputAllPins(i).PinsName).Gate
            
            Select Case lclBool_GateState
                 Case True: lclStr_GateState = "GATE ON"
                 Case False: lclStr_GateState = "GATE OFF"
            End Select

            If UCase(InputAllPins(i).ChanMapType) = glbConstIns_HEXVS Then
                InputHEXVSPins = InputHEXVSPins & "," & InputAllPins(i).PinsName
            ElseIf UCase(InputAllPins(i).ChanMapType) = glbConstIns_VHDVS Then
                InputUVSPins = InputUVSPins & "," & InputAllPins(i).PinsName
            ElseIf UCase(InputAllPins(i).ChanMapType) = glbConstIns_VSM Then
                InputVSMPins = InputVSMPins & "," & InputAllPins(i).PinsName
''==============================================================================================================
''                                              UFP instruments.                                               =
''==============================================================================================================
            ElseIf UCase(InputAllPins(i).ChanMapType) = glbConstIns_VS800MA Then
                InputUFPPins = InputUFPPins & "," & InputAllPins(i).PinsName
            ElseIf UCase(InputAllPins(i).ChanMapType) = glbConstIns_VS5A Then
                InputUFPPins = InputUFPPins & "," & InputAllPins(i).PinsName
            End If
''==============================================================================================================
            lclStr_Temp = FormatNumericDatalog(InputAllPins(i).PinsName, 15, False) & "= " & lclStr_GateState
            TheExec.Datalog.WriteComment lclStr_Temp
            lclStr_Temp = ""
        End If
    Next i

    Exit Sub
errHandler:
    TheExec.Datalog.WriteComment "<Error> " + funcName + ":: please check it out."
    If AbortTest Then Exit Sub Else Resume Next
End Sub
''==============================================================================================================
''RestoreIRange
''==============================================================================================================
Public Sub RestoreIRange(InputAllPins() As PinsInfo)
    On Error GoTo errHandler:: Dim funcName As String:: funcName = "RestoreIRange"
    
    Dim i As Long

    For i = 0 To UBound(InputAllPins)
        If Not InputAllPins(i).PinsName = glbConstKeyWD_PRESERVE Then
            With TheHdw.DCVS.pins(InputAllPins(i).PinsName)
                .Meter.mode = tlDCVSMeterCurrent
                .SetCurrentRanges InputAllPins(i).Range_List(UBound(InputAllPins(i).Range_List)), InputAllPins(i).Range_List(UBound(InputAllPins(i).Range_List))
            End With
        End If
    Next i

    Exit Sub
errHandler:
    TheExec.Datalog.WriteComment "<Error> " + funcName + ":: please check it out."
    If AbortTest Then Exit Sub Else Resume Next
End Sub
