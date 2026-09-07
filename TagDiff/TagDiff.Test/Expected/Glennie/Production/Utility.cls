VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Utility"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
#If VBA7 Then
    Private Declare PtrSafe Function CryptAcquireContext Lib "advapi32.dll" Alias "CryptAcquireContextA" (phProv As Long, pszContainer As String, pszProvider As String, ByVal dwProvType As Long, ByVal dwFlags As Long) As Long
    Private Declare PtrSafe Function CryptGenRandom Lib "advapi32.dll" (ByVal hProv As Long, ByVal dwLen As Long, pbBuffer As Byte) As Long
    Private Declare PtrSafe Function CryptReleaseContext Lib "advapi32.dll" (ByVal hProv As Long, ByVal dwFlags As Long) As Long
#Else
'    Private Declare Function CryptAcquireContext Lib "C:\windows\system32\advapi32_AES.dll" Alias "CryptAcquireContextA" (phProv As Long, pszContainer As String, pszProvider As String, ByVal dwProvType As Long, ByVal dwFlags As Long) As Long
'    Private Declare Function CryptGenRandom Lib "C:\windows\system32\advapi32_AES.dll" (ByVal hProv As Long, ByVal dwLen As Long, pbBuffer As Byte) As Long
'    Private Declare Function CryptReleaseContext Lib "C:\windows\system32\advapi32_AES.dll" (ByVal hProv As Long, ByVal dwFlags As Long) As Long
#End If

'For Random Codes Generation
Private Const CRYPT_VERIFYCONTEXT   As Long = &HF0000000
Private Const PROV_RSA_FULL         As Long = 1
Private Const RNG_BUFFER_SIZE As Integer = 4096
Private Type RNG_RandomBuffer_type
    Buffer(1 To RNG_BUFFER_SIZE) As Byte
    pos As Integer
End Type
Private RNG_RandomBuffer As RNG_RandomBuffer_type

Private stagesList As Variant
Public dicStageList As Dictionary
Public testedStages As Dictionary
Public currStage As String
Public preStage As String
Private ReferenceTime As Double
'Private Declare Function GetTickCount Lib "kernel32.dll" () As Long
Private LotIdBits As Integer, WfrIdBits As Integer, XcoorBits As Integer, YcoorBits As Integer
Private site As Variant
Public Timer As Boolean
Const HexPrefix = "0x"
Const binFmt = "^b[0|1]+"
Const hexFmt = "^0x[\d+,A-F]|^x[\d+,A-F]"
Const decFmt = "^\d+"
Public OnlineMode As Boolean

Public Function Initialize(stage As String) As Boolean
On Error GoTo errHandler
Dim pos As Variant, iStage As Integer

    If Len(currStage) <> 0 And currStage = UCase(stage) Then Initialize = True: Exit Function
    OnlineMode = False
    If TheExec.TesterMode = testModeOnline Then OnlineMode = True
    Timer = True
    stagesList = Split(Replace(UCase(stagesStr), " ", ""), ",")
    Set dicStageList = New Dictionary
    dicStageList.compareMode = TextCompare
    
    ''----------------------------------------------------
    ''Get previous stage
    ''Test job sequence : CP1-> CP2-> WLFT1-> FT2-> FT3
    ''ex: current job = "WLFT1" -> preStage = "CP2"
    ''----------------------------------------------------
    For iStage = 0 To UBound(stagesList)
        dicStageList.Add stagesList(iStage), True
        If iStage > 0 And UCase(stage) = stagesList(iStage) Then
            preStage = stagesList(iStage - 1)
        End If
    Next
    
    ''Check current job exists in stagelist(stagelist set in VBT declare)
    pos = Application.Match(stage, stagesList, False)  ' result will count from 1
    
    If LCase(CStr(pos)) Like "error*" Then
        MessageBox "Wrong Stage!! " & stage
        TheExec.flow.TestLimit resultVal:=999, lowVal:=1, hiVal:=1, Tname:="GlbUtility_Initialize"
        Initialize = False
        Exit Function
    End If
    
    currStage = UCase(stage)
    Set testedStages = New Dictionary
    testedStages.compareMode = TextCompare
    
    'Add tested stage
    'ex: current job : CP1 -> testedStages = "" (nothing)
    '    current job : FT2 -> testedStages = "CP1", "CP2", "WLFT1"
    For iStage = 1 To pos - 1
        testedStages.Add stagesList(iStage - 1), True
    Next
    Initialize = True

Exit Function
errHandler:
    Call Print_Error_Message(Error_Info, "Utility", "Initialize")
    If AbortTest Then Exit Function Else Resume Next
End Function

Public Function BinCheck() As SiteLong
On Error GoTo errHandler
Dim ReadBin As Long, result As New SiteLong: result = -1
Dim PASSBIN As Variant, iPassBin As Long

    PASSBIN = Array(1) '(1,19,109), listed your pass bins here
    TheHdw.DSSC.MoveMode = tlDSSCMoveModeDatabus
    TheExec.Datalog.WriteComment (vbCrLf & "*****Bincheck: Checking Bin Number*****" & vbCrLf)
    
    'get the current bin for each site available 'If the bin is NOT = to -1 then it means failed at some trim
    For Each site In TheExec.sites
        ReadBin = TheExec.sites(site).BinNumber
        If ReadBin = -1 Then
             result = 1
        Else
            result = ReadBin
            For iPassBin = 0 To UBound(PASSBIN)
                If ReadBin = PASSBIN(iPassBin) Then
                     result = 1: Exit For
                End If
            Next
        End If
    Next site
    Set BinCheck = result

Exit Function
errHandler:
    Call Print_Error_Message(Error_Info, "Utility", "BinCheck")
    If AbortTest Then Exit Function Else Resume Next
End Function

Public Function IsFutureStage(chkStage As String) As Boolean
On Error GoTo errHandler
Dim pos As Variant, poscurr As Variant

    poscurr = Application.Match(currStage, stagesList, False)
    pos = Application.Match(chkStage, stagesList, False)
    If LCase(CStr(pos)) Like "error*" Then
        MessageBox "Wrong Stage!! " & chkStage
        IsFutureStage = False
        Exit Function
    End If
    IsFutureStage = IIf(CLng(pos) > poscurr, True, False)

Exit Function
errHandler:
    Call Print_Error_Message(Error_Info, "Utility", "BinCheck")
    If AbortTest Then Exit Function Else Resume Next
End Function

Public Function IsOnline() As Boolean
On Error GoTo errHandler

    If TheExec.TesterMode = testModeOnline Then IsOnline = True: Exit Function
    IsOnline = False
    
Exit Function
errHandler:
    Call Print_Error_Message(Error_Info, "Utility", "IsOnline")
    If AbortTest Then Exit Function Else Resume Next
End Function

Public Sub IniSiteVar(var As String, v As Long)
On Error GoTo errHandler

    For Each site In TheExec.sites
        TheExec.sites(site).SiteVariableValue(var) = v
    Next site
    
Exit Sub
errHandler:
    Call Print_Error_Message(Error_Info, "Utility", "IsOnline")
    If AbortTest Then Exit Sub Else Resume Next
End Sub

Public Function GetStoreWaveName(bankstr As String, Optional getPreStage As Boolean = False) As String
On Error GoTo errHandler
Dim CurrSite As Long, stagename As String

    CurrSite = TheExec.sites.siteNumber
    stagename = IIf(getPreStage, preStage, currStage)
    GetStoreWaveName = lotId & "_" & WaferID & "_" & XCoord(CurrSite) & "_" & YCoord(CurrSite) & "_" & stagename & "_" & bankstr & ".txt"  'These're global variables

Exit Function
errHandler:
    Call Print_Error_Message(Error_Info, "Utility", "GetStoreWaveName")
    If AbortTest Then Exit Function Else Resume Next
End Function

Public Function GetStoreWaveform(bankstr As String, mypath As String, ByRef oWave As DSPWave) As SiteBoolean
On Error GoTo errHandler
Dim filename As String, bWaveGet As New SiteBoolean: bWaveGet = False

    If testedStages.Count = 0 Then GoTo result
    For Each site In TheExec.sites
        filename = mypath & lotId & "_" & WaferID & "_" & XCoord & "_" & YCoord & "_" & preStage & ".txt"
        If Dir(filename) <> "" Then
            oWave.FileImport filename
            bWaveGet = True
        End If
    Next
result:
    Set GetStoreWaveform = bWaveGet

Exit Function
errHandler:
    Call Print_Error_Message(Error_Info, "Utility", "GetStoreWaveName")
    If AbortTest Then Exit Function Else Resume Next
End Function

Public Sub MessageBox(msg As String, Optional level As String = "Error")
On Error GoTo errHandler
Dim err As String

    Select Case level
        Case "Warning": err = "[Warning]"
        Case Else
            err = "[Error]"
    End Select
    MsgBox err & msg

Exit Sub
errHandler:
    Call Print_Error_Message(Error_Info, "Utility", "MessageBox")
    If AbortTest Then Exit Sub Else Resume Next
End Sub

Public Function EncodeEcidBinString(inLotId As String, inWfrId As Integer, inXcoor As Integer, inYcoor As Integer, _
Optional inLotIdBits As Integer = 36, Optional inWfrIdBits As Integer = 5, Optional inXcoorBits As Integer = 6, Optional inYcoorBits As Integer = 6) As String
On Error GoTo errHandler

    LotIdBits = inLotIdBits
    WfrIdBits = inWfrIdBits
    XcoorBits = inXcoorBits
    YcoorBits = inYcoorBits
    EncodeEcidBinString = GetLotIdBin(inLotId, inLotIdBits) & Dec2Bin(inWfrId, CLng(WfrIdBits)) & Dec2Bin(inXcoor, CLng(XcoorBits)) & Dec2Bin(inYcoor, CLng(YcoorBits))

Exit Function
errHandler:
    Call Print_Error_Message(Error_Info, "Utility", "EncodeEcidBinString")
    If AbortTest Then Exit Function Else Resume Next
End Function

Public Function DecodeEcidBinString(ByVal ecid As String) As String
On Error GoTo errHandler
Dim mylotid As String, wfr As String, x As String, y As String

    mylotid = DecodeLotId(mid(ecid, 1, LotIdBits), False)
    wfr = Bin2Dec(mid(ecid, 1 + LotIdBits, WfrIdBits))
    x = Bin2Dec(mid(ecid, 1 + LotIdBits + WfrIdBits, XcoorBits))
    y = Bin2Dec(mid(ecid, 1 + LotIdBits + WfrIdBits + XcoorBits, YcoorBits))
    DecodeEcidBinString = mylotid & "-" & wfr & "-" & x & "-" & y

Exit Function
errHandler:
    Call Print_Error_Message(Error_Info, "Utility", "DecodeEcidBinString")
    If AbortTest Then Exit Function Else Resume Next
End Function

Public Sub TimerStarter()
On Error GoTo errHandler

    If Not Timer Then Exit Sub
    Call WriteDlg("Time Elapse > " & TheExec.DataManager.instancename)
    ReferenceTime = TheExec.Timer

Exit Sub
errHandler:
    Call Print_Error_Message(Error_Info, "Utility", "TimerStarter")
    If AbortTest Then Exit Sub Else Resume Next
End Sub

Public Sub PersonalTimerLog(ByVal DlogInfo As String, Optional toDp As Boolean = False)
On Error GoTo errHandler

    If Not Timer Then Exit Sub
    Dim log As String, elapse As Double
    elapse = TheExec.Timer(ReferenceTime)
    If elapse < 32 Then
        log = "Time Elapse >" & TxtFmt("(" & DlogInfo & ") ", 24, " ") & CLng(elapse * 1000000) / 1000 & " mSec"
    Else
        log = "Time Elapse >" & TxtFmt("(" & DlogInfo & ") ", 24, " ") & elapse * 1000 & " mSec"
    End If
    Call WriteDlg(log)
    If toDp Then Debug.Print log
    ReferenceTime = TheExec.Timer

Exit Sub
errHandler:
    Call Print_Error_Message(Error_Info, "Utility", "PersonalTimerLog")
    If AbortTest Then Exit Sub Else Resume Next
End Sub

Public Sub WriteDlg(ByVal DlogInfo As String)
On Error GoTo errHandler
    TheExec.Datalog.WriteComment (DlogInfo)
Exit Sub
errHandler:
    Call Print_Error_Message(Error_Info, "Utility", "WriteDlg")
    If AbortTest Then Exit Sub Else Resume Next
End Sub

Public Sub WriteErrDlg(ByVal DlogInfo As String)
On Error GoTo errHandler
    TheExec.Datalog.WriteComment (DlogInfo)
    TheExec.ErrorLogMessage (DlogInfo & ", Caused by Instance>" & TheExec.DataManager.instancename)
Exit Sub
errHandler:
    Call Print_Error_Message(Error_Info, "Utility", "WriteErrDlg")
    If AbortTest Then Exit Sub Else Resume Next
End Sub

Public Sub Put2IEDA(identifyItem As String, objItem As SiteVariant, Optional bankName As String = "")
On Error GoTo errHandler
Dim regKey As String: regKey = vbNullString
Dim m_DataPrint As String
Dim opbank As eFuseBdfBank

    If bankName <> "" Then
        Set opbank = GetBdfBank(bankName)
    End If
    For Each site In TheExec.sites.Existing
        If bankName <> "" Then
            If objItem(site) = "" Or opbank.ReadPatResut(site) = False Then
                regKey = regKey & ""
            Else
                regKey = regKey & objItem(site)
            End If
        Else
            If objItem(site) = "" Then
                regKey = regKey & ""
            Else
                regKey = regKey & objItem(site)
            End If
        End If

        If site <> TheExec.sites.Existing.Count - 1 Then regKey = regKey & ","
    Next
    m_DataPrint = "print: Set IEDA registry ( " & identifyItem & " ) = " & regKey
    TheExec.Datalog.WriteComment m_DataPrint
    RegKeySave identifyItem, regKey

Exit Sub
errHandler:
    Call Print_Error_Message(Error_Info, "Utility", "Put2IEDA")
    If AbortTest Then Exit Sub Else Resume Next
End Sub

Public Function DecodeLotId(binstr As String, Optional needReverse As Boolean = True) As String
On Error GoTo errHandler
Dim EcidBin As String, lotChr As String, mylotid As String, iChar As Integer, iDec As Integer

    EcidBin = binstr
    If needReverse Then EcidBin = StrReverse(binstr)
    mylotid = vbNullString
    For iChar = 0 To Len(binstr) / LotIdCharBits - 1
        iDec = Bin2Dec(mid(EcidBin, iChar * LotIdCharBits + 1, LotIdCharBits))
        If iDec < 10 Then
            iDec = iDec + 48
        Else
            iDec = iDec + 55
        End If
        lotChr = UCase(Chr(iDec))
        mylotid = mylotid & lotChr
    Next iChar
    DecodeLotId = mylotid

Exit Function
errHandler:
    Call Print_Error_Message(Error_Info, "Utility", "DecodeLotId")
    If AbortTest Then Exit Function Else Resume Next
End Function

Public Function GetLotIdBin(mylotid As String, Optional MaxBit As Integer = 36, Optional toHex As Boolean = False) As String
On Error GoTo errHandler
Dim iChar As Integer, lotChr As String, binstr As String

    binstr = vbNullString
    If Len(mylotid) > MaxBit / LotIdCharBits Then
        TheExec.Datalog.WriteComment "lotid string is too long for " & MaxBit & " bits"
        GoTo skip
    End If
    
    For iChar = 1 To Len(mylotid)
        lotChr = mid(mylotid, iChar, 1)
        If IsNumeric(lotChr) Then
            binstr = binstr & Dec2Bin(val(lotChr), LotIdCharBits)
        Else
            binstr = binstr & Dec2Bin(Asc(lotChr) - 55, LotIdCharBits)
        End If
    Next iChar
    If UCase(binstr) Like "*ERROR101*" Then GoTo skip
    If toHex Then binstr = Bin2HexStr(binstr, (MaxBit - 1) \ 4 + 1)
    GetLotIdBin = binstr
    
Exit Function
skip:
    GetLotIdBin = "Error101"
    Exit Function
errHandler:
    Call Print_Error_Message(Error_Info, "Utility", "GetLotIdBin")
    If AbortTest Then Exit Function Else Resume Next
End Function

Public Function Dec2HexStr(value As Double, Optional hBytes As Long = 8) As String
On Error GoTo errHandler
Dim fmt As String, iByte As Integer

    fmt = vbNullString
    For iByte = 0 To hBytes - 1
        fmt = fmt & "@"
    Next
    Dec2HexStr = HexPrefix & Replace(Format(Hex(value), fmt), " ", "0")

Exit Function
errHandler:
    Call Print_Error_Message(Error_Info, "Utility", "Dec2HexStr")
    If AbortTest Then Exit Function Else Resume Next
End Function

Public Function oDec2HexStr(value As SiteVariant, Optional hBytes As Long = 8) As SiteVariant
On Error GoTo errHandler
Dim fmt As String, iByte As Integer, o As New SiteVariant

    If IsHex(value) Then
        For Each site In TheExec.sites
            o(site) = String2Hex(value, hBytes:=hBytes):
        Next
    Else
        fmt = vbNullString
        For iByte = 0 To hBytes - 1
            fmt = fmt & "@"
        Next
        For Each site In TheExec.sites
            o(site) = HexPrefix & Replace(Format(Hex(value), fmt), " ", "0")
        Next
    End If
    Set oDec2HexStr = o

Exit Function
errHandler:
    Call Print_Error_Message(Error_Info, "Utility", "Dec2HexStr")
    If AbortTest Then Exit Function Else Resume Next
End Function

Public Function Bin2HexStr(value As String, Optional hBytes As Long = 8) As String
On Error GoTo errHandler
Dim fmt As String, iByte As Integer, rt As String, tmp As String, ipadding As Integer, ipad As Integer

    fmt = vbNullString
    value = Replace(Replace(LCase(value), "b", ""), "_", "")
    ipadding = hBytes * 4 - Len(value)
    If ipadding < 0 Then
        If InStr(value, "1") = 0 Then
            ipadding = 0
        Else
            MessageBox "Error " & value & " is overflow. Max bit size is " & hBytes * 4
            Exit Function
        End If
    End If

    tmp = String(ipadding, "0") & value
    For iByte = 0 To hBytes - 1
        Select Case mid(tmp, iByte * 4 + 1, 4)
            Case "0000": fmt = fmt + "0"
            Case "1111": fmt = fmt + "F"
            Case "1110": fmt = fmt + "E"
            Case "1101": fmt = fmt + "D"
            Case "1100": fmt = fmt + "C"
            Case "1011": fmt = fmt + "B"
            Case "1010": fmt = fmt + "A"
            Case "1001": fmt = fmt + "9"
            Case "1000": fmt = fmt + "8"
            Case "0111": fmt = fmt + "7"
            Case "0110": fmt = fmt + "6"
            Case "0101": fmt = fmt + "5"
            Case "0100": fmt = fmt + "4"
            Case "0011": fmt = fmt + "3"
            Case "0010": fmt = fmt + "2"
            Case "0001": fmt = fmt + "1"
        End Select
    Next
    Bin2HexStr = HexPrefix & fmt

Exit Function
errHandler:
    Call Print_Error_Message(Error_Info, "Utility", "Bin2HexStr")
    If AbortTest Then Exit Function Else Resume Next
End Function

Public Function xHex2BinStr(value As String, Optional Nbit As Long = -1) As String
On Error GoTo errHandler
Dim fmt As String, iChr As Integer, tmp As String

    If Nbit <> -1 Then
        tmp = Replace(value, "0", "")
        If (tmp = "" Or tmp = "x" Or tmp = "X") Then xHex2BinStr = String(Nbit, "0"): Exit Function
    End If
    tmp = UCase(value)
    tmp = Replace(tmp, "0X", "")
    tmp = Replace(tmp, "X", "")
    fmt = vbNullString

    For iChr = 0 To Len(tmp) - 1
        Select Case mid(tmp, iChr + 1, 1)
            Case "0": fmt = fmt + "0000"
            Case "F": fmt = fmt + "1111"
            Case "E": fmt = fmt + "1110"
            Case "D": fmt = fmt + "1101"
            Case "C": fmt = fmt + "1100"
            Case "B": fmt = fmt + "1011"
            Case "A": fmt = fmt + "1010"
            Case "9": fmt = fmt + "1001"
            Case "8": fmt = fmt + "1000"
            Case "7": fmt = fmt + "0111"
            Case "6": fmt = fmt + "0110"
            Case "5": fmt = fmt + "0101"
            Case "4": fmt = fmt + "0100"
            Case "3": fmt = fmt + "0011"
            Case "2": fmt = fmt + "0010"
            Case "1": fmt = fmt + "0001"
            Case Else
                MessageBox "Error text in " & value
        End Select
    Next
    
    If Nbit <> -1 Then fmt = right(fmt, Nbit)
    If Len(fmt) < Nbit Then
        fmt = String(Nbit - Len(fmt), "0") & fmt
    End If
    xHex2BinStr = fmt

Exit Function
errHandler:
    Call Print_Error_Message(Error_Info, "Utility", "xHex2BinStr")
    If AbortTest Then Exit Function Else Resume Next
End Function

Public Function Hex2BinStr(value As SiteVariant, Optional Nbit As Long = -1) As SiteVariant
On Error GoTo errHandler
Dim fmt As New SiteVariant, iChr As Integer, tmp As New SiteVariant

    tmp = value
    For Each site In TheExec.sites
        fmt = xHex2BinStr(tmp(site), Nbit)
    Next
    Set Hex2BinStr = fmt

Exit Function
errHandler:
    Call Print_Error_Message(Error_Info, "Utility", "Hex2BinStr")
    If AbortTest Then Exit Function Else Resume Next
End Function

Public Function IsHex(v As SiteVariant) As Boolean
On Error GoTo errHandler

    IsHex = False
    For Each site In TheExec.sites
        If Not IsHex Then IsHex = xIsHex(v)
    Next

Exit Function
errHandler:
    Call Print_Error_Message(Error_Info, "Utility", "IsHex")
    If AbortTest Then Exit Function Else Resume Next
End Function

Public Function xIsHex(v As Variant) As Boolean
On Error GoTo errHandler

    xIsHex = False
    If v Like "0x*" Or v Like "x*" Or v Like "X*" Then xIsHex = True: Exit Function

Exit Function
errHandler:
    Call Print_Error_Message(Error_Info, "Utility", "xIsHex")
    If AbortTest Then Exit Function Else Resume Next
End Function

Public Function ReverseBits(value As Double, bitCount As Long) As Double
On Error GoTo errHandler
Dim bin As String, binRev As String

    bin = Dec2Bin(value, bitCount)
    binRev = StrReverse(bin)
    ReverseBits = Bin2Dec(binRev)
    Debug.Print bin & "=>" & binRev

Exit Function
errHandler:
    Call Print_Error_Message(Error_Info, "Utility", "ReverseBits")
    If AbortTest Then Exit Function Else Resume Next
End Function

Function Bin2Dec(BinaryString As String, Optional b2sComp As Boolean = False) As Double
On Error GoTo errHandler
Dim x As Integer
Dim BinaryStringTmp As String

    BinaryStringTmp = Replace(Replace(BinaryString, "_", ""), "b", "")
    If b2sComp Then
        For x = 0 To Len(BinaryStringTmp) - 2
            Bin2Dec = (Bin2Dec) + val(mid(BinaryStringTmp, Len(BinaryStringTmp) - x, 1)) * 2# ^ x
        Next
        Bin2Dec = (Bin2Dec) - val(mid(BinaryStringTmp, Len(BinaryStringTmp) - x, 1)) * 2# ^ x
    Else
        For x = 0 To Len(BinaryStringTmp) - 1
            Bin2Dec = (Bin2Dec) + CDbl(val(mid(BinaryStringTmp, Len(BinaryStringTmp) - x, 1)) * 2# ^ x)
        Next
    End If

Exit Function
errHandler:
    Call Print_Error_Message(Error_Info, "Utility", "Bin2Dec")
    If AbortTest Then Exit Function Else Resume Next
End Function

Public Function Dec2Bin(ByVal value As Double, MaxBit As Long, Optional b2sComp As Boolean = True) As String
On Error GoTo errHandler
Dim n As Integer
Dim Sign As String

    Dec2Bin = vbNullString
    If 2# ^ MaxBit > value Then
        'If b2sComp Then
        If value < 0 Then
            If value < 0 Then
                value = value + 2 ^ (MaxBit - 1)
                Sign = "1"
            Else
                Sign = "0"
            End If
            
            For n = 1 To MaxBit - 1
                Dec2Bin = MyMod(value, 2#) & Dec2Bin
                value = Fix(value / 2#)
            Next
            Dec2Bin = Sign & Dec2Bin
        Else
            For n = 1 To MaxBit
                Dec2Bin = MyMod(value, 2#) & Dec2Bin
                value = Fix(value / 2#)
            Next
        End If
    Else
        MsgBox "DecNumber too big for " & MaxBit & " bits"
        Dec2Bin = "Error101"
    End If

Exit Function
errHandler:
    Call Print_Error_Message(Error_Info, "Utility", "Dec2Bin")
    If AbortTest Then Exit Function Else Resume Next
End Function

Public Function MyMod(v As Double, Base As Double) As Double
On Error GoTo errHandler

    MyMod = v - (Int(v / Base) * Base)

Exit Function
errHandler:
    Call Print_Error_Message(Error_Info, "Utility", "MyMod")
    If AbortTest Then Exit Function Else Resume Next
End Function

Public Function Hex2Dec(ByVal value As String) As Double
On Error GoTo errHandler

    Hex2Dec = CDbl(Replace(Replace(Replace(Replace(value, "_", ""), "0x", "&H"), "x", "&H"), "X", "&H"))

Exit Function
errHandler:
    Call Print_Error_Message(Error_Info, "Utility", "Hex2Dec")
    If AbortTest Then Exit Function Else Resume Next
End Function

Public Function Hex2Dbl(HexString As String) As Double
On Error GoTo errHandler
Dim x As Integer

    If Len(HexString) > 32 Then
        HexString = "0x" & Hex(Replace(Replace(Replace(Replace(HexString, "_", ""), "0x", "&H"), "x", "&H"), "X", "&H"))
    End If
    HexString = Replace(HexString, "0x", "")

    For x = 0 To Len(HexString) - 1
        Hex2Dbl = (Hex2Dbl) + MapNum(mid(HexString, Len(HexString) - x, 1)) * 16# ^ x
    Next

Exit Function
errHandler:
    Call Print_Error_Message(Error_Info, "Utility", "Hex2Dbl")
    If AbortTest Then Exit Function Else Resume Next
End Function

Private Function MapNum(v As String) As Double
On Error GoTo errHandler

    Select Case v
        Case "F": MapNum = 15
        Case "E": MapNum = 14
        Case "D": MapNum = 13
        Case "C": MapNum = 12
        Case "B": MapNum = 11
        Case "A": MapNum = 10
        Case Else
            MapNum = val(v)
    End Select
 
Exit Function
errHandler:
    Call Print_Error_Message(Error_Info, "Utility", "MapNum")
    If AbortTest Then Exit Function Else Resume Next
End Function

Public Function String2Long(value As String, Optional Default As Long = 0) As Double
On Error GoTo errHandler

    If IsStrMatch(value, binFmt) Then
       String2Long = Bin2Dec(value)
       Exit Function
    ElseIf IsStrMatch(value, hexFmt) Then
       String2Long = Hex2Dec(value)
       Exit Function
    ElseIf IsStrMatch(value, decFmt) Then
       String2Long = CDbl(value)
       Exit Function
    End If
    String2Long = Default
 
Exit Function
errHandler:
    Call Print_Error_Message(Error_Info, "Utility", "String2Long")
    If AbortTest Then Exit Function Else Resume Next
End Function

Public Function String2Dbl(value As String, Optional Default As Double = 0#) As Double
On Error GoTo err:
Dim str As String
         If IsStrMatch(value, binFmt) Then
            str = Replace(Replace(value, "_", ""), "b", "")
            If (Len(Trim(str)) > 32) Then
                String2Dbl = -1
            Else
                String2Dbl = CDbl(Bin2Dec(value))
            End If
            Exit Function
         ElseIf IsStrMatch(value, hexFmt) Then
            str = Replace(Replace(Replace(Replace(Replace(value, "_", ""), "0x", ""), "x", ""), "X", ""), " ", "")
            If (Len(Trim(str)) > 8) Then
                String2Dbl = -1
            Else
                String2Dbl = CDbl(Hex2Dec(value))
            End If
       Exit Function
    ElseIf IsStrMatch(value, decFmt) Then
       String2Dbl = CDbl(value)
       Exit Function
    End If
err:
    String2Dbl = Default
     
End Function

Public Function String2Hex(ByVal value As String, Optional Default As Double = 0#, Optional hBytes As Long = 8) As String
On Error GoTo skip:

    If IsStrMatch(value, "^n/a|^na") Then value = "0"
    If IsStrMatch(value, binFmt) Then
       String2Hex = Bin2HexStr(value, hBytes)
       Exit Function
    ElseIf IsStrMatch(value, hexFmt) Then
       String2Hex = UCase(Replace(Replace(Replace(Replace(value, "_", ""), "0x", ""), "x", ""), "X", ""))
       If hBytes < Len(String2Hex) Then
           String2Hex = HexPrefix & UCase(Replace(Replace(Replace(Replace(value, "_", ""), "0x", ""), "x", ""), "X", ""))
           TheExec.Datalog.WriteComment "Error:HexValue string Size doesn't match field size, Error!!!."
       Else
           String2Hex = HexPrefix & String(hBytes - Len(String2Hex), "0") & UCase(Replace(Replace(Replace(Replace(value, "_", ""), "0x", ""), "x", ""), "X", ""))
       End If
       Exit Function
    ElseIf IsStrMatch(value, decFmt) Then
       String2Hex = Dec2HexStr(CDbl(value), hBytes)
       Exit Function
    End If
    String2Hex = IIf(Default = 0, Dec2HexStr(Default, hBytes), Default): Exit Function
skip:

End Function

Function ShortHexFmt(ByVal prtStr As String, fmt As Long, Optional padding As String = " ") As String
On Error GoTo errHandler 'overflow, deal <32 bits value

    If Len(prtStr) <= fmt Then
        prtStr = "0x" & Hex(Replace(prtStr, "0x", "&H"))
    End If
    ShortHexFmt = TxtFmt(prtStr, fmt, padding)
    
Exit Function
errHandler:
    Call Print_Error_Message(Error_Info, "Utility", "ShortHexFmt")
   
        If AbortTest Then Exit Function Else Resume Next

End Function

Function TxtFmt(ByVal prtStr As String, fmt As Long, Optional padding As String = " ", Optional left As Boolean = False) As String
On Error GoTo errHandler
Dim strLen As Long

    strLen = Len(prtStr)
    If strLen < fmt Then
        If left Then
           TxtFmt = prtStr + String(fmt - strLen, padding)
        Else
           TxtFmt = String(fmt - strLen, padding) + prtStr
        End If
        Exit Function
    End If
    TxtFmt = prtStr
      
Exit Function
errHandler:
    Call Print_Error_Message(Error_Info, "Utility", "TxtFmt")
    If AbortTest Then Exit Function Else Resume Next
End Function

Public Function Hex2BinArray(ByVal v As SiteVariant, Nbit As Long) As SiteLong()
On Error GoTo errHandler
Dim binA() As New SiteLong, binstr As String
Dim iBit As Integer, bitwidth As Integer: iBit = 0

    ReDim binA(Nbit - 1)
    For Each site In TheExec.sites
        binstr = StrReverse(xHex2BinStr(v(site), Nbit))
        For iBit = 0 To Nbit - 1
            binA(iBit) = val(mid(binstr, iBit + 1, 1))
        Next
    Next
    Hex2BinArray = binA ' Dec9 5bits => result = binA[0..4] = [10010], lSB=>[0] MSB=>[nbit-1]
    
Exit Function
errHandler:
    Call Print_Error_Message(Error_Info, "Utility", "Hex2BinArray")
    If AbortTest Then Exit Function Else Resume Next
End Function

Public Sub PrintSiteBitVar(site As Integer, binA() As SiteLong)
On Error GoTo errHandler
Dim item As Variant, bin As SiteLong, bitStr As String

    bitStr = vbNullString
    For Each item In binA
        Set bin = item
        bitStr = bitStr & item(site)
    Next
    Debug.Print "Site#" & site & " >> " & bitStr

Exit Sub
errHandler:
    Call Print_Error_Message(Error_Info, "Utility", "PrintSiteBitVar")
    If AbortTest Then Exit Sub Else Resume Next
End Sub

Public Sub CRC_0Array(ByRef crc() As Byte)
On Error GoTo errHandler
Dim i As Integer

    For i = 0 To UBound(crc)
        crc(i) = 0
    Next

Exit Sub
errHandler:
    Call Print_Error_Message(Error_Info, "Utility", "CRC_0Array")
    If AbortTest Then Exit Sub Else Resume Next
End Sub

Public Sub CalculateCRC16(ByRef crc() As Byte, bit As Byte)
On Error GoTo errHandler
Dim inv As Byte

    inv = bit Xor crc(15)
    crc(15) = crc(14)
    crc(14) = crc(13)
    crc(13) = crc(12) Xor inv
    crc(12) = crc(11) Xor inv
    crc(11) = crc(10) Xor inv
    crc(10) = crc(9) Xor inv
    crc(9) = crc(8)
    crc(8) = crc(7) Xor inv
    crc(7) = crc(6)
    crc(6) = crc(5) Xor inv
    crc(5) = crc(4) Xor inv
    crc(4) = crc(3)
    crc(3) = crc(2)
    crc(2) = crc(1) Xor inv
    crc(1) = crc(0)
    crc(0) = inv

Exit Sub
errHandler:
    Call Print_Error_Message(Error_Info, "Utility", "CalculateCRC16")
    If AbortTest Then Exit Sub Else Resume Next
End Sub

Public Sub CalculateCRC32(ByRef crc() As Byte, bit As Byte)
On Error GoTo errHandler
Dim inv As Byte
    
    inv = bit Xor crc(31)
    crc(31) = crc(30)
    crc(30) = crc(29)
    crc(29) = crc(28)
    crc(28) = crc(27)
    crc(27) = crc(26)
    crc(26) = crc(25) Xor inv
    crc(25) = crc(24)
    crc(24) = crc(23)
    crc(23) = crc(22) Xor inv
    crc(22) = crc(21) Xor inv
    crc(21) = crc(20)
    crc(20) = crc(19)
    crc(19) = crc(18)
    crc(18) = crc(17)
    crc(17) = crc(16)
    crc(16) = crc(15) Xor inv
    crc(15) = crc(14)
    crc(14) = crc(13)
    crc(13) = crc(12)
    crc(12) = crc(11) Xor inv
    crc(11) = crc(10) Xor inv
    crc(10) = crc(9) Xor inv
    crc(9) = crc(8)
    crc(8) = crc(7) Xor inv
    crc(7) = crc(6) Xor inv
    crc(6) = crc(5)
    crc(5) = crc(4) Xor inv
    crc(4) = crc(3) Xor inv
    crc(3) = crc(2)
    crc(2) = crc(1) Xor inv
    crc(1) = crc(0) Xor inv
    crc(0) = inv

Exit Sub
errHandler:
    Call Print_Error_Message(Error_Info, "Utility", "CalculateCRC32")
    If AbortTest Then Exit Sub Else Resume Next
End Sub

Public Function CeilingValue(ByVal x As Double, Optional ByVal Factor As Double = 1) As Variant
On Error GoTo errHandler
' X is the value you want to round
Dim m_lngX As Long

    m_lngX = CLng(x)

    ''''20161215 update, By this way to avoid any unsuitable Ceiling value
    If ((CDbl(x) - m_lngX) > 0.000000001) Then
        CeilingValue = m_lngX + 1
    Else
        CeilingValue = m_lngX
    End If

Exit Function
errHandler:
    Call Print_Error_Message(Error_Info, "Utility", "CeilingValue")
    If AbortTest Then Exit Function Else Resume Next
End Function

Public Function FloorValue(ByVal x As Double, Optional ByVal Factor As Double = 1) As Variant
On Error GoTo errHandler

    ' X is the value you want to round
    FloorValue = Int(x / Factor) * Factor

Exit Function
errHandler:
    Call Print_Error_Message(Error_Info, "Utility", "FloorValue")
    If AbortTest Then Exit Function Else Resume Next
End Function

Private Function ReverseArray(arr() As Long) As Variant
On Error GoTo errHandler
Dim val As Variant

    With CreateObject("System.Collections.ArrayList") '<-- create a "temporary" array list with late binding
        For Each val In arr '<--| fill arraylist
            .Add val
        Next val
        .Reverse '<--| reverse it
        ReverseArray = .Toarray '<--| write it into an array
    End With

Exit Function
errHandler:
    Call Print_Error_Message(Error_Info, "Utility", "ReverseArray")
    If AbortTest Then Exit Function Else Resume Next
End Function

Public Function DicKeysSort(oriDic As Dictionary) As Dictionary 'MSB first
On Error GoTo errHandler
Dim tmpLng() As Long: ReDim tmpLng(UBound(oriDic.Keys))
Dim newDic As New Dictionary, item As Variant
Dim idx As Integer

    idx = 0
    For Each item In oriDic.Keys
        tmpLng(idx) = item
        idx = idx + 1
    Next
    
    Call Sort(tmpLng, True)
    For idx = 0 To UBound(tmpLng)
        newDic.Add tmpLng(idx), True
    Next
    Set DicKeysSort = newDic

Exit Function
errHandler:
    Call Print_Error_Message(Error_Info, "Utility", "DicKeysSort")
    If AbortTest Then Exit Function Else Resume Next
End Function

Public Sub IniDictionary(ByRef dic As Dictionary, Optional forceReset As Boolean = False)
On Error GoTo errHandler

    If dic Is Nothing Or forceReset Then
        Set dic = New Dictionary
        dic.compareMode = TextCompare
    End If

Exit Sub
errHandler:
    Call Print_Error_Message(Error_Info, "Utility", "IniDictionary")
    If AbortTest Then Exit Sub Else Resume Next
End Sub

Public Sub DicCleaner(ByRef dic As Dictionary)
On Error GoTo errHandler

    If dic.Exists(Empty) Then dic.Remove (Empty)
    If dic.Exists("") Then dic.Remove ("")

Exit Sub
errHandler:
    Call Print_Error_Message(Error_Info, "Utility", "DicCleaner")
    If AbortTest Then Exit Sub Else Resume Next
End Sub

Public Function GetSampleSize(ByRef InWave As DSPWave) As Long
On Error GoTo errHandler
Dim size As Long: size = 0

    For Each site In TheExec.sites
        If size = 0 Then size = InWave.SampleSize
    Next
    GetSampleSize = size

Exit Function
errHandler:
    Call Print_Error_Message(Error_Info, "Utility", "GetSampleSize")
    If AbortTest Then Exit Function Else Resume Next
End Function

Public Function IsStrMatch(myStr As String, regStr As String, Optional IgnoreCase As Boolean = True) As Boolean
On Error GoTo errHandler
Dim reParse As New RegExp

    reParse.IgnoreCase = IgnoreCase
    reParse.Pattern = regStr
    IsStrMatch = reParse.test(myStr)

Exit Function
errHandler:
    Call Print_Error_Message(Error_Info, "Utility", "IsStrMatch")
    If AbortTest Then Exit Function Else Resume Next
End Function

Public Function IsInstLike(regStr As String, Optional IgnoreCase As Boolean = True) As Boolean
On Error GoTo errHandler
Dim instance As String: instance = TheExec.DataManager.instancename

    IsInstLike = IsStrMatch(instance, regStr, IgnoreCase)

Exit Function
errHandler:
    Call Print_Error_Message(Error_Info, "Utility", "IsInstLike")
    If AbortTest Then Exit Function Else Resume Next
End Function

Public Function xHexCompare(value As String, criteria As String, chk As ChkType, Optional Nbit As Long = -1) As Boolean
On Error GoTo err
Dim bValue As String, bCriteria As String, iBit As Long
Dim vbit As Long, cbit As Long

    value = Replace(value, "0x", "")
    criteria = Replace(criteria, "0x", "")
    If Nbit = -1 Then Nbit = (IIf(Len(value) > Len(criteria), Len(value), Len(criteria))) * 4 ' Hex chars, 1 chr = 4 bits
    bValue = xHex2BinStr(value, Nbit)
    bCriteria = xHex2BinStr(criteria, Nbit)
    xHexCompare = False
    For iBit = 0 To Nbit - 1
        vbit = CLng(mid(bValue, iBit + 1, 1))
        cbit = CLng(mid(bCriteria, iBit + 1, 1))
        Select Case chk:
            Case ChkType.ChkGreaterThan, ChkType.ChkGreaterEqualThan
                If vbit > cbit Then
                     xHexCompare = True:  Exit Function
                ElseIf vbit < cbit Then
                     Exit Function
                End If
            Case ChkType.ChkLessThan, ChkType.ChkLessEqualThan
                If vbit < cbit Then
                    xHexCompare = True:  Exit Function
                ElseIf vbit > cbit Then
                    Exit Function
                End If
            Case ChkType.ChkEqualTo
                If vbit <> cbit Then Exit Function
            Case Else
                MessageBox ("xHexCompare function error")
        End Select
    Next
    If chk = ChkType.ChkGreaterEqualThan Or chk = ChkType.ChkLessEqualThan Or chk = ChkType.ChkEqualTo Then xHexCompare = True  ' means exactly equal
Exit Function
err:
    xHexCompare = False
End Function

Public Function UpdateDLogColumns(tsNameWidth As Variant) As Long
On Error GoTo errHandler

    tsNameWidth = CLng(tsNameWidth)
    TheExec.Datalog.Setup.Shared.ascii.Columns.EnableCustomWidths = True
    With TheExec.Datalog.Setup.Shared.ascii.Columns.Parametric
        .TestName.Enable = True
        .TestName.Width = tsNameWidth
        .pin.Enable = True
        .pin.Width = 25
    End With
    With TheExec.Datalog.Setup.Shared.ascii.Columns.Functional
        .Pattern.Enable = True
        .Pattern.Width = 128 '.Pattern.DefaultWidth
        .TestName.Enable = True
        .TestName.Width = tsNameWidth
    End With
    TheExec.Datalog.ApplySetup

Exit Function
errHandler:
    Call Print_Error_Message(Error_Info, "Utility", "UpdateDLogColumns")
    If AbortTest Then Exit Function Else Resume Next
End Function

Public Function ErrorHandler(funcName As String)
    GlbUtility.WriteDlg "<Error> " + funcName + ":: please check it out."
    If AbortTest Then Exit Function Else Resume Next
End Function

Public Function RNG_cryptoRandomByte() As Byte
On Error GoTo errHandler
Dim RandomVal As Byte
Dim Error As Boolean
Dim i As Integer
Dim Sum As Double

    If RNG_RandomBuffer.pos = 0 Then
        ' either this is first run, assuming global variables are zero initilised
        ' or the buffer has wrapped
        
        ' refill the buffer
        Error = RNG_GetRandom()
        If (Error <> True) Then
            MsgBox ("RNG_cryptoRandomByte: Error getting random buffer")
            End
        End If
    
        ' perform a basic sanity test that the buffer isn't full of zeros
            Sum = 0
        For i = 1 To RNG_BUFFER_SIZE
            Sum = Sum + RNG_RandomBuffer.Buffer(i)
        Next i
        If (Sum = 0) Then
            MsgBox ("RNG_cryptoRandomByte: Random buffer is zero")
            End
        End If
    End If
    
    ' plus one since the array starts at index 1
    RandomVal = RNG_RandomBuffer.Buffer(RNG_RandomBuffer.pos + 1)
    
    ' progress the buffer read position
    RNG_RandomBuffer.pos = (RNG_RandomBuffer.pos + 1) Mod RNG_BUFFER_SIZE
    
    RNG_cryptoRandomByte = RandomVal

Exit Function
errHandler:
    Call Print_Error_Message(Error_Info, "Utility", "RNG_cryptoRandomByte")
    If AbortTest Then Exit Function Else Resume Next
End Function

Private Function RNG_GetRandom() As Boolean
On Error GoTo errHandler
Dim phProv As Long
Dim Result1 As Long, Result2 As Long, result3 As Long
Dim MsgStr As String
 
    Result1 = CryptAcquireContext(phProv, ByVal vbNullChar, ByVal vbNullChar, PROV_RSA_FULL, CRYPT_VERIFYCONTEXT)
    Result2 = CryptGenRandom(phProv, RNG_BUFFER_SIZE, RNG_RandomBuffer.Buffer(1))
    result3 = CryptReleaseContext(phProv, 0)
    If (Result1 <> 1) Or (Result2 <> 1) Or (result3 <> 1) Then
        MsgStr = "RNG_GetRandom: Call to " & "CryptAcquireContext32" & " FAILED with error: 0x" & "Hex(Err.LastDllError)"
        MsgBox MsgStr
    End If

    RNG_GetRandom = Result1 And Result2 And result3
Exit Function
errHandler:
    Call Print_Error_Message(Error_Info, "Utility", "RNG_GetRandom")
    If AbortTest Then Exit Function Else Resume Next
End Function

Public Sub Sort(ByRef a() As Long, Optional isDescend As Boolean = False)
On Error GoTo errHandler
    QuickSort a, LBound(a), UBound(a), isDescend
    InsertionSort a, LBound(a), UBound(a), isDescend
Exit Sub
errHandler:
    Call Print_Error_Message(Error_Info, "Utility", "Sort")
    If AbortTest Then Exit Sub Else Resume Next
End Sub

Private Sub QuickSort(ByRef a() As Long, ByVal L As Long, ByVal R As Long, Optional isDescend As Boolean = False)
On Error GoTo errHandler
Dim m As Long, i As Long, j As Long, v As Long

    m = 4
    If ((R - L) > m) Then
        i = (R + L) / 2
        If isDescend Then
            If (a(L) < a(i)) Then swap a, L, i '// Tri-Median Methode!'
            If (a(L) < a(R)) Then swap a, L, R
            If (a(i) < a(R)) Then swap a, i, R
        Else
            If (a(L) > a(i)) Then swap a, L, i '// Tri-Median Methode!'
            If (a(L) > a(R)) Then swap a, L, R
            If (a(i) > a(R)) Then swap a, i, R
        End If
        j = R - 1
        swap a, i, j
        i = L
        v = a(j)
        Do
            If isDescend Then
                Do: i = i + 1: Loop While (a(i) > v)
                Do: j = j - 1: Loop While (a(j) < v)
            Else
                Do: i = i + 1: Loop While (a(i) < v)
                Do: j = j - 1: Loop While (a(j) > v)
            End If
            If (j < i) Then Exit Do
            swap a, i, j
        Loop
        swap a, i, R - 1
        QuickSort a, L, j, isDescend
        QuickSort a, i + 1, R, isDescend
    End If

Exit Sub
errHandler:
    Call Print_Error_Message(Error_Info, "Utility", "QuickSort")
    If AbortTest Then Exit Sub Else Resume Next
End Sub

Private Sub InsertionSort(ByRef a() As Long, ByVal L As Long, ByVal R As Long, Optional isDescend As Boolean = False)
On Error GoTo errHandler
Dim i As Long, j As Long, v As Long

    For i = L + 1 To R
        v = a(i)
        j = i
        Do While j > L
           If isDescend Then
                If Not a(j - 1) < v Then Exit Do
           Else
                If Not a(j - 1) > v Then Exit Do
           End If
           a(j) = a(j - 1)
           j = j - 1
        Loop
        a(j) = v
    Next i

Exit Sub
errHandler:
    Call Print_Error_Message(Error_Info, "Utility", "InsertionSort")
    If AbortTest Then Exit Sub Else Resume Next
End Sub

Private Sub swap(ByRef a() As Long, ByVal i As Long, ByVal j As Long)
On Error GoTo errHandler
Dim t As Long

    t = a(i)
    a(i) = a(j)
    a(j) = t

Exit Sub
errHandler:
    Call Print_Error_Message(Error_Info, "Utility", "swap")
    If AbortTest Then Exit Sub Else Resume Next
End Sub

Public Sub AddDictionary(keyname As String, ByRef obj As Variant, ByRef dic As Dictionary)
On Error GoTo errHandler
    
    If dic.Exists(keyname) Then
        dic.Remove (keyname)
    End If
    dic.Add keyname, obj

Exit Sub
errHandler:
    Call Print_Error_Message(Error_Info, "Utility", "AddDictionary")
    If AbortTest Then Exit Sub Else Resume Next
End Sub


Public Function GetDictionary(keyname As String, dic As Dictionary) As String
On Error GoTo errHandler

    If Not dic.Exists(keyname) Then
        TheExec.ErrorLogMessage "Stored capture data " & keyname & " not found."
    Else
        GetDictionary = dic(keyname)
    End If

Exit Function
errHandler:
    Call Print_Error_Message(Error_Info, "Utility", "GetDictionary")
    If AbortTest Then Exit Function Else Resume Next
End Function

Public Function Get_Sheet_Info(sheetName As String, SheetMaxRow As Long, SheetMaxCol As Long, ParsingArr() As Variant)
On Error GoTo errHandler
Dim MySheet As New Worksheet
Dim m_StartRow As Long: m_StartRow = 1

    Set MySheet = Sheets(sheetName)         ''Active sheet
    MySheet.Activate

    SheetMaxRow = MySheet.UsedRange.Rows.Count    ''Get the max used rows and columns of BDF.
    SheetMaxCol = MySheet.UsedRange.Columns.Count

    ParsingArr = MySheet.range(Cells(m_StartRow, 1), Cells(SheetMaxRow, SheetMaxCol)).value   ''Parsing BDF sheet to array

Exit Function
errHandler:
    Call Print_Error_Message(Error_Info, "Utility", "Get_Sheet_Info")
    If AbortTest Then Exit Function Else Resume Next
End Function

Public Function IsNumber(value As String) As Boolean
On Error GoTo errHandler

    If IsStrMatch(value, "^b[0|1|\_]+$") Or IsStrMatch(value, "(^0x[A-F0-9|\_]+|^x[A-F0-9|\_]+)$") Or IsStrMatch(value, "^\d+$") Then
        IsNumber = True
    Else
        IsNumber = False
    End If

Exit Function
errHandler:
    Call Print_Error_Message(Error_Info, "Utility", "IsNumber")
    If AbortTest Then Exit Function Else Resume Next
End Function

Public Sub SetEfuseErrorBin(bin As Long, sortBin As Long)
On Error GoTo errHandler
Dim site As Variant

    For Each site In TheExec.sites.Selected
         TheExec.sites.site(site).testResult = siteFail
         TheExec.sites.site(site).result = tlResultFail
         TheExec.sites.site(site).BinNumber = bin
         TheExec.sites.site(site).SortNumber = sortBin
    Next

Exit Sub
errHandler:
    Call Print_Error_Message(Error_Info, "Utility", "SetErrorBin", "runMode:" + CStr(TheExec.RunMode) + " errDes:" + CStr(err.Description))
    
        If AbortTest Then Exit Sub Else Resume Next

End Sub
