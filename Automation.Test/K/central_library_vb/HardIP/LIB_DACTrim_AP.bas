Attribute VB_Name = "LIB_DACTrim_AP"
Option Explicit
Enum SearchMethod
    DoAll = 0
    linear = 1
    Binary = 2
    Interpolation = 3
    Transitions = 4
    'PS:
    '1. SearchMethod Max num do not over 9 because of it will split searchmethod into unit decimal
    '2. Boundary = 9 reserve 9 for Boundary search
End Enum
Enum TargetCondition
    Equal = 0
    GreaterThanEqual = 1
    LessThanEqual = 2
    Transition = 3
End Enum
Enum EventSourceTerminateMode
     VOH_HIZ = 5
     VOL_HIZ = 6
     BOTH_HIZ = 4
     VOH_VT = 2
     VOL_VT = 3
     BOTH_VT = 1
End Enum
Enum Freq_TerminateMode
    HIZ = 0
    VT = 1
End Enum
Enum TrimMeasType
    NoTrimCalcName = 0
    SPinListData = 1
    SDSPWave = 2
End Enum
'Enum InstrumentSetup
'     DEFAULT_SETUP = 0
'     DigitalConnectPPMU = 1
'End Enum
Type MeasIConditions
    PinName As String
    CurrentRange As Double
End Type
Public Const LSBFirst = True
Public StepIdx As Long
Public RV As Boolean
Public RPIndex As New DSPWave
Public G_TrimWave As New DSPWave
Public TrimDebug As Boolean
Public SearchDone As New SiteBoolean 'for transition used
Public TrimStoreType As TrimMeasType ' Check trimCalcName store type

Public Sub BinStr2DWave(InputStr As SiteVariant, outwave As DSPWave, Optional WaveName As String)
On Error GoTo errHandler 'Add ErrHandler 2023/05/29
Dim site As Variant
Dim a As New DSPWave
Dim i As Long
Dim SampleSize As Long
    For Each site In TheExec.sites
        SampleSize = Len(InputStr)
        If SampleSize > 0 Then Exit For
    Next site
    a.CreateConstant 0, SampleSize
    
    For Each site In TheExec.sites
        For i = 1 To SampleSize
            a.Element(i - 1) = mid(InputStr, i, 1)
        Next i
        If WaveName <> "" Then a.info.WaveName = WaveName
    Next site
    outwave = a
Exit Sub 'Add ErrHandler 2023/05/29
errHandler: 'Add ErrHandler 2023/05/29
    Call Print_Error_Message(Error_Info, "LIB_DACTrim_AP", "BinStr2DWave") 'Add ErrHandler 2023/05/29
    If AbortTest Then Exit Sub Else Resume Next 'Add ErrHandler 2023/05/29
End Sub

Public Function GetInstrumentType(PinList As String, site As Variant) As String
On Error GoTo errHandler 'Add ErrHandler 2023/05/29
    Dim chanString As String
    Dim PinName() As String
    Dim NumberPins As Long
    Call TheExec.DataManager.DecomposePinList(PinList, PinName(), NumberPins)
    Call TheExec.DataManager.GetChannelStringFromPinAndSite(PinName(0), site, chanString)
    Dim slotstr() As String
    Dim slot As Long
    If chanString = "" Then
        MsgBox ("Please check pin type of  " & PinList & " in channel map")
    Else
        slotstr = Split(chanString, ".")
        slot = CLng(slotstr(0))
        GetInstrumentType = TheHdw.config.Slots(slot).type
    End If
Exit Function 'Add ErrHandler 2023/05/29
errHandler: 'Add ErrHandler 2023/05/29
    Call Print_Error_Message(Error_Info, "LIB_DACTrim_AP", "GetInstrumentType") 'Add ErrHandler 2023/05/29
    If AbortTest Then Exit Function Else Resume Next 'Add ErrHandler 2023/05/29
End Function

Public Function Dec2BinStr(DecVal As Long, width As Long) As Variant
On Error GoTo errHandler 'Add ErrHandler 2023/05/29
'53,7=0110101'
'8,3=overflow'
'-1,3=overflow'
    If 2 ^ width - 1 < DecVal Or DecVal < 0 Then Dec2BinStr = "OverFlow": Exit Function
    If width > 0 Then
        Dec2BinStr = CStr(Dec2BinStr(Fix(DecVal / 2), width - 1)) & CStr((DecVal Mod 2))
    End If
Exit Function 'Add ErrHandler 2023/05/29
errHandler: 'Add ErrHandler 2023/05/29
    Call Print_Error_Message(Error_Info, "LIB_DACTrim_AP", "Dec2BinStr") 'Add ErrHandler 2023/05/29
    If AbortTest Then Exit Function Else Resume Next 'Add ErrHandler 2023/05/29
End Function

Public Sub optInv(IntervalF As Double, TargetF As Double)
On Error GoTo errHandler 'Add ErrHandler 2023/05/29

IntervalF = 1 / TargetF * Fix(0.010485602 / (1 / TargetF))

Exit Sub 'Add ErrHandler 2023/05/29
errHandler: 'Add ErrHandler 2023/05/29
    Call Print_Error_Message(Error_Info, "LIB_DACTrim_AP", "optInv") 'Add ErrHandler 2023/05/29
    If AbortTest Then Exit Sub Else Resume Next 'Add ErrHandler 2023/05/29
End Sub

Public Sub GetTrimCalcNameType(TrimCalcName As String, Meas_StoreName As String, CalcEq As String)
On Error GoTo errHandler 'Add ErrHandler 2023/05/29
Dim Temp1 As String
Dim Temp2 As String
Dim Temp3 As String
Dim Temp4() As String
Dim i As Integer

    Temp1 = LCase(TrimCalcName)
    Temp2 = LCase(Meas_StoreName)
    Temp3 = LCase(CalcEq)
        If CalcEq Like "*Calc_2S_Complement_To_SignDec_Modified*" Then
                TrimStoreType = SDSPWave: Exit Sub
    End If
    If TrimCalcName = "" Then
        TrimStoreType = NoTrimCalcName
    ElseIf TrimCalcName <> "" And InStr(1, Meas_StoreName, TrimCalcName) > 0 Then
        TrimStoreType = SPinListData
    ElseIf CalcEq <> "" Then
        Temp4 = Split(CalcEq, ";")
        For i = 0 To UBound(Temp4)
            If InStr(1, Temp4(i), TrimCalcName) > 0 Then
                Temp4 = Split(Temp4(i), ":")
                If Temp4(0) <> "C" Then
                    TrimStoreType = SPinListData: Exit Sub
                Else
                    TrimStoreType = SDSPWave: Exit Sub
                End If
            End If
        Next i
        TrimStoreType = SPinListData
    Else
        TrimStoreType = SDSPWave
    End If
Exit Sub 'Add ErrHandler 2023/05/29
errHandler: 'Add ErrHandler 2023/05/29
    Call Print_Error_Message(Error_Info, "LIB_DACTrim_AP", "GetTrimCalcNameType") 'Add ErrHandler 2023/05/29
    If AbortTest Then Exit Sub Else Resume Next 'Add ErrHandler 2023/05/29
End Sub

Public Sub StoreDoAll(DefaultSet As Boolean)
On Error GoTo errHandler 'Add ErrHandler 2023/05/29
    If TheExec.RunOptions.DoAll = True Then
        DefaultSet = True
    Else
        TheExec.RunOptions.DoAll = True
        DefaultSet = False
    End If
Exit Sub 'Add ErrHandler 2023/05/29
errHandler: 'Add ErrHandler 2023/05/29
    Call Print_Error_Message(Error_Info, "LIB_DACTrim_AP", "StoreDoAll") 'Add ErrHandler 2023/05/29
    If AbortTest Then Exit Sub Else Resume Next 'Add ErrHandler 2023/05/29
End Sub
Public Sub RestoreDoAll(DefaultSet As Boolean)
On Error GoTo errHandler 'Add ErrHandler 2023/05/29
    If TheExec.RunOptions.DoAll <> DefaultSet Then TheExec.RunOptions.DoAll = DefaultSet
Exit Sub 'Add ErrHandler 2023/05/29
errHandler: 'Add ErrHandler 2023/05/29
    Call Print_Error_Message(Error_Info, "LIB_DACTrim_AP", "RestoreDoAll") 'Add ErrHandler 2023/05/29
    If AbortTest Then Exit Sub Else Resume Next 'Add ErrHandler 2023/05/29
End Sub
