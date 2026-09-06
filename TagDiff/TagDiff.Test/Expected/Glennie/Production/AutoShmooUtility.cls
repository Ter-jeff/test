VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "AutoShmooUtility"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'========== const =========='
Private Const COREPOWER = "CorePower"

'========== fields =========='
Private Instrument As instrumentUtility

'========== constructor =========='
Sub Class_Initialize()
    Set Instrument = New instrumentUtility
    Instrument.Initialize COREPOWER
End Sub

'========== properties =========='
Public Function IsRunningShmoo() As Boolean
    IsRunningShmoo = TheExec.DevChar.Setups.IsRunning
End Function
Public Function GetCurrentCharSetupName() As String
    GetCurrentCharSetupName = TheExec.DevChar.Setups.ActiveSetupName
End Function
'========== methods =========='
Public Function appendToArray(ary As Variant, value As Variant)
    If Len(Join(ary, "")) = 0 Then
        ReDim ary(0)
    Else
        ReDim Preserve ary(UBound(ary) + 1)
    End If
    
    ary(UBound(ary)) = value
End Function

Public Function IsInitPattern(Pattern As Variant) As Boolean
    Dim info() As String
    
    info = Split(UCase(Pattern), "\") ' get the pattern name without relative path
    IsInitPattern = IIf(Split(info(UBound(info)), "_")(3) Like "IN??", True, False)
End Function

Public Function HasKey(ByRef data As Variant, key As String) As Boolean
    On Error Resume Next
    key = data(key)
    HasKey = (err.number = 0)
End Function

Public Function GetIndex(ByRef data As Variant, val As Variant) As Integer
    Dim index As Integer: index = -1
    
    For index = 1 To data.Count
        If data.item(index) = val Then Exit For
    Next index
    
    GetIndex = index
End Function

Public Function GetPatternGenericName(Pattern As Variant) As String
    Dim RegEx As New RegExp
    Dim matches As Object
    Dim info() As String
    Dim name As String
    Dim index As Integer

    ''' default set to input pattern name
    GetPatternGenericName = Pattern
    ''' get the pattern without path (ex. .\pattern\soc_scan\PP_.......*.pat)
    info = Split(UCase(Pattern), "\")
    ''' get the pattern without module name (ex. PP_TD...*pat:PP_TD...)
    name = Split(info(UBound(info)), ":")(0)
    
    ''' setup regular expression
    With RegEx
        .Global = True ' find all match
        .IgnoreCase = True ' ingore case
        .Pattern = "_\d{1,2}_" ' pattern version (ex. _1_/_12_)
    End With

    ''' start to match the case
    ''' if no any match, it may be generic already.
    ''' if match, then we always get the last match result (ex. PP_TD_3_20990231/PP_TD_1_2_20990231)
    Set matches = RegEx.Execute(name)
    If matches.Count = 0 Then Exit Function
    index = matches(matches.Count - 1).FirstIndex
    GetPatternGenericName = UCase(mid(name, 1, index))
End Function

Public Function GetPatsFromPatSets(patternSet As Variant, Optional isGenericName As Boolean = False) As String()
    Dim PatternAry() As String
    Dim Pattern As String
    Dim PatternCnt As Long
    Dim queue As Collection
    Dim name As Variant
     
    Set queue = New Collection
    queue.Add patternSet

    While queue.Count > 0
        Pattern = queue(1): queue.Remove 1 ' pop value from first one
        PatternAry = TheExec.DataManager.Raw.GetPatternsInSet(Pattern, PatternCnt)
        If PatternCnt = 0 Then
            autoShmooUtility.appendToArray GetPatsFromPatSets, IIf(isGenericName, GetPatternGenericName(Pattern), UCase(Pattern))
            GoTo NextOne
        End If
        
        For Each name In PatternAry
            queue.Add name
        Next name
NextOne:
    Wend
    
    Set queue = Nothing
End Function

Public Function SwitchVmainValt(Output As tlDCVSVoltageOutput)
    TheHdw.DCVS.pins(Instrument.GetDCVSPins).Voltage.Output = Output
End Function

Public Function GetGlobalValue(name As String) As SiteDouble
    Dim site As Variant
    
    Set GetGlobalValue = TheExec.Specs.Globals(name).CurrentValue
    For Each site In TheExec.sites.Selected
        GetGlobalValue = IIf(GetGlobalValue < 0, 0, GetGlobalValue)
    Next site
End Function

Public Function GetVoltageFromHW(pin As String, Output As tlDCVSVoltageOutput) As SiteDouble
    Select Case gl_GetInstrumentType_Dic(LCase(pin))
        Case "DCVS", "DCVSMerged2", "DCVSMerged4", "DCVSMerged8":
            If Output = tlDCVSVoltageMain Then
                Set GetVoltageFromHW = TheHdw.DCVS.pins(pin).Voltage.Main.ValuePerSite
            Else
                Set GetVoltageFromHW = TheHdw.DCVS.pins(pin).Voltage.Alt.ValuePerSite
            End If
        Case "DCVI":
            Set GetVoltageFromHW = TheHdw.DCVI.pins(pin).Voltage.ValuePerSite
        Case Else:
            ' something wrong here
            Exit Function
    End Select
End Function

Public Function GetInitSequence(Pattern As String) As Integer
    Dim info() As String
    Dim VERSION As String
    
    ''' get the pattern version located at index 3
    ''' ex. PP_THRA0_S_IN01_...
    info = Split(Pattern, "_")
    VERSION = mid(info(3), 3)
    
    If IsNumeric(VERSION) Then
        GetInitSequence = CInt(VERSION)
    Else
        GetInitSequence = -1
    End If
End Function

Public Function GetClusterName(Pattern As String) As String
    Dim info() As String
    
    info = Split(Pattern, "_")
    If UBound(info) > 4 Then GetClusterName = info(5)
End Function

Public Function GetServerName(Pattern As String, Optional length As Integer = 4) As String
    Dim info() As String
    
    info = Split(Pattern, "_")
    If UBound(info) > 10 Then GetServerName = mid(info(11), 1, length)
End Function

Public Function DecomposePin(pin As String) As String()
    Dim queue As Collection
    Dim pinAry() As String
    Dim pinCnt As Long
    Dim val As Variant
    
    Set queue = New Collection: queue.Add pin
    
    While queue.Count > 0
        pin = queue.item(1): queue.Remove 1
        TheExec.DataManager.DecomposePinList pin, pinAry, pinCnt
        If pinCnt = 1 Then
            appendToArray DecomposePin, pin
            GoTo nextLoop
        End If
        
        For Each val In pinAry
            queue.Add val
        Next val
nextLoop:
    Wend
End Function

Public Function PrintOutVoltage(Optional pins As String = COREPOWER)
    Dim val As Variant
    Dim pin As Variant
    Dim pinAry() As String
    
    For Each val In Split(pins, ",")
        For Each pin In DecomposePin(Trim(val))
            If Not gl_GetInstrumentType_Dic.Exists(LCase(pin)) Then GoTo nextPin
                
            If LCase(gl_GetInstrumentType_Dic(LCase(pin))) Like "*dcvs*" Then
                Debug.Print UCase(pin) + " -> " + Format(TheHdw.DCVS.pins(pin).Voltage.value, "0.000")
            ElseIf LCase(gl_GetInstrumentType_Dic(LCase(pin))) Like "*dcvi*" Then
                Debug.Print UCase(pin) + " -> " + Format(TheHdw.DCVI.pins(pin).Voltage.value, "0.000")
            Else
                ' something wrong here
            End If
nextPin:
        Next pin
    Next val
End Function
