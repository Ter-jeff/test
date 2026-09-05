VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "AutoShmooDSSC"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'========== const =========='
Private Const SOURCE_PIN = "JTAG_TDI"
Private Const SOURCE_LABEL = "AutoShmooLabel"
Private Const SOURCE_SIGNAL = "AutoShmooSignal"


'========== fields =========='
Private dsscPatterns_ As Dictionary
Private selsramMappingIndex_ As Integer

'========== constructor =========='
Sub Class_Initialize()
    Set dsscPatterns_ = New Dictionary
End Sub

'========== properties =========='
Private Property Get PatternLabelName(Pattern As String)
    Dim labelNames() As String
    
    labelNames = TheHdw.DSSC.pins(SOURCE_PIN).Pattern(Pattern).Source.Labels.list
    PatternLabelName = IIf(Join(labelNames) = "", SOURCE_LABEL, labelNames(0))
End Property
Public Property Get SelSramMappingIndex() As Integer
    SelSramMappingIndex = selsramMappingIndex_
End Property

'========== methods =========='
Private Function SetupDSSCByPattern(Pattern As String, sourceBits() As Long, WaveName As String)
    Dim dsp As DSPWave
    Dim labelName As String
    
    Set dsp = New DSPWave: dsp.data = sourceBits
    Pattern = autoShmooUtility.GetPatsFromPatSets(Pattern)(0) ' must use pattern name with relative path for dssc setup
    labelName = PatternLabelName(Pattern)
    
    TheExec.WaveDefinitions.CreateWaveDefinition WaveName, dsp, True
    With TheHdw.DSSC.pins(SOURCE_PIN).Pattern(Pattern).Source
        .Signals.Add labelName
        .Signals(labelName).WaveDefinitionName = WaveName
        .Signals(labelName).SampleSize = UBound(sourceBits) + 1
        .Signals(labelName).Amplitude = 1
        .Signals(labelName).LoadSamples
        .Signals(labelName).LoadSettings '' this should only run on flex
        .Signals.DefaultSignal = labelName
    End With
    
    Set dsp = Nothing
End Function
Private Function FindIndexFromMappingTable(Pattern As String) As Integer
    ''' find index out from mapping table to judge the selsram later
    Dim index As Integer
    
    For index = LBound(SelsramMapping) To UBound(SelsramMapping)
        If Pattern Like SelsramMapping(index).blockName And Pattern Like SelsramMapping(index).Pattern Then
            selsramMappingIndex_ = index ' for autoShmooItem usage
            FindIndexFromMappingTable = index
            Exit Function
        End If
    Next index
    
    ' can't not find corresponding selsram in mapping table
    FindIndexFromMappingTable = -1
End Function
Public Function IsDSSCPattern(Pattern As String) As Boolean
    ''' check if pattern is belong to dssc
    If Pattern Like "*DSSC*" And Not dsscPatterns_.Exists(Pattern) Then
        dsscPatterns_.Add Pattern, FindIndexFromMappingTable(Pattern)
        IsDSSCPattern = True
    Else
        IsDSSCPattern = False
    End If
End Function

Public Function SetupDSSC()
    ''' setup dssc based on current condition during shmoo
    Dim Pattern As Variant
    Dim index As Integer
    Dim sramZero As Integer
    Dim sramOne As Integer
    Dim logicPin As String
    Dim SramPin As String
    Dim selsramBits() As Long
    
    For Each Pattern In dsscPatterns_.Keys
        If Pattern Like "*XORDSSC*" Then
            ' MFSTP case
            ' maybe we don't need to handle mfstp since we already source it
        Else
            ' selsram case
            For Each site In TheExec.sites.Selected
                ReDim selsramBits(UBound(SelsramMapping(dsscPatterns_.item(Pattern)).bitCount))
                
                For index = LBound(SelsramMapping(dsscPatterns_.item(Pattern)).bitCount) To UBound(SelsramMapping(dsscPatterns_.item(Pattern)).bitCount)
                    logicPin = SelsramMapping(dsscPatterns_.item(Pattern)).logic_Pin(index)
                    SramPin = SelsramMapping(dsscPatterns_.item(Pattern)).SRAM_PIN(index)
                    sramZero = CInt(SelsramMapping(dsscPatterns_.item(Pattern)).SelSrm0(index))
                    sramOne = CInt(SelsramMapping(dsscPatterns_.item(Pattern)).SelSrm1(index))
                    
                    ' need to reverse selsram0 and selsram1
                    selsramBits(index) = IIf(autoShmooUtility.GetVoltageFromHW(logicPin, tlDCVSVoltageAlt)(site) > autoShmooUtility.GetVoltageFromHW(SramPin, tlDCVSVoltageAlt)(site), sramZero, sramOne)
                Next index
                
                SetupDSSCByPattern CStr(Pattern), selsramBits, SOURCE_SIGNAL + "_" + CStr(site)
            Next site
        End If
    Next Pattern
    
End Function
