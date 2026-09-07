VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "VbumpInfos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Enum parTypes
    pWriteGlobal
    pReadGlobal
    pInitVoltage
    pWriteVoltage
    pReadVoltage
End Enum

Private Enum Suffixes
    sNone = 0
    sTyp
    sMax
    sMin
End Enum

Private Const vrsPins As String = "CorePower"
Private Const srcPin As String = "JTAG_TDI"

''' arguments
Private patterns_ As Collection
Private sweepPins_ As Collection
Private writePars_ As Collection
Private readPars_ As Collection
Private initVol_ As Collection
Private writeVol_ As Collection
Private readVol_ As Collection
Private digSrc_ As String
Private pmodeVol_ As String
Private blockName_ As String
''' internal variables
Private sweepPinsString_ As String
Private writePat_ As String
Private readPat_ As String
Private srcPin_ As New PinList
Private vrsPins_ As Collection
Private vrsPinsVol_ As New PinListData
Private isAutoSwitch_ As Scripting.Dictionary
Private srcPatName_ As New Scripting.Dictionary
Private sweepPinsWritePar_ As Scripting.Dictionary
Private sweepPinsReadPar_ As Scripting.Dictionary
Private sweepPinsInitVol_ As Scripting.Dictionary
Private sweepPinsWriteVol_ As Scripting.Dictionary
Private sweepPinsReadVol_ As Scripting.Dictionary

Sub Class_Initialize()
    Set patterns_ = New Collection
    Set sweepPins_ = New Collection
    Set writePars_ = New Collection
    Set readPars_ = New Collection
    Set initVol_ = New Collection
    Set writeVol_ = New Collection
    Set readVol_ = New Collection
    Set isAutoSwitch_ = New Scripting.Dictionary
    Set srcPatName_ = New Scripting.Dictionary
    Set sweepPinsWritePar_ = New Scripting.Dictionary
    Set sweepPinsReadPar_ = New Scripting.Dictionary
    Set sweepPinsInitVol_ = New Scripting.Dictionary
    Set sweepPinsWriteVol_ = New Scripting.Dictionary
    Set sweepPinsReadVol_ = New Scripting.Dictionary
    Set vrsPins_ = getVrsPins ''' get vrs pins from string
    srcPin_.value = srcPin
End Sub

Public Sub Initialize(patterns As Pattern, SweepPins As String, GlobalForWrite As String, GlobalForRead As String, _
                            VolForInit As String, VolForWrite As String, VolForRead As String, digSrc As String, pmodeVol As String, blockName As String)
    
    If SweepPins <> "" Then sweepPinsString_ = SweepPins: If multiInputs(SweepPins) Then Set sweepPins_ = splitString(SweepPins) Else sweepPins_.Add Trim(SweepPins)
    If GlobalForWrite <> "" Then If multiInputs(GlobalForWrite) Then Set writePars_ = splitString(GlobalForWrite) Else writePars_.Add Trim(GlobalForWrite)
    If GlobalForRead <> "" Then If multiInputs(GlobalForRead) Then Set readPars_ = splitString(GlobalForRead) Else readPars_.Add Trim(GlobalForRead)
    If VolForInit <> "" Then If multiInputs(VolForInit) Then Set initVol_ = splitString(VolForInit) Else initVol_.Add Trim(VolForInit)
    If VolForWrite <> "" Then If multiInputs(VolForWrite) Then Set writeVol_ = splitString(VolForWrite) Else writeVol_.Add Trim(VolForWrite)
    If VolForRead <> "" Then If multiInputs(VolForRead) Then Set readVol_ = splitString(VolForRead) Else readVol_.Add Trim(VolForRead)
    If digSrc <> "" Then digSrc_ = digSrc
    If pmodeVol <> "" Then pmodeVol_ = pmodeVol
    If blockName <> "" Then blockName_ = blockName
    If patterns.value <> "" Then getPatternSet patterns
End Sub

''' ========== Get Property ==========
Public Function IsSelectSramCase() As Boolean
    IsSelectSramCase = IIf(pmodeVol_ <> "", True, False)
End Function
Public Function IsAutoSwitchCase() As Boolean
    IsAutoSwitchCase = IIf(blockName <> "", True, False)
End Function
Public Function IsExistSourceCode() As Boolean
    IsExistSourceCode = IIf(digSrc_ <> "", True, False)
End Function
Public Function IsRunningShmoo() As Boolean
    IsRunningShmoo = TheExec.DevChar.Setups.IsRunning
End Function
Public Function IsShmooFirstRun() As Boolean
    IsShmooFirstRun = TheExec.DevChar.results(TheExec.DevChar.Setups.ActiveSetupName).StartTime Like "1/1/0001*" Or TheExec.DevChar.results(TheExec.DevChar.Setups.ActiveSetupName).StartTime Like "0001/1/1*"
End Function
Public Function IsExistInitVoltage() As Boolean
    IsExistInitVoltage = IIf(initVol_.Count > 0, True, False)
End Function
Public Function IsExistWriteVoltage() As Boolean
    IsExistWriteVoltage = IIf(writeVol_.Count > 0, True, False)
End Function
Public Function IsExistReadVoltage() As Boolean
    IsExistReadVoltage = IIf(readVol_.Count > 0, True, False)
End Function
Public Function patterns() As Collection
    Set patterns = patterns_
End Function
Public Function blockName() As String
    blockName = blockName_
End Function
Public Property Get vrsPinsVol() As PinListData
    Set vrsPinsVol = vrsPinsVol_
End Property
Public Property Get SweepPinsString() As String
    SweepPinsString = sweepPinsString_
End Property

''' ========== Private Function ==========
Private Function multiInputs(str As String) As Boolean
    If InStr(1, str, ",") Then
        multiInputs = True
    Else
        multiInputs = False
    End If
End Function

Private Function splitString(str As String) As Collection
    Dim cols As Collection
    Dim ary() As String
    Dim val As Variant
    
    Set cols = New Collection
    
    ary = Split(str, ",")
    For Each val In ary
        cols.Add Trim(val)
    Next val
    Set splitString = cols
    Set cols = Nothing
End Function

Private Function getPatternSet(pats As Pattern) As Collection
    Dim patset() As String
    Dim patcnt As Long
    Dim index As Integer
    Dim i As Long
    
    If Not IsAutoSwitchCase Then
        index = 1 ''' not autoswitch, means no dssc before read
    Else
        index = 2 ''' autoswtich, means there is additional dssc before read
    End If
    
    GetPatsFromPatSets pats.value, patset, patcnt, False
    For i = LBound(patset) To UBound(patset)
        If i = UBound(patset) - index Then
            writePat_ = patset(i)
        ElseIf i = UBound(patset) Then
            readPat_ = patset(i)
        End If
        patterns_.Add patset(i)
    Next i
End Function

Private Function getVrsPins() As Collection
    Dim cols As Collection
    Dim pinAry() As String
    Dim pinCnt As Long
    Dim pin As Variant
    
    Set cols = New Collection
    
    TheExec.DataManager.DecomposePinList vrsPins, pinAry, pinCnt
    For Each pin In pinAry
        cols.Add pin
    Next pin
    
    Set getVrsPins = cols
    Set cols = Nothing
End Function

Private Function getPreviousPar(pars As Collection, ByVal index As Integer) As String
    Do While True
        If index = 0 Then Exit Function ''' should not happen
        If pars(index) <> "" Then
            getPreviousPar = pars(index)
            Exit Do
        Else
            index = index - 1
        End If
    Loop
End Function

Private Function setParToPins(pars As Collection, ByRef dist As Scripting.Dictionary, ByVal index As Integer, parType As parTypes) As Boolean

On Error GoTo errhandle
    ''' check first parameter
    If index = 1 Then
        If pars(index) = "" Then
            Select Case parType
                Case parTypes.pWriteGlobal:
                    TheExec.ErrorLogMessage "First write global is empty, Please check it"
                Case parTypes.pReadGlobal:
                    TheExec.ErrorLogMessage "First read global is empty, Please check it"
                Case parTypes.pInitVoltage:
                    TheExec.ErrorLogMessage "First init voltage is empty, Please check it"
                Case parTypes.pWriteVoltage:
                    TheExec.ErrorLogMessage "First write voltage is empty, Please check it"
                Case parTypes.pReadVoltage:
                    TheExec.ErrorLogMessage "First read voltage is empty, Please check it"
            End Select
            setParToPins = False
            Exit Function
        End If
    End If
    
    If index <= pars.Count Then
        If pars(index) <> "" Then
            dist.Add sweepPins_(index), pars(index)
        Else
            dist.Add sweepPins_(index), getPreviousPar(pars, index)
        End If
    Else
        dist.Add sweepPins_(index), dist(sweepPins_(index - 1))
    End If
    setParToPins = True
    Exit Function
errhandle:
    TheExec.ErrorLogMessage "Vbump set parameter to pins failed, Please check it"
    Stop
    Resume Next
    setParToPins = False
End Function

''' ========== Public Function ==========
Public Function IsInitPattern(pat As String) As Boolean
    If UCase(pat) Like "*_IN??_*" Then
        IsInitPattern = True
    Else
        IsInitPattern = False
    End If
End Function

Public Function IsSourcePattern(pat As String) As Boolean
    IsSourcePattern = srcPatName_.Exists(pat)
End Function

Public Function switchToVmainVol()
    If TheHdw.DCVS.pins(vrsPins).Voltage.Output <> tlDCVSVoltageMain Then
        TheHdw.DCVS.pins(vrsPins).Voltage.Output = tlDCVSVoltageMain
        'PrintRSVoltage False, "VDD_SOC,VDD_SRAM_SOC"
        TheHdw.Wait 0.01
    End If
End Function

Public Function switchToValtVol()
    If TheHdw.DCVS.pins(vrsPins).Voltage.Output <> tlDCVSVoltageAlt Then
        TheHdw.DCVS.pins(vrsPins).Voltage.Output = tlDCVSVoltageAlt
        'PrintRSVoltage False, "VDD_SOC,VDD_SRAM_SOC"
        TheHdw.Wait 0.01
    End If
End Function

Public Function IsWritePattern(pat As String) As Boolean
    IsWritePattern = IIf(pat = writePat_, True, False)
End Function

Public Function IsReadPattern(pat As String) As Boolean
    IsReadPattern = IIf(pat = readPat_, True, False)
End Function

Public Function testPattern(pat As String)
    Dim relays As Integer
    ''' test pattern
    relays = TheHdw.DCVS.pins(vrsPins).Voltage.Output ''' prevent auto switch by pattern
    TheHdw.patterns(pat).test pfAlways, 0
    TheHdw.DCVS.pins(vrsPins).Voltage.Output = relays
End Function

Public Function PreProcess() As Boolean

On Error GoTo errhandle
    ''' check basic parameters
    If patterns_.Count < 3 Then
        TheExec.ErrorLogMessage "Pattern count is lower than 3, Please check it"
        PreProcess = False
        Exit Function
    End If
    
    Dim index As Integer
    ''' set global to each sweep pin
    If sweepPins_.Count > 0 Then
        For index = 1 To sweepPins_.Count
            ''' check write global
            If Not setParToPins(writePars_, sweepPinsWritePar_, index, pWriteGlobal) Then PreProcess = False: Exit Function
            
            ''' check read global
            If Not setParToPins(readPars_, sweepPinsReadPar_, index, pReadGlobal) Then PreProcess = False: Exit Function
            
            ''' set voltage for init pattern
            If initVol_.Count > 0 Then
                If Not setParToPins(initVol_, sweepPinsInitVol_, index, pInitVoltage) Then PreProcess = False: Exit Function
            End If
            
            ''' set voltage for write pattern
            If writeVol_.Count > 0 Then
                If Not setParToPins(writeVol_, sweepPinsWriteVol_, index, pWriteVoltage) Then PreProcess = False: Exit Function
            End If
            
            ''' set voltage for read
            If readVol_.Count > 0 Then
                If Not setParToPins(readVol_, sweepPinsReadVol_, index, pReadVoltage) Then PreProcess = False: Exit Function
            End If
        Next index
    End If
    
    PreProcess = True
    Exit Function
errhandle:
    TheExec.ErrorLogMessage "Vbump preprocess failed, Please check it"
    PreProcess = False
End Function

Public Function ApplyTimingLevel()
    TheHdw.Digital.ApplyLevelsTiming True, True, True, tlPowered
End Function

Public Function setVrsPinsVoltage() As Boolean
    Dim vol As Double
    Dim pin As Variant
    Dim ary() As String
    Dim specName As String
    Dim suffix As Suffixes
    
    ''' apply pmode voltage to valt
    For Each pin In vrsPins_
                If TheExec.DataManager.ChannelType(pin) = "N/C" Then GoTo nextP
        If IsSelectSramCase Then
            ''' get pmode name and suffix
            If suffix = Suffixes.sNone Then
                If InStr(1, pmodeVol_, ":") Then
                    ary = Split(pmodeVol_, ":")
                    Select Case Trim(UCase(ary(1)))
                        Case "HV":
                            suffix = Suffixes.sMax
                        Case "NV":
                            suffix = Suffixes.sTyp
                        Case "LV":
                            suffix = Suffixes.sMin
                    End Select
                    pmodeVol_ = Trim(ary(0))
                Else
                    suffix = Suffixes.sTyp
                End If
            End If
            
            ''' check if spec exist VOP or not
            If TheExec.Specs.Contains(pin & "_VOP_VAR") Then
                specName = pin & "_VOP_VAR"
            ElseIf TheExec.Specs.Contains(pin & "_VAR") Then
                specName = pin & "_VAR"
            Else
                TheExec.ErrorLogMessage "Can not find DC spec name: " & pin & "_VAR or " & pin & "_VOP_VAR" & ", Please check it!!"
                setVrsPinsVoltage = False: Exit Function
            End If
            
            Select Case suffix
                Case Suffixes.sMax:
                    vol = Format(CStr(TheExec.Specs.DC.item(specName).Categories(pmodeVol_).Max.value), "0.000")
                Case Suffixes.sTyp:
                    vol = Format(CStr(TheExec.Specs.DC.item(specName).Categories(pmodeVol_).Typ.value), "0.000")
                Case Suffixes.sMin:
                    vol = Format(CStr(TheExec.Specs.DC.item(specName).Categories(pmodeVol_).Min.value), "0.000")
            End Select
            ''' apply
            If IsRunningShmoo Then
                If IsShmooFirstRun Then
                    TheHdw.DCVS.pins(pin).Voltage.Alt.value = vol
                End If
            Else
                TheHdw.DCVS.pins(pin).Voltage.Alt.value = vol
            End If
            ''' store for auto switch if necessary
            vrsPinsVol_.AddPin pin
            vrsPinsVol_.pins(pin).value = vol
        Else
            vrsPinsVol_.AddPin pin
            If IsRunningShmoo Then
                If IsShmooFirstRun Then
                    vrsPinsVol_.pins(pin).value = TheHdw.DCVS.pins(pin).Voltage.Main.value
                    TheHdw.DCVS.pins(pin).Voltage.Alt.value = vrsPinsVol_.pins(pin).value ''' store init voltage to valt for non select sram case
                Else
                    vrsPinsVol_.pins(pin).value = TheHdw.DCVS.pins(pin).Voltage.Alt.value ''' get init voltage
                End If
            Else
                vrsPinsVol_.pins(pin).value = TheHdw.DCVS.pins(pin).Voltage.Main.value ''' store init voltage for non select sram case
            End If
        End If
nextP:
    Next pin
    
    setVrsPinsVoltage = True
End Function

Public Function SetupDigSrc(Optional autoSwitchInfo As AutoSwitchInfos) As Boolean

    Dim switchCase As Boolean
    Dim patIndex As Long
    Dim Pattern As String
    Dim SrcSeqs() As String
    Dim SrcSeq() As String
    Dim SrcBits_tmp() As String
    Dim SrcStr As String
    Dim SrcSize As Long
    Dim srcWave() As New DSPWave
    Dim patSets() As String
    Dim patCnts As Long
    Dim site As Variant
    Dim Src As Variant
    
    SrcSeqs = Split(digSrc_, ",")
    ReDim srcWave(UBound(SrcSeqs))
    For Each Src In SrcSeqs
        ''' split source string
        SrcSeq = Split(Src, ":")
        patIndex = Trim(SrcSeq(0)) ''' patindex
        Pattern = patterns_(patIndex) ''' pattern
        SrcStr = Trim(SrcSeq(1)) ''' source string
        SrcSize = Len(SrcStr) ''' source size
        
'        ''' if auto switch case then store pattern for voltage sweeping
        If InStr(1, SrcStr, "s", vbTextCompare) Then
            If Not isAutoSwitch_.Exists(Pattern) Then
                isAutoSwitch_.Add Pattern, True
            End If
            switchCase = True
        Else
            If Not isAutoSwitch_.Exists(Pattern) Then
                isAutoSwitch_.Add Pattern, False
            End If
            switchCase = False
'            If isAutoSwitch_.Exists(pattern) Then
'                theexec.ErrorLogMessage "Using auto switch on same pattern: " & pattern & ", Please check it"
'                SetupDigSrc = False
'                Exit Function
'            End If
       End If
        
        ''' avoid src pattern is not in patset
        If patIndex > patterns_.Count Then
            TheExec.ErrorLogMessage "Source pattern index is higher than total pattern count, Please check it"
            SetupDigSrc = False
            Exit Function
        End If
        
        ''' avoid source different source code on same pattern
        If srcPatName_.Exists(Pattern) Then
            If Not switchCase Then
                If Not IsAutoSwitchCase Then
                    TheExec.ErrorLogMessage "Source different source-code on same pattern: " & Pattern & ", Please check it"
                    SetupDigSrc = False
                    Exit Function
                Else
                    GoTo nextSrc ''' hard code dssc pattern, no need to setup again
                End If
            Else
                ''' already setup dssc pattern, just exit
                SetupDigSrc = True
                Exit Function
            End If
        End If
        
        ''' convert srcbit string to srcbit array to apply into dspwave
        SrcBits_tmp = Split(StrConv(SrcStr, vbUnicode), Chr$(0))
        ''' remove last useless element
        ReDim Preserve SrcBits_tmp(UBound(SrcBits_tmp) - 1)
        ''' to make sure srcbit size equal srcbit length
        If UBound(SrcBits_tmp) + 1 <> SrcSize Then
            TheExec.ErrorLogMessage "Source bit count is not equal to source bit length, Please check it"
            SetupDigSrc = False
            Exit Function
        Else
            For Each site In TheExec.sites.Active
                ''' transfer to long from string
                Dim SrcStr_Temp As New Scripting.Dictionary
                Dim SrcBits() As Long
                Dim switchBits As String
                Dim sramBit As String
                Dim siteIndex As Integer
                Dim index As Integer
                Dim i As Long
                
                ReDim SrcBits(UBound(SrcBits_tmp))
                For i = LBound(SrcBits_tmp) To UBound(SrcBits_tmp)
                    ''' source code
                    If IsNumeric(SrcBits_tmp(i)) Then
                        SrcBits(i) = CLng(SrcBits_tmp(i))
                    '''auto switch
                    ElseIf UCase(SrcBits_tmp(i)) = "S" Then
                        ''' check if block name in mapping table
                        If Not autoSwitchInfo.Blocks.Exists(UCase(blockName_)) Then
                            TheExec.ErrorLogMessage "Can not find block name: " & blockName_ & " in mapping table, Please check it"
                            SetupDigSrc = False
                            Exit Function
                        End If
                        ''' check if select sram pin higher than table
                        If index > autoSwitchInfo.getMaxPinCount Then
                            TheExec.ErrorLogMessage "Select sram pin count is higher than pin count in mapping table, Please check it"
                            SetupDigSrc = False
                            Exit Function
                        End If
                        
'                        SrcBits_tmp(i) = autoSwitchInfo.getSelectSramBit(blockName_, vrsPinsVol_, index)
'                        switchBits = switchBits & CStr(SrcBits_tmp(i))
                        sramBit = CStr(autoSwitchInfo.getVbumpSelectSramBit(blockName_, vrsPinsVol_, index, site))
                        SrcBits(i) = CLng(sramBit)
                        switchBits = switchBits & sramBit
                        index = index + 1
                    Else
                        TheExec.ErrorLogMessage "Can not recognize source bit: " & SrcBits_tmp(i) & ", Please check it"
                        SetupDigSrc = False
                        Exit Function
                    End If
                Next i
            
                If switchCase Then
                    ''' prevent select sram pin count is lower than mapping table
                    If index - 1 < autoSwitchInfo.getMaxPinCount Then
                        TheExec.ErrorLogMessage "Select sram pin count is lower than pin count in mapping table, Please check it"
                        SetupDigSrc = False
                        Exit Function
                    End If
                    ''' store source string on each site
'                    SrcStr_Temp.Add CStr(site)
                    SrcStr_Temp(site) = Replace(SrcStr, String(Len(switchBits), "s"), switchBits, compare:=vbTextCompare)
                Else
                    SrcStr_Temp(site) = SrcStr
                End If
                ''' apply to dspwave
                srcWave(srcPatName_.Count).data = SrcBits
                
                ''' initial
                index = 0
                Erase SrcBits
                switchBits = ""
                siteIndex = siteIndex + 1
            Next site
            ''' store soruce string for each site to print datalog
            Set srcPatName_(Pattern) = SrcStr_Temp
            Set SrcStr_Temp = New Scripting.Dictionary
        End If

        ''' setup into dsp wave
        Call GetPatsFromPatSets(Pattern, patSets, patCnts)
        Call SetupDigSrcDspWave(patSets(0), srcPin_, "BETA_SRC" + CStr(srcPatName_.Count), SrcSize, srcWave(srcPatName_.Count - 1))
nextSrc:
    Next Src
    
    SetupDigSrc = True
End Function

Public Function setVoltageToAxis(patType As patTypes, Optional pat As String, Optional ByRef preType As patTypes)
    
    Dim site As Variant
    Dim index As Integer
    
    ''' switch to another side to change voltage
    Select Case patType
        Case patTypes.pInit:
            If Not patterns_(1) = pat And Not srcPatName_.Exists(pat) Then
                If Not IsAutoSwitchCase Then
                    Exit Function ''' only apply voltage at frist init pattern
                Else
                    If Not isAutoSwitch_(pat) Then Exit Function ''' not hit second dssc pattern, exit function
                End If
            End If
            If IsSelectSramCase And IsExistInitVoltage Then switchToValtVol ''' force condition in select sram, switch to valt to apply voltage
        Case patTypes.pWrite, patTypes.pRead:
            If IsSelectSramCase Then switchToVmainVol ''' select sram case, switch to vmain to apply voltage
    End Select
    
    For index = 1 To sweepPins_.Count
        For Each site In TheExec.sites.Active
            Select Case patType
                ''' set init voltage
                Case patTypes.pInit:
                    If IsExistInitVoltage Then
                        ''' init force condition, replace init voltage in dc spec
                        TheHdw.DCVS.pins(sweepPins_(index)).Voltage.Main.value = CDbl(sweepPinsInitVol_(sweepPins_(index)))
                        vrsPinsVol_.pins(sweepPins_(index)).value = TheHdw.DCVS.pins(sweepPins_(index)).Voltage.Main.value ''' store for auto switch
                        'vrsPinsVol_.pins(sweepPins_(index)).Value = CDbl(sweepPinsInitVol_(sweepPins_(index)))
                    Else
                        ''' set to init voltage from dc spec
                        If Not IsSelectSramCase Then
                            TheHdw.DCVS.pins(sweepPins_(index)).Voltage.Main.value = vrsPinsVol_.pins(sweepPins_(index)).value
                        Else
                            If Not IsAutoSwitchCase Then
                                switchToVmainVol
                                vrsPinsVol_.pins(sweepPins_(index)).value = TheHdw.DCVS.pins(sweepPins_(index)).Voltage.Main.value ''' store for auto switch
                            Else
                                Select Case preType
                                    Case patTypes.pWrite:
                                        If IsExistWriteVoltage Then
                                            vrsPinsVol_.pins(sweepPins_(index)).value = CDbl(sweepPinsWriteVol_(sweepPins_(index)))
                                        Else
                                            vrsPinsVol_.pins(sweepPins_(index)).value = TheExec.Specs.Globals(sweepPinsWritePar_(sweepPins_(index))).CurrentValue(site)
                                        End If
                                    Case patTypes.pRead:
                                        If IsExistReadVoltage Then
                                            vrsPinsVol_.pins(sweepPins_(index)).value = CDbl(sweepPinsReadVol_(sweepPins_(index)))
                                        Else
                                            vrsPinsVol_.pins(sweepPins_(index)).value = TheExec.Specs.Globals(sweepPinsReadPar_(sweepPins_(index))).CurrentValue(site)
                                        End If
                                End Select
                            End If
                        End If
                    End If
                ''' set write voltage
                Case patTypes.pWrite:
                    If IsSelectSramCase Then
                        ''' set voltage to Valt since select sram used valt for payload
                        If IsExistWriteVoltage Then
                            ''' write force condition, replace write voltage in shmoo
                            TheHdw.DCVS.pins(sweepPins_(index)).Voltage.Alt.value = CDbl(sweepPinsWriteVol_(sweepPins_(index)))
                            vrsPinsVol_.pins(sweepPins_(index)).value = CDbl(sweepPinsWriteVol_(sweepPins_(index)))
                        Else
                            ''' set to write voltage from shmoo
                            TheHdw.DCVS.pins(sweepPins_(index)).Voltage.Alt.value = TheExec.Specs.Globals(sweepPinsWritePar_(sweepPins_(index))).CurrentValue(site)
                            vrsPinsVol_.pins(sweepPins_(index)).value(site) = TheExec.Specs.Globals(sweepPinsWritePar_(sweepPins_(index))).CurrentValue(site)
                        End If
                    Else
                        ''' set voltage to Vmain
                        If IsExistWriteVoltage Then
                            ''' write force condition, replace write voltage in shmoo
                            TheHdw.DCVS.pins(sweepPins_(index)).Voltage.Main.value = CDbl(sweepPinsWriteVol_(sweepPins_(index)))
                            vrsPinsVol_.pins(sweepPins_(index)).value = CDbl(sweepPinsWriteVol_(sweepPins_(index)))
                        Else
                            ''' set to write voltage from shmoo
                            TheHdw.DCVS.pins(sweepPins_(index)).Voltage.Main.value = TheExec.Specs.Globals(sweepPinsWritePar_(sweepPins_(index))).CurrentValue(site)
                            vrsPinsVol_.pins(sweepPins_(index)).value(site) = TheExec.Specs.Globals(sweepPinsWritePar_(sweepPins_(index))).CurrentValue(site)
                        End If
                    End If
                ''' set read voltage
                Case patTypes.pRead:
                    If IsSelectSramCase Then
                        ''' set voltage to Valt since select sram used valt for payload
                        If IsExistReadVoltage Then
                            ''' read force condition, replace read voltage in shmoo
                            TheHdw.DCVS.pins(sweepPins_(index)).Voltage.Alt.value = CDbl(sweepPinsReadVol_(sweepPins_(index)))
                            vrsPinsVol_.pins(sweepPins_(index)).value = CDbl(sweepPinsReadVol_(sweepPins_(index)))
                        Else
                            ''' set to read voltage from shmoo
                            TheHdw.DCVS.pins(sweepPins_(index)).Voltage.Alt.value = TheExec.Specs.Globals(sweepPinsReadPar_(sweepPins_(index))).CurrentValue(site)
                            vrsPinsVol_.pins(sweepPins_(index)).value(site) = TheExec.Specs.Globals(sweepPinsReadPar_(sweepPins_(index))).CurrentValue(site)
                        End If
                    Else
                        ''' set voltage to Vmain
                        If IsExistReadVoltage Then
                            ''' read force condition, replace read voltage in shmoo
                            TheHdw.DCVS.pins(sweepPins_(index)).Voltage.Main.value = CDbl(sweepPinsReadVol_(sweepPins_(index)))
                            vrsPinsVol_.pins(sweepPins_(index)).value = CDbl(sweepPinsReadVol_(sweepPins_(index)))
                        Else
                            ''' set to read voltage from shmoo
                            TheHdw.DCVS.pins(sweepPins_(index)).Voltage.Main.value = TheExec.Specs.Globals(sweepPinsReadPar_(sweepPins_(index))).CurrentValue(site)
                            vrsPinsVol_.pins(sweepPins_(index)).value(site) = TheExec.Specs.Globals(sweepPinsReadPar_(sweepPins_(index))).CurrentValue(site)
                        End If
                    End If
            End Select
        Next site
    Next index
    
    ''' switch to response side to run pattern
    Select Case patType
        Case patTypes.pInit:
            If Not IsAutoSwitchCase Then
                If IsSelectSramCase And IsExistInitVoltage Then switchToVmainVol
            Else
                switchToVmainVol
            End If
        Case patTypes.pWrite, patTypes.pRead:
            If IsSelectSramCase Then switchToValtVol
    End Select
    If srcPatName_.Exists(pat) And preType = patTypes.pWrite Then preType = patTypes.pRead ''' switch to read for next autoswitch
    If TheExec.enableWord("PrintRSVoltage") Then PrintRSVoltage False, sweepPinsString_
End Function

Public Function PrintDigSrcInfo(pat As String)
    Dim site As Variant
    For Each site In TheExec.sites.Active
        TheExec.Datalog.WriteComment "Site:" & CStr(site) & ", DigSrc pattern: " & pat & ",Src Bits = " & CStr(Len(srcPatName_(pat)(site))) & ",Output String [ LSB(L) ==> MSB(R) ]:" & srcPatName_(pat)(site)
    Next site
    If isAutoSwitch_(pat) Then srcPatName_.Remove pat
End Function
