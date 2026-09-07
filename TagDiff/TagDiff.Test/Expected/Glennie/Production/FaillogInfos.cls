VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "FaillogInfos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Enum Suffixes
    sNone = 0
    sTyp
    sMax
    sMin
End Enum

''' select sram pins
'Private Const vrsPins As String = "VDD_AVE,VDD_DCS_DDR,VDD_DISP,VDD_ECPU,VDD_GPU,VDD_PCPU,VDD_SOC,VDD_CPU_SRAM,VDD_SRAM_GPU,VDD_SRAM_SOC,VDD_FIXED,VDD_LOW"
'Private Const vrsPins As String = "VDD_CPU,VDD_SRAM_CPU,VDD_SOC,VDD_SRAM,VDD_FIXED,VDD_LOW,VDD_DISP" ''' this is for debug on skua
Private Const vrsPins As String = "CorePower"
Private Const srcPin As String = "JTAG_TDI"
''' pattern
Private patterns_ As Collection
''' x-axis parameter
Private xAxis_ As Collection
Private xAxisCount_ As Collection
Private xBegs_ As Collection
Private xEnds_ As Collection
Private xStep_ As Double
''' y-axis parameter
Private yAxis_ As Collection
Private yAxisCount_ As Collection
Private yBegs_ As Collection
Private yEnds_ As Collection
Private yStep_ As Double
''' option
Private disComparePins_ As Collection 'set discompare pins
Private capXfailCount_ As Long 'capture how many fail count in each x-axis
Private capYfailCount_ As Long 'capture how many fail count in each y-axis
Private capLoopsCount_ As Long 'capture how many fail count in the looping
Private showVolPins_ As Collection 'show voltage of pins
Private loopTimes_ As Long 'set loop times
Private pinTypes_ As pinTypes 'pin type on y-axis
Private volTypes_ As volTypes 'voltage type in sweeping
Private loopTypes_ As loopTypes 'loop type while sweeping
Private pmodeVol_ As String 'safe voltage for init patterns
Private digSrc_ As String 'digital source pattern index and source code
Private blockName_ As String 'block name in select sram table
''' internal variables
Private sweepPinsString_ As String
Private axis_ As axes
Private vrsPins_ As Collection
Private isAutoSwitch_ As Scripting.Dictionary
''' frequency measurement
Private WaitTime_ As String
Private FreqHighLimit_ As Double
Private FreqLowLimit_ As Double
''' digital source
Private srcPin_ As New PinList
Private srcPatName_ As Scripting.Dictionary
''' ig-xl variables
Private vrsPinsVol_ As New PinListData
Private showFrqPins_ As New PinList 'show frequency of pins
Private xAxisFail_ As New SiteBoolean ''' this is used for detect if x axis exist fail point then X axis count 1
Private yAxisFail_ As New SiteBoolean ''' this is used for detect if x axis exist fail ooint then y axis count 1
Private pfResults_ As New SiteBoolean ''' used for pattern pass/fail result
Private testNumbers_ As New SiteLong ''' used for incremental testnumber
Private activeSites_ As New SiteBoolean ''' used for storing active site from program
Private XactiveSites_ As New SiteBoolean ''' used for storing x axis active site
Private YactiveSites_ As New SiteBoolean ''' used for storing y axis active site
Private LactiveSites_ As New SiteBoolean ''' used for storing loop active site
Private xAxisFailCount_ As New SiteLong ''' used for record x axis fail count
Private yAxisFailCount_ As New SiteLong ''' used for record y axis fail count
Private loopFailCount_ As New SiteLong ''' used for record looping fail count


Sub Class_Initialize()
    Set patterns_ = New Collection
    Set xAxis_ = New Collection
    Set xAxisCount_ = New Collection
    Set xBegs_ = New Collection
    Set xEnds_ = New Collection
    Set yAxis_ = New Collection
    Set yAxisCount_ = New Collection
    Set yBegs_ = New Collection
    Set yEnds_ = New Collection
    Set disComparePins_ = New Collection
    Set showVolPins_ = New Collection
    Set srcPatName_ = New Scripting.Dictionary
    Set isAutoSwitch_ = New Scripting.Dictionary
'    Set vrsPins_ = getVrsPins ''' get vrs pins from string
    Set vrsPins_ = New Collection ''' get vrs pins from string
    pfResults_ = InitPFresult ''' reset pass/fail result
    activeSites_ = InitActiveSites ''' enbale default active sites
    XactiveSites_ = InitActiveSites ''' enbale default active sites
    YactiveSites_ = InitActiveSites ''' enbale default active sites
    LactiveSites_ = InitActiveSites ''' enable default active sites
    xAxisFailCount_ = InitFailCount ''' reset to zero
    yAxisFailCount_ = InitFailCount ''' reset to zero
    loopFailCount_ = InitFailCount ''' reset to zero
    capYfailCount_ = 1 ''' at least one count
    capLoopsCount_ = 0 ''' at least one count
    loopTimes_ = 1 ''' at least one loop time
    pinTypes_ = pinTypes.VolType ''' default shmoo voltage
    volTypes_ = volTypes.mainType ''' default shmoo main voltage
    loopTypes_ = loopTypes.lFixType ''' default fix point loop
    FreqHighLimit_ = 0 ''' high limit for frequency measurement
    FreqLowLimit_ = 0 ''' low limit for frequency measurement
    srcPin_.value = srcPin ''' source pin
End Sub

''' ========== get property ==========
Public Property Get IsSelectSramCase() As Boolean
    IsSelectSramCase = IIf(pmodeVol_ <> "", True, False)
End Property
Public Property Get IsFrequencyPin() As Boolean
    Select Case pinTypes_
        Case pinTypes.VolType:
            IsFrequencyPin = False
        Case pinTypes.frqType:
            IsFrequencyPin = True
    End Select
End Property
Public Property Get IsExistDisComparePins() As Boolean
    IsExistDisComparePins = IIf(disComparePins_.Count > 0, True, False)
End Property
Public Property Get IsExistShowVolPins() As Boolean
    IsExistShowVolPins = IIf(showVolPins_.Count > 0, True, False)
End Property
Public Property Get IsExistShowFrqPins() As Boolean
    IsExistShowFrqPins = IIf(showFrqPins_ <> "", True, False)
End Property
Public Property Get IsExistSourceCode() As Boolean
    IsExistSourceCode = IIf(digSrc_ <> "", True, False)
End Property
Public Property Get IsAutoSwitchCase() As Boolean
    IsAutoSwitchCase = IIf(blockName_ <> "", True, False)
End Property
Public Property Get vrsPinsVol() As PinListData
    Set vrsPinsVol = vrsPinsVol_
End Property
Public Property Get blockName() As String
    blockName = blockName_
End Property
Public Property Get yAxisStepCount() As Integer
    ''' first index is master pin
    If yBegs_(1) <> "" Then
        yAxisStepCount = Round(Abs(CDbl(yEnds_(1)) - CDbl(yBegs_(1))) / yStep_)
    Else
        ''' no need to sweep the axis
        yAxisStepCount = 0
    End If
End Property
Public Property Get xAxisStepCount() As Integer
    ''' first index is master pin
    If xBegs_(1) <> "" Then
        xAxisStepCount = Round(Abs(CDbl(xEnds_(1)) - CDbl(xBegs_(1))) / xStep_)
    Else
        ''' no need to sweep the axis
        xAxisStepCount = 0
    End If
End Property
Public Property Get patterns() As Collection
    Set patterns = patterns_
End Property
Public Property Get LoopTimes() As Long
    LoopTimes = loopTimes_
End Property
Public Property Get SweepPinsString() As String
    SweepPinsString = sweepPinsString_
End Property
Public Property Get LoopCountNumber() As Long
    Select Case loopTypes_
        Case loopTypes.lFixType:
            Select Case axis_
                Case axes.aXaxis:
                    LoopCountNumber = 1
                Case axes.aYaxis:
                    LoopCountNumber = loopTimes_
            End Select
        Case loopTypes.lVolType:
            Select Case axis_
                Case axes.aXaxis:
                    LoopCountNumber = loopTimes_
                Case axes.aYaxis:
                    LoopCountNumber = 1
            End Select
    End Select
End Property

''' ========== private funciton ==========
Private Function concatenateString(str As String) As String
    If sweepPinsString_ = "" Then
        sweepPinsString_ = sweepPinsString_ & str
    Else
        sweepPinsString_ = sweepPinsString_ & "," & str
    End If
End Function
Private Function InitPFresult() As SiteBoolean
    Dim site As Variant
    Dim result As New SiteBoolean
    
    For Each site In TheExec.sites.Active
        result(site) = True
    Next site
    
    Set InitPFresult = result
End Function

Private Function InitActiveSites() As SiteBoolean
    Dim site As Variant
    Dim activeSites As New SiteBoolean
    
    ''' default set to true, then if pat fail, set to false
    For Each site In TheExec.sites.Active
        activeSites(site) = True
    Next site
    
    ''' init select sites since default selected site is zero
    TheExec.sites.Selected = activeSites
    
    Set InitActiveSites = activeSites
End Function

Private Function InitFailCount() As SiteLong
    Dim site As Variant
    Dim Count As New SiteLong
    
    For Each site In TheExec.sites.Active
        Count(site) = 0
    Next site
    
    Set InitFailCount = Count
End Function

Private Function CurrentPinType(pin As String) As pinTypes
    If InStr(1, pin, "VDD", vbTextCompare) Then
        CurrentPinType = pinTypes.VolType
    Else
        CurrentPinType = pinTypes.frqType
    End If
End Function

Private Function CurrentVolOnPin(pin As String) As Double
    CurrentVolOnPin = TheHdw.DCVS.pins(pin).Voltage.value
End Function

Private Function CurrentLoopOrder(Count As Integer) As String
    Select Case Count
        Case 1:
            CurrentLoopOrder = "st"
        Case 2:
            CurrentLoopOrder = "nd"
        Case 3:
            CurrentLoopOrder = "rd"
        Case Else:
            CurrentLoopOrder = "th"
    End Select
End Function

Private Function DecomposeString(str As String) As Collection
    Dim cols As Collection
    Dim pinAry() As String
    Dim pinCnt As Long
    Dim pin As Variant
    
    Set cols = New Collection
    If Not LCase(str) Like "*shift*in*" And Not LCase(str) Like "*tck*" And Not LCase(str) Like "*xo0*" And Not LCase(str) Like "*xi0*" And Not str = "" Then
        TheExec.DataManager.DecomposePinList Trim(str), pinAry, pinCnt
        If Not pinCnt = 0 Then
            For Each pin In pinAry
                cols.Add pin
            Next pin
        Else
            ''' must have at least one value
            cols.Add ""
        End If
    Else
        cols.Add str
    End If
    
    Set DecomposeString = cols
    Set cols = Nothing
End Function

Private Function splitString(str As String, Optional pinType As Boolean = False, Optional axisType As axes) As Collection
    Dim cols As Collection
    Dim ary() As String
    Dim val As Variant
    Dim pinAry() As String
    Dim pinCnt As Long
    Dim pin As Variant
    Dim index As Integer
    Dim Count As Integer
    Dim i As Integer
    
    Set cols = New Collection
    
    ary = Split(str, ",")
    index = 1
    For Each val In ary
        If pinType Then
            TheExec.DataManager.DecomposePinList Trim(val), pinAry, pinCnt
            If Not pinCnt = 0 Then
                For Each pin In pinAry
                    cols.Add Trim(pin)
                Next pin
                If axisType = axes.aXaxis Then
                    xAxisCount_.Add pinCnt, CStr(index)
                ElseIf axisType = axes.aYaxis Then
                    yAxisCount_.Add pinCnt, CStr(index)
                End If
            End If
        Else
            If axisType = axes.aXaxis Then
                Count = xAxisCount_(CStr(index))
            ElseIf axisType = axes.aYaxis Then
                Count = yAxisCount_(CStr(index))
            End If
            For i = 1 To Count
                cols.Add Trim(val)
            Next i
        End If
        index = index + 1
    Next val
    
    ''' must have at least one value
    If cols.Count = 0 Then
        cols.Add ""
    End If
    
    Set splitString = cols
    Set cols = Nothing
End Function

Private Function multiInputs(str As String) As Boolean
    If InStr(1, str, ",") Then
        multiInputs = True
    Else
        multiInputs = False
    End If
End Function

Private Function getVrsPins() As Collection
    Dim cols As Collection
    Dim pins As Dictionary
    Dim pinAry() As String
    Dim pinCnt As Long
    Dim pin As Variant
    Dim p As Variant
    
    Set cols = New Collection
    Set pins = New Dictionary
    
    TheExec.DataManager.DecomposePinList vrsPins, pinAry, pinCnt
    If Not pinCnt = 0 Then
        For Each pin In pinAry
            cols.Add pin
            pins.Add pin, True
        Next pin
    End If
    
    ''' store x-axis pin if pin is not in core power
    For Each pin In xAxis_
        If pin = "" Then GoTo nextX
        'TheExec.DataManager.DecomposePinList pin, pinAry, pinCnt
        'For Each p In pinAry
        If Not pins.Exists(pin) And InStr(1, pin, "VDD", vbTextCompare) Then
            cols.Add pin
            pins.Add pin, True
        End If
        'Next p
nextX:
    Next pin

    ''' store y-axis pin if pin is not in core power
    For Each pin In yAxis_
        If pin = "" Then GoTo nextY
'        TheExec.DataManager.DecomposePinList pin, pinAry, pinCnt
'        For Each p In pinAry
        If Not pins.Exists(pin) And InStr(1, pin, "VDD", vbTextCompare) Then
            cols.Add pin
            pins.Add pin, True
        End If
'        Next p
nextY:
    Next pin
    
    Set getVrsPins = cols
    Set cols = Nothing
    Set pins = Nothing
End Function

Private Function getPatternSet(pats As Pattern) As Collection
    Dim patset() As String
    Dim patcnt As Long
    Dim pat As Variant
    
    GetPatsFromPatSets pats.value, patset, patcnt, False
    For Each pat In patset
        patterns_.Add pat
    Next pat
End Function

Private Function assignActiveSites(ref As SiteBoolean) As SiteBoolean
    Dim activeSites As New SiteBoolean
    Dim site As Variant
    
    For Each site In TheExec.sites.Active
        activeSites(site) = ref(site)
    Next site
    
    Set assignActiveSites = activeSites
End Function

Private Function applyFrqValToPin(ACspecPin As String, FreqValue As Double)

    '/**** Change Frequency by different type ****/'
    If UCase(ACspecPin) Like "*XI0*" Or UCase(ACspecPin) Like "*PCIE*" Then
        Call VaryFreq("XI0_Diff_Port", FreqValue, ACspecPin)
'        Call VaryFreq("XI0_1_Diff_Port", FreqValue, ACspecPin)
'        Call VaryFreq("PCIE_XI0_0_Diff_Port", FreqValue, ACspecPin)
    ElseIf UCase(ACspecPin) Like "*FREF*" Then
        Call VaryFreq("FREF_Port", FreqValue, ACspecPin)
    Else
        Call TheExec.Overlays.ApplyUniformSpecToHW(ACspecPin, FreqValue)
        TheHdw.Wait 0.005
    End If
End Function

Private Function AxisCoefficient(ByVal index As Integer) As Integer
    Dim begs As String
    Dim ends As String
    
    Select Case axis_
        Case axes.aXaxis:
            begs = xBegs_(index)
            ends = xEnds_(index)
        Case axes.aYaxis:
            begs = yBegs_(index)
            ends = yEnds_(index)
    End Select
    
    ''' if same voltage, keep original value
    If CDbl(begs) = CDbl(ends) Then
        AxisCoefficient = 0
    Else
        AxisCoefficient = IIf(CDbl(begs) < CDbl(ends), 1, -1)
    End If
End Function

Private Function AxisStep(ByVal index As Integer, ByRef outIndex As Integer) As Double
    Dim begs As String
    Dim ends As String
    Dim step As Double
    Dim StepCnt As Integer
    
    ''' initial
    outIndex = index
    
    Select Case axis_
        Case axes.aXaxis:
            ''' first pin directly follow original step
            If outIndex = 1 Then
                AxisStep = xStep_
                Exit Function
            End If
            ''' since value is empty, find out last voltage appeared
            If outIndex <= xBegs_.Count Then
                Do While xBegs_(outIndex) = ""
                    outIndex = outIndex - 1
                Loop
                begs = xBegs_(outIndex)
                ends = xEnds_(outIndex)
            Else
                ''' since value is null, used last one value
                outIndex = xBegs_.Count
                begs = xBegs_(outIndex)
                ends = xEnds_(outIndex)
            End If
            step = xStep_
            StepCnt = xAxisStepCount
        Case axes.aYaxis:
            ''' first pin directly follow original step
            If outIndex = 1 Then
                AxisStep = yStep_
                Exit Function
            End If
            ''' since value is empty, find out last voltage appeared
            If outIndex <= yBegs_.Count Then
                Do While yBegs_(outIndex) = ""
                    outIndex = outIndex - 1
                    Exit Do
                Loop
                begs = yBegs_(outIndex)
                ends = yEnds_(outIndex)
            Else
                ''' since value is null, used last one value
                outIndex = yBegs_.Count
                begs = yBegs_(outIndex)
                ends = yEnds_(outIndex)
            End If
            step = yStep_
            StepCnt = yAxisStepCount
    End Select
    
    If StepCnt = 0 Then
        AxisStep = 0
    Else
        AxisStep = Abs((CDbl(ends) - CDbl(begs))) / StepCnt
    End If
End Function

Private Function MeasFrequencyOnPins(MeasPins As PinList)
    Dim MeasPinsData As New PinListData
    
    TheHdw.Digital.pins(MeasPins).Levels.DriverMode = tlDriverModeVt
    Call Freq_MeasFreqSetup(MeasPins, 0.01, BOTH)      '' 20150621 - default d_FreqMeasInterval = 0.01
    '' 20150623 - Add Customize Wait Time
    Call HardIP_Freq_MeasFreqStart(MeasPins, 0.01, MeasPinsData, WaitTime_)          '' 20150621 - default d_FreqMeasInterval = 0.01
    If FreqLowLimit_ <> 0 Or FreqHighLimit_ <> 0 Then
        TheExec.flow.TestLimit MeasPinsData, FreqLowLimit_, FreqHighLimit_, , , , unitHz, , ForceResults:=tlForceNone
    Else
        TheExec.flow.TestLimit MeasPinsData, , , , , , unitHz, , ForceResults:=tlForceNone
    End If
    TheHdw.Digital.pins(MeasPins).Levels.DriverMode = tlDriverModeHiZ
End Function

''' ========== public function ==========
Public Sub initialFaillog(pat As Pattern, xAxis As String, xBegs As String, xEnds As String, xStep As Double, _
                            yAxis As String, yBegs As String, yEnds As String, yStep As Double, _
                            capXfailcount As Long, capYfailcount As Long, LoopTimes As Long, capLoopsCount As Long, _
                            showFrqPins As PinList, showVolPins As String, disComparePins As String, _
                            pinType As pinTypes, VolType As volTypes, loopType As loopTypes, pmodeVol As String, digSrc As String, blockName As String)
                                 
    If pat.value <> "" Then getPatternSet pat
    xStep_ = Trim(xStep)
    yStep_ = Trim(yStep)
    pmodeVol_ = Trim(pmodeVol)
    If multiInputs(xAxis) Then Set xAxis_ = splitString(xAxis, True, axes.aXaxis) Else Set xAxis_ = DecomposeString(xAxis)
    If multiInputs(xBegs) Then Set xBegs_ = splitString(xBegs, False, axes.aXaxis) Else xBegs_.Add Trim(xBegs)
    If multiInputs(xEnds) Then Set xEnds_ = splitString(xEnds, False, axes.aXaxis) Else xEnds_.Add Trim(xEnds)
    If multiInputs(yAxis) Then Set yAxis_ = splitString(yAxis, True, axes.aYaxis) Else Set yAxis_ = DecomposeString(yAxis)
    If multiInputs(yBegs) Then Set yBegs_ = splitString(yBegs, False, axes.aYaxis) Else yBegs_.Add Trim(yBegs)
    If multiInputs(yEnds) Then Set yEnds_ = splitString(yEnds, False, axes.aYaxis) Else yEnds_.Add Trim(yEnds)
    Set vrsPins_ = getVrsPins
    If showVolPins <> "" Then If multiInputs(showVolPins) Then Set showVolPins_ = splitString(showVolPins) Else showVolPins_.Add Trim(showVolPins)
    If disComparePins <> "" Then If multiInputs(disComparePins) Then Set disComparePins_ = splitString(disComparePins) Else disComparePins_.Add Trim(disComparePins)
    If showFrqPins <> "" Then showFrqPins_ = showFrqPins
    If LoopTimes <> 0 Then loopTimes_ = LoopTimes
    If capXfailcount <> 0 Then capXfailCount_ = capXfailcount
    If capYfailcount <> 0 Then capYfailCount_ = capYfailcount
    If capLoopsCount <> 0 Then capLoopsCount_ = capLoopsCount
    If pinType <> pinTypes.None Then pinTypes_ = pinType
    If VolType <> volTypes.None Then volTypes_ = VolType
    If loopType <> loopTypes.None Then loopTypes_ = loopType
    If digSrc <> "" Then digSrc_ = digSrc
    If blockName <> "" Then blockName_ = blockName
    If xAxis <> "" Then concatenateString xAxis
    If yAxis <> "" Then concatenateString yAxis
End Sub

Public Function getVrsPinsToString() As String
    Dim pin As Variant
    Dim index As Integer
    Dim size As Integer
    Dim pinAry() As String
    
    size = vrsPins_.Count - 1
    ReDim pinAry(size)
    index = 0
    For Each pin In vrsPins_
        pinAry(index) = pin
        index = index + 1
    Next pin
    getVrsPinsToString = Join(pinAry, ",")
End Function
Public Function resetXaxisFailCount()
    Dim site As Variant
    
    For Each site In TheExec.sites.Selected
        xAxisFailCount_(site) = 0
    Next site
End Function
Public Function resetYaxisFailCount()
    Dim site As Variant
    
    For Each site In TheExec.sites.Selected
        yAxisFailCount_(site) = 0
    Next site
End Function

Public Function resetToActiveSites()
    TheExec.sites.Selected = activeSites_
End Function
Public Function resetToXactiveSites()
    TheExec.sites.Selected = XactiveSites_
End Function
Public Function resetToYactiveSites()
    Set LactiveSites_ = assignActiveSites(YactiveSites_)
    Set XactiveSites_ = assignActiveSites(YactiveSites_)
    TheExec.sites.Selected = YactiveSites_
End Function
Public Function resetLoopactiveSites()
    Set LactiveSites_ = assignActiveSites(YactiveSites_)
End Function

Public Function IsInitPattern(pat As String) As Boolean
    If UCase(pat) Like "*_IN??_*" Then
        IsInitPattern = True
    Else
        IsInitPattern = False
    End If
End Function

Public Function IsExistAxis(axis As axes) As Boolean
    Dim pin As String

    Select Case axis
        Case axes.aXaxis:
            pin = xAxis_(1)
        Case axes.aYaxis:
            pin = yAxis_(1)
    End Select
    ''' force to judge if it's voltage or frequency type
    pinTypes_ = CurrentPinType(pin)
    ''' set current axis
    axis_ = axis
    
    If pin <> "" Then
        IsExistAxis = True
    Else
        IsExistAxis = False
    End If
End Function

Public Function IsSourcePattern(pat As String) As Boolean
    IsSourcePattern = IIf(srcPatName_.Exists(pat), True, False)
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
            If suffix = Suffixes.sNone Then
                ''' get pmode name and suffix
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
            If Not UCase(pin) Like "*GRP*" And Not UCase(pin) Like "*GROUP*" Then
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
            End If
            ''' apply
            TheHdw.DCVS.pins(pin).Voltage.Alt.value = vol
            ''' store for auto switch if necessary
            vrsPinsVol_.AddPin pin
            vrsPinsVol_.pins(pin).value = vol
            
            ''' select sram case, force to apply on Valt
            If volTypes_ <> volTypes.altType Then volTypes_ = volTypes.altType
        Else
            vrsPinsVol_.AddPin pin
            vrsPinsVol_.pins(pin).value = TheHdw.DCVS.pins(pin).Voltage.Main.value
        End If
nextP:
    Next pin
    
    setVrsPinsVoltage = True
End Function

Public Function PrintInfos(info As Infos, Optional PinName As String, Optional val As Double, Optional Count As Long, Optional pat As String)
    Select Case info
        ''' print faillog start
        Case Infos.iStart:
            TheExec.Datalog.WriteComment ""
            TheExec.Datalog.WriteComment "************** < FAILLOG START > **************"
        ''' print faillog end
        Case Infos.iEnd:
            TheExec.Datalog.WriteComment "************** < FAILLOG END > **************"
        ''' print blank line
        Case Infos.iBlank:
            TheExec.Datalog.WriteComment "***********************************************"
        ''' print count during looping
        Case Infos.iLoop:
            TheExec.Datalog.WriteComment "***********************************************"
            TheExec.Datalog.WriteComment " <<<<<< Loop Voltage Step, " & (Count) & " Times >>>>>> "
        ''' print x axis start
        Case Infos.iXstart:
            Select Case pinTypes_:
                Case pinTypes.VolType:
                    TheExec.Datalog.WriteComment "***********************************************"
                    TheExec.Datalog.WriteComment "**** The x-axis " & (Count + 1) & CurrentLoopOrder(Count + 1) & " Time to Change Voltage *****"
                    TheExec.Datalog.WriteComment "*************** Voltage Setting ***************"
                    TheExec.Datalog.WriteComment "***********************************************"
                Case pinTypes.frqType:
                    TheExec.Datalog.WriteComment ""
                    TheExec.Datalog.WriteComment "***********************************************"
                    TheExec.Datalog.WriteComment "*** The x-axis " & (Count + 1) & CurrentLoopOrder(Count + 1) & " Time to Change Frequency ***"
'                    TheExec.Datalog.WriteComment "********** " & PinName & " = " & val & " Hz *********"
            End Select
        ''' Print y axis start
        Case Infos.iYstart:
            Select Case pinTypes_:
                Case pinTypes.VolType:
                    TheExec.Datalog.WriteComment "***********************************************"
                    TheExec.Datalog.WriteComment "**************Shmoo Y axis Voltage*************"
                Case pinTypes.frqType:
                    TheExec.Datalog.WriteComment ""
                    TheExec.Datalog.WriteComment "***********************************************"
                    TheExec.Datalog.WriteComment "*** The y-axis " & (Count + 1) & CurrentLoopOrder(Count + 1) & " Time to Change Frequency ***"
            End Select
        ''' print calculation voltage on each pin
        Case Infos.iCalVol:
            Select Case pinTypes_:
                Case pinTypes.VolType:
                    Select Case volTypes_
                        Case volTypes.mainType:
                            TheExec.Datalog.WriteComment vbTab & vbTab & PinName & " = " & Format(val, "0.000") & " V"
                            TheExec.Datalog.WriteComment vbTab & vbTab & "VOLTAGE" & " = " & Format(val, "0.000") & " V"
                        Case volTypes.altType:
                            TheExec.Datalog.WriteComment vbTab & vbTab & PinName & " = " & Format(val, "0.000") & " V"
                            TheExec.Datalog.WriteComment vbTab & vbTab & "VOLTAGE(Alt)" & " = " & Format(val, "0.000") & " V"
                    End Select
                Case pinTypes.frqType:
                    TheExec.Datalog.WriteComment vbTab & PinName & " = " & val & " Hz"
            End Select
        ''' print voltage on specific pin
        Case Infos.iPinVol:
                TheExec.Datalog.WriteComment vbTab & vbTab & PinName & " = " & Format(val, "0.000") & " V"
                TheExec.Datalog.WriteComment vbTab & vbTab & "VOLTAGE" & " = " & Format(val, "0.000") & " V"
                TheExec.Datalog.WriteComment "***********************************************"
        ''' print frequency on specific pin
        Case Infos.iPinFrq:
            TheExec.Datalog.WriteComment "***********Meas Freq Before Init Pat***********"
        ''' print source code on specific pattern
        Case Infos.iSrcStr:
            TheExec.Datalog.WriteComment "DigSrc pattern: " & pat & ",Src Bits = " & CStr(Len(srcPatName_(pat))) & ",Output String [ LSB(L) ==> MSB(R) ]:" & srcPatName_(pat)
            srcPatName_.Remove pat
    End Select
End Function

Public Function setTestNumbers()
    Dim site As Variant
    
    For Each site In TheExec.sites.Active
        testNumbers_(site) = TheExec.sites.item(site).TestNumber
    Next site
    
End Function

Public Function setValueToAxis(Count As Integer)
    Dim outIndex As Integer
    Dim index As Integer
    Dim pins As Collection
    Dim begs As Collection
    Dim CalValue As Double
    Dim step As Double
    Dim coef As Double
    
    ''' check which axis is
    Select Case axis_
        Case axes.aXaxis:
            Set pins = xAxis_
            Set begs = xBegs_
            PrintInfos iXstart, , , CLng(Count)
        Case axes.aYaxis:
            Set pins = yAxis_
            Set begs = yBegs_
            PrintInfos iYstart, , , CLng(Count)
    End Select
    
    For index = 1 To pins.Count
        ''' get step value
        step = AxisStep(index, outIndex)
        ''' get coefficient determined by HV->LV or LV->HV
        coef = AxisCoefficient(outIndex)
        ''' calculate voltage
        CalValue = begs(outIndex) + step * Count * coef
        ''' print calculation voltage by pin
        PrintInfos iCalVol, pins(index), CalValue
        ''' apply sweep voltage to valt
        Select Case pinTypes_
            Case pinTypes.VolType:
                Select Case volTypes_:
                    Case volTypes.mainType:
                        TheHdw.DCVS.pins(pins(index)).Voltage.Main.value = CalValue
                    Case volTypes.altType:
                        TheHdw.DCVS.pins(pins(index)).Voltage.Alt.value = CalValue
                End Select
            Case pinTypes.frqType:
                applyFrqValToPin pins(index), CalValue
        End Select
        
        ''' store for autoswitch
        If IsAutoSwitchCase And Not pinTypes_ = pinTypes.frqType Then
            vrsPinsVol_.pins(pins(index)).value = CalValue
        End If
    Next index
    
    PrintInfos iBlank
End Function

Public Function ApplyTimingLevel()
    TheHdw.Digital.ApplyLevelsTiming True, True, True, tlPowered
End Function
Public Function switchToVmainVol()
    If TheHdw.DCVS.pins(getVrsPinsToString).Voltage.Output <> tlDCVSVoltageMain Then
        TheHdw.DCVS.pins(getVrsPinsToString).Voltage.Output = tlDCVSVoltageMain
        TheHdw.Wait 0.01
    End If
End Function

Public Function switchToValtVol()
    If TheHdw.DCVS.pins(getVrsPinsToString).Voltage.Output <> tlDCVSVoltageAlt Then
        TheHdw.DCVS.pins(getVrsPinsToString).Voltage.Output = tlDCVSVoltageAlt
        If TheExec.enableWord("PrintRSVoltage") Then PrintRSVoltage False, sweepPinsString_
        TheHdw.Wait 0.01
    End If
End Function

Public Function testPattern(pat As String)
    Dim site As Variant
    Dim relays As Integer
    
    ''' test pattern
    relays = TheHdw.DCVS.pins(getVrsPinsToString).Voltage.Output ''' prevent auto switch by pattern
    TheHdw.patterns(pat).test pfAlways, 0
    TheHdw.DCVS.pins(getVrsPinsToString).Voltage.Output = relays
    
    For Each site In TheExec.sites.Selected
        ''' get result
        pfResults_(site) = pfResults_(site) And TheHdw.Digital.Patgen.PatternBurstPassed
        
'        ''' print result on datalog
'        If pfResults_(site) Then
'            theexec.Datalog.WriteFunctionalResult site, testNumbers_(site), logTestPass
'        Else
'            theexec.Datalog.WriteFunctionalResult site, testNumbers_(site), logTestFail
'        End If
        
        ''' increment testnumber
        testNumbers_(site) = testNumbers_(site) + 1
        TheExec.sites.item(site).TestNumber = testNumbers_(site)
    Next site

End Function

Public Function judgePFResult(pfAxis As pfAxes)
    
    Dim site As Variant
    Dim activeSites As New SiteBoolean
    
    For Each site In TheExec.sites.Active
        ''' initial
        'If TheExec.sites.Selected(site) Then activeSites(site) = True
        If XactiveSites_(site) Then activeSites(site) = True
        
        Select Case pfAxis
            ''' x axis pf judge
            Case pfAxes.pXaxis:
                ''' reset inner fail loop count
                If loopTypes_ = lFixType Then loopFailCount_(site) = 0: LactiveSites_(site) = XactiveSites_(site)
                ''' count fail from pattern test
                If Not pfResults_(site) Then xAxisFailCount_(site) = xAxisFailCount_(site) + 1: yAxisFail_(site) = True: pfResults_(site) = True
                ''' judge fail count
                If capXfailCount_ <> 0 Then
                    If xAxisFailCount_(site) = capXfailCount_ Then
                        activeSites(site) = False
                        XactiveSites_(site) = False
                        xAxisFailCount_(site) = 0
                    End If
                End If
            ''' y axis pf judge
            ''' reset site for next x axis search
            Case pfAxes.pYaxis:
                ''' reset x-axis/loop fail count
                xAxisFailCount_(site) = 0
                If loopTypes_ = lVolType Then loopFailCount_(site) = 0: LactiveSites_(site) = YactiveSites_(site)
                
                If capYfailCount_ <> 0 Then
                    ''' count fail from x axis
                    If yAxisFail_(site) Then yAxisFailCount_(site) = yAxisFailCount_(site) + 1: yAxisFail_(site) = False
                    ''' judge fail count
                    If yAxisFailCount_(site) = capYfailCount_ Then
                        activeSites(site) = False
                        YactiveSites_(site) = False
                        yAxisFailCount_(site) = 0
                    End If
                End If
            ''' inner/outer loop pf judge
            ''' since x axis will reset it, directly disable site which match fail count
            Case pfAxes.pILoops, pfAxes.pOLoops:
                If Not pfResults_(site) Then xAxisFail_(site) = True
                If capLoopsCount_ <> 0 Then
                    ''' count fail
                    If (pfAxis = pILoops And loopTypes_ = lFixType) Or (pfAxis = pOLoops And loopTypes_ = lVolType) Then
                        If xAxisFail_(site) Then loopFailCount_(site) = loopFailCount_(site) + 1: xAxisFail_(site) = False
                    End If
                    ''' judge fail count
                    If loopFailCount_(site) = capLoopsCount_ Then
                        'XactiveSites_(site) = False
                        LactiveSites_(site) = False
                        activeSites(site) = False
                        loopFailCount_(site) = 0
                    End If
                End If
        End Select
        ''' disable loop count match site
        activeSites(site) = activeSites(site) And LactiveSites_(site)
    Next site
    
    ''' enalbe/disable active site
    TheExec.sites.Selected = activeSites
End Function

Public Function AnyEnableSites(axis As pfAxes, Optional Count As Long) As Boolean
    Dim site As Variant
    
    For Each site In TheExec.sites.Selected
        Select Case axis
            Case pfAxes.pXaxis:
                If XactiveSites_(site) Then
                    AnyEnableSites = True
                    Exit Function
                End If
            Case pfAxes.pYaxis:
                If YactiveSites_(site) Then
                    AnyEnableSites = True
                    Exit Function
                End If
            Case pfAxes.pILoops, pfAxes.pOLoops:
                If Count = loopTimes_ Then If (axis = pILoops And loopTypes_ = lFixType) Or (axis = pOLoops And loopTypes_ = lVolType) Then LactiveSites_(site) = True: loopFailCount_(site) = 0
                If TheExec.sites.Selected(site) Then
                    AnyEnableSites = True
                    Exit Function
                End If
        End Select
    Next site

    AnyEnableSites = False
End Function

Public Function DisableComparePins(disable As Boolean)
    Dim pin As Variant
    
    For Each pin In disComparePins_
        TheHdw.Digital.pins(pin).DisableCompare = disable
    Next pin
End Function

Public Function ShowVoltageOnPins()
    Dim pin As Variant
    
    For Each pin In showVolPins_
        PrintInfos iPinVol, CStr(pin), TheHdw.DCVS.pins(pin).Voltage.value
    Next pin
End Function

Public Function ShowFrequencyOnPins()
    PrintInfos iPinFrq
    MeasFrequencyOnPins showFrqPins_
    PrintInfos iBlank
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
            switchCase = True
        Else
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
                TheExec.ErrorLogMessage "Source different source-code on same pattern: " & Pattern & ", Please check it"
                SetupDigSrc = False
                Exit Function
            End If
        Else
            srcPatName_.Add Pattern, SrcStr
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
            ''' transfer to long from string
            Dim switchBits As String
            Dim SrcBits() As Long
            Dim index As Integer
            Dim i As Long
            ReDim SrcBits(UBound(SrcBits_tmp))
            
            For i = LBound(SrcBits_tmp) To UBound(SrcBits_tmp)
                ''' source code
                If IsNumeric(SrcBits_tmp(i)) Then
                    SrcBits(i) = CLng(SrcBits_tmp(i))
                '''auto switch
                ElseIf UCase(SrcBits_tmp(i)) = "S" Then
                    If Not isAutoSwitch_.Exists(Pattern) Then
                        isAutoSwitch_.Add Pattern, True
                    End If
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
                    
                    SrcBits_tmp(i) = autoSwitchInfo.getSelectSramBit(UCase(blockName_), vrsPinsVol_, index)
                    switchBits = switchBits & CStr(SrcBits_tmp(i))
                    SrcBits(i) = CLng(SrcBits_tmp(i))
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
                
                ''' replace source string on auto switch pattern
                If isAutoSwitch_.Exists(Pattern) Then
                    SrcStr = Replace(SrcStr, String(Len(switchBits), "s"), switchBits, compare:=vbTextCompare)
                    srcPatName_(Pattern) = SrcStr
                End If
            End If
        End If

        ''' apply to dspwave
        srcWave(srcPatName_.Count - 1).data = SrcBits
        ''' setup into dsp wave
        Call GetPatsFromPatSets(Pattern, patSets, patCnts)
        Call SetupDigSrcDspWave(patSets(0), srcPin_, "BETA_SRC" + CStr(srcPatName_.Count), SrcSize, srcWave(srcPatName_.Count - 1))
    Next Src
    
    SetupDigSrc = True
End Function




