VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "AutoShmooItem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Enum TestDomains
    Soc = 1
    Cpu = 2
    Gfx = 3
    unknown = 0
End Enum
Private Enum TestBlocks
    Scan = 1
    Mbist = 2
    unknown = 0
End Enum

'========== const =========='
Private Const CPU_PIN = "*CPU"
Private Const LOW_PIN = "VDD_LOW"
Private Const CPU_SRAM = "VDD_CPU_SRAM"
Private Const SRAM_PIN = "VDD_SRAM_SOC"

Private Const SHIFT_IN_SOC = "ShiftIn_Freq_VAR"
Private Const SHIFT_IN_CPU = "ShiftIn_Freq_VAR"
Private Const SHIFT_IN_GFX = "TSETSHIFTFREQ"
Private Const NWire = "XI0_Diff_Freq_VAR"

Private Const RET_ENT = "*_ENT"
Private Const RET_EXT = "*_EXT"

'========== fields =========='
Private autoShmooDSSC_ As AutoShmooDSSC
Private autoShmooContext_ As AutoShmooContext

Private autoShmooSetupsOnName_ As Dictionary
Private autoShmooSetupsOnSite_ As Dictionary
Private autoShmooSetupsOnPL_ As Dictionary

Private operationPoint_ As SiteDouble ' current lvcc condtion for shmoo later
Private secondOperationPoint_ As SiteDouble
Private patternSets_ As Collection
Private initPatterns_ As Dictionary ' full init pattern set
Private plPatterns_ As Dictionary ' full paylaod pattern set
Private testInitPatterns_ As Collection ' fail init pattern set
Private testPlPatterns_ As Collection ' fail payload pattern set
Private testBlock_ As TestBlocks
Private testDomain_ As TestDomains
Private plSetups_ As Collection ' store fail payload setup

Private powerPin_ As String ' power pin for x-axis (power pin only)
Private secondPin_ As String ' second pin for y-axis (could be AC pin)

Private isShmooNeeded_ As Boolean

Private initPatternIndex_ As Integer ' to store the init pattern set count from argument prepattern

'========== constructor =========='
Sub Class_Initialize()
    Set autoShmooDSSC_ = New AutoShmooDSSC
    Set autoShmooContext_ = New AutoShmooContext ' for context
    Set autoShmooSetupsOnSite_ = New Dictionary ' shmoo setups based on site variable
    Set autoShmooSetupsOnName_ = New Dictionary ' shmoo setups based on char setup name
    Set autoShmooSetupsOnPL_ = New Dictionary
    
    Set initPatterns_ = New Dictionary
    Set plPatterns_ = New Dictionary
    Set patternSets_ = New Collection
    Set plSetups_ = New Collection
    'Set testInitPatterns_ = New Collection
    'Set testPlPatterns_ = New Collection
End Sub

'========== properties =========='
Public Property Get OperationPoint() As SiteDouble
    Set OperationPoint = operationPoint_
End Property
Public Property Get SecondOperationPoint() As SiteDouble
    Set SecondOperationPoint = secondOperationPoint_
End Property
Public Property Get powerPin() As String
    powerPin = powerPin_
End Property
Public Property Get SecondPin() As String
    SecondPin = secondPin_
End Property
Public Property Get InitPatterns() As Dictionary
    Set InitPatterns = initPatterns_
End Property
Public Property Get PlPatterns() As Dictionary
    Set PlPatterns = plPatterns_
End Property
Public Property Get TestinitPatterns() As Collection
    Set TestinitPatterns = testInitPatterns_
End Property
Public Property Get TestPlPatterns() As Collection
    Set TestPlPatterns = testPlPatterns_
End Property
Public Property Get IsShmooNeeded() As Boolean
    IsShmooNeeded = isShmooNeeded_
End Property
Public Property Get CharShmooSetupsOnSite() As Dictionary
    Set CharShmooSetupsOnSite = autoShmooSetupsOnSite_
End Property
Public Property Get CharShmooSetupsOnName() As Dictionary
    Set CharShmooSetupsOnName = autoShmooSetupsOnName_
End Property
Public Property Get IsMbistBlock() As Boolean
    IsMbistBlock = (testBlock_ = TestBlocks.Mbist)
End Property
Public Property Let powerPin(pin As String)
    powerPin_ = UCase(pin)
End Property

'========== methods =========='
Private Property Get IsSTCKPattern(Pattern As String) As Boolean
    Dim Server As String
    Server = autoShmooUtility.GetServerName(Pattern, 8)
    IsSTCKPattern = (mid(Server, 5) = "STCK")
End Property
Private Function GetTestDomain(Domain As String) As TestDomains
    Select Case Domain
        Case "S":
            GetTestDomain = TestDomains.Soc
        Case "C":
            GetTestDomain = TestDomains.Cpu
        Case "L":
            GetTestDomain = TestDomains.Gfx
        Case Else
            ' something wrong here
            GetTestDomain = TestDomains.unknown
    End Select
End Function
Private Function GetTestBlock(Block As String) As TestBlocks
    Select Case Block
        Case "CH", "SC":
            GetTestBlock = TestBlocks.Scan
        Case "BI":
            GetTestBlock = TestBlocks.Mbist
        Case Else:
            ' something wrong here
            GetTestBlock = TestBlocks.unknown
    End Select
End Function
Private Function GetTestRail(rail As String, Optional isSingleRail As Boolean = True) As String
    Dim index As Integer
    Dim rails() As String
    Dim powerPin As String

    For index = 0 To Len(rail)
        powerPin = ""
        Select Case mid(rail, index + 1, 1)
            Case "S", "R":
                powerPin = "VDD_SOC"
            Case "C":
                powerPin = "VDD_CPU"
            Case "E":
                powerPin = "VDD_ECPU"
            Case "P":
                powerPin = "VDD_PCPU"
            Case "G":
                powerPin = "VDD_GPU"
            Case "L":
                powerPin = "VDD_LOW"
            Case "I":
                powerPin = "VDD_DISP"
            Case "A":
                powerPin = "VDD_AVE"
            Case "D"
                powerPin = "VDD_DCS_DDR"
        End Select
        
        If powerPin <> "" Then
            autoShmooUtility.appendToArray rails, powerPin
            If isSingleRail Then GoTo breakLoop
        End If
    Next index

breakLoop:
    GetTestRail = Join(rails, ",")
End Function
Private Function GetSecondRailFromPowerRail() As String
    Dim pin As Variant
    Dim rail As String
    Dim suffix As String
    
    autoShmooContext_.class_initialzie ' need to think how to make it better
    If UCase(autoShmooContext_.TimeSetSheet) Like "*_ND" Then suffix = "_1"
    
    Select Case testBlock_
        Case TestBlocks.Scan:
            ''' default set to shift in rail
            Select Case testDomain_
                Case TestDomains.Soc:
                    GetSecondRailFromPowerRail = SHIFT_IN_SOC + suffix
                Case TestDomains.Cpu:
                    GetSecondRailFromPowerRail = SHIFT_IN_CPU + suffix
                Case TestDomains.Gfx:
                    GetSecondRailFromPowerRail = SHIFT_IN_GFX + suffix
            End Select
        Case TestBlocks.Mbist
            ''' try to get sram power for second rail
            For Each pin In Split(powerPin_, ",")
                GetSecondRailFromPowerRail = GetSramPinFromLogic(CStr(pin))
                If GetSecondRailFromPowerRail <> "" Then Exit Function
            Next pin
        Case Else:
            ' something wrong here
    End Select
End Function
Public Function AddPattern(ParamArray patternSets() As Variant)

    Dim patternSet As Variant
    Dim patterns() As String
    Dim Pattern As Variant
    
    For Each patternSet In patternSets
        If patternSet = "" Then GoTo nextPatSet
        initPatternIndex_ = IIf(initPatterns_.Count = 0, -1, initPatterns_.Count) ' store init pattern count from prepattern
        patterns = autoShmooUtility.GetPatsFromPatSets(patternSet, True)
        For Each Pattern In patterns
            ''' store pattern set separated by init and payload
            If autoShmooUtility.IsInitPattern(Pattern) Then
                If Not initPatterns_.Exists(Pattern) Then initPatterns_.Add Pattern, New SiteBoolean
                autoShmooDSSC_.IsDSSCPattern CStr(Pattern)
            Else
                If Not plPatterns_.Exists(Pattern) Then plPatterns_.Add Pattern, New SiteBoolean
            End If
            
            ''' store the full pattern set
            patternSets_.Add Pattern
        Next Pattern
nextPatSet:
    Next patternSet

End Function
Public Function SetPatternPassFail(Pattern As String, passFail As SiteBoolean)
    If initPatterns_.Exists(Pattern) Then
        Set initPatterns_.item(Pattern) = passFail
    ElseIf plPatterns_.Exists(Pattern) Then
        Set plPatterns_.item(Pattern) = passFail
        
        If passFail.Any(False) Then isShmooNeeded_ = True
    Else
        ' something wrong here
    End If
End Function
Public Function ApplyTimingLevel()
    autoShmooContext_.ApplyTimingLevel
End Function
Private Function SetCurrentCondition(pins As String) As SiteDouble
    ''' get the fail item condition on its corepower '''
    Dim pin As Variant
    
    For Each pin In Split(pins, ",")
        If UCase(pin) Like "VDD*" Then
            If Not gl_GetInstrumentType_Dic.Exists(LCase(pin)) Then Exit Function ' may need to handle the case later
             
            Select Case gl_GetInstrumentType_Dic(LCase(pin))
                Case "DCVS", "DCVSMerged2", "DCVSMerged4", "DCVSMerged8":
                    Set SetCurrentCondition = TheHdw.DCVS.pins(pin).Voltage.ValuePerSite
                Case "DCVI":
                    Set SetCurrentCondition = TheHdw.DCVI.pins(pin).Voltage.ValuePerSite
                Case Else:
                    ' something wrong here
                    Exit Function
            End Select
        Else
            Set SetCurrentCondition = TheExec.Specs.AC(pin).CurrentValue
        End If
        Exit Function
    Next pin
End Function
Private Function GetSramPinFromLogic(pin As String) As String
    ''' get sram pin based on logic for mbist item '''
    Dim index As Integer
    
    For index = LBound(SelsramMapping(autoShmooDSSC_.SelSramMappingIndex).bitCount) To UBound(SelsramMapping(autoShmooDSSC_.SelSramMappingIndex).bitCount)
        If pin = SelsramMapping(autoShmooDSSC_.SelSramMappingIndex).logic_Pin(index) Then
            GetSramPinFromLogic = SelsramMapping(autoShmooDSSC_.SelSramMappingIndex).SRAM_PIN(index)
            Exit For
        End If
    Next index
    
    ''' check if it's non selsram pin
    If GetSramPinFromLogic = "" Then
        If pin Like CPU_PIN Then
            GetSramPinFromLogic = CPU_SRAM
        ElseIf pin = LOW_PIN Then
            ' no corresponding sram pin for VDD_LOW
            ' default set to VDD_SRAM_SOC
            GetSramPinFromLogic = SRAM_PIN
        Else
            ' something wrong here
        End If
    End If
End Function
Private Function SetupTestPatterns(Pattern As String) As String
    Dim index As Integer
    Dim pivot As Integer
    Dim seq As Integer
    Dim pat As Variant
    Dim VERSION As Integer
    Dim IsInit As Boolean
    Dim patternSet() As String
    Dim targetCluster As String
    Dim currentCluster As String
    Dim targetServer As String
    Dim currentServer As String
    Dim patternHash As Collection
    Dim isSTCK As Boolean
    
    'TODO need to think how to handle mbist sequence
    
    ''' initialize
    Set testInitPatterns_ = Nothing: Set testInitPatterns_ = New Collection
    Set testPlPatterns_ = Nothing: Set testPlPatterns_ = New Collection
    Set patternHash = New Collection
    
    index = autoShmooUtility.GetIndex(patternSets_, Pattern)
    targetCluster = autoShmooUtility.GetClusterName(Pattern)
    targetServer = autoShmooUtility.GetServerName(Pattern)
    
    For pivot = index To 1 Step -1
        If testBlock_ = TestBlocks.Scan Then
            ''' scan block
            VERSION = autoShmooUtility.GetInitSequence(patternSets_.item(pivot)) ' get init sequence number
            
            If patternHash.Count = 0 Then patternHash.Add patternSets_.item(pivot) ' need to include failing payload for TD
            If Not initPatterns_.Exists(patternSets_.item(pivot)) Then GoTo nextPivot  ' only check init pattern for TD
            If testDomain_ = TestDomains.Gfx Then
                ''' gpu domain
                If (Not currentCluster = targetCluster And Not currentCluster = "CXXX") Or _
                    autoShmooUtility.HasKey(patternHash, patternSets_.item(pivot)) Then GoTo nextPivot
                patternHash.Add patternSets_.item(pivot), patternSets_.item(pivot)
            Else
                ''' soc/cpu domain
                If autoShmooUtility.HasKey(patternHash, CStr(VERSION)) Then GoTo nextPivot
                patternHash.Add patternSets_.item(pivot), CStr(VERSION)
            End If
        Else
            ''' mbist block
            currentCluster = autoShmooUtility.GetClusterName(patternSets_.item(pivot)) ' get cluster name
            currentServer = autoShmooUtility.GetServerName(patternSets_.item(pivot)) ' get server name
            
            ''' we shouldn't store same pattern twice for both init and payload
            If autoShmooUtility.HasKey(patternHash, patternSets_.item(pivot)) Then Exit For
            ''' we shoudln't have init pattern b/t two payload pattern except retention wait time init
            ''' PP_PL_1
            ''' PP_INIT_0
            ''' PP_PL_2
            If plPatterns_.Exists(patternSets_.item(pivot)) And IsInit Then Exit For
            ''' need to handle init pattern b/t different cluster
            If initPatterns_.Exists(patternSets_.item(pivot)) Then
                ''' special handle for soc mbist here
                If patternSets_.item(pivot) Like RET_EXT Then GoTo nextPivot ' always skip the retention exit pattern
                patternHash.Add patternSets_.item(pivot), patternSets_.item(pivot): IsInit = True
                GoTo nextPivot
            End If

            If testDomain_ = TestDomains.Soc Then
                ''' skip if the cluster or server is different to the failing payload
                ''' but we need to include the wait time payload pattern for once
                If (Not currentCluster = targetCluster Or Not currentServer = targetServer) And Not currentCluster = "XXXX" Then GoTo nextPivot
                ''' only add wait time pattern for STCK bist
                ''' PP_BIST_SRVXSETP
                ''' PP_WFC_WAITTIME
                ''' PP_BIST_SRVXSTCK
                If currentCluster = "XXXX" And Not isSTCK Then GoTo nextPivot
                If IsSTCKPattern(patternSets_.item(pivot)) Then isSTCK = True
            Else
                If plPatterns_.Exists(patternSets_.item(pivot)) And Not patternHash.Count = 0 Then GoTo nextPivot
            End If
            patternHash.Add patternSets_.item(pivot), patternSets_.item(pivot)
        End If
nextPivot:
    Next pivot
    
    ''' scan:  check if there is any not included in test pattern set
    ''' mbist: force to add all the init pattern unless the pattern already exist
    If testBlock_ = TestBlocks.Scan Then
        For Each pat In initPatterns_.Keys
            VERSION = autoShmooUtility.GetInitSequence(CStr(pat))
            If autoShmooUtility.HasKey(patternHash, CStr(VERSION)) Or autoShmooUtility.HasKey(patternHash, CStr(pat)) Then Exit For
            testInitPatterns_.Add pat
            autoShmooUtility.appendToArray patternSet, pat
        Next pat
    Else
        For pivot = 0 To initPatternIndex_ - 1
            If autoShmooUtility.HasKey(patternHash, CStr(initPatterns_.Keys(pivot))) Then Exit For
            testInitPatterns_.Add initPatterns_.Keys(pivot)
            autoShmooUtility.appendToArray patternSet, initPatterns_.Keys(pivot)
        Next pivot
    End If
    
    ''' reverse the remaining pattern into test pattern set
    For pivot = patternHash.Count To 1 Step -1
        If initPatterns_.Exists(patternHash.item(pivot)) Then
            testInitPatterns_.Add patternHash.item(pivot)
        Else
            testPlPatterns_.Add patternHash.item(pivot)
        End If
        autoShmooUtility.appendToArray patternSet, patternHash.item(pivot)
    Next pivot
    
    ''' for printout shmoo test sequence
    SetupTestPatterns = Join(patternSet, ",")
End Function
Private Function SetupSpecFromPattern(Pattern As String)
    Dim info() As String
    Dim rails As String
    
    info = Split(Pattern, "_") ' split the pattern name
    testDomain_ = GetTestDomain(info(2)) ' get test domain (ex. S/C/G)
    testBlock_ = GetTestBlock(info(4)) ' get test block (ex. CH/TD/BI)
    powerPin_ = IIf(powerPin_ = "", GetTestRail(info(9), True), powerPin_) ' get power pin from pattern if not exist (ex. MSAIXX -> VDD_SOC/VDD_AVE/VDD_DISP)
    secondPin_ = GetSecondRailFromPowerRail ' get second rail (ex. Mbist->SRAM/Scan->ShiftIn)
    'TODO how to handle multiple pin
    Set operationPoint_ = SetCurrentCondition(powerPin_) ' set current condition from powerPin
    Set secondOperationPoint_ = SetCurrentCondition(secondPin_) ' set current condition from secondPin
End Function
Public Function SetupItem(Optional powerPin As String = "") As Boolean
    ''' setup shmoo item based on failing pattern '''
    
    Dim Pattern As Variant
    Dim site As Variant
    Dim shmooSetup As AutoShmooSetup
    
    For Each Pattern In plPatterns_.Keys
        If plPatterns_.item(Pattern).All(True) Then GoTo nextPat ' no any site fail continue
        For Each site In TheExec.sites.Selected
            If plPatterns_.item(Pattern)(site) Or autoShmooSetupsOnSite_.Exists(site) Then GoTo nextSite  ' site not fail or already exist shmoo setup
            If autoShmooUtility.HasKey(plSetups_, CStr(Pattern)) Then
                autoShmooSetupsOnSite_.Add site, autoShmooSetupsOnPL_.item(CStr(Pattern)) ' bypass same failed pattern in same item
                GoTo nextSite
            End If
            plSetups_.Add Pattern, CStr(Pattern)
            Set shmooSetup = New AutoShmooSetup
            SetupSpecFromPattern CStr(Pattern) ' set power pin and condition from fail item
            shmooSetup.ShmooTestPattern = SetupTestPatterns(CStr(Pattern)) ' set pattern sequence for failing item
            shmooSetup.SetupShmoo Me, CStr(site) ' setup shmoo based on the fail item
            If Not shmooSetup.IsValid Then GoTo nextSite
            autoShmooSetupsOnSite_.Add site, shmooSetup ' store shmoo setup based on its failed site
            autoShmooSetupsOnName_.Add shmooSetup.CharSetupName, shmooSetup ' store shmoo setup its char setup name
            autoShmooSetupsOnPL_.Add CStr(Pattern), shmooSetup
nextSite:
        Next site
nextPat:
    Next Pattern
    
    Set shmooSetup = Nothing
    SetupItem = autoShmooSetupsOnSite_.Count
End Function
Public Function SetShmooCondition(CharSetupName As String)

    Dim charSetup As AutoShmooSetup
    Dim pin As Variant
    
    Set charSetup = autoShmooSetupsOnName_.item(CharSetupName)

    '''TODO how to prevent global value is smaller than zero

    ''' setup X-axis pin
    For Each pin In Split(powerPin_, ",")
        Select Case gl_GetInstrumentType_Dic(LCase(pin))
            Case "DCVS", "DCVSMerged2", "DCVSMerged4", "DCVSMerged8":
                TheHdw.DCVS.pins(pin).Voltage.Alt.ValuePerSite = autoShmooUtility.GetGlobalValue(charSetup.GetXaxisGlobalName)
            Case "DCVI":
                TheHdw.DCVI.pins(pin).Voltage.ValuePerSite = autoShmooUtility.GetGlobalValue(charSetup.GetXaxisGlobalName)
            Case Else:
                ' something wrong here
                Exit Function
        End Select
    Next pin
    
    ''' we only need to setup voltage whenever it is 2-D and mbist block
    If Not charSetup.IsOneDimensionOnly And IsMbistBlock Then
        ''' setup Y-axis pin
        For Each pin In Split(secondPin_, ",")
            Select Case gl_GetInstrumentType_Dic(LCase(pin))
                Case "DCVS", "DCVSMerged2", "DCVSMerged4", "DCVSMerged8":
                    TheHdw.DCVS.pins(pin).Voltage.Alt.ValuePerSite = autoShmooUtility.GetGlobalValue(charSetup.GetYaxisGlobalName)
                Case "DCVI":
                    TheHdw.DCVI.pins(pin).Voltage.ValuePerSite = autoShmooUtility.GetGlobalValue(charSetup.GetYaxisGlobalName)
                Case Else:
                    ' something wrong here
                    Exit Function
            End Select
        Next pin
    End If
    
    ''' setup dssc based on Valt value
    autoShmooDSSC_.SetupDSSC
End Function









