Attribute VB_Name = "Lib_SSN"
Option Explicit
Public Const glb_SSN_Enabled = True
Public ssnPatternsDict As New Dictionary        ''Save SSN patterns into this dict during first run
Public Const glb_SSN_Failflag = "F_SSN_INIT_Fail"
Public Const glb_SSN_CaptureLimit = 307200

Private CorePFResults As PinListData
Private CorePF As SiteBoolean
Private CoreScanPins As String
Public DebugMsgEnable As Boolean
Public glb_maxReburst As SiteBoolean
Public nonSSNFailOnSSN As SiteBoolean
Public ssnResultsGlobal As IScanNetworkResults

'Public dic_SSNMapping_UF As New Dictionary
'Public F_Parsing_SSNMapping_UF_Done As Boolean
Public reburst As Boolean

'SSN info for current instance
Public Type Inst_SSN
    bSSNTest As Boolean
    bSSNCoreHarvest As Boolean
    bSSNCoreMask As Boolean
    patkeyword() As String ' CJR: 4/16/24 Defined as array
    PatName() As String
    coreName() As String
    Failflag() As String
    mappingFileName() As String
    maskedCores() As New SiteVariant
    idxMatchedPattern As Long
End Type

''20231120: SSN info from harvest table
Public Type ssnMappingType
    patternKeyword As String
    coreName As String
    Failflag As String
End Type
Public SSNMapping() As ssnMappingType

Function SetupSSNCentralCMEM() As Long
On Error GoTo errHandler
Dim funcName As String:: funcName = "SetupSSNCentralCMEM"
    Dim CaptSize As Long
    ' Set up Central CMEM to record Scan Indexs
    TheHdw.Digital.CMEM.CentralFields = tlCMEMScanIndex + tlCMEMAbsoluteCycle + tlCMEMVMVectorOffset + tlCMEMPatternName
    ' Set Capture size to 16M
    CaptSize = -1 '(2 ^ 18) ' (2 ^ 24)
    ' Using CmemCaptAll allows counting cycles to associate cycles and failures
    Call TheHdw.Digital.CMEM.SetCaptureConfig(CaptSize, CmemCaptFail)
Exit Function
errHandler:
    TheExec.Datalog.WriteComment "<Error> " + funcName + ":: please check it out."
    If AbortTest Then Exit Function Else Resume Next
End Function

'Check if SSN is enabled from harvest fail flag and pattern keyword for this instance
Public Function CheckInstForSSN(PatArray() As String, inData As Inst_SSN, SSNMapping() As ssnMappingType, _
                                Optional strHarvFlag As String = vbNullString, Optional isCharInst As Boolean = False)
On Error GoTo errHandler
Dim funcName As String:: funcName = "CheckInstForSSN"
    Dim i As Long
    Dim pat As Variant
    Dim isSSNHarvestInst As Boolean, isSSNCoreMaskInst As Boolean
    Dim countMatchedSSN As Long
    
    ''Initialize
    ReDim inData.PatName(0)
    ReDim inData.mappingFileName(0)
    ReDim inData.maskedCores(0)
    ReDim inData.PatName(0)
    ReDim inData.patkeyword(0) 'CJR: 4/16/24
    ReDim inData.coreName(0) 'CJR: 4/20/24
    ReDim inData.Failflag(0) 'CJR: 4/20/24
    
    inData.bSSNTest = False
    inData.bSSNCoreHarvest = False
    inData.bSSNCoreMask = False
    
    inData.idxMatchedPattern = 0
    
    pat = vbNullString
    isSSNHarvestInst = False
    isSSNCoreMaskInst = False
    countMatchedSSN = 0
    
    'To avoid pin group harvest
    If TheExec.flow.IsCharacterizing = True Then
        inData.bSSNTest = True
        inData.bSSNCoreHarvest = False
        inData.bSSNCoreMask = False
        Exit Function
    End If
    
    ''Added for SSN char
    If isCharInst = True Then
        strHarvFlag = "enablecoreharvest:true;enablecoremask:false"
    End If
    
    If strHarvFlag = "" Then Exit Function
    
    ''Find SSN pattern by keyword from table
    ''Added to save multiple SSN pattern
    For Each pat In PatArray
        
        If ssnPatternsDict.Exists(UCase(pat)) = True Then
            For i = 0 To UBound(SSNMapping)
                If UCase(pat) Like UCase(SSNMapping(i).patternKeyword) Then
                    inData.bSSNTest = True
                    
                    ReDim Preserve inData.PatName(countMatchedSSN)
                    ReDim Preserve inData.patkeyword(countMatchedSSN) 'CJR: 4/16/24
                    ReDim Preserve inData.coreName(countMatchedSSN) 'CJR: 4/20/24
                    ReDim Preserve inData.Failflag(countMatchedSSN) 'CJR: 4/20/24
                    
                    inData.coreName(countMatchedSSN) = SSNMapping(i).coreName  ' CJR: 4/16/24
                    inData.Failflag(countMatchedSSN) = SSNMapping(i).Failflag  ' CJR: 4/16/24
                    inData.PatName(countMatchedSSN) = pat
                    inData.patkeyword(countMatchedSSN) = SSNMapping(i).patternKeyword 'CJR: 4/16/24
                    countMatchedSSN = countMatchedSSN + 1
                    
                    'Exit For
                End If
            Next i
        End If
        ''Found matched pattern keyword, break
        'If inData.bSSNTest = True Then Exit For
    Next pat
    
    ''No matched pattern, exit
    If inData.bSSNTest = False Then Exit Function
    
    ''Check SSN feature on/off by input argument
    Dim splitHarvestSetup As Variant, harvestSetup As Variant
    splitHarvestSetup = Split(Replace(LCase(strHarvFlag), " ", ""), ";")
    
    For Each harvestSetup In splitHarvestSetup
        If harvestSetup = "enablecoreharvest:true" Then
            isSSNHarvestInst = True
        ElseIf harvestSetup = "enablecoreharvest:false" Then
            isSSNHarvestInst = False
        ElseIf harvestSetup = "enablecoremask:true" Then
            ''Core mask
            isSSNCoreMaskInst = True
        ElseIf harvestSetup = "enablecoremask:false" Then
            ''Disable core mask
            isSSNCoreMaskInst = False
        Else
            'Call Print_Error_Message(Error_Info, "", "", "Can't find correct harvest argument setup from instance:" & glb_TestInstance & ", harvest flag:" & strHarvFlag)
        End If
    Next harvestSetup

    ''Update core harvest / core mask into instSSNinfo
    inData.bSSNCoreHarvest = isSSNHarvestInst
    inData.bSSNCoreMask = isSSNCoreMaskInst
    
    ''Find and save SSN mapping name
    ReDim inData.mappingFileName(UBound(inData.PatName))
    For i = 0 To UBound(inData.PatName)
        Dim patternCount As Long, patternPath As Variant
        patternPath = TheExec.DataManager.Raw.GetPatternsInSet(inData.PatName(i), patternCount)
        inData.mappingFileName(i) = Split(UCase(patternPath(0)), ".PAT")(0)
    Next i
    
    ''Apply core mask using matched pattern, core name, fail flag
    If inData.bSSNCoreMask = True Then
        For i = 0 To UBound(inData.mappingFileName)
            Call MaskCoreForSSN(CStr(inData.coreName(i)), CStr(inData.Failflag(i)), CStr(inData.mappingFileName(i)))
        Next i
    End If

    '===20240516
    ''Save masked core then we can restore it if SSN reburst happen
    If inData.bSSNCoreHarvest = True Then
        'Stop
        ReDim inData.maskedCores(UBound(inData.PatName))
        For i = 0 To UBound(inData.mappingFileName)
            inData.maskedCores(i) = TheHdw.Digital.ScanNetworks(inData.mappingFileName(i)).maskedCores
        Next i
    End If
    
Exit Function
errHandler:
    TheExec.Datalog.WriteComment "Error encountered in VBT Function of CheckInstForSSN"
    If AbortTest Then Exit Function Else Resume Next
End Function

' Converts whatever is passed as InValues to a Byte Array
' Useful for CMEM cases, as bitwise data is often returned as Longs, which are signed, or Doubles, which cannot have bitwise operators applied to them (throws overflow exception)
' InValues could be a Long, Double, or Array of Doubles
Function ValuesToByteArray(ByVal InValues As Variant) As Variant
On Error GoTo errHandler
Dim funcName As String:: funcName = "ValuesToByteArray"
    Dim InValuesType As Integer
    Dim Inc As Integer
    Dim Denominator As Double
    InValuesType = VarType(InValues)
    Dim ReturnArray() As Byte
    If InValuesType = 3 Then ' If InValuesType = 3, then InValues is a Long
        ReDim ReturnArray(3) ' 4 Bytes for 1 32 bit Long
        Dim bitInc As Integer
        For bitInc = 0 To 31
            Dim CompVal As Double
            ' CompVal is a 32 bit Long with a single left shifting bit each cycle
            ' Since Longs are signed, to flip the 32nd bit CompVal is set to -1
            If bitInc < 31 Then CompVal = 2 ^ bitInc Else CompVal = -1
            If (CompVal And InValues) = CompVal Then
                 ReturnArray(bitInc \ 8) = ReturnArray(bitInc \ 8) Or 2 ^ (bitInc Mod 8)
            End If
        Next bitInc
        
    ElseIf InValuesType = 5 Then ' If InValuesType = 5, then InValues is a Double
        ReDim ReturnArray(7) ' 8 Bytes for 1 64 bit Double
        For Inc = 0 To 7
            If Inc = 0 Then Denominator = 1 Else Denominator = 2 ^ (8 * Inc)
            ' Isolates each byte from InValues by:
                ' Starting with InValues, shifting right in 8 bit chunks, dividing by 2 ^ 8*(byte #), then using Int rounding (this is not casting to an integer)
                    ' This cuts off the least siginificant byte(s)
                ' subtract InValues, shifted right by 1 more byte than trying to isolate, dividing by 2 ^ 8*(byte # + 1), using Int rounding, then left shifting by 1 byte
            ReturnArray(Inc) = CByte(Int(InValues / Denominator) - (Int(InValues / (2 ^ (8 * (Inc + 1)))) * 2 ^ 8))
        Next Inc
        
    ElseIf (InValuesType And 8192) = 8192 Then ' If InValuesType is not 3 or 5, it is a Double()
        ReDim ReturnArray(8 * UBound(InValues) - 1) ' 8 Bytes for each 64 bit Long
        Dim ValInc As Long
        For ValInc = 0 To UBound(InValues)
            For Inc = 0 To 7
                If Inc = 0 Then Denominator = 1 Else Denominator = 2 ^ (8 * Inc)
                ' Isolates each byte from InValues by:
                    ' Starting with InValues, shifting right in 8 bit chunks, dividing by 2 ^ 8*(byte #), then using Int rounding (this is not casting to an integer)
                        ' This cuts off the least siginificant byte(s)
                    ' subtract InValues, shifted right by 1 more byte than trying to isolate, dividing by 2 ^ 8*(byte # + 1), using Int rounding, then left shifting by 1 byte
                ReturnArray(ValInc * 8 + Inc) = CByte(Int(InValues(ValInc) / Denominator) - (Int(InValues(ValInc) / (2 ^ (8 * (Inc + 1)))) * 2 ^ 8))
            Next Inc
        Next ValInc
    End If
    ValuesToByteArray = ReturnArray
Exit Function
errHandler:
    TheExec.Datalog.WriteComment "<Error> " + funcName + ":: please check it out."
    If AbortTest Then Exit Function Else Resume Next
End Function

Public Function NumExtract(inData As String) As String
On Error GoTo errHandler
Dim funcName As String:: funcName = "NumExtract"
    Dim i As Long
    Dim num As String
    For i = 1 To Len(inData)
        If IsNumeric(mid(inData, i, 1)) = True Then
            num = num & mid(inData, i, 1)
        End If
    Next i
    NumExtract = num
Exit Function
errHandler:
    TheExec.Datalog.WriteComment "<Error> " + funcName + ":: please check it out."
    If AbortTest Then Exit Function Else Resume Next
End Function

' Reads all lines of a file and returns the lines as a string array
Public Function FileToStringArray(FilePath As String) As Variant
On Error GoTo errHandler
Dim funcName As String:: funcName = "FileToStringArray"
    
    Dim fso As FileSystemObject: Set fso = New FileSystemObject
    Dim txtStream As TextStream
    Dim arr() As String
    Set txtStream = fso.OpenTextFile(FilePath, ForReading, False)
    
    Dim Inc As Long
    Inc = 0
    
    ' Loop until end of stream
    Do While Not txtStream.AtEndOfStream
        Dim line As String
        line = txtStream.ReadLine
        line = Split(line, "//")(0)
        ' Ignore empty lines, and split at "//" to ensure C-Style EOL comments are ignored
        If Not StrComp(line, "") = 0 Then
            ReDim Preserve arr(Inc) As String
            arr(Inc) = line
        End If
        Inc = Inc + 1
    Loop
    txtStream.Close
    
    FileToStringArray = arr
Exit Function
errHandler:
    TheExec.Datalog.WriteComment "<Error> " + funcName + ":: please check it out."
    If AbortTest Then Exit Function Else Resume Next

End Function

Public Function ReadMappingFile(mappingfile As String) As Class_TCSSN
On Error GoTo errHandler
Dim funcName As String:: funcName = "ReadMappingFile"
    Dim i As Long
    Dim MappingFileText() As String
    Dim SplitData As Variant, splittmp As Variant
    
    Dim SSNMapping As New Class_TCSSN
    'Dim TotalCycle As Long
    Dim coreData As String
    'Dim PinCnt As Long
    'Dim Packet() As Variant
    Dim index As Long
    'Dim dic_index2core As New Dictionary, dic_core2Index As New Dictionary
    
    If Dir(mappingfile) = "" Then
        TheExec.Datalog.WriteComment mappingfile & ", it doesn't exist. Error!!!"
    Else
        'initial
        MappingFileText = FileToStringArray(mappingfile)
        index = -1
        SSNMapping.TotalCoreCnt = 0
        'parsing data
        For i = 0 To UBound(MappingFileText)
            SplitData = Split(MappingFileText(i), ",")
            If i = 0 Then
                SSNMapping.TotalCycle = SplitData(1)
            End If
            
            If Not SSNMapping.PinData Like "*" & SplitData(0) & "*" Then
                SSNMapping.PinData = SSNMapping.PinData & SplitData(0) & ","
            End If
            
            splittmp = Split(SplitData(3), "/")
            coreData = NumExtract(CStr(splittmp(UBound(splittmp))))
            
            If Not SSNMapping.dic_core2Index.Exists(coreData) Then
                index = index + 1
                SSNMapping.TotalCoreCnt = SSNMapping.TotalCoreCnt + 1
            End If
            
            If SSNMapping.dic_index2core.Exists(index) Then
                SSNMapping.Packet(SSNMapping.dic_core2Index(coreData)) = SSNMapping.Packet(SSNMapping.dic_core2Index(coreData)) + 1
            Else
                SSNMapping.dic_index2core.Add index, coreData
                SSNMapping.dic_core2Index.Add coreData, index
                SSNMapping.ResizePacketArr (index)
                SSNMapping.Packet(index) = SSNMapping.Packet(index) + 1
            End If
        Next i
        
        SSNMapping.PinData = left(SSNMapping.PinData, Len(SSNMapping.PinData) - 1)
        SSNMapping.pinCnt = UBound(Split(SSNMapping.PinData, ",")) + 1
        SSNMapping.GenPacketMap (SSNMapping.pinCnt)
        SSNMapping.GetBusRepeatMap (SSNMapping.pinCnt)
        Set ReadMappingFile = SSNMapping
    End If
Exit Function
errHandler:
    TheExec.Datalog.WriteComment "<Error> " + funcName + ":: please check it out."
    If AbortTest Then Exit Function Else Resume Next

End Function




'===============================================================================
''BEGIN UFP SSN
'===============================================================================
''Attribute VB_Name = "VB_Scan_Utils"
''Option Explicit

'===============================================================================
' Setup the datalog for logging scan fail diagnostic data to STDF
'===============================================================================
Public Sub SetupForScanLogging(logPins As String, _
                               scanFailsPerPin As Long, _
                               logFormat As tlDatalogScanCaptureFormat, _
                               logResultMode As tlDatalogScanResultMode)
On Error GoTo errHandler
Dim funcName As String:: funcName = "SetupForScanLogging"

    TheExec.Datalog.Setup.ScanSetup.EnableScanLogging = True
    TheExec.Datalog.Setup.ScanSetup.CaptureFormat = logFormat
    TheExec.Datalog.Setup.ScanSetup.ResultMode = logResultMode
    'Need GPIO pins from ssn mapping table
    TheExec.Datalog.Setup.ScanSetup.PinList = logPins
    TheExec.Datalog.Setup.ScanSetup.CMEMCaptureLimit = scanFailsPerPin

    TheExec.Datalog.ApplySetup
Exit Sub
errHandler:
    TheExec.Datalog.WriteComment "<Error> " + funcName + ":: please check it out."
    If AbortTest Then Exit Sub Else Resume Next
End Sub


'==============================================================================
' Insert random fails into a scan-based SSN pattern
'==============================================================================
Public Function InsertRandomScanFails(targetPattern As Pattern, _
                                      scanLength As Integer, _
                                      failsPerSite As Long, _
                                      scanInPins As String) As Long
On Error GoTo errHandler
Dim funcName As String:: funcName = "InsertRandomScanFails"
    '___ Init _________________________________________________________________
    Dim maxLabelNum As Integer
    maxLabelNum = 119
    
    '___ Ensure pattern is fresh to start with _________________________________________
    '    Ignore error incase pattern has not already been loaded
    '-----------------------------------------------------------------------------------
    On Error Resume Next
    TheHdw.patterns(targetPattern).Reload
    
    '___ Get the individual target pins to change _____________________________________
    Dim targetInPins() As String
    Dim targetPinCnt As Long
    Call TheExec.DataManager.DecomposePinList(scanInPins, targetInPins, targetPinCnt)
    
    '___ Make N random fails per site __________________________________________________
    Dim origData As Integer
    
Exit Function
errHandler:
    TheExec.Datalog.WriteComment "<Error> " + funcName + ":: please check it out."
    If AbortTest Then Exit Function Else Resume Next
End Function

''Attribute VB_Name = "VB_Utils"
''Option Explicit

'======================================================================
' Quick sort algo for numberic or string 1 dimensional array
'======================================================================
Public Sub QuickSort(vArray As Variant, inLow As Long, inHi As Long)
On Error GoTo errHandler
Dim funcName As String:: funcName = "QuickSort"
  Dim pivot   As Variant
  Dim tmpSwap As Variant
  Dim tmpLow  As Long
  Dim tmpHi   As Long

  tmpLow = inLow
  tmpHi = inHi

  pivot = vArray((inLow + inHi) \ 2)

  While (tmpLow <= tmpHi)
     While (vArray(tmpLow) < pivot And tmpLow < inHi)
        tmpLow = tmpLow + 1
     Wend

     While (pivot < vArray(tmpHi) And tmpHi > inLow)
        tmpHi = tmpHi - 1
     Wend

     If (tmpLow <= tmpHi) Then
        tmpSwap = vArray(tmpLow)
        vArray(tmpLow) = vArray(tmpHi)
        vArray(tmpHi) = tmpSwap
        tmpLow = tmpLow + 1
        tmpHi = tmpHi - 1
     End If
  Wend

  If (inLow < tmpHi) Then QuickSort vArray, inLow, tmpHi
  If (tmpLow < inHi) Then QuickSort vArray, tmpLow, inHi
Exit Sub
errHandler:
    TheExec.Datalog.WriteComment "<Error> " + funcName + ":: please check it out."
    If AbortTest Then Exit Sub Else Resume Next
End Sub

'===============================================================================
' Reburst an SSN pattern applying mask until all failing cores identified
'===============================================================================
Private Function ssnReburst(PatName As String, _
                            ssnResults As IScanNetworkResults, _
                            PerCorePFResults As PinListData, _
                            logScanFails As Boolean, _
                            PreTestF As InterposeName, _
                            PreTestFArgs As String, _
                            PostTestF As InterposeName, _
                            PostTestFArgs As String, _
                            reburstCnt_Site As SiteVariant) As IScanNetworkResults
On Error GoTo errHandler
Dim funcName As String:: funcName = "ssnReburst"
    Dim errMsg As String
    
    '___ Init __________________________________________________________________
     On Error GoTo errHandler
    
    '___ Make sure reburst is needed ___________________________________________
    If (TheExec.sites.Selected.Count = 0) Then
        TheExec.flow.TestLimitIndex = TheExec.flow.TestLimitIndex + 2
        Exit Function
    End If
        
    '___ Reburst loop __________________________________________________________
    Dim coreNames() As String
    Dim coreName As Variant
    Dim siteMasksByCore As New Dictionary
    Dim coreIsMasked As SiteBoolean
    Dim maxReburst As Long
    Dim reburstCnt As Long
    Dim ssnName As String
    Dim newCoreMask As New SiteVariant
    Dim reburstPFResults As New SiteBoolean
        
    ssnName = ssnResults.name
    coreNames = ssnResults.coreNames
    For Each coreName In coreNames
        Set coreIsMasked = New SiteBoolean
        siteMasksByCore.Add coreName, coreIsMasked
    Next coreName

    newCoreMask = ssnResults.FailedCores
    maxReburst = 10
    reburstCnt = 1
    
    While ((TheExec.sites.Selected.Count > 0) And (reburstCnt <= maxReburst))
    
        reburstCnt_Site = reburstCnt  '240603, Jayden, Hidra
        
        If (TheExec.enableWord("ENABLE_DEMO_STOP") = True) Then
            '************************************************************************************
            '*** This is the start of the Reburst loop                                        ***
            '***    - Previous burst failing cores will be applied as a mask for next reburst ***
            '************************************************************************************
            Stop
        End If
        
        '___ Initialize CMEM before each reburst _______________________________
        If glb_TesterType = "UltraFLEXplus" Then
            Call TheHdw.Digital.CMEM.SetCaptureConfig(-1, CmemCaptNone, tlCMEMCaptureSourcePassFailData, False)
            Call TheHdw.Digital.CMEM.SetCaptureConfig(True, CmemCaptFail, tlCMEMCaptureSourcePassFailData, True)

            TheHdw.Digital.CMEM.CaptureLimitMode = tlDigitalCMEMCaptureLimitMode_Enable
            TheHdw.Digital.CMEM.CaptureLimit = glb_SSN_CaptureLimit
            TheHdw.Digital.Patgen.ScanBurstEnabled = True
        Else
            Call TheHdw.Digital.CMEM.SetCaptureConfig(0, CmemCaptNone, tlCMEMCaptureSourcePassFailData, False)
            Call TheHdw.Digital.CMEM.SetCaptureConfig(-1, CmemCaptFail, tlCMEMCaptureSourcePassFailData, False)
            TheHdw.Digital.CMEM.CentralFields = tlCMEMPatternName + tlCMEMModCycle 'CJR: 5/14/24: require this new code for UF
        End If
        
        '___ Apply a mask of known failing cores _______________________________
        With TheHdw.Digital.ScanNetworks(ssnName).CoreMasks
            Call .AddPerSite(newCoreMask)
            Call .Apply
        End With
        
        If glb_TesterType = "UltraFLEXplus" Then
        '___ Update the masking info in the COND_LIST __________________________
        SSNUpdateCondLst ssnName, siteMasksByCore
        End If
                
        '___ Do the rebusrt ____________________________________________________
'        theexec.Flow.TestLimit resultVal:=reburstCnt, _
'                               scaletype:=scaleNoScaling, formatStr:="%d", _
'                               PinName:=CStr(reburstCnt), _
'                               ForceResults:=tlForceFlow
'        '___ Clear SSN masks and setup again_________20240521 Debug_________
'
'        Call TheHdw.Digital.CMEM.SetCaptureConfig(-1, CmemCaptNone, _
'                                                  tlCMEMCaptureSourcePassFailData, True)
'
'
'        '___ Setup CMEM ____________________________________________________________
'        Call TheHdw.Digital.CMEM.SetCaptureConfig(glb_SSN_CaptureLimit, CmemCaptFail, _
'                                                  tlCMEMCaptureSourcePassFailData, True)
'
'
'        TheHdw.Digital.CMEM.CaptureLimitMode = tlDigitalCMEMCaptureLimitMode_Enable
'        TheHdw.Digital.CMEM.CaptureLimit = glb_SSN_CaptureLimit
'        TheHdw.Digital.Patgen.ScanBurstEnabled = True
'        '___ Clear SSN masks and setup again________________________________

        
        
        If (logScanFails = True) Then
            Call TheHdw.patterns(PatName).test(pfAlways, 0, tlResultModeDomain)
            
        Else
            TheHdw.patterns(PatName).start
            TheHdw.patterns(PatName).HaltWait
            
            reburstPFResults = TheHdw.Digital.Patgen.PatternBurstPassedPerSite
            TheExec.flow.FunctionalTestLimit resultVal:=reburstPFResults, _
                                             PatName:=PatName, _
                                             ForceResults:=tlForceNone ' CJR: 6/21/24: Modified to prevent error in case logScanFails = False
                                             'ForceResults:=tlForceFlow
        End If
        
        '___Process the results _______________________________________________
        Dim reburstSites As New SiteBoolean
        Dim ssnReburstResults As IScanNetworkResults
        
        Set ssnReburstResults = TheHdw.Digital.Patgen.ReadScanNetworkResults
        
        '___ Call PostTest interpose __________________________________________
        'Call Interpose(PostTestF, PostTestFArgs)
        LogPinCoreFailsBySite ssnReburstResults, PostTestFArgs
        
        reburstSites = ProcessSSNResults(ssnReburstResults)
        'reburstSites = False    'cj suggests
        If (TheExec.enableWord("ENABLE_DEMO_STOP") = True) Then
            '**********************************************************************
            '*** Dataloged reburst # and results                                ***
            '***   - this is not required but done for demo visibility          ***
            '***   - can also optionally log scan fail data if needed           ***
            '***       - can use STR's COND_LST field to indicate masking setup ***
            '*** Next line merges reburst results into overall results          ***
            '**********************************************************************
            Stop
        End If
        
        ssnResults.MergeWith ssnReburstResults
        
        '___ Check if another reburst required _________________________________
        If (reburstSites.Any(True) = True) Then
            
            '___ Get new mask setup ____________________________________________
            newCoreMask = ssnReburstResults.FailedCores
            
            '___ Reuse the reburst limit sets values for all reburst ___________
            TheExec.flow.TestLimitIndex = TheExec.flow.TestLimitIndex - 2
            
        End If
        
        '___ Check if any site fail but no core fail _________T-The 20241126____
        Set nonSSNFailOnSSN = New SiteBoolean
        For Each site In TheExec.sites.Active
            If reburstSites(site) = False And TheHdw.Digital.Patgen.PatternBurstPassedPerSite(site) = False Then
                '___ Will turn on all harvest flag in ssn_body _________________
                nonSSNFailOnSSN(site) = logicTrue
            Else
                nonSSNFailOnSSN(site) = logicFalse
            End If
        Next site
        
        
        '___ Update the sites selected for reburst _____________________________
        TheExec.sites.Selected = reburstSites
                    
        reburstCnt = reburstCnt + 1
    Wend
    
    '___ Adjust limits index if max reburst count hit __________________________
    '___ and record on glb_maxReburst ________________________T-The 20241126____
    Set glb_maxReburst = New SiteBoolean
    If (reburstCnt >= maxReburst) Then
        
        For Each site In TheExec.sites.Active
            If reburstSites(site) = True Then
                glb_maxReburst(site) = logicTrue
            Else
                glb_maxReburst(site) = logicFalse
            End If
        Next site
        
        TheExec.flow.TestLimitIndex = TheExec.flow.TestLimitIndex '+ 2
    Else
        For Each site In TheExec.sites.Active
            glb_maxReburst(site) = logicFalse
        Next site
    End If
    
    '___ Cleanup _______________________________________________________________
    Set ssnReburst = ssnResults
Exit Function
errHandler:
    Stop
    Resume Next

    If (Len(errMsg) = 0) Then
        errMsg = err.Description
    End If
    TheExec.ErrorLogMessage ("*E* ssnReburst: " + errMsg)
    Set ssnReburst = ssnResults
        
    Exit Function
                            
End Function

' ==============================================================================
' Update the scan conditions (STR: COND_LST) with the SSN masking info
'      NOTE: Currently ScanSetup.Conditions is not site aware (IGXL-81548)
'            COND_LST will be the same for all sites
'            SSN_CORE MASK @ <core name> = <# #>
'                 where <core_name>   name from core map
'                       <# #>         space delimited list of sites
'                                     where t<core name> is masked
'===============================================================================
Private Function SSNUpdateCondLst(ssnName As String, siteMasksByCore As Dictionary) As Long
On Error GoTo errHandler
Dim funcName As String:: funcName = "SSNUpdateCondLst"
        '___ Update the masking info in the COND_LIST __________________________
        Dim site As Variant
        Dim coreName As Variant
        Dim maskedCores() As String
        
        For Each site In TheExec.sites.Active
            If (IsEmpty(TheHdw.Digital.ScanNetworks(ssnName).maskedCores(site)) = False) Then
                maskedCores = TheHdw.Digital.ScanNetworks(ssnName).maskedCores(site)
                For Each coreName In maskedCores
                    siteMasksByCore(coreName)(site) = True
                Next coreName
            End If
        Next site
        
        Dim condValue As String
        If glb_TesterType = "UltraFLEXplus" Then
            With TheExec.Datalog.Setup.ScanSetup.Conditions
                .RemoveAll
                For Each coreName In siteMasksByCore.Keys
                    condValue = ""
                    For Each site In TheExec.sites.Active
                        If (siteMasksByCore(coreName)(site) = True) Then
                            condValue = condValue + " " + CStr(site)
                        End If
                    Next site
                    If (Len(condValue) > 0) Then
                        .Add ("SSN_CORE_MASK @ " + coreName), condValue
                    End If
                Next coreName
            End With
        End If
Exit Function
errHandler:
    TheExec.Datalog.WriteComment "<Error> " + funcName + ":: please check it out."
    If AbortTest Then Exit Function Else Resume Next
End Function


'==========================================================================================
' Process the SSN results from the last burst
'    - Mark failing cores
'    - Check for fails outside of SSN scan blocks
'         - Mark all cores as failed
'         - Remove from reburst site list
'    - Determine the sites that require reburst
'==========================================================================================
Public Function ProcessSSNResults(ssnResults As IScanNetworkResults) As SiteBoolean
On Error GoTo errHandler
Dim funcName As String:: funcName = "ProcessSSNResults"
    Dim errMsg As String
    
    '___ Update per core results for identified core fails ________________________________
    '    Assumption is per core results was initialized as passing
    '    Keep track if a site has any failing cores
    '--------------------------------------------------------------------------------------
    Dim coreName As Variant
    Dim currentResults As New SiteBoolean
    Dim newResults As New SiteBoolean
    Dim anyCoreFailsBySite As New SiteBoolean
    
    For Each coreName In ssnResults.coreNames
        anyCoreFailsBySite = anyCoreFailsBySite.LogicalOr(ssnResults.core(coreName).Failed)
        currentResults = CorePFResults.pins(coreName)
        newResults = ssnResults.core(coreName).Failed
        CorePFResults.pins(coreName).value = currentResults.LogicalAnd(newResults.LogicalNot)
    Next coreName
    
    '___ Check for fails that only occur outside the SSN scan blocks ______________________
    Dim burstPF As New SiteBoolean
    Dim forceFailSites As New SiteBoolean
    Dim site As Variant
    
    burstPF = TheHdw.Digital.Patgen.PatternBurstPassedPerSite.LogicalNot
    forceFailSites = burstPF.LogicalAnd(anyCoreFailsBySite.LogicalNot)
        
    If (forceFailSites.Any(True) = True) Then
        For Each site In TheExec.sites.Selected
            If (forceFailSites(site) = True) Then
                'For Each coreName In ssnResults.coreNames
                If (UBound(ssnResults.FailedCores(site)) <> -1) Then ' CJR: 6/27/24: When failures outside scan happens, prevent to datalog cores
                    For Each coreName In ssnResults.FailedCores ' CJR: 6/27/24: When failures outside scan happens, prevent to datalog cores
                        CorePFResults.pins(coreName).value = False
                    Next coreName
                End If
            End If
        Next site
    End If

    '___ Determine which sites need to be reburst __________________________________________
    Dim reburstSites As New SiteBoolean
    reburstSites = ssnResults.ReburstNeeded.LogicalAnd(forceFailSites.LogicalNot)
    
    '___ Cleanup _______________________________________________________________
    Set ProcessSSNResults = reburstSites
    Exit Function
                            
errHandler:

    If (Len(errMsg) = 0) Then
        errMsg = err.Description
    End If
    TheExec.ErrorLogMessage ("*E* ProcessSSNResults: " + errMsg)
    Set ProcessSSNResults = reburstSites
End Function
Public Sub SSN_PreBody(EnableScanLogging As Boolean, scanLogPins As String, _
    maxFailsPerPin As Long, scanLogMode As tlDatalogScanResultMode, _
    enableFullSpeedMode As Boolean, ssnMapName As String)
On Error GoTo errHandler
Dim funcName As String:: funcName = "SSN_PreBody"
    
    CoreScanPins = scanLogPins
    '~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
    '~~~ START OF PRE_BODY ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
    '~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
    ''If (Step_ = subAllBody Or Step_ = subPrebody) Then
    
        ''TheHdw.Digital.ApplyLevelsTiming True, True, True, tlPowered
        ''TheHdw.Patterns(PatName).Load
        
        '___ Clear all existing SSN masks ______________________________________
        'Remove clear mask
        'TheHdw.Digital.ScanNetworks.ClearAllMasks
        
        '___ Setup for scan fail logging (if enabled) __________________________
        If (EnableScanLogging = True) And (glb_TesterType = "UltraFLEXplus") Then
            Call SetupForScanLogging(scanLogPins, _
                                     maxFailsPerPin, _
                                     tl_DCScanCaptureFormat_Cycle, _
                                     scanLogMode)
                                     
            
            '___ Set COND_LIST _________________________________________________
            With TheExec.Datalog.Setup.ScanSetup.Conditions
                .RemoveAll
                .Add "SSN_MASKING", "False"
            End With
            
        End If
        
        If glb_TesterType = "UltraFLEXplus" Then
        '___ Setup CMEM ____________________________________________________________
        Call TheHdw.Digital.CMEM.SetCaptureConfig(maxFailsPerPin, CmemCaptFail, _
                                                  tlCMEMCaptureSourcePassFailData, _
                                                  enableFullSpeedMode)
                                                  
        TheHdw.Digital.CMEM.CaptureLimitMode = tlDigitalCMEMCaptureLimitMode_Enable
        TheHdw.Digital.CMEM.CaptureLimit = maxFailsPerPin
        TheHdw.Digital.Patgen.ScanBurstEnabled = True
        Else
            'CJR: 5/14/24: In UF, enableFullSpeedMode has to be false. Since it's not supported in UF
            Call TheHdw.Digital.CMEM.SetCaptureConfig(-1, CmemCaptFail, _
                                                      tlCMEMCaptureSourcePassFailData, _
                                                      False)
            TheHdw.Digital.CMEM.CentralFields = tlCMEMPatternName + tlCMEMModCycle 'CJR: 5/14/24: require this new code for UF
        End If

                                                      
        '___ Build the per core PF PLD _________________________________________
        '    Core names are dummy pins in PLD
        '    Init all cores to pass
        '-----------------------------------------------------------------------
        
        '===Need to be delete 240517
        Dim coreNames() As String
        coreNames = TheHdw.Digital.ScanNetworks(ssnMapName).coreNames
        QuickSort coreNames, 0, UBound(coreNames)
                    
        Set CorePFResults = New PinListData
        
        Set CorePF = New SiteBoolean
        CorePF = True
        
        Dim coreName As Variant
        For Each coreName In coreNames
            CorePFResults.AddPin coreName
            CorePFResults.pins(coreName).value = CorePF
        Next coreName
       
    ''End If
    '~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
    '~~~ END OF PRE_BODY ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
    '~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'    If theexec.DataManager.instanceName Like "*_COREMASK" Then 'maxFailsPerPin = 50000
'        ssnApplyCoreMask ssnResultsGlobal
'    End If
Exit Sub
errHandler:
    TheExec.Datalog.WriteComment "<Error> " + funcName + ":: please check it out."
    If AbortTest Then Exit Sub Else Resume Next
End Sub

Public Sub SSN_Body(instSSNinfo As Inst_SSN, enableReburst As Boolean, enableReburstScanLogging As Boolean, _
    shortPatName As String, _
                            Optional PreTestF As InterposeName, _
                            Optional PreTestFArgs As String, _
                            Optional PostTestF As InterposeName, _
                            Optional PostTestFArgs As String)
                            PostTestFArgs = CoreScanPins
                            
On Error GoTo errHandler
Dim funcName As String:: funcName = "SSN_Body"
    '___ Reburst loop __________________________________________________________
    Dim coreName As Variant
    
     On Error GoTo errHandler
    '~~~ START OF BODY ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
    '~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
    ''If Step_ = subAllBody Or Step_ = subBody Then
    
        '___ Call StartOfBody interpose if defined _____________________________
        ''Call Interpose(StartOfBodyF, StartOfBodyFArgs)
        
        '___ Cache the sites selected for start of test ________________________
        Dim startOfTestSites As New SiteBoolean
        startOfTestSites = TheExec.sites.Selected
        'If startOfTestSites.All(False) Then Exit Sub 'pg 1007 not sure why erring now, add this
        
        '___ 'test' the pattern ________________________________________________
        'Call TheHdw.Patterns(ShortPatName).Test(pfAlways, 0, tlResultModeModule)
        
        ''Call Interpose(PostTestF, PostTestFArgs)
        ''LogPinCoreFailsBySite ssnResults, CoreScanPins
        
        '___ Process the initial SSN results / determine reburst sites _________
        Dim reburstSites As New SiteBoolean
        Dim ssnResults As IScanNetworkResults
        Dim ReburstCount_SV As New SiteVariant: ReburstCount_SV = 0
        
        Set ssnResults = TheHdw.Digital.Patgen.ReadScanNetworkResults
        reburstSites = ProcessSSNResults(ssnResults)
        Set ssnResultsGlobal = ssnResults 'TheHdw.Digital.Patgen.ReadScanNetworkResults
        
        LogPinCoreFailsBySite ssnResults, CoreScanPins
        
        If (TheExec.enableWord("ENABLE_DEMO_STOP") = True) Then
            '****************************************************************************
            '*** Review SSN results for initial burst ***
            '***    - Mapping file name in effect
            '***    - All core names contained in map
            '***    - Failed cores by site (SiteVariant - array of strings)
            '***    - Reburst needs by site (SiteBoolean)
            '***    - Optionally can query by core name
            '***         ssnResults.Core("CoreName").Failed|.Masked
            '****************************************************************************
            Stop
        End If
        'enableReburst = False 'pg force no reburst
        '___ Is Reburst Enabled ________________________________________________
        If (enableReburst = True) Then
            
            '___ Is rebust needed ______________________________________________
            If (reburstSites.Any(True) = True) Then
            
                '___ Cache the selected sites __________________________________
                Dim origSelectedSites As New SiteBoolean
                origSelectedSites = TheExec.sites.Selected
                TheExec.sites.Selected = reburstSites
                
                '___ Setup Rebusrt scan logging ________________________________
                If (enableReburstScanLogging = False) Then
                    TheExec.Datalog.Setup.ScanSetup.EnableScanLogging = False
                    TheExec.Datalog.ApplySetup
                End If

                '___ Do the reburst ____________________________________________
                Set ssnResults = ssnReburst(shortPatName, ssnResults, CorePFResults, _
                                            enableReburstScanLogging, _
                                            PreTestF, PreTestFArgs, _
                                            PostTestF, PostTestFArgs, ReburstCount_SV)
                                            
                '___ Re-enable original selected sites _________________________
                TheExec.sites.Selected = origSelectedSites
                                
                If reburstSites.Any(True) Then
                    TheHdw.Digital.ScanNetworks(ssnResults.name).CoreMasks.Clear
                    TheExec.Datalog.WriteComment "Clear SSN core mask after reburst" ', SSN=" & ssnResults.Name

                    ''If there is masked cores from previous instance exist, restore masked cores
                    For Each site In TheExec.sites                                                  'Modify on 240628
                        If Not IsEmpty(instSSNinfo.maskedCores(instSSNinfo.idxMatchedPattern)(site)) Then
                            Dim strCore As Variant
                            For Each strCore In instSSNinfo.maskedCores(instSSNinfo.idxMatchedPattern)(site)
                                TheHdw.Digital.ScanNetworks(ssnResults.name).CoreMasks.Add strCore
                                TheExec.Datalog.WriteComment "site:" & site & ", restore original core mask=" & strCore
                            Next strCore
                            TheHdw.Digital.ScanNetworks(ssnResults.name).CoreMasks.Apply ' CJ: 3/6/24: Missing CoreMask Apply
                        End If
                    Next site
                End If
            Else
                Set nonSSNFailOnSSN = New SiteBoolean
                Set glb_maxReburst = New SiteBoolean
                For Each site In TheExec.sites.Active
                ' No need this code'
                    'If reburstSites(site) = False And TheHdw.Digital.Patgen.PatternBurstPassedPerSite(site) = False Then
                        '___ Will turn on all harvest flag in ssn_body _________________
                        'nonSSNFailOnSSN(site) = logicTrue
                    'Else
                    nonSSNFailOnSSN(site) = logicFalse
                    glb_maxReburst(site) = logicFalse
                    'End If
                Next site
            
                '___ No reburst - skip entries in Limits Set _______________________
'                TheExec.Flow.TestLimitIndex = TheExec.Flow.TestLimitIndex + 2
                TheExec.flow.TestLimitIndex = TheExec.flow.TestLimitIndex 'pg+ 2
            
            End If  '<- Reburst needed?
        
        Else
            '___ If reburst is needed but not enabled ___________________________
            '    Fail all cores for only those sites where reburst needed
            '       - Don't know they are good so assume they are bad
            '--------------------------------------------------------------------
            Dim currentResults As New SiteBoolean
            Dim newResults As New SiteBoolean
            
            For Each coreName In ssnResults.coreNames
                If (ssnResults.ReburstNeeded.Any(True) = True) Then
                    
                    currentResults = CorePFResults.pins(coreName)
                    newResults = ssnResults.ReburstNeeded
                    CorePFResults.pins(coreName).value = currentResults.LogicalAnd(newResults.LogicalNot)
                    
                End If
                
            Next coreName
                
            '___ No reburst - skip entries in Limits Set _______________________
'pg             TheExec.Flow.TestLimitIndex = TheExec.Flow.TestLimitIndex + 2
             TheExec.flow.TestLimitIndex = TheExec.flow.TestLimitIndex 'pg+ 2


                        
        End If  '<- Reburst enabled
   
        TheExec.flow.TestLimit resultVal:=ReburstCount_SV, lowVal:=0, hiVal:=0, ForceResults:=tlForceNone, Tname:="Reburst Count"  '240603, Jayden, Hidra
   
        '___ Datalog per core P/F results ______________________________________

        Dim printPF As New SiteLong
        Dim core As Variant
        Dim CoreArray As Variant
        Dim FlagArray As Variant
        Dim i As Long, j As Long
        Dim TempFlag As String
        Dim TempFlagStr As String
        For j = 0 To UBound(instSSNinfo.patkeyword) 'CJR:4/20/24
            If UCase(shortPatName) Like instSSNinfo.patkeyword(j) Then
            
                '___ Turn on all harvest flag if "pattern fail but no core fail" or _
                '___ "Reburst time exceed the upper limit" _________T-The 20241126___
                For Each site In TheExec.sites.Selected
                    
                    '___ Get harvest fail flag without any duplicate flag____________
                    FlagArray = Split(instSSNinfo.Failflag(j), ";")
                    TempFlagStr = ";"
                    For i = 0 To UBound(FlagArray)
                        If Not TempFlagStr Like "*;" & FlagArray(i) & ";*" Then
                            TempFlagStr = TempFlagStr & FlagArray(i) & ";"
                        End If
                    Next i
                    TempFlagStr = mid(TempFlagStr, 1, Len(TempFlagStr) - 2)
                    FlagArray = Split(TempFlagStr, ";")
                    
                    '___ Check pattern fail but no core fail_________________________
                    If nonSSNFailOnSSN(site) Then
                        For i = 0 To UBound(FlagArray)
                            If FlagArray(i) <> "" Then
                                TheExec.sites(site).FlagState(FlagArray(i)) = logicTrue
                                TheExec.flow.TestLimit resultVal:=-1, lowVal:=0, hiVal:=0, Tname:=glb_TestInstance & " : Fail on non-SSN vector", PinName:=FlagArray(i)
                            End If
                        Next i
                        
                    '___ Check Reburst time exceed the upper limit___________________
                    ElseIf glb_maxReburst(site) Then
                        For i = 0 To UBound(FlagArray)
                            If FlagArray(i) <> "" Then
                                TheExec.sites(site).FlagState(FlagArray(i)) = logicTrue
                                TheExec.flow.TestLimit resultVal:=-1, lowVal:=0, hiVal:=0, Tname:=glb_TestInstance & " : Reburst time exceed the upper limit", PinName:=FlagArray(i)
                            End If
                        Next i
                    End If
                Next
                
                For Each core In CorePFResults.pins
                    CorePF = CorePFResults.pins(core)
                    printPF = CorePFResults.pins(core)
                    printPF = printPF.Add(1)
'pg            TheExec.Flow.TestLimit resultVal:=CorePF, _
                                   scaletype:=scaleNoScaling, formatStr:="%d", _
                                   unit:=unitCustom, customUnit:="P/F", _
                                   ForceResults:=tlForceFlow
                    'Modified low/hi limit = 0 with resultPF = PASS(0)/FAIL(1)
                    'CorePF=0 means fail, CorePF=-1 means pass
                    'TheExec.Flow.TestLimit resultVal:=printPF, lowVal:=0, hiVal:=0, Tname:=core
                    
            
                    CoreArray = Split(instSSNinfo.coreName(j), ";")
                    FlagArray = Split(instSSNinfo.Failflag(j), ";")
                    
                    For Each site In TheExec.sites.Selected
                        
                        '___ Skip "pattern fail but no core fail" ____T-The 20241126_
                        '___ and "Reburst time exceed the upper limit" case__________
                        If Not (nonSSNFailOnSSN(site) Or glb_maxReburst(site)) Then
                            If CorePF(site) = False Then
                                For i = 0 To UBound(CoreArray)
                                    If LCase(core) = LCase(CoreArray(i)) Then
            '                        If LCase(core) Like LCase("*" & CoreArray(i) & "*") Then
                                        TheExec.sites(site).FlagState(FlagArray(i)) = logicTrue
                                        TempFlag = FlagArray(i)
                                        Exit For
                                    End If
                                Next i
                                TheExec.flow.TestLimit resultVal:=printPF(site), lowVal:=0, hiVal:=0, Tname:=core, PinName:=TempFlag
                            End If
                        End If
                    Next site
                Next core
        End If
    Next j
    
        
        'Update index of matched SSN patterns
        instSSNinfo.idxMatchedPattern = instSSNinfo.idxMatchedPattern + 1
        
         '___ Call StartOfBody interpose if defined ____________________________
        ''Call Interpose(EndOfBodyF, EndOfBodyFArgs)
        
        ''If (TheExec.enableWord("ENABLE_DEMO_STOP") = True) Then
            '*************************************************************************
            '*** Per site P/F results were datalogged                              ***
            '***    CorePFResults is PLD where "pin" is the core name              ***
            '***    Could do a single test and eliminate loop                      ***
            '***       - using loop for demo to allow unique TName per core        ***
            '***       - without loop would need to log "pin" to see core name     ***
            '*************************************************************************
        ''    Stop
        ''End If
   
    ''End If
    '~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
    '~~~ END OF BODY ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
    '~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
Exit Sub
errHandler:
    TheExec.Datalog.WriteComment "<Error> " + funcName + ":: please check it out."
    If AbortTest Then Exit Sub Else Resume Next
End Sub

Public Sub SSN_PostBody(enableFullSpeedMode As Boolean)
On Error GoTo errHandler
Dim funcName As String:: funcName = "SSN_PostBody"
    '~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
    '~~~ START OF POST-BODY ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
    '~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
    ''If Step_ = subAllBody Or Step_ = subPostbody Then
    
    
        '___ Clear SSN masks __________________________________________________
        'TheHdw.Digital.ScanNetworks.ClearAllMasks
        
        If glb_TesterType = "UltraFLEXplus" Then
        Call TheHdw.Digital.CMEM.SetCaptureConfig(-1, CmemCaptNone, _
                                                  tlCMEMCaptureSourcePassFailData, _
                                                  enableFullSpeedMode)
                                                      
        '___ Set COND_LIST ____________________________________________________
        With TheExec.Datalog.Setup.ScanSetup.Conditions
            .RemoveAll
            .Add "SSN_MASKING", "False"
        End With
        Else
            Call TheHdw.Digital.CMEM.SetCaptureConfig(0, CmemCaptNone, _
                                                  tlCMEMCaptureSourcePassFailData, _
                                                  False)
        End If
        
        '___ Release the global core PLD ______________________________________
        Set CorePFResults = Nothing
     
     ''End If
    '~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
    '~~~ END OF POST-BODY ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
    '~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
Exit Sub
errHandler:
    TheExec.Datalog.WriteComment "<Error> " + funcName + ":: please check it out."
    If AbortTest Then Exit Sub Else Resume Next
End Sub



''Attribute VB_Name = "VBT_Interpose"
''Option Explicit

''Public DebugMsgEnable As Boolean

'===============================================================================
' Disable compares so no fails are logged
'===============================================================================
Public Function LogPeriod(argc As Long, argv() As String) As Long
On Error GoTo errHandler
Dim funcName As String:: funcName = "LogPeriod"

    '___ Check if this is globally enabled _____________________________________
    If (DebugMsgEnable = False) Then
        Exit Function
    End If
    
    '___ Log the current t_period spec value ___________________________________
    Dim period As New SiteDouble
    period = TheExec.Specs.AC("t_period").ContextValue
    
    Dim site As Variant
    For Each site In TheExec.sites.Starting
        TheExec.Datalog.WriteComment vbCrLf + "*I* Digital Period = " + Format(period(site) / 0.000000001, "##0.000 nsec") + vbCrLf
        Exit For
    Next site
Exit Function
errHandler:
    TheExec.Datalog.WriteComment "<Error> " + funcName + ":: please check it out."
    If AbortTest Then Exit Function Else Resume Next
End Function


'===============================================================================
' Log the failing cores and fail count per pin per site for the specified pins
'===============================================================================
'Public Function LogPinCoreFailsBySite(argc As Long, argv() As String) As Long
Public Function LogPinCoreFailsBySite(ssnResults As IScanNetworkResults, ScanPins As String) As Long
On Error GoTo errHandler
Dim funcName As String:: funcName = "LogPinCoreFailsBySite"
    Dim argv() As String
    argv = Split(ScanPins, ",")
    DebugMsgEnable = True

    '___ Check if this is globally enabled _____________________________________
    If (DebugMsgEnable = False) Then
        Exit Function
    End If
    
    '___ Log fail count per pin per site ______________________________________
    Dim pinGrp As Variant
    Dim site As Variant
    Dim failPin As Variant
    Dim pinFails As New PinListData
    Dim msg As String
    
    '=====240619 Need to check=====
    If glb_TesterType = "UltraFLEXplus" Then
        TheExec.Datalog.WriteComment vbCrLf + _
                                     "*I* Max Fails Per Pin: " + CStr(TheExec.Datalog.Setup.ScanSetup.CMEMCaptureLimit)
    Else
        'CMEMCaptureLimit does not support on FLEX
        Dim CMEMCaptureSize As Long
        Dim CMEMCaptureType As Long
        TheHdw.Digital.CMEM.GetCaptureConfig CMEMCaptureSize, CMEMCaptureType
        TheExec.Datalog.WriteComment vbCrLf + _
                                     "*I* Max Fails Per Pin: " + CStr(CMEMCaptureSize)
    End If
    '=====240619 Need to check=====
    
'    theexec.Datalog.WriteComment "*I* Detected fail per pin per site: " + theexec.DataManager.instanceName
'
'    For Each pinGrp In argv
'
'        pinFails = TheHdw.Digital.Pins(pinGrp).FailCount
'
'        For Each site In theexec.sites.Selected
'            msg = "    Pins: " + Format(pinGrp, "@@@@@@@@@@@@@@@@") + "   Site: " + Format(CStr(site), "@@") + "  "
'            For Each failPin In pinFails.Pins
'                msg = msg + failPin + ":" + Format(CStr(pinFails.Pins(failPin).value(site)), "@@@@") + "  "
'            Next failPin
'            theexec.Datalog.WriteComment msg
'        Next site
'    Next pinGrp
    
    '___ Log failing core per site _____________________________________________
    'Dim ssnResults As IScanNetworkResults
    Dim failingCores As String
    Dim failingCore As Variant
    
    'Set ssnResults = TheHdw.Digital.Patgen.ReadScanNetworkResults

    msg = vbCrLf + "*I* Failing Cores Per Site: " + TheExec.DataManager.instancename + vbCrLf
    For Each site In TheExec.sites.Selected
        failingCores = ""
        If (UBound(ssnResults.FailedCores(site)) >= 0) Then
            For Each failingCore In ssnResults.FailedCores(site)
                failingCores = failingCores + " " + failingCore
            Next failingCore
        Else
            failingCores = "None "
        End If
        msg = msg + "    Site: " + CStr(site) + " Failing Cores: " + failingCores
        
        '___ Check if all fails captured _______________________________________
        If (ssnResults.ReburstNeeded(site) = False) Or (UBound(ssnResults.FailedCores(site)) = -1) Then
            msg = msg + "  All fails captured - No Reburst Required" + vbCrLf
        Else
            msg = msg + "  All fails not captured - Reburst Required" + vbCrLf
        End If
        
    Next site
    TheExec.Datalog.WriteComment msg

    '___ Log the current t_period spec value ___________________________________
    'Dim period As New SiteDouble
    'period = TheExec.Specs.AC("t_period").ContextValue
    
    'Dim Site As Variant
    'For Each Site In TheExec.sites.Starting
    '    TheExec.Datalog.WriteComment vbCrLf + "*I* Digital Period = " + Format(period(Site) / 0.000000001, "##0.000 nsec") + vbCrLf
    '    Exit For
    'Next Site

Exit Function
errHandler:
    TheExec.Datalog.WriteComment "<Error> " + funcName + ":: please check it out."
    If AbortTest Then Exit Function Else Resume Next
End Function




'===============================================================================
' Mask SSN Cores by Name
' Expected argument...... "#:XX+YY+ZZ, #:AA+BB" ...
'    # is a specific site number or "EVEN" or "ODD"
'           note: if site not selected then ignored
'    XX,YY,ZZ,AA,BB are core names to mask for that site
'===============================================================================
Public Function MaskSSNCoresBySite(argc As Long, argv() As String) As Long
On Error GoTo errHandler
Dim funcName As String:: funcName = "MaskSSNCoresBySite"
    If (argc > 0) Then
    
        Dim argIdx As Long
        Dim ssnName As String
        ssnName = Trim(argv(0))
    
        For argIdx = 1 To (argc - 1)
        
            '___ Get the masking data from the args ____________________________
            Dim Fields() As String
            Dim maskedCoreNames() As String
            Dim CoreIdx As Long
            
            Fields = Split(argv(argIdx), ":")
            
            maskedCoreNames = Split(Fields(1), "+")
            
            For CoreIdx = 0 To UBound(maskedCoreNames)
                maskedCoreNames(CoreIdx) = Trim(maskedCoreNames(CoreIdx))
            Next CoreIdx
            
            '___ Build the mask ________________________________________________
            Dim newCoreMask As New SiteVariant
            Dim site As Variant
            
            If (Trim(Fields(0)) = "EVEN") Then
                '___ Mask applied to all even sites ___________________________
                For Each site In TheExec.sites.Selected
                    If ((site Mod 2) = 0) Then
                        newCoreMask = maskedCoreNames
                    End If
                Next site
                
            
            ElseIf (Trim(Fields(0)) = "ODD") Then
                '___ Mask applied to all odd sites ____________________________
                For Each site In TheExec.sites.Selected
                    If ((site Mod 2) <> 0) Then
                        newCoreMask = maskedCoreNames
                    End If
                Next site
            
            Else
                '___ Mask applied to specific _________________________________
                Dim siteNum As Long
                siteNum = CLng(Trim(Fields(0)))
                
                For Each site In TheExec.sites.Selected
                    If (site = siteNum) Then
                        newCoreMask = maskedCoreNames
                    End If
                Next site

                
            End If
        
        Next argIdx
        
        '___ Apply the mask _________________________________________________
        With TheHdw.Digital.ScanNetworks(ssnName).CoreMasks
            .Clear
            .AddPerSite newCoreMask
            .Apply
        End With
        
    
    End If
    
Exit Function
errHandler:
    TheExec.Datalog.WriteComment "<Error> " + funcName + ":: please check it out."
    If AbortTest Then Exit Function Else Resume Next
End Function

Public Sub ssnApplyCoreMask(ssnResults As IScanNetworkResults)
On Error GoTo errHandler
Dim funcName As String:: funcName = "ssnApplyCoreMask"
    Dim coreNames() As String
    Dim coreName As Variant
    Dim siteMasksByCore As New Dictionary
    Dim coreIsMasked As SiteBoolean
    Dim ssnName As String
    Dim newCoreMask As New SiteVariant
    newCoreMask = ssnResults.FailedCores
    'newCoreMask.
        
    ssnName = ssnResults.name
    coreNames = ssnResults.coreNames
    For Each coreName In coreNames
        Set coreIsMasked = New SiteBoolean
        siteMasksByCore.Add coreName, coreIsMasked
    Next coreName

    newCoreMask = ssnResults.FailedCores
        
        '___ Apply a mask of known failing cores _______________________________
        With TheHdw.Digital.ScanNetworks(ssnResults.name).CoreMasks
            Call .AddPerSite(newCoreMask)
            Call .Apply
        End With

        '___ Update the masking info in the COND_LIST __________________________
        SSNUpdateCondLst ssnResults.name, siteMasksByCore
Exit Sub
errHandler:
    TheExec.Datalog.WriteComment "<Error> " + funcName + ":: please check it out."
    If AbortTest Then Exit Sub Else Resume Next
End Sub

Public Function SPN(ByVal pat As String) As String
On Error GoTo errHandler
Dim funcName As String:: funcName = "SPN"
    Dim shortPatName As String
    Dim ShortPos As Long: ShortPos = 1
    While InStr(ShortPos, pat, "\") > 0
        ShortPos = InStr(ShortPos, pat, "\") + 1
    Wend
    shortPatName = mid(pat, ShortPos, InStr(UCase(pat), ".PAT") - ShortPos)
    SPN = shortPatName
Exit Function
errHandler:
    TheExec.Datalog.WriteComment "<Error> " + funcName + ":: please check it out."
    If AbortTest Then Exit Function Else Resume Next
End Function
'Find patterns contains "SSC" and "SSU"
Public Function FindSSNPattern(PatName As String)
On Error GoTo errHandler
Dim funcName As String:: funcName = "FindSSNPattern"
Dim SplitPat() As String
Dim TempPat As String
Dim i As Byte

PatName = UCase(PatName)

If (PatName Like "*_PL*_SSC_*" Or PatName Like "*_PL*_SSU_*") And (InStr(PatName, "_LPB_") = 0 Or InStr(PatName, "_CON_") = 0) And glb_TesterType = "UltraFLEXplus" Then

    If InStr(PatName, ":") > 0 Then ' 20240605, Jayden, Hidra
        SplitPat = Split(Split(PatName, ":")(1), "_")
        For i = 0 To UBound(SplitPat) - 3
            If i = 0 Then
                TempPat = SplitPat(i)
            Else
                TempPat = TempPat + "_" + SplitPat(i)
            End If
        Next
        PatName = TempPat
    End If
ElseIf UCase(PatName) Like "*_PL*_SSC_*" Or UCase(PatName) Like "*_PL*_SSU_*" And glb_TesterType <> "UltraFLEXplus" Then
    PatName = Split(PatName, ":")(0)
    If ssnPatternsDict.Exists(UCase(PatName)) = False Then
        ssnPatternsDict.Add UCase(PatName), ""
    End If
End If

Exit Function
errHandler:
    TheExec.Datalog.WriteComment "<Error> " + funcName + ":: please check it out."
    If AbortTest Then Exit Function Else Resume Next
End Function

Public Function MaskCoreForSSN(strCoreName As String, strFailFlag As String, ssnMappingName As String, Optional ByRef isAnyCoreMasked As Boolean)
On Error GoTo errHandler
Dim funcName As String:: funcName = "MaskCoreForSSN"
Dim coreToMask As New SiteVariant, strCoreToMask As String
Dim i As Long
Dim site As Variant

Dim failFlagArr As Variant, coreNameArr As Variant
failFlagArr = Split(strFailFlag, ";")
coreNameArr = Split(strCoreName, ";")

''Check valid
If UBound(failFlagArr) <> UBound(coreNameArr) Then
    Call Print_Error_Message(Error_Info, "Lib_SSN", "ssnCoreMask", "The quantity of SSN core name:" & strCoreName & ", is not the same as SSN fail flag:" & strFailFlag)
    Exit Function
End If

''Concatenate core string by site
With TheHdw.Digital.ScanNetworks(ssnMappingName).CoreMasks      ''Add core mask by site
    For Each site In TheExec.sites
        For i = 0 To UBound(failFlagArr)
            If TheExec.sites(site).FlagState(failFlagArr(i)) = logicTrue Then
                
                Call .Add(coreNameArr(i))       ''Add core mask by core
                TheExec.flow.TestLimit resultVal:=1, lowVal:=1, hiVal:=1, Tname:=coreNameArr(i) & "_masked on SSN:" & ssnMappingName, ForceResults:=tlForceNone
                If isAnyCoreMasked = False Then
                    isAnyCoreMasked = True
                End If
            End If
        Next i
    Next site
        Call .Apply     ''Apply core mask by site
End With

Exit Function
errHandler:
    TheExec.Datalog.WriteComment "<Error> " + funcName + ":: please check it out."
    If AbortTest Then Exit Function Else Resume Next
End Function

Public Function SSNConfigSetting(pat As String, bSSNTest As Boolean, bSSNCoreHarvest As Boolean, instSSNinfo As Inst_SSN)
    On Error GoTo errHandler
    Dim funcName As String:: funcName = "SSNConfigSetting"
    Dim shortPatName As String
    Dim i As Integer
    'If bSSNTest = True And LCase(Pat) Like "*_pl*" Then     '20231120: Added pl only
    For i = 0 To UBound(instSSNinfo.patkeyword)
        If bSSNTest = True And LCase(pat) Like "*_pl*" And UCase(pat) Like "*" & instSSNinfo.patkeyword(i) & "*" Then     'CJR: 4/18/24
        shortPatName = SPN(UCase(pat))     ''Split pattern name with path
        If (TheExec.flow.IsCharacterizing = False) Then
            If bSSNCoreHarvest = True Then
                '===Need to be delete 240517
                'Stop
                Dim SSN_ScanPins As String: SSN_ScanPins = ""
                Dim maxFailsPerPin As Long: maxFailsPerPin = glb_SSN_CaptureLimit
'                                                If LCase(TheExec.DataManager.instanceName) Like "*_final*" Then
                SSN_PreBody True, SSN_ScanPins, maxFailsPerPin, tlDatalogScanResultMode_Module, True, shortPatName
'                                                Else
'                                                    Stop
'                                                    'SSN_PreBody True, SSN_ScanPins, maxFailsPerPin, tlDatalogScanResultMode_Module, True, SPN(pat) & ".SSNMapping"
'                                                End If
            Else
                '=====240625 Update for SFC=====
                Call TheHdw.Digital.CMEM.SetCaptureConfig(-1, CmemCaptFail, _
                                                          tlCMEMCaptureSourcePassFailData)
                TheHdw.Digital.CMEM.CentralFields = tlCMEMPatternName + _
                                                    tlCMEMVMVectorOffset + _
                                                    tlCMEMModCycle
            End If
        Else
            '=====240625 Update for SFC=====
            Call TheHdw.Digital.CMEM.SetCaptureConfig(-1, CmemCaptFail, _
                                                      tlCMEMCaptureSourcePassFailData)
            TheHdw.Digital.CMEM.CentralFields = tlCMEMPatternName + _
                                                tlCMEMVMVectorOffset + _
                                                tlCMEMModCycle
            ' Running Char flow
        End If
    Else
        '=====240625 Update for SFC=====
        Call TheHdw.Digital.CMEM.SetCaptureConfig(-1, CmemCaptFail, _
                                                  tlCMEMCaptureSourcePassFailData)
        TheHdw.Digital.CMEM.CentralFields = tlCMEMPatternName + _
                                            tlCMEMVMVectorOffset + _
                                            tlCMEMModCycle
    End If
    Next i
                                    
    Exit Function
errHandler:
    TheExec.Datalog.WriteComment "<Error> " + funcName + ":: please check it out."
    If AbortTest Then Exit Function Else Resume Next
End Function

Public Function SSNDataAnalysis(pat As String, instSSNinfo As Inst_SSN, SCAN_Site_Blooean As SiteBoolean, Optional SecondSSN As Boolean = False)
    On Error GoTo errHandler
    Dim funcName As String:: funcName = "SSNDataAnalysis"
    Dim site As Variant
    Dim shortPatName As String
    Dim i As Integer
    
    Dim isPatternSSN As Boolean: isPatternSSN = False
    If instSSNinfo.bSSNTest = True And instSSNinfo.bSSNCoreHarvest = True Then 'And SCAN_Site_Blooean.Any(False) Then
        For i = 0 To UBound(instSSNinfo.patkeyword)
             If LCase(pat) Like "*_pl*" And UCase(pat) Like "*" & instSSNinfo.patkeyword(i) & "*" Then          'CJR: 4/15/24: SSC AND PRC in same Patset 4/15/24 Then
                If (TheExec.flow.IsCharacterizing = False) Then
                    Dim SSN_InterposeFunc As InterposeName ': SSN_InterposeFunc = "LogPinCoreFailsBySite"
                    shortPatName = SPN(UCase(pat))
                    SSN_Body instSSNinfo, True, True, shortPatName, , , SSN_InterposeFunc  ', SSN_ScanPins
                    SSN_PostBody False
                    SecondSSN = True
                Else
                    '___ Clear SSN masks __________________________________________________
                'theHdw.Digital.ScanNetworks.ClearAllMasks
                End If
                
                isPatternSSN = True
                
                Exit For        '' 20240502: Added to stop finding new pattern keyword
'
            End If
        Next i
        
        If isPatternSSN = False Then ''CJR: Added this boolean to avoid 2nd SSN PL being recognized as init pattern
            For Each site In TheExec.sites
                If SCAN_Site_Blooean(site) = False Then
                    TheExec.sites(site).FlagState(glb_SSN_Failflag) = logicTrue
                    TheExec.flow.TestLimit resultVal:=1, lowVal:=0, hiVal:=0, Tname:=CStr(glb_SSN_Failflag), ForceResults:=tlForceNone
                    TheExec.Datalog.WriteComment "Non-'SSN payload pattern' fails in SSN instance!"     '' 20240502: Added to clarify
                End If
            Next site
        End If
        
    End If
     Call TheHdw.Digital.CMEM.SetCaptureConfig(-1, CmemCaptNone, _
                                               tlCMEMCaptureSourcePassFailData, _
                                               False)


    Exit Function
errHandler:
    TheExec.Datalog.WriteComment "<Error> " + funcName + ":: please check it out."
    If AbortTest Then Exit Function Else Resume Next
End Function

Public Function SSNMainProcedure(patset() As String, ReportResult As PFType, TL_C_YES As Long, ResultMode As tlResultMode, _
            ConcurrentMode As tlPatConcurrentMode, ByRef SCAN_Site_Blooean As SiteBoolean, instSSNinfo As Inst_SSN, Optional ApplyVoltageFromBinCut As String = vbNullString, _
            Optional Harv_FailFlag As String = vbNullString, Optional HarvestPinGrpOtherFail As String = vbNullString)
    
    On Error GoTo errHandler
    Dim funcName As String:: funcName = "SSNMainProcedure"
    Dim pat As Variant
    Dim inst_info As Instance_Info
'    Dim instrumentUtility As New Instrument_Utility
    Dim i, j As Long
    Dim SecondSSN As Boolean: SecondSSN = False
    '=====240619 Need to check===== Cannot set CMEM capture limit on FLEX
    'Harvest_CMEM_InitSetup
    '=====240619 Need to check=====
    Dim sBool_PatternPass As New SiteBoolean
    
    For Each pat In patset
        If ApplyVoltageFromBinCut <> "" Then
'           'T-Col TTR purpose for the scenario w/o selsrm pattern, 20230531
'           If LCase(Pat) Like "*_pl??_*" Then TheHdw.DCVS.Pins(Join(instrumentUtility.GetDCVSPinsFromCorePower, ",")).Voltage.Output = tlDCVSVoltageAlt
           
           For i = 0 To UBound(selsramLogicPingroup)
               If UCase(selsramLogicPingroup(i)) <> "PRESERVED" And UCase(selsramLogicPingroup(i)) <> "RESERVED" Then
                   If (TheHdw.DCVS.pins(selsramLogicPingroup(i)).Voltage.Output = tlDCVSVoltageAlt) Then
                       inst_info.currentDcvsOutput = tlDCVSVoltageAlt
                       Exit For
                   End If
               End If
           Next i
           
           'SycnUp
           If Flag_SyncUp_DCVS_Output_enable Then
               Call SyncUp_DCVS_Output(inst_info.p_mode, inst_info.currentDcvsOutput, SyncUp_PowerPin_Group) '''This is to sync up logic powers and sram powers on the same DCVS output (for TD testing)
           End If
        End If
        
        
        If TheExec.TesterMode = testModeOffline Then
            Call ATPG_offline(CStr(pat), ResultMode)
        Else
            If gl_bTTRDisableAlarm = False Then     'T-Col TTR approve by Si -- 230413
                TheHdw.Alarms.Check
            End If
            
            '===20240516
            ''Save masked core then we can restore it if SSN reburst happen
            If instSSNinfo.bSSNCoreHarvest = True And glb_TesterType <> "UltraFLEXplus" And SecondSSN Then
                ''Apply core mask using matched pattern, core name, fail flag
                
                For i = 0 To UBound(instSSNinfo.mappingFileName)
                    Call MaskCoreForSSN(CStr(instSSNinfo.coreName(i)), CStr(instSSNinfo.Failflag(i)), CStr(instSSNinfo.mappingFileName(i)))
                Next i
                
                '===20240516
                ''Save masked core then we can restore it if SSN reburst happen
                If instSSNinfo.bSSNCoreHarvest = True Then
                
                    ReDim instSSNinfo.maskedCores(UBound(instSSNinfo.PatName))
                    For i = 0 To UBound(instSSNinfo.mappingFileName)
                        instSSNinfo.maskedCores(i) = TheHdw.Digital.ScanNetworks(instSSNinfo.mappingFileName(i)).maskedCores
                    Next i
                    
                End If
                
            End If
            '''Check for 240702
            
            Call SSNConfigSetting(CStr(pat), instSSNinfo.bSSNTest, instSSNinfo.bSSNCoreHarvest, instSSNinfo)
            
            Call TheHdw.patterns(CStr(pat)).test(ReportResult, CLng(TL_C_YES), ResultMode, ConcurrentMode)
        End If
        
        '230711 swtich to Valt after SC Selsram pattern
'        If PrintVolatgeOutput And (UCase(CStr(Pat)) Like "*SC*") And (UCase(CStr(Pat)) Like "*_SRMDSSC*") Then
'            If TheHdw.DCVS.Pins("VDD_SOC_S1").Voltage.Output = tlDCVSVoltageMain Then
'                TheHdw.DCVS.Pins(Join(instrumentUtility.GetDCVSPinsFromCorePower, ",")).Voltage.Output = tlDCVSVoltageAlt
'                TheExec.Datalog.WriteComment "Switch to Valt after selsram pattern (by VBT)"
'                IsSwitch2Valt = True
'            Else
'                TheExec.Datalog.WriteComment "Selsram pSSattern switch to Valt"
'            End If
'        End If
        
        ''===20240516 SSN on UF(T-Har
        If glb_TesterType = "UltraFLEXplus" Then
            SCAN_Site_Blooean = TheHdw.Digital.Patgen.PatternBurstPassedPerSite ' CJR: 4/19/24 It's not a logical AND. It needs to be check per pattern
        Else
            For Each site In TheExec.sites
                SCAN_Site_Blooean(site) = TheHdw.Digital.Patgen.PatternBurstPassedPerSite(site)
            Next site
        End If

        Call SSNDataAnalysis(CStr(pat), instSSNinfo, SCAN_Site_Blooean, SecondSSN)

    Next pat
    
    Exit Function
                                                                                                                                                                                                                                                               
errHandler:
    TheExec.Datalog.WriteComment "<Error> " + funcName + ":: please check it out."
    If AbortTest Then Exit Function Else Resume Next
End Function

'Function ReadCMEM(pinlist As String, pattern As String) As Long
'    Dim captureData() As Long
'    Dim currentSite As Long
'    Dim i As Long
'    Dim result As Boolean
''
'logdssccapture = FreeFile
'Open "logdssc.out" For Output As #logdssccapture
'Print #logdssccapture, " ************DSSC Capture data************ "
'
'    For currentSite = 0 To 3
'   ' pinlist = "IO1,IO2" -->pinlist format
'   'pattern = "C:\Users\thirupah\Desktop\DSSC_SignalSheet_Issue\DSSC_SignalSheet_Issue\DSSC_solution.PAT"   -->pattern
'
'      TheHdw.DSSC.Raw.Pins(pinlist).pattern(pattern).Capture.site(currentSite).GetCaptureData "", captureData
'
'     For i = LBound(captureData) To UBound(captureData)
'        Print #logdssccapture, "addr " & i & " captureData = " & captureData(i); " currentSite =" & currentSite
'     Next i
'
'    Next currentSite
'    Close #logdssccapture
'    Exit Function
'errHandler:
'    Stop
'    Resume Next
'    End Function
