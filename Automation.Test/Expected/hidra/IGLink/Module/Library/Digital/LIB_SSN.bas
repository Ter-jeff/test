Attribute VB_Name = "Lib_SSN"
Option Explicit
Public Const glb_SSN_Enabled = True
Public ssnPatternsDict As New Dictionary        ''Save SSN patterns into this dict during first run
Public Const glb_SSN_Failflag = "F_SSN_INIT_Fail"
Public Const glb_SSN_CaptureLimit = 16777216

Private CorePFResults As PinListData
Private CorePF As SiteBoolean
Private CoreScanPins As String
Public DebugMsgEnable As Boolean
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
    theexec.Datalog.WriteComment "<Error> " + funcName + ":: please check it out."
    If AbortTest Then Exit Function Else Resume Next
End Function

'Check if SSN is enabled from harvest fail flag and pattern keyword for this instance
Public Function CheckInstForSSN(PatArray() As String, inData As Inst_SSN, SSNMapping() As ssnMappingType, _
                                Optional strHarvFlag As String = vbNullString, Optional isCharInst As Boolean = False)
On Error GoTo errHandler
Dim funcName As String:: funcName = "CheckInstForSSN"
    Dim i As Long
    Dim Pat As Variant
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
    Pat = vbNullString
    isSSNHarvestInst = False
    isSSNCoreMaskInst = False
    countMatchedSSN = 0
    
    'To avoid pin group harvest
    If theexec.Flow.IsCharacterizing = True Then
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
    For Each Pat In PatArray
        
        If ssnPatternsDict.Exists(UCase(Pat)) = True Then
            For i = 0 To UBound(SSNMapping)
                If UCase(Pat) Like UCase(SSNMapping(i).patternKeyword) Then
                    inData.bSSNTest = True
                    
                    ReDim Preserve inData.PatName(countMatchedSSN)
                    ReDim Preserve inData.patkeyword(countMatchedSSN) 'CJR: 4/16/24
                    ReDim Preserve inData.coreName(countMatchedSSN) 'CJR: 4/20/24
                    ReDim Preserve inData.Failflag(countMatchedSSN) 'CJR: 4/20/24
                    
                    inData.coreName(countMatchedSSN) = SSNMapping(i).coreName  ' CJR: 4/16/24
                    inData.Failflag(countMatchedSSN) = SSNMapping(i).Failflag  ' CJR: 4/16/24
                    inData.PatName(countMatchedSSN) = Pat
                    inData.patkeyword(countMatchedSSN) = SSNMapping(i).patternKeyword 'CJR: 4/16/24
                    countMatchedSSN = countMatchedSSN + 1
                    
                    'Exit For
                End If
            Next i
        End If
        ''Found matched pattern keyword, break
        'If inData.bSSNTest = True Then Exit For
    Next Pat
    
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
        patternPath = theexec.DataManager.Raw.GetPatternsInSet(inData.PatName(i), patternCount)
        inData.mappingFileName(i) = Split(UCase(patternPath(0)), ".PAT")(0)
    Next i
    
    ''Apply core mask using matched pattern, core name, fail flag
    If inData.bSSNCoreMask = True Then
        For i = 0 To UBound(inData.mappingFileName)
            Call MaskCoreForSSN(CStr(inData.coreName(i)), CStr(inData.Failflag(i)), CStr(inData.mappingFileName(i)))
        Next i
    End If
    
    ''Save masked core then we can restore it if SSN reburst happen
    If inData.bSSNCoreHarvest = True Then
        ReDim inData.maskedCores(UBound(inData.PatName))
        For i = 0 To UBound(inData.mappingFileName)
            inData.maskedCores(i) = TheHdw.Digital.ScanNetworks(inData.mappingFileName(i)).maskedCores
        Next i
    End If
    
Exit Function
errHandler:
    theexec.Datalog.WriteComment "Error encountered in VBT Function of CheckInstForSSN"
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
    theexec.Datalog.WriteComment "<Error> " + funcName + ":: please check it out."
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
    theexec.Datalog.WriteComment "<Error> " + funcName + ":: please check it out."
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
    theexec.Datalog.WriteComment "<Error> " + funcName + ":: please check it out."
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
        theexec.Datalog.WriteComment mappingfile & ", it doesn't exist. Error!!!"
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
        SSNMapping.PinCnt = UBound(Split(SSNMapping.PinData, ",")) + 1
        SSNMapping.GenPacketMap (SSNMapping.PinCnt)
        SSNMapping.GetBusRepeatMap (SSNMapping.PinCnt)
        Set ReadMappingFile = SSNMapping
    End If
Exit Function
errHandler:
    theexec.Datalog.WriteComment "<Error> " + funcName + ":: please check it out."
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

    theexec.Datalog.Setup.ScanSetup.EnableScanLogging = True
    theexec.Datalog.Setup.ScanSetup.CaptureFormat = logFormat
    theexec.Datalog.Setup.ScanSetup.ResultMode = logResultMode
    'Need GPIO pins from ssn mapping table
    theexec.Datalog.Setup.ScanSetup.PinList = logPins
    theexec.Datalog.Setup.ScanSetup.CMEMCaptureLimit = scanFailsPerPin

    theexec.Datalog.ApplySetup
Exit Sub
errHandler:
    theexec.Datalog.WriteComment "<Error> " + funcName + ":: please check it out."
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
    TheHdw.Patterns(targetPattern).Reload
    
    '___ Get the individual target pins to change _____________________________________
    Dim targetInPins() As String
    Dim targetPinCnt As Long
    Call theexec.DataManager.DecomposePinList(scanInPins, targetInPins, targetPinCnt)
    
    '___ Make N random fails per site __________________________________________________
    Dim origData As Integer
    
Exit Function
errHandler:
    theexec.Datalog.WriteComment "<Error> " + funcName + ":: please check it out."
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
    theexec.Datalog.WriteComment "<Error> " + funcName + ":: please check it out."
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
                            PostTestFArgs As String) As IScanNetworkResults
On Error GoTo errHandler
Dim funcName As String:: funcName = "ssnReburst"
    Dim errMsg As String
    
    '___ Init __________________________________________________________________
     On Error GoTo errHandler
    
    '___ Make sure reburst is needed ___________________________________________
    If (theexec.sites.Selected.Count = 0) Then
        theexec.Flow.TestLimitIndex = theexec.Flow.TestLimitIndex + 2
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
    
    While ((theexec.sites.Selected.Count > 0) And (reburstCnt <= maxReburst))
    
        If (theexec.enableWord("ENABLE_DEMO_STOP") = True) Then
            '************************************************************************************
            '*** This is the start of the Reburst loop                                        ***
            '***    - Previous burst failing cores will be applied as a mask for next reburst ***
            '************************************************************************************
            Stop
        End If
            
        '___ Apply a mask of known failing cores _______________________________
        With TheHdw.Digital.ScanNetworks(ssnName).CoreMasks
            Call .AddPerSite(newCoreMask)
            Call .Apply
        End With
        
        '___ Update the masking info in the COND_LIST __________________________
        SSNUpdateCondLst ssnName, siteMasksByCore
                
        '___ Do the rebusrt ____________________________________________________
'        theexec.Flow.TestLimit resultVal:=reburstCnt, _
'                               scaletype:=scaleNoScaling, formatStr:="%d", _
'                               PinName:=CStr(reburstCnt), _
'                               ForceResults:=tlForceFlow
        
        If (logScanFails = True) Then
            Call TheHdw.Patterns(PatName).test(pfAlways, 0, tlResultModeModule)
            
        Else
            TheHdw.Patterns(PatName).start
            TheHdw.Patterns(PatName).HaltWait
            
            reburstPFResults = TheHdw.Digital.Patgen.PatternBurstPassedPerSite
            theexec.Flow.FunctionalTestLimit resultVal:=reburstPFResults, _
                                             PatName:=PatName, _
                                             ForceResults:=tlForceFlow
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
        If (theexec.enableWord("ENABLE_DEMO_STOP") = True) Then
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
            theexec.Flow.TestLimitIndex = theexec.Flow.TestLimitIndex - 2
            
        End If
                
        '___ Update the sites selected for reburst ____________________________
        theexec.sites.Selected = reburstSites
                    
        reburstCnt = reburstCnt + 1
    Wend
    
    '___ Adjust limits index if max reburst count hit __________________________
    If (reburstCnt >= maxReburst) Then
        theexec.Flow.TestLimitIndex = theexec.Flow.TestLimitIndex '+ 2
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
    theexec.ErrorLogMessage ("*E* ssnReburst: " + errMsg)
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
        
        For Each site In theexec.sites.Active
            If (IsEmpty(TheHdw.Digital.ScanNetworks(ssnName).maskedCores(site)) = False) Then
                maskedCores = TheHdw.Digital.ScanNetworks(ssnName).maskedCores(site)
                For Each coreName In maskedCores
                    siteMasksByCore(coreName)(site) = True
                Next coreName
            End If
        Next site
        
        Dim condValue As String
        With theexec.Datalog.Setup.ScanSetup.Conditions
            .RemoveAll
            For Each coreName In siteMasksByCore.keys
                condValue = ""
                For Each site In theexec.sites.Active
                    If (siteMasksByCore(coreName)(site) = True) Then
                        condValue = condValue + " " + CStr(site)
                    End If
                Next site
                If (Len(condValue) > 0) Then
                    .Add ("SSN_CORE_MASK @ " + coreName), condValue
                End If
            Next coreName
        End With
Exit Function
errHandler:
    theexec.Datalog.WriteComment "<Error> " + funcName + ":: please check it out."
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
        currentResults = CorePFResults.Pins(coreName)
        newResults = ssnResults.core(coreName).Failed
        CorePFResults.Pins(coreName).value = currentResults.LogicalAnd(newResults.LogicalNot)
    Next coreName
    
    '___ Check for fails that only occur outside the SSN scan blocks ______________________
    Dim burstPF As New SiteBoolean
    Dim forceFailSites As New SiteBoolean
    Dim site As Variant
    
    burstPF = TheHdw.Digital.Patgen.PatternBurstPassedPerSite.LogicalNot
    forceFailSites = burstPF.LogicalAnd(anyCoreFailsBySite.LogicalNot)
        
    If (forceFailSites.Any(True) = True) Then
        For Each site In theexec.sites.Selected
            If (forceFailSites(site) = True) Then
                For Each coreName In ssnResults.coreNames
                    CorePFResults.Pins(coreName).value = False
                Next coreName
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
    theexec.ErrorLogMessage ("*E* ProcessSSNResults: " + errMsg)
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
        If (EnableScanLogging = True) Then
            Call SetupForScanLogging(scanLogPins, _
                                     maxFailsPerPin, _
                                     tl_DCScanCaptureFormat_Cycle, _
                                     scanLogMode)
                                     
            
            '___ Set COND_LIST _________________________________________________
            With theexec.Datalog.Setup.ScanSetup.Conditions
                .RemoveAll
                .Add "SSN_MASKING", "False"
            End With
            
        End If
        
        '___ Setup CMEM ____________________________________________________________
        Call TheHdw.Digital.CMEM.SetCaptureConfig(maxFailsPerPin, CmemCaptFail, _
                                                  tlCMEMCaptureSourcePassFailData, _
                                                  enableFullSpeedMode)
                                                  
        TheHdw.Digital.CMEM.CaptureLimitMode = tlDigitalCMEMCaptureLimitMode_Enable
        TheHdw.Digital.CMEM.CaptureLimit = maxFailsPerPin
        TheHdw.Digital.Patgen.ScanBurstEnabled = True

                                                      
        '___ Build the per core PF PLD _________________________________________
        '    Core names are dummy pins in PLD
        '    Init all cores to pass
        '-----------------------------------------------------------------------
        Dim coreNames() As String
        coreNames = TheHdw.Digital.ScanNetworks(ssnMapName).coreNames
        QuickSort coreNames, 0, UBound(coreNames)
                    
        Set CorePFResults = New PinListData
        
        Set CorePF = New SiteBoolean
        CorePF = True
        
        Dim coreName As Variant
        For Each coreName In coreNames
            CorePFResults.AddPin coreName
            CorePFResults.Pins(coreName).value = CorePF
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
    theexec.Datalog.WriteComment "<Error> " + funcName + ":: please check it out."
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
    Dim site As Variant
    
    On Error GoTo errHandler
    '~~~ START OF BODY ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
    '~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
    ''If Step_ = subAllBody Or Step_ = subBody Then
    
        '___ Call StartOfBody interpose if defined _____________________________
        ''Call Interpose(StartOfBodyF, StartOfBodyFArgs)
        
        '___ Cache the sites selected for start of test ________________________
        Dim startOfTestSites As New SiteBoolean
        startOfTestSites = theexec.sites.Selected
        'If startOfTestSites.All(False) Then Exit Sub 'pg 1007 not sure why erring now, add this
        
        '___ 'test' the pattern ________________________________________________
        'Call TheHdw.Patterns(ShortPatName).Test(pfAlways, 0, tlResultModeModule)
        
        ''Call Interpose(PostTestF, PostTestFArgs)
        ''LogPinCoreFailsBySite ssnResults, CoreScanPins
        
        '___ Process the initial SSN results / determine reburst sites _________
        Dim reburstSites As New SiteBoolean
        Dim ssnResults As IScanNetworkResults
        
        Set ssnResults = TheHdw.Digital.Patgen.ReadScanNetworkResults
        reburstSites = ProcessSSNResults(ssnResults)
        Set ssnResultsGlobal = ssnResults 'TheHdw.Digital.Patgen.ReadScanNetworkResults
        
        LogPinCoreFailsBySite ssnResults, CoreScanPins
        
        If (theexec.enableWord("ENABLE_DEMO_STOP") = True) Then
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
                origSelectedSites = theexec.sites.Selected
                theexec.sites.Selected = reburstSites
                
                '___ Setup Rebusrt scan logging ________________________________
                If (enableReburstScanLogging = False) Then
                    theexec.Datalog.Setup.ScanSetup.EnableScanLogging = False
                    theexec.Datalog.ApplySetup
                End If

                '___ Do the reburst ____________________________________________
                Set ssnResults = ssnReburst(shortPatName, ssnResults, CorePFResults, _
                                            enableReburstScanLogging, _
                                            PreTestF, PreTestFArgs, _
                                            PostTestF, PostTestFArgs)
                                            
                '___ Re-enable original selected sites _________________________
                theexec.sites.Selected = origSelectedSites
                                
                If reburstSites.Any(True) Then
                    TheHdw.Digital.ScanNetworks(ssnResults.name).CoreMasks.Clear
                    theexec.Datalog.WriteComment "Clear SSN core mask after reburst" ', SSN=" & ssnResults.Name

                    ''If there is masked cores from previous instance exist, restore masked cores
                    For Each site In theexec.sites
                        If Not IsEmpty(instSSNinfo.maskedCores(instSSNinfo.idxMatchedPattern)(site)) Then
                            Dim strCore As Variant
                            For Each strCore In instSSNinfo.maskedCores(instSSNinfo.idxMatchedPattern)(site)
                                TheHdw.Digital.ScanNetworks(ssnResults.name).CoreMasks.Add strCore
                                theexec.Datalog.WriteComment "site:" & site & ", restore original core mask=" & strCore
                            Next strCore
                            TheHdw.Digital.ScanNetworks(ssnResults.name).CoreMasks.Apply ' CJ: 3/6/24: Missing CoreMask Apply
                        End If
                    Next site
                End If
            Else
            
                '___ No reburst - skip entries in Limits Set _______________________
'                TheExec.Flow.TestLimitIndex = TheExec.Flow.TestLimitIndex + 2
                theexec.Flow.TestLimitIndex = theexec.Flow.TestLimitIndex 'pg+ 2
            
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
                    
                    currentResults = CorePFResults.Pins(coreName)
                    newResults = ssnResults.ReburstNeeded
                    CorePFResults.Pins(coreName).value = currentResults.LogicalAnd(newResults.LogicalNot)
                    
                End If
                
            Next coreName
                
            '___ No reburst - skip entries in Limits Set _______________________
'pg             TheExec.Flow.TestLimitIndex = TheExec.Flow.TestLimitIndex + 2
             theexec.Flow.TestLimitIndex = theexec.Flow.TestLimitIndex 'pg+ 2


                        
        End If  '<- Reburst enabled
   
        '___ Datalog per core P/F results ______________________________________
       '___ Datalog per core P/F results ______________________________________
        Dim printPF As New SiteLong
        Dim core As Variant
        Dim CoreArray As Variant
        Dim FlagArray As Variant
        Dim i As Long, j As Long
        For j = 0 To UBound(instSSNinfo.patkeyword) 'CJR:4/20/24
            If UCase(shortPatName) Like instSSNinfo.patkeyword(j) Then
                For Each core In CorePFResults.Pins
                    CorePF = CorePFResults.Pins(core)
                    printPF = CorePFResults.Pins(core)
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
                    
                    For Each site In theexec.sites.Selected
                        If CorePF(site) = False Then
                            theexec.Flow.TestLimit resultVal:=printPF(site), lowVal:=0, hiVal:=0, Tname:=core
                            For i = 0 To UBound(CoreArray)
                                If LCase(core) = LCase(CoreArray(i)) Then
        '                        If LCase(core) Like LCase("*" & CoreArray(i) & "*") Then
                                    theexec.sites(site).FlagState(FlagArray(i)) = logicTrue
                                    Exit For
                                End If
                            Next i
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
    theexec.Datalog.WriteComment "<Error> " + funcName + ":: please check it out."
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
        
        Call TheHdw.Digital.CMEM.SetCaptureConfig(-1, CmemCaptNone, _
                                                  tlCMEMCaptureSourcePassFailData, _
                                                  enableFullSpeedMode)
                                                      
        '___ Set COND_LIST ____________________________________________________
        With theexec.Datalog.Setup.ScanSetup.Conditions
            .RemoveAll
            .Add "SSN_MASKING", "False"
        End With
        
        '___ Release the global core PLD ______________________________________
        Set CorePFResults = Nothing
     
     ''End If
    '~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
    '~~~ END OF POST-BODY ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
    '~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
Exit Sub
errHandler:
    theexec.Datalog.WriteComment "<Error> " + funcName + ":: please check it out."
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
    period = theexec.Specs.AC("t_period").ContextValue
    
    Dim site As Variant
    For Each site In theexec.sites.Starting
        theexec.Datalog.WriteComment vbCrLf + "*I* Digital Period = " + Format(period(site) / 0.000000001, "##0.000 nsec") + vbCrLf
        Exit For
    Next site
Exit Function
errHandler:
    theexec.Datalog.WriteComment "<Error> " + funcName + ":: please check it out."
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
    
    theexec.Datalog.WriteComment vbCrLf + _
                                 "*I* Max Fails Per Pin: " + CStr(theexec.Datalog.Setup.ScanSetup.CMEMCaptureLimit)
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

    msg = vbCrLf + "*I* Failing Cores Per Site: " + theexec.DataManager.instancename + vbCrLf
    For Each site In theexec.sites.Selected
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
        If (ssnResults.ReburstNeeded(site) = False) Then
            msg = msg + "  All fails captured - No Reburst Required" + vbCrLf
        Else
            msg = msg + "  All fails not captured - Reburst Required" + vbCrLf
        End If
        
    Next site
    theexec.Datalog.WriteComment msg

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
    theexec.Datalog.WriteComment "<Error> " + funcName + ":: please check it out."
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
            Dim coreIdx As Long
            
            Fields = Split(argv(argIdx), ":")
            
            maskedCoreNames = Split(Fields(1), "+")
            
            For coreIdx = 0 To UBound(maskedCoreNames)
                maskedCoreNames(coreIdx) = Trim(maskedCoreNames(coreIdx))
            Next coreIdx
            
            '___ Build the mask ________________________________________________
            Dim newCoreMask As New SiteVariant
            Dim site As Variant
            
            If (Trim(Fields(0)) = "EVEN") Then
                '___ Mask applied to all even sites ___________________________
                For Each site In theexec.sites.Selected
                    If ((site Mod 2) = 0) Then
                        newCoreMask = maskedCoreNames
                    End If
                Next site
                
            
            ElseIf (Trim(Fields(0)) = "ODD") Then
                '___ Mask applied to all odd sites ____________________________
                For Each site In theexec.sites.Selected
                    If ((site Mod 2) <> 0) Then
                        newCoreMask = maskedCoreNames
                    End If
                Next site
            
            Else
                '___ Mask applied to specific _________________________________
                Dim siteNum As Long
                siteNum = CLng(Trim(Fields(0)))
                
                For Each site In theexec.sites.Selected
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
    theexec.Datalog.WriteComment "<Error> " + funcName + ":: please check it out."
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
    theexec.Datalog.WriteComment "<Error> " + funcName + ":: please check it out."
    If AbortTest Then Exit Sub Else Resume Next
End Sub

Public Function SPN(ByVal Pat As String) As String
On Error GoTo errHandler
Dim funcName As String:: funcName = "SPN"
    Dim shortPatName As String
    Dim ShortPos As Long: ShortPos = 1
    While InStr(ShortPos, Pat, "\") > 0
        ShortPos = InStr(ShortPos, Pat, "\") + 1
    Wend
    shortPatName = mid(Pat, ShortPos, InStr(Pat, ".PAT") - ShortPos)
    SPN = shortPatName
Exit Function
errHandler:
    theexec.Datalog.WriteComment "<Error> " + funcName + ":: please check it out."
    If AbortTest Then Exit Function Else Resume Next
End Function
'Find patterns contains "SSC" and "SSU"
Public Function FindSSNPattern(PatName As String)
On Error GoTo errHandler
Dim funcName As String:: funcName = "FindSSNPattern"

If UCase(PatName) Like "*_PL*_SSC_*" Or UCase(PatName) Like "*_PL*_SSU_*" Then
    PatName = Split(PatName, ":")(0)
    If ssnPatternsDict.Exists(UCase(PatName)) = False Then
        ssnPatternsDict.Add UCase(PatName), ""
    End If
End If

Exit Function
errHandler:
    theexec.Datalog.WriteComment "<Error> " + funcName + ":: please check it out."
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
For Each site In theexec.sites
    With TheHdw.Digital.ScanNetworks(ssnMappingName).CoreMasks      ''Add core mask by site
        For i = 0 To UBound(failFlagArr)
            If theexec.sites(site).FlagState(failFlagArr(i)) = logicTrue Then
                
                Call .Add(coreNameArr(i))       ''Add core mask by core
                theexec.Flow.TestLimit resultVal:=1, lowVal:=1, hiVal:=1, Tname:=coreNameArr(i) & "_masked on SSN:" & ssnMappingName, ForceResults:=tlForceNone
                If isAnyCoreMasked = False Then
                    isAnyCoreMasked = True
                End If
            End If
        Next i
        
        Call .Apply     ''Apply core mask by site
    End With
Next site

Exit Function
errHandler:
    theexec.Datalog.WriteComment "<Error> " + funcName + ":: please check it out."
    If AbortTest Then Exit Function Else Resume Next
End Function

Public Function SSNConfigSetting(Pat As String, bSSNTest As Boolean, bSSNCoreHarvest As Boolean, instSSNinfo As Inst_SSN)
    On Error GoTo errHandler
    Dim funcName As String:: funcName = "SSNConfigSetting"
    Dim shortPatName As String
    Dim i As Integer
    'If bSSNTest = True And LCase(Pat) Like "*_pl*" Then     '20231120: Added pl only
    For i = 0 To UBound(instSSNinfo.patkeyword)
        If bSSNTest = True And LCase(Pat) Like "*_pl*" And UCase(Pat) Like instSSNinfo.patkeyword(i) Then     'CJR: 4/18/24
        shortPatName = SPN(UCase(Pat))     ''Split pattern name with path
        If (theexec.Flow.IsCharacterizing = False) Then
            If bSSNCoreHarvest = True Then
                Dim SSN_ScanPins As String: SSN_ScanPins = ""
                Dim maxFailsPerPin As Long: maxFailsPerPin = glb_SSN_CaptureLimit
'                                                If LCase(TheExec.DataManager.instanceName) Like "*_final*" Then
                SSN_PreBody True, SSN_ScanPins, maxFailsPerPin, tlDatalogScanResultMode_Module, True, shortPatName
'                                                Else
'                                                    Stop
'                                                    'SSN_PreBody True, SSN_ScanPins, maxFailsPerPin, tlDatalogScanResultMode_Module, True, SPN(pat) & ".SSNMapping"
'                                                End If
            End If
        Else
            ' Running Char flow
        End If
    End If
    Next i
                                    
    Exit Function
errHandler:
    theexec.Datalog.WriteComment "<Error> " + funcName + ":: please check it out."
    If AbortTest Then Exit Function Else Resume Next
End Function

Public Function SSNDataAnalysis(Pat As String, instSSNinfo As Inst_SSN, SCAN_Site_Blooean As SiteBoolean)
    On Error GoTo errHandler
    Dim funcName As String:: funcName = "SSNConfigSetting"
    Dim site As Variant
    Dim shortPatName As String
    Dim i As Integer
    
    Dim isPatternSSN As Boolean: isPatternSSN = False
    If instSSNinfo.bSSNTest = True And instSSNinfo.bSSNCoreHarvest = True Then 'And SCAN_Site_Blooean.Any(False) Then
        For i = 0 To UBound(instSSNinfo.patkeyword)
             If LCase(Pat) Like "*_pl*" And UCase(Pat) Like instSSNinfo.patkeyword(i) Then          'CJR: 4/15/24: SSC AND PRC in same Patset 4/15/24 Then
                If (theexec.Flow.IsCharacterizing = False) Then
                    Dim SSN_InterposeFunc As InterposeName ': SSN_InterposeFunc = "LogPinCoreFailsBySite"
                    shortPatName = SPN(UCase(Pat))
                    SSN_Body instSSNinfo, True, True, shortPatName, , , SSN_InterposeFunc  ', SSN_ScanPins
                    SSN_PostBody False
                Else
                    '___ Clear SSN masks __________________________________________________
                    'theHdw.Digital.ScanNetworks.ClearAllMasks
                End If
                
                isPatternSSN = True
                
                Exit For        '' 20240502: Added to stop finding new pattern keyword
'            Else 'If LCase(Pat) Like "*_in*" Then     ''Init SSN fail. ''CJR: 4/15/24: Modified this condition for scenario#2 & 3. This needs update!
'                ''20240502: Removed "*_in*" keyword, to binout "non-SSN pl pattern fail" in SSN instance
'                For Each site In TheExec.sites
'                    If SCAN_Site_Blooean(site) = False Then
'                        TheExec.sites(site).FlagState(glb_SSN_Failflag) = logicTrue
'                        TheExec.Flow.TestLimit resultVal:=1, lowVal:=0, hiVal:=0, Tname:=CStr(glb_SSN_Failflag), ForceResults:=tlForceNone
'                        TheExec.Datalog.WriteComment "Non-'SSN payload pattern' fails in SSN instance!"     '' 20240502: Added to clarify
'                    End If
'                Next site
'
'                Exit For        '' 20240502: Added to stop finding new pattern keyword
            End If
        Next i
        
        If isPatternSSN = False Then            ''Added this boolean to avoid 2nd SSN PL being recognized as init pattern
            For Each site In theexec.sites
                If SCAN_Site_Blooean(site) = False Then
                    theexec.sites(site).FlagState(glb_SSN_Failflag) = logicTrue
                    theexec.Flow.TestLimit resultVal:=1, lowVal:=0, hiVal:=0, Tname:=CStr(glb_SSN_Failflag), ForceResults:=tlForceNone
                    theexec.Datalog.WriteComment "Non-'SSN payload pattern' fails in SSN instance!"     '' 20240502: Added to clarify
                End If
            Next site
        End If
        
    End If

    Exit Function
errHandler:
    theexec.Datalog.WriteComment "<Error> " + funcName + ":: please check it out."
    If AbortTest Then Exit Function Else Resume Next
End Function

Public Function SSNMainProcedure(patset() As String, ReportResult As PFType, TL_C_YES As Long, ResultMode As tlResultMode, _
            ConcurrentMode As tlPatConcurrentMode, ByRef SCAN_Site_Blooean As SiteBoolean, instSSNinfo As Inst_SSN, Optional ApplyVoltageFromBinCut As String = vbNullString, _
            Optional Harv_FailFlag As String = vbNullString, Optional HarvestPinGrpOtherFail As String = vbNullString)
    
    On Error GoTo errHandler
    Dim funcName As String:: funcName = "SSNMainProcedure"
    Dim Pat As Variant
    Dim inst_info As Instance_Info
'    Dim instrumentUtility As New Instrument_Utility
    Dim i As Long
    
    Dim sBool_PatternPass As New SiteBoolean
        
    For Each Pat In patset
        If ApplyVoltageFromBinCut <> "" Then
'           'T-Col TTR purpose for the scenario w/o selsrm pattern, 20230531
'           If LCase(Pat) Like "*_pl??_*" Then TheHdw.DCVS.Pins(Join(instrumentUtility.GetDCVSPinsFromCorePower, ",")).Voltage.Output = tlDCVSVoltageAlt
           
           For i = 0 To UBound(selsramLogicPingroup)
               If UCase(selsramLogicPingroup(i)) <> "PRESERVED" And UCase(selsramLogicPingroup(i)) <> "RESERVED" Then
                   If (TheHdw.DCVS.Pins(selsramLogicPingroup(i)).Voltage.Output = tlDCVSVoltageAlt) Then
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

        If theexec.TesterMode = testModeOffline Then
            Call ATPG_offline(CStr(Pat), ResultMode)
        Else
            If gl_bTTRDisableAlarm = False Then     'T-Col TTR approve by Si -- 230413
                TheHdw.Alarms.Check
            End If
            
            Call SSNConfigSetting(CStr(Pat), instSSNinfo.bSSNTest, instSSNinfo.bSSNCoreHarvest, instSSNinfo)
        
            Call TheHdw.Patterns(CStr(Pat)).test(ReportResult, CLng(TL_C_YES), ResultMode, ConcurrentMode)
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
        
        'SCAN_Site_Blooean = SCAN_Site_Blooean.LogicalAnd(thehdw.Digital.Patgen.PatternBurstPassedPerSite)
        SCAN_Site_Blooean = TheHdw.Digital.Patgen.PatternBurstPassedPerSite ' CJR: 4/19/24 It's not a logical AND. It needs to be check per pattern

        Call SSNDataAnalysis(CStr(Pat), instSSNinfo, SCAN_Site_Blooean)

    Next Pat

    Exit Function
                                                                                                                                                                                                                                                               
errHandler:
    theexec.Datalog.WriteComment "<Error> " + funcName + ":: please check it out."
    If AbortTest Then Exit Function Else Resume Next
End Function

