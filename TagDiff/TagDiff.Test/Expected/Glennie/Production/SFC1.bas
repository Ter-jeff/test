Attribute VB_Name = "SFC1"
Public ScanFail_ReadHRAM_Count As Long


'Private Const SFC1_Enable As String = "A_SFC1" ' enable word

'Private Const SFC1_ReadInput As String = "A_SFC1_Input" ' enable word
'
'Private Const SFC1PatternsSheet As String = "SFC1Patterns"
'
'Private Is_SFC1_Enabled As Boolean

Private rowCount As Long



Public Type TableScanPatterns

    instancename As String

    FailCycleCount As Integer

    ScanPatternName As String

End Type



Private ScanPatternsList() As TableScanPatterns



Public Function ScanFail_ReadHRAM(pmCnt As Long, pmNames() As String)

    If TheExec.enableWord("A_FFC1") = False Then Exit Function

    Dim hramsize As Long: hramsize = TheHdw.Digital.hram.size

    TheExec.Datalog.WriteComment "ScanFail_ReadHRAM Size = " & CStr(hramsize)

    

: ScanFail_ReadHRAM_Count = 2

    Dim site As Variant

    Dim CapCnt As Long, i As Long, VecCnt As Long, VecCycle As Long, VecCyclePrev As Long

'    Dim ModName As String, ModNamePrev As String

    Dim InsName As String, VecStep As Long, failpins() As String

    Dim ModName As String, ModNamePrev As String

    If pmCnt > 0 Then

        'pmNames(0)

                ModName = TheHdw.Digital.hram.PatGenInfo(i, pgPattern)

        If pmNames(0) <> ModName Then Exit Function

    End If

    InsName = TheExec.DataManager.instancename

    For Each site In TheExec.sites

        CapCnt = TheHdw.Digital.hram.CapturedCycles

                    VecCnt = 0: VecCyclePrev = -1

        If CapCnt > 0 Then

            failpins = TheHdw.Digital.FailedPins(site) ''' get fail pin

            Dim FailedPins As String, idx As Long

            'For idx = 0 To UBound(FailPins)

                If UBound(failpins) >= 0 Then FailedPins = Join(failpins, "/")

                'Else: FailedPins = FailPins(0)

            'Next idx

            For i = 0 To CapCnt - 1

                ModName = TheHdw.Digital.hram.PatGenInfo(i, pgPattern)

                If ModName <> ModNamePrev Then

                    ModNamePrev = ModName

                    VecCnt = 0

                End If

                If VecCnt < ScanFail_ReadHRAM_Count Then

                VecCycle = TheHdw.Digital.hram.PatGenInfo(i, pgCycle)

                VecStep = TheHdw.Digital.hram.PatGenInfo(i, pgVector)

                'If VecCycle <> VecCyclePrev Then

                'TheExec.Datalog.WriteComment "Site[" & CStr(Site) & "] " & ModName & " Cycle = " & VecCycle

                'theexec.Datalog.WriteComment "FFC1,1," & CStr(Site) & "," & InsName & "," & ModName & "," & VecStep & "," & VecCycle & "," & FailedPins

                Dim FlowName As String: FlowName = TheExec.flow.CurrentFlowSheetName

                TheExec.Datalog.WriteComment "SFC1,1," & CStr(site) & "," & mid(FlowName, 6, Len(FlowName)) & "," & InsName & "," & ModName & "," & VecStep & "," & VecCycle & "," & FailedPins

                VecCnt = VecCnt + 1: VecCyclePrev = VecCycle

                'End If

                End If

            Next i

        End If

    Next site

    

'    Get_CMEM_Capture "ScanFail_ReadHRAM"

End Function



Public Sub Get_CMEM_Capture(Optional argF As String = "Debug")

    Dim CMEMCaptureSize As Long

    Dim CMEMCaptureType As Long

    TheHdw.Digital.CMEM.GetCaptureConfig CMEMCaptureSize, CMEMCaptureType

    

'    TheExec.Datalog.WriteComment argF & " >>> CMEMCaptureSize: " & CStr(CMEMCaptureSize)

'    Debug.Print argF & " >>> CMEMCaptureSize: " & CStr(CMEMCaptureSize)

'    Debug.Print argF & " >>> CMEMCaptureType: " & CMEMCaptureType



End Sub



Public Function ifcmemSetCapture(argc As Long, argv() As String)

    If glb_isSFC_Enabled = False Then Exit Function

    ' Set up CMEM to capture cycle count.

'    TheHdw.Digital.hram.CaptureType = CaptType.captNone

    TheHdw.Digital.CMEM.CentralFields = tlCMEMCaptureFields.tlCMEMPatternName + tlCMEMCaptureFields.tlCMEMAbsoluteCycle + tlCMEMCaptureFields.tlCMEMVMVectorOffset

'    TheHdw.Digital.CMEM.SetCaptureConfig 1024, CmemCaptFail 'CmemCaptType.CmemCaptFail

    TheHdw.Digital.CMEM.SetCaptureConfig 5, CmemCaptType.CmemCaptFail, tlCMEMCaptureSource.tlCMEMCaptureSource_PatPassFailData 'tlCMEMCaptureSource_PatPassFailData 'CmemCaptType.CmemCaptFail

'    TheHdw.Digital.CMEM.SetCaptureConfig 0, CmemCaptNone 'CmemCaptType.CmemCaptFail



    Harvest_CMEM_InitSetup

        

''    Get_CMEM_Capture "PreBody"

End Function



Public Function ifcmemSetCaptureOff(argc As Long, argv() As String)

    If glb_isSFC_Enabled = False Then Exit Function

    ' Set up CMEM to capture cycle count.

'    TheHdw.Digital.CMEM.CentralFields = tlCMEMCaptureFields.tlCMEMPatternName + tlCMEMCaptureFields.tlCMEMAbsoluteCycle

'    TheHdw.Digital.CMEM.SetCaptureConfig 1024, CmemCaptFail 'CmemCaptType.CmemCaptFail

    TheHdw.Digital.CMEM.SetCaptureConfig 0, CmemCaptNone 'CmemCaptType.CmemCaptFail

    TheHdw.Digital.CMEM.CentralFields = tlCMEMCaptureFields.tlCMEMNone

    

     ''''' New request for pattern pin group '''''

    If glb_TesterType = "UltraFLEXplus" Then

        TheHdw.Digital.Patgen.ScanBurstEnabled = False

''        TheExec.Datalog.Setup.ScanSetup.EnableScanLogging = False

    End If

    

End Function



Public Function ifcmemGetFailCycles(pmCnt As Long, pmNames() As String)

    If glb_isSFC_Enabled = False Then Exit Function

    On Error GoTo errHandler



    ' Or if pattern site threading

    Dim slNumCycles As New SiteLong

'    TheHdw.Digital.CMEM.SetCaptureConfig -1, CmemCaptType.CmemCaptFail 'CmemCaptType.CmemCaptFail

''    slNumCycles.value = TheHdw.Digital.hram.CapturedCyclesPerSite

'    AbsoluteCycle = TheHdw.Digital.CMEM.PatGenInfo(tlCMEMCaptureFields.tlCMEMAbsoluteCycle)

'    AbsoluteCycle is empty if a site has no fails; do in site loop

: ScanFail_ReadHRAM_Count = 2

: ScanFail_ReadHRAM_Count = CLng(pmNames(0))

    Dim site As Variant, VecCycle As Long, i As Long

    Dim ModName As String, ModNamePrev As String

    Dim patIdx() As Long, patNames() As String

'    If pmCnt > 0 Then

'        Call TheHdw.Digital.CMEM.PatternName(patIdx, patNames)

'        'pmNames(0)

'                ModName = TheHdw.Digital.hram.PatGenInfo(i, pgPattern)

'        If pmNames(0) <> ModName Then Exit Function

'    End If

    InsName = TheExec.DataManager.instancename

    slNumCycles.value = TheHdw.Digital.CMEM.CapturedCyclesPerSite

    

''    Call TheHdw.Digital.CMEM.PatternName(patIdx, patNames)

    

    For Each site In TheExec.sites

'                    VecCnt = TheHdw.Digital.CMEM.CapturedCycles

'                    VecCnt = TheHdw.Digital.hram.CapturedCycles

        VecCnt = 0: VecCyclePrev = -1

        If slNumCycles > 0 Then

            Call TheHdw.Digital.CMEM.PatternName(patIdx, patNames)

            ModName = patNames(0)

'    If pmCnt > 0 Then _

'        If pmNames(0) <> ModName Then Exit Function

            If pmCnt > 1 Then If Not (ModName Like pmNames(1)) Then Exit Function

            

            failpins = TheHdw.Digital.FailedPins(site) ''' get fail pin

            Dim FailedPins As String, idx As Long

            If UBound(failpins) >= 0 Then FailedPins = Join(failpins, "/")

                'Else: FailedPins = FailPins(0)

'    RelativeCycle = TheHdw.Digital.CMEM.PatGenInfo(tlCMEMCaptureFields.tlCMEMScanIndex) ' tlCMEMVMVectorOffset

'    RelativeCycle = TheHdw.Digital.CMEM.PatGenInfo(tlCMEMCaptureFields.tlCMEMVMVectorOffset) ' tlCMEMVMVectorOffset

'    AbsoluteCycle = TheHdw.Digital.CMEM.PatGenInfo(tlCMEMCaptureFields.tlCMEMAbsoluteCycle)

'    Call TheHdw.Digital.CMEM.PatternName(patIdx, patNames)

'    ModName = patNames(0)

'    If pmCnt > 0 Then _

'        If pmNames(0) <> ModName Then Exit Function

                For i = 0 To slNumCycles - 1

'    RelativeCycle = TheHdw.Digital.CMEM.PatGenInfo(tlCMEMCaptureFields.tlCMEMScanIndex) ' tlCMEMVMVectorOffset

                    RelativeCycle = TheHdw.Digital.CMEM.PatGenInfo(tlCMEMCaptureFields.tlCMEMVMVectorOffset, ScanFail_ReadHRAM_Count) ' tlCMEMVMVectorOffset

                    AbsoluteCycle = TheHdw.Digital.CMEM.PatGenInfo(tlCMEMCaptureFields.tlCMEMAbsoluteCycle, ScanFail_ReadHRAM_Count)

'                ModName = TheHdw.Digital.CMEM.PatGenInfo(tlCMEMPatternName)

'ModName = TheHdw.Digital.CMEM.PatGenInfo(tlCMEMPatternName, 1)

''                If ModName <> ModNamePrev Then

''                    ModNamePrev = ModName

''                    VecCnt = 0

''                End If

                    If VecCnt < ScanFail_ReadHRAM_Count Then

                        VecCycle = AbsoluteCycle(i) 'TheHdw.Digital.CMEM.PatGenInfo(i, pgVector)

                        VecStep = RelativeCycle(i) 'TheHdw.Digital.CMEM.PatGenInfo(i, pgVector)

                'If VecCycle <> VecCyclePrev Then

                        Dim FlowName As String: FlowName = TheExec.flow.CurrentFlowSheetName

                        TheExec.Datalog.WriteComment "SFC1,1," & CStr(site) & "," & mid(FlowName, 6, Len(FlowName)) & "," & InsName & "," & ModName & "," & VecStep & "," & VecCycle & "," & FailedPins

'                TheExec.Datalog.WriteComment "CMEM Capture Site[" & CStr(Site) & "] " & ModName & " Fail Cycle = " & VecCycle

'                TheExec.AddOutput ("CMEM Capture Site[" & CStr(Site) & "] " & ModName & " Fail Cycle = " & VecCycle)

                        VecCnt = VecCnt + 1: VecCyclePrev = VecCycle

                'End If

                    Else

                        i = slNumCycles

                    End If

            Next i

        End If

    Next site

    Exit Function

errHandler:

    TheExec.AddOutput ("Error in ifcmemGetFailCycles function")

End Function





'''''''''''''''''''''''''''''''''
'
''''>>>>> SFC1
'
'Public Function Read_SFC1_Table() As Long
'
'    ' Read worksheet and saves info to ScanPatternsList
'
'    On Error GoTo errHandler
'
'    If Is_SFC1_Enabled = False Then
'
'        If ActiveWorkbook.Sheets(SFC1PatternsSheet).index = 0 Then
'
'errHandler:
'
'            TheExec.AddOutput "Reading SFC1 Input ... FAILED"
'
'            TheExec.AddOutput "        SFC1       ... DISABLED"
'
'            Read_SFC1_Table = 0
'
'    '        SFC1_Disable
'
'            Exit Function
'
'        End If
'
'
'
'     '   ScanPatternsList(0).InstanceName = ""
'
'        Dim ScanFailPatterns As Worksheet
'
'        Set ScanFailPatterns = Sheets(SFC1PatternsSheet)
'
'
'
'        If ScanFailPatterns.Cells(1, Columns.Count).End(xlToLeft).Column < 1 Then
'
'            TheExec.AddOutput ScanFailPatternsSheet & " -> Invalid Data"
'
'            Exit Function
'
'        End If
'
'        TheExec.AddOutput ScanFailPatterns.name & " -> Column Count: " & ScanFailPatterns.Cells(1, Columns.Count).End(xlToLeft).Column
'
'
'
'        'Dim RowCount As Long
'
'        rowCount = ScanFailPatterns.Cells(Rows.Count, 1).End(xlUp).Row
'
'        TheExec.AddOutput " ->Row Count: " & rowCount
'
'
'
'        ReDim ScanPatternsList(rowCount)
'
'        Dim idx As Long
'
'        Dim SFC1_Record() As String
'
'        For idx = 1 To rowCount
'
'        SFC1_Record = Split(ScanFailPatterns.Cells(idx, 1).value, ",")
'
'            ScanPatternsList(idx).FailCycleCount = SFC1_Record(2) 'ScanFailPatterns.Cells(idx, 1).Value
'
'            ScanPatternsList(idx).ScanPatternName = SFC1_Record(1) & "*" 'ScanFailPatterns.Cells(idx, 2).Value
'
'            ScanPatternsList(idx).instanceName = SFC1_Record(0) & "*" 'ScanFailPatterns.Cells(idx, 2).Value
'
'    '    theexec.AddOutput " ->Pattern[" & CStr(idx) & "]: " & ScanPatternsList(idx).ScanPatternName
'
'        Next idx
'
'        Is_SFC1_Enabled = True
'
'        SFC1_Show
'
'    End If
'
'End Function
'
'
'
Public Function SFC1_Show()

    Dim idx As Long

    For idx = 1 To rowCount

        'ScanPatternsList(idx).FailCycleCount '= ScanFailPatterns.Cells(idx, 1).Value

        'ScanPatternsList(idx).ScanPatternName '= ScanFailPatterns.Cells(idx, 2).Value

    TheExec.AddOutput " ->Instance[" & CStr(idx) & "]: " & ScanPatternsList(idx).instancename

    TheExec.AddOutput " ->Pattern[" & CStr(idx) & "]: " & ScanPatternsList(idx).ScanPatternName

    TheExec.AddOutput " ->CycleCount[" & CStr(idx) & "]: " & ScanPatternsList(idx).FailCycleCount

    Next idx

End Function



Public Function SFC1_InstanceMatch(ByRef PrePatF As String, ByRef PostPatF As String, ByRef PostPatFArg As String) As Boolean

    SFC1_InstanceMatch = False

    Dim idx As Long

    Dim instance_name As String

    instance_name = TheExec.DataManager.instancename

    For idx = 1 To rowCount

'        If ScanPatternsList(idx).InstanceName Like Instance_Name Then

        If instance_name Like ScanPatternsList(idx).instancename Then

            PostPatFArg = ScanPatternsList(idx).FailCycleCount & "," & ScanPatternsList(idx).ScanPatternName

            PostPatF = "SFC1_PostPatF"

            PrePatF = "SFC1_PrePatF"

            SFC1_InstanceMatch = True

        End If

    Next idx

End Function



'Public Function Is_SFC1_ReadInput() As Boolean
'
'    Is_SFC1_ReadInput = TheExec.enableWord(SFC1_ReadInput)
'
'    TheExec.enableWord(SFC1_ReadInput) = False
'
'End Function



'Private Function SFC1_Disable()
'
'    TheExec.enableWord(SFC1_Enable) = False
'
'    TheExec.enableWord(SFC1_ReadInput) = False
'
'    Is_SFC1_Enabled = False
'
'End Function



Public Function SFC1_PostPatF(pmCnt As Long, pmNames() As String)

    ifcmemGetFailCycles pmCnt, pmNames

'    TheHdw.Digital.CMEM.SetCaptureConfig 0, CmemCaptNone 'CmemCaptType.CmemCaptFail

'    TheHdw.Digital.CMEM.CentralFields = tlCMEMCaptureFields.tlCMEMNone

    If glb_isSFC_Enabled = False Or Glb_Bool_CheckSFCInitPat = True Then Exit Function

     ''''' New request for pattern pin group '''''

    If glb_TesterType = "UltraFLEXplus" Then

'        TheHdw.Digital.Patgen.ScanBurstEnabled = False ' cannot be called in PostPatF

''        TheExec.Datalog.Setup.ScanSetup.EnableScanLogging = False

    End If

End Function

Public Function SFC1_PrePatF(pmCnt As Long, pmNames() As String)

    TheHdw.Digital.CMEM.CentralFields = tlCMEMCaptureFields.tlCMEMPatternName + tlCMEMCaptureFields.tlCMEMAbsoluteCycle + tlCMEMCaptureFields.tlCMEMVMVectorOffset

    TheHdw.Digital.CMEM.SetCaptureConfig 5, CmemCaptType.CmemCaptFail, tlCMEMCaptureSource.tlCMEMCaptureSource_PatPassFailData 'tlCMEMCaptureSource_PatPassFailData 'CmemCaptType.CmemCaptFail

    Harvest_CMEM_InitSetup

End Function



Public Function SFC1_Main(ByRef StartOfBodyF As InterposeName, ByRef PostPatF As InterposeName, ByRef PostPatFArg As String)

        ''''>>>>> SFC1

'        If Is_SFC1_ReadInput = True Then Read_SFC1_Table

        Dim SFC1_StartOfBody As String, SFC1_PostPat As String

        If glb_isSFC_Enabled Then

            If SFC1_InstanceMatch(SFC1_StartOfBody, SFC1_PostPat, PostPatFArg) = True Then

            StartOfBodyF = SFC1_StartOfBody

            PostPatF = SFC1_PostPat

''            TheExec.AddOutput "SFC1 " & TheExec.DataManager.InstanceName
''
''            TheExec.AddOutput "SFC1 " & PostPatFArgs

            End If

        End If

End Function



