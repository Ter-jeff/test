Attribute VB_Name = "VBT_ZeFuseRead"
#Const isUFP = True
Option Explicit
Private TimeCheck As Boolean
'''This is designed for DAA read(the bits are in plain)
Public Function Bank_Read(ReadPatSet As Pattern, PinRead As PinList, bank As String, earlyfuse As Boolean, ecid As Boolean, blankCheck As Boolean, _
                    blankCheckAll As Boolean, Optional RvOnly As Boolean, Optional DvOnly As Boolean, Optional printdecode As Boolean, Optional PrintDspWave As Boolean, Optional PreRead As Boolean, Optional InitPinsHi As PinList, Optional InitPinsLo As PinList, Optional InitPinsHiZ As PinList, _
                    Optional PrePatSet As Pattern, Optional DigSource As String = vbNullString, Optional MultitypJob As String = vbNullString, Optional Validating_ As Boolean)
On Error GoTo errHandler
'Debug.Print vbCrLf & "<" & TheExec.DataManager.InstanceName & ">, SiteCnt = " & TheExec.sites.Selected.Count
Dim funcName As String: funcName = "Bank_Read"
Dim dicTrimmed As Dictionary, Tname As String
Dim allBlank As New SiteLong, SiteVarValue As New SiteLong, SignalName As String, patFailCntRt As New SiteLong, putWave2Db As New SiteBoolean
Dim opbank As eFuseBdfBank, ReadPatt As String, PattArray() As String, PatCount As Long, m_siteVar As String, capWave As New DSPWave, cntBlankWave As New DSPWave
Dim tmp As New DSPWave
Dim PrePatResult As New SiteBoolean
Dim onlyPrintDecodeField As Boolean: onlyPrintDecodeField = False
Dim tempChkVar As New SiteVariant
Dim instNameKey As String: instNameKey = vbNullString
Dim m_value As New SiteLong
Dim BlankChk_CMPresult As New SiteLong
Dim SplitDspWave_Key As Variant
Dim i As Long
Dim testType As ReadTestType
Dim site As Variant
Dim CapPinArr() As String
Dim CapPinCnt As Long
Dim initSiteStatus As New SiteBoolean
Dim storeActiveSiteDic As New Dictionary: storeActiveSiteDic.compareMode = TextCompare

    If Validating_ Then
        If ReadPatSet.value <> "" Then Call PrLoadPattern(ReadPatSet.value)
        If PrePatSet.value <> "" Then Call PrLoadPattern(PrePatSet.value)
        Exit Function    ' Exit after validation
    End If

    '20210812,Add for enable word control printing
    If gB_eFuse_Disable_DecodeDataPrint_Flag = True Then printdecode = False
    If gB_eFuse_Disable_DSPwavePrint_Flag = True Then PrintDspWave = False

    ''202005xx for ap
    Call RunDspSet

    Set opbank = GetBdfBank(bank)
    
    If Not blankCheck And opbank.NeedJTAGRead And opbank.pgmMode = pgm_DAA Then
        theexec.Flow.TestLimit resultVal:=1, lowVal:=0, hiVal:=0, Tname:="ErrorPgmMode"
        GoTo BankRead_ErrorPgmMode
    End If
    
    If opbank.SplitReadDspWaveFlag Then
        i = 0
        For Each SplitDspWave_Key In Dic_SplitDspWave.Keys()
            For Each site In theexec.sites
                If i = 0 Then
                    capWave(site) = Dic_SplitDspWave(SplitDspWave_Key)(site).Copy
                Else
                    capWave(site) = capWave(site).Concatenate(Dic_SplitDspWave(SplitDspWave_Key)(site))
                End If
            Next site
            i = i + 1
        Next SplitDspWave_Key
        
        GlbUtility.IniDictionary Dic_SplitDspWave
        Dic_SplitDspWave.RemoveAll
    Else
        TheHdw.Digital.ApplyLevelsTiming True, True, True, tlPowered, InitPinsHi, InitPinsLo, InitPinsHiZ
        If PrePatSet.value <> "" Then
            initSiteStatus = TheExec.sites.Active
            PrePatResult = EfuseExecInitPattern(PrePatSet.value, DigSource)
            If Not PrePatResult.Any(True) Then
                Exit Function
            ElseIf PrePatResult.Any(False) Then
                TheExec.sites.Selected = PrePatResult
            End If
        End If
        Call PATT_GetPatListFromPatternSet(ReadPatSet.value, PattArray, PatCount)
        ReadPatt = PattArray(0)
    End If

GlbUtility.TimerStarter

    SignalName = theexec.DataManager.instancename
    If opbank.IsExistMultiTypeKey(SignalName, instNameKey) Or MultitypJob <> "" Then
        If opbank.IsMultiTypInst(instNameKey, MultitypJob) Then
            onlyPrintDecodeField = True
        Else
            m_value = 1
            theexec.Flow.TestLimit resultVal:=m_value, lowVal:=0, hiVal:=0, Tname:="Multi-type_SettingError"
            Exit Function
        End If
    End If
    If blankCheckAll Then blankCheck = True ' blankCheckAll is the type of blankCheck
    '****** Check Pattern Fail Count******
    'Set patFailCntRt = AllBanksPatFailCnt  'If you need a blocker in case any bank contains pattern failure, then uncomment this and set a limit for it.
    Set patFailCntRt = PatFailCnt(opbank)
    theexec.Flow.TestLimit resultVal:=patFailCntRt, lowVal:=0, hiVal:=0, Tname:=m_siteVar, PinName:="PatFailCnt"
    If theexec.sites.Active.Count = 0 Then Exit Function

    '****** Check early fused status******
    If earlyfuse Then
        ''202005xx for ap
        If Not opbank.HadEarlyFuse Then
                SiteVarValue = 2
        End If
        'If Not opbank.HadEarlyFuse Then GoTo errHandler:
        'If opbank.IsEarlyFused Then GlbUtility.MessageBox bank & " bank had already done early fused! Please confirmed!": GoTo errHandler:
    Else
        Call EarlyCheck4MultipleStage(opbank)
    End If
    
    '****** Initialize Site Varaible/ Getting categories as Needs******
    If blankCheck Then
            m_siteVar = bank + "Chk_Var"
            If Not blankCheckAll Then
                For Each site In theexec.sites
                    tempChkVar = theexec.sites(site).SiteVariableValue(m_siteVar)
                Next site
                opbank.BlankCheckCount = opbank.BlankCheckCount + 1
            End If
            GlbUtility.IniSiteVar m_siteVar, -1
            Set dicTrimmed = ObtainCatDictionary(opbank, ecid, earlyfuse, blankCheckAll, RvOnly, DvOnly, MultitypJob) 'if dicTrimmed is nothing means "blank check ALL"
    Else
        ''when execute the read item, set BlankCheckCount and IsSameCheckVar to initial value
        Call opbank.InitialMultiCheckVarVariable
        If MultitypJob <> "" Then
            Set dicTrimmed = ObtainCatDictionary(opbank, ecid, earlyfuse, blankCheckAll, RvOnly, DvOnly, MultitypJob)
        End If
    End If

    If g_Rvenable Then
        Set dicTrimmed = ObtainCatDictionary(opbank, ecid, earlyfuse, blankCheckAll, RvOnly, DvOnly, MultitypJob)
    End If

    If Not dicTrimmed Is Nothing Then 'Margin read, the dicTrimmed = Nothing
        If dicTrimmed.Count = 0 Then GoTo SkipInstance:
    End If
     '****** Runs pattern and reterieve the dspwave******  'capture result is double bits result, if it's in double bit mode
    'If GlbUtility.OnlineMode Then
    'If GlbUtility.IsOnline Then
    theexec.DataManager.DecomposePinList PinRead, CapPinArr, CapPinCnt
    opbank.CapturePinCnt = CapPinCnt
    Erase CapPinArr()
    If Not opbank.SplitReadDspWaveFlag Then
            Set opbank.CapturePin = PinRead
            For Each site In TheExec.sites
                storeActiveSiteDic.Add site, True
            Next
            testType = IIf(GlbUtility.IsStrMatch(opbank.name, "UDR"), ReadTestType.JTAG_Or_Result, ReadTestType.DefaultMode)
            eFuse_DSSC_CapSetup ReadPatt, opbank, SignalName, capWave, testType
            Call TheHdw.Patterns(ReadPatt).test(pfAlways, 0, tlResultModeDomain)    'run read pattern and capture
            Call UpdateReadPatternResult(storeActiveSiteDic, TheHdw.Digital.Patgen.PatternBurstPassedPerSite, opbank.ReadPatResut)
    Else
        opbank.SplitReadDspWaveFlag = False
    End If

    If Not GlbUtility.OnlineMode Then
            If opbank.pgmMode = pgm_JTAG Then PreRead = True ' set this will enable "putWave2Db", let offline log gets pass result
            If earlyfuse Then opbank.GetEarlyFuseOnly = True
            Set capWave = OfflineDspWave(opbank, ecid, earlyfuse, blankCheck, blankCheckAll)  'capture result is double bits result
            If earlyfuse Then opbank.GetEarlyFuseOnly = False
    End If
    'If debugprint Then Debug.Print "samplesize1 " & capWave.SampleSize
    
    ''tmp = capWave.Copy
GlbUtility.PersonalTimerLog "Retrieve Capture Wave"
'Debug.Print TheHdw.DSP.ExecutionMode
    ''202006xx dsp modify------------------------------------------------------------------------------------------
    'Dim testwave As New DSPWave
    'testwave = capWave.Copy
    Dim forceDbit As Boolean
    Dim nonEarly As Boolean: nonEarly = False
    If Not earlyfuse Then nonEarly = True
    forceDbit = IIf(GlbUtility.IsStrMatch(opbank.name, "UDR") And GlbUtility.OnlineMode, False, True)
    'forceDbit = IIf(GlbUtility.IsStrMatch(opbank.name, "UDR") And GlbUtility.IsOnline, False, True)
    If Not blankCheckAll Then
        opbank.DigCapWaveProcess blankCheck, capWave, allBlank, True, dicTrimmed, forceDbit, chkNonEarly:=nonEarly, MultitypJob:=MultitypJob
        'DigCapWaveProcess blankCheck, opbank, capWave, allBlank, True, dicTrimmed, forceDbit, chkNonEarly:=nonEarly
    
    ''202010xx, verify
    Else
        Ze_DspwaveBlankCheck capWave, allBlank
    End If

    '20230313, Support check capture waves of two blank_check items
    If blankCheck And opbank.BlankCheckCount = 1 Then
        Set BlankCheck_capWave = New DSPWave
        Set BlankCheck_capWave = capWave
    End If
   
    'TheHdw.DSP.EnableHostThreadMode = False
    '---------------------------------------------------------------------------------------------------------------
 'TheExec.Flow.TestLimit resultVal:=allBlank, LowVal:=0, HiVal:=2, Tname:=m_siteVar, PinName:=Tname
 'If debugprint Then Debug.Print "samplesize2 " & capWave.SampleSize
'Debug.Print TheHdw.DSP.ExecutionMode

    '****** Don't put the wave to parse the field result if "blankCheck=True" and result is blank actually.
''''    If Not blankCheckAll Then
''''            If opbank.pgmMode = pgm_DAA Then
''''                    opbank.PutDaaCapToDb CapWave
''''                    If blankCheck Then
''''                             If earlyfuse Then Set cntBlankWave = opbank.DspWaveByConditions(opbank.DaaCapWaveSerial, True, dicTrimmed)
''''                             If Not earlyfuse Then Set cntBlankWave = opbank.DspWaveByConditions(opbank.DaaCapWaveSerial, True, dicTrimmed, chkNonEarly:=True)
''''                    End If
''''            Else
''''                    Dim forceDbit As Boolean
''''                    forceDbit = IIf(GlbUtility.IsStrMatch(opbank.name, "UDR") And GlbUtility.IsOnline, False, True)
''''                    opbank.PutJtagCapToDb CapWave, forceDbit
''''                    If blankCheck Then
''''                            If earlyfuse Then Set cntBlankWave = opbank.DspWaveByConditions(opbank.JtagCapturedSerial, True, dicTrimmed)
''''                            If Not earlyfuse Then Set cntBlankWave = opbank.DspWaveByConditions(opbank.JtagCapturedSerial, True, dicTrimmed, chkNonEarly:=True)
''''                    End If
''''            End If
''''    Else
''''            Set cntBlankWave = CapWave
''''    End If
'TheHdw.Wait 5
GlbUtility.PersonalTimerLog "Obtain blkWave from Capwave(Discard)"
    '****** Blank Check Result******
    putWave2Db = IIf(blankCheckAll, False, True)
    
    If RunPartType = AP _
        And (opbank.DoubleBits And Not printdecode) And Not GlbUtility.IsStrMatch(opbank.name, "UDR") And Not ForceDecodeEnable _
        And GlbUtility.OnlineMode Then
            putWave2Db = False ' move the decode to SingleDoubleBit
    End If
    
    '20211210, Modify for offline jtag bank all blank check
    If PreRead And Not blankCheckAll Then putWave2Db = True
    
    If blankCheck Then
    Tname = GetTName(earlyfuse, ecid)
            ''20210630, Modify for blankCheckAll could support not in now job
            If Not blankCheckAll Then
                For Each site In theexec.sites
                    If (allBlank(site) > 0) Or (opbank.JobExistInBank = False) Or (SiteVarValue = 2) Then 'If not blank
                        opbank.IsBlank = 0
                        SiteVarValue = 2
                    Else ' It's means this Fuse area is all balnk.
                        SiteVarValue = 1
                        If opbank.DicOthers.Count <> 0 And Not printdecode Then putWave2Db(site) = False
                    End If
                    
                    If EFUSE_REFUSE_FOR_PTE And EFUSE_POWER_OFF_SETTING Then
                        SiteVarValue = 1
                        If opbank.DicOthers.Count <> 0 And Not printdecode Then putWave2Db(site) = False
                    End If
                    
                    If Not GlbUtility.OnlineMode Then SiteVarValue = 1 'And Site = 1 Then SiteVarValue = 2
                    'If Not GlbUtility.IsOnline Then SiteVarValue = 1 'And Site = 1 Then SiteVarValue = 2
                    If tempChkVar(site) > 0 And tempChkVar(site) <> SiteVarValue And opbank.BlankCheckCount = 2 Then
                        opbank.IsSameCheckVar = 0
                    End If
                    theexec.sites(site).SiteVariableValue(m_siteVar) = CLng(SiteVarValue)
                Next site
                theexec.Flow.TestLimit resultVal:=SiteVarValue, lowVal:=1, hiVal:=2, Tname:=m_siteVar, PinName:=Tname
                If opbank.BlankCheckCount = 2 Then
                    theexec.Flow.TestLimit resultVal:=opbank.IsSameCheckVar, lowVal:=1, hiVal:=1, Tname:=bank + "BlankCheckSummary", PinName:=Tname
                    rundsp.CompareBlkChkData capWave, BlankCheck_capWave, BlankChk_CMPresult
                    theexec.Flow.TestLimit resultVal:=BlankChk_CMPresult, lowVal:=0, hiVal:=0, Tname:=bank + "_BlkChk_HV/LV_Dspwave_Compare"
                    Set BlankCheck_capWave = New DSPWave
                    Call opbank.InitialMultiCheckVarVariable
                End If
            Else
                For Each site In theexec.sites
                    If (allBlank(site) > 0) Or (SiteVarValue = 2) Then 'If not blank
                        opbank.IsBlank = 0
                        SiteVarValue = 2
                    Else ' It's means this Fuse area is all balnk.
                        SiteVarValue = 1
                        If opbank.DicOthers.Count <> 0 And Not printdecode Then putWave2Db(site) = False
                    End If
                    
                    If EFUSE_REFUSE_FOR_PTE And EFUSE_POWER_OFF_SETTING Then
                        SiteVarValue = 1
                        If opbank.DicOthers.Count <> 0 And Not printdecode Then putWave2Db(site) = False
                    End If
                    
                    If Not GlbUtility.OnlineMode Then SiteVarValue = 1 'And Site = 1 Then SiteVarValue = 2
                    'If Not GlbUtility.IsOnline Then SiteVarValue = 1 'And Site = 1 Then SiteVarValue = 2
                    theexec.sites(site).SiteVariableValue(m_siteVar) = CLng(SiteVarValue)
                Next site
                theexec.Flow.TestLimit resultVal:=SiteVarValue, lowVal:=1, hiVal:=1, Tname:=m_siteVar, PinName:=Tname
            End If
    End If
   ' If bank = "UDRP" Or bank = "UDR_P" Then Stop
     If opbank.pgmMode = pgm_DAA Then
        tmp = opbank.DaaCapWaveSerial
     Else
        tmp = opbank.JtagCapturedSerial
     End If
     opbank.PutDsscRtToDb tmp, putWave2Db, IIf(onlyPrintDecodeField, dicTrimmed, Nothing), earlyfuse
     
    '20210716, Add for FT1 PseudoFuse Read from Device
    If PseudoFuseEnable And Not MixPseudoFuseEnable Then
        Dim tmpDSP As New DSPWave
        Dim tmpDSP2 As New DSPWave
        
        For Each site In theexec.sites
        tmpDSP2 = tmp.ConvertDataTypeTo(DspLong)
        tmpDSP = opbank.PseudoFuseDspWave
        Next
        For Each site In theexec.sites
            If tmpDSP2.sampleSize = tmpDSP.sampleSize Then
                tmpDSP = tmpDSP.BitwiseOr(tmpDSP2)
            End If
        Next
        
        'Next
        Set opbank.PseudoFuseDspWave = tmpDSP
    End If
'    If opbank.pgmMode = pgm_DAA Then
'           opbank.PutDsscRtToDb opbank.DaaCapWaveSerial, putWave2Db
'    Else
'           opbank.PutDsscRtToDb opbank.JtagCapturedSerial, putWave2Db
'    End If
    'If debugprint Then Debug.Print "samplesize3 " & capWave.SampleSize
    If LCase(bank) Like "udr*" Then UpdateCmpFields (bank)
    ''202004xx for ap
    If Not earlyfuse And Not blankCheckAll And opbank.HadVddBinFuse Then
        Call ReadEfuseDataFromBinCut
    ElseIf Not earlyfuse And Not blankCheckAll And opbank.HadIdsFuse Then
        Call GetIdsValues
    End If
    'If opbank.cmpFlag = True Then Stop
    If opbank.cmpFlag = True Then opbank.PatName = ReadPatt
GlbUtility.PersonalTimerLog "Blank Check(+Put DB)"
    '****** Print eFuse Data Infomation******
    'Debug.Print TheHdw.DSP.ExecutionMode
    'If debugprint Then Debug.Print "samplesize4 " & capWave.SampleSize
    If Not blankCheck Then '' for collect pte
'      If earlyfuse And PrintDspWave Then
'            If opbank.pgmMode = pgm_DAA Then
'                    opbank.DumpDspWave capWave, "Early Fuse Data"
'            Else
'                    opbank.DumpDspWave capWave, "Early Fuse Data", 32
'            End If
'    ElseIf Not blankCheckAll Then
        If opbank.DoubleBits = True Then
            If onlyPrintDecodeField Then
                opbank.PrintReadFuseDatalog PrintDoubleBitWave:=PrintDspWave, dicItem4Print:=dicTrimmed, printdecode:=printdecode, isEarly:=earlyfuse
            Else
                opbank.PrintReadFuseDatalog PrintDoubleBitWave:=PrintDspWave, printdecode:=printdecode, isEarly:=earlyfuse
            End If
        Else
            If onlyPrintDecodeField Then
                opbank.PrintReadFuseDatalog PrintSingleBitWave:=PrintDspWave, dicItem4Print:=dicTrimmed, printdecode:=printdecode, isEarly:=earlyfuse
            Else
                opbank.PrintReadFuseDatalog PrintSingleBitWave:=PrintDspWave, printdecode:=printdecode, isEarly:=earlyfuse
            End If
        End If
        
'     ''202004xx for ap
'     ElseIf Not blankCheckAll And (Not UCase(opbank.Name) Like "*CFG*") Then
'     'ElseIf Not blankCheckAll And (Not UCase(opbank.Name) Like "*CONFIG*") Then
'            opbank.PrintReadFuseDatalog PrintDspWave, False, PrintDecode:=PrintDecode
'     ElseIf Not blankCheckAll Then
'            opbank.PrintReadFuseDatalog PrintDspWave, False, dicItem4Print:=dicTrimmed, PrintDecode:=PrintDecode ' Read data pase and print out  the result
'     End If
     'If debugprint Then Debug.Print "samplesize5 " & capWave.SampleSize
     End If
     'Debug.Print "----------------------"
     DebugPrintFunc ReadPatt
     If (PrePatSet.value <> "") Then
        TheExec.sites.Selected = initSiteStatus
     End If
GlbUtility.PersonalTimerLog "Print Out Cost"
Exit Function
BankRead_ErrorPgmMode:
    Call Print_Error_Message(Error_Info, "VBT_ZeFuseRead", "Bank_Read", "The instance can not support the bank of DAA_mode read by JTAG_mode")
    Exit Function
SkipInstance:
    Call Print_Error_Message(Warning_Info, "VBT_ZeFuseRead", "Bank_Read", theexec.DataManager.instancename & " was skipped!")
    Exit Function
errHandler:
    Call Print_Error_Message(Error_Info, "VBT_ZeFuseRead", "Bank_Read")
    If AbortTest Then Exit Function Else Resume Next
End Function

'''This is designed for APB serial read(the bits are in plain)
Public Function Bank_ApbRead(ReadPatSet As Pattern, PinRead As PinList, bank As String, _
                    Optional InitPinsHi As PinList, Optional InitPinsLo As PinList, Optional InitPinsHiZ As PinList, _
                    Optional Validating_ As Boolean)
On Error GoTo errHandler
Dim SignalName As String
Dim ReadPatt As String, PattArray() As String, PatCount As Long
Dim capWave As New DSPWave
Dim opbank As eFuseBdfBank
Dim tPgmMode As ProgramMode
Dim result As New SiteLong
Dim out1 As New DSPWave, out2 As New DSPWave

    If Validating_ Then Call PrLoadPattern(ReadPatSet.value):  Exit Function    ' Exit after validation
    Call PATT_GetPatListFromPatternSet(ReadPatSet.value, PattArray, PatCount)
    ReadPatt = PattArray(0)
    
    SignalName = theexec.DataManager.instancename
    Set opbank = GetBdfBank(bank)
     If GlbUtility.OnlineMode Then
    'If GlbUtility.IsOnline Then
            Set opbank.CapturePin = PinRead
            TheHdw.Digital.ApplyLevelsTiming True, True, True, tlPowered, InitPinsHi, InitPinsLo, InitPinsHiZ
            eFuse_DSSC_CapSetup ReadPatt, opbank, SignalName, capWave, JTAG_APB
            Call TheHdw.Patterns(ReadPatt).test(pfAlways, 0, tlResultModeDomain)   'run read pattern and capture
    Else
            'capture result is double bit data (plain bits)
            tPgmMode = opbank.pgmMode
            opbank.pgmMode = pgm_JTAG
            Set capWave = FakeBankDsscResult(opbank, AteTrimData, True, True)
            opbank.pgmMode = tPgmMode
    End If
   
    Call Ze_SplitWaveSerial32(capWave, out1, out2)
    Ze_TwoDspWaveCompare out1, opbank.DaaCapWaveSerial, result
    theexec.Flow.TestLimit resultVal:=result, lowVal:=0, hiVal:=0, Tname:="Bank_TapRead_Half1", PinName:="Value"
    Ze_TwoDspWaveCompare out2, opbank.DaaCapWaveSerial, result
    theexec.Flow.TestLimit resultVal:=result, lowVal:=0, hiVal:=0, Tname:="Bank_TapRead_Half2", PinName:="Value"
    
    'opbank.DumpDspWave CapWave, "APB Read Out", 32
    If LCase(bank) Like "udr*" Then UpdateCmpFields (bank)
    
    If theexec.sites.Active.Count = 0 Then Exit Function
Exit Function
errHandler:
    Call Print_Error_Message(Error_Info, "VBT_ZeFuseRead", "Bank_ApbRead")
    If AbortTest Then Exit Function Else Resume Next
End Function

'''This is designed for JTAG serial read(the bits are in result "OR")
Public Function Bank_TapRead(ReadPatSet As Pattern, PinRead As PinList, bank As String, _
                    Optional InitPinsHi As PinList, Optional InitPinsLo As PinList, Optional InitPinsHiZ As PinList, _
                    Optional PrePatSet As Pattern, Optional DigSource As String = vbNullString, Optional Validating_ As Boolean)
On Error GoTo errHandler
Dim SignalName As String
Dim ReadPatt As String, PattArray() As String, PatCount As Long
Dim capWave As New DSPWave
Dim opbank As eFuseBdfBank
Dim result As New SiteLong
Dim PrePatResult As New SiteBoolean
Dim paraBit As Long
Dim initSiteStatus As New SiteBoolean

    If Validating_ Then
        If ReadPatSet.value <> "" Then Call PrLoadPattern(ReadPatSet.value)
        If PrePatSet.value <> "" Then Call PrLoadPattern(PrePatSet.value)
        Exit Function    ' Exit after validation
    End If

    ''202005xx for ap
    Call RunDspSet
    
    TheHdw.Digital.ApplyLevelsTiming True, True, True, tlPowered, InitPinsHi, InitPinsLo, InitPinsHiZ
    If PrePatSet.value <> "" Then
        initSiteStatus = TheExec.sites.Active
        PrePatResult = EfuseExecInitPattern(PrePatSet.value, DigSource)
        If Not PrePatResult.Any(True) Then
            Exit Function
        ElseIf PrePatResult.Any(False) Then
            TheExec.sites.Selected = PrePatResult
        End If
    End If
    
    Call PATT_GetPatListFromPatternSet(ReadPatSet.value, PattArray, PatCount)
    ReadPatt = PattArray(0)
    
    SignalName = theexec.DataManager.instancename
    Set opbank = GetBdfBank(bank)
    
    If opbank.HadDeidFuse Then
        If (EFUSE_ECID_SORTING_ENABLE) And (Not opbank.NeedJTAGRead) Then
            Call Print_Error_Message(Error_Info, "VBT_ZeFuseRead", "Bank_TapRead", "Please turn on ***EFUSE_" & UCase(opbank.name) & "_READ_BYJTAG***, if user want to run ecid sorting.")
            theexec.Flow.TestLimit resultVal:=0, lowVal:=1, hiVal:=1, Tname:="EFUSE_" & UCase(opbank.name) & "_READ_BYJTAG"
        End If
    End If
    
    ''when execute the read item, set BlankCheckCount and IsSameCheckVar to initial value
    Call opbank.InitialMultiCheckVarVariable
    Set opbank.CapturePin = PinRead
    eFuse_DSSC_CapSetup ReadPatt, opbank, SignalName, capWave, JTAG_Or_Result
    Call TheHdw.Patterns(ReadPatt).test(pfAlways, 0, tlResultModeDomain)   'run read pattern and capture
    
    If Not GlbUtility.OnlineMode Then
        'capture result is single bit data (the or result)
        Set capWave = FakeBankDsscResult_Serial(opbank, AteTrimData, True)
        'CapWave(0).ElementLite(0) = 1
    End If
    
    '20211214, Modify for ECID JTAG Read decode
    If False = opbank.NeedJTAGRead Then
        If opbank.pgmMode = pgm_DAA Then
            opbank.PutJtagCapToDb capWave
            Ze_TwoDspWaveCompare opbank.DaaCapWaveSerial, opbank.JtagCapturedSerial, result
        Else
           'opbank.PutJtagCapToDb CapWave
           Ze_TwoDspWaveCompare capWave, opbank.JtagCapturedSerial, result
        End If
        If Not GlbUtility.OnlineMode Then
            theexec.Flow.TestLimit resultVal:=0, lowVal:=0, hiVal:=0, Tname:="Bank_TapRead", PinName:="Value"
        Else
            theexec.Flow.TestLimit resultVal:=result, lowVal:=0, hiVal:=0, Tname:="Bank_TapRead", PinName:="Value"
        End If
    ElseIf True = opbank.NeedJTAGRead Then
        If opbank.pgmMode = pgm_DAA Then
            opbank.PutJtagCapToDb capWave
            opbank.PrintTapReadFuseDatalog opbank.NeedJTAGRead
            If opbank.HadVddBinFuse Then
                Call ReadEfuseDataFromBinCut
            ElseIf opbank.HadIdsFuse Then
                Call GetIdsValues
            End If
        End If
    End If
    
    paraBit = IIf(opbank.DoubleBits, 16, 32)
    'opbank.DumpDspWave CapWave, "TAP Read Out", paraBit
    If LCase(bank) Like "udr*" Then UpdateCmpFields (bank)
    DebugPrintFunc ReadPatt
    If (PrePatSet.value <> "") Then
        TheExec.sites.Selected = initSiteStatus
    End If
    If theexec.sites.Active.Count = 0 Then Exit Function

Exit Function
errHandler:
    Call Print_Error_Message(Error_Info, "VBT_ZeFuseRead", "Bank_TapRead")
    If AbortTest Then Exit Function Else Resume Next
End Function

'udr <-> udr_cmp. It must run either Bank_Read or Bank_TapRead in first then the comparison can be made
Public Function Bank_Udrcmp(ReadPatSet As Pattern, PinRead As PinList, bank As String, _
                    Optional InitPinsHi As PinList, Optional InitPinsLo As PinList, Optional InitPinsHiZ As PinList, _
                    Optional Validating_ As Boolean)
On Error GoTo errHandler
Dim SignalName As String
Dim values As New SiteDouble
Dim ReadPatt As String, PattArray() As String, PatCount As Long
Dim capWave As New DSPWave
Dim opbank As eFuseBdfBank
Dim putWave2Db As New SiteBoolean
Dim fieldStr As Variant, field As eFuseBdfField, value As Double
Dim site As Variant

    If Validating_ Then Call PrLoadPattern(ReadPatSet.value):  Exit Function    ' Exit after validation
    Call PATT_GetPatListFromPatternSet(ReadPatSet.value, PattArray, PatCount)
    ReadPatt = PattArray(0)
    
    SignalName = theexec.DataManager.instancename

    Set opbank = GetBdfBank(bank)
    If GlbUtility.OnlineMode Then
        Set opbank.CapturePin = PinRead
        TheHdw.Digital.ApplyLevelsTiming True, True, True, tlPowered, InitPinsHi, InitPinsLo, InitPinsHiZ
        eFuse_DSSC_CapSetup ReadPatt, opbank, SignalName, capWave
        Call TheHdw.Patterns(ReadPatt).test(pfAlways, 0, tlResultModeDomain)   'run read pattern and capture
    Else
        'capture result is single bit data (the or result)
        Set capWave = FakeBankDsscResult_Serial(opbank, AteTrimData, True)
    End If
    
    putWave2Db = True
    opbank.PutJtagCapToDb capWave
    opbank.PutDsscRtToDb opbank.JtagCapturedSerial, putWave2Db

    For Each site In theexec.sites
        For Each fieldStr In opbank.Fields.Keys
            Set field = opbank.Fields(fieldStr)
            value = GlbUtility.Hex2Dbl(field.TrimAteValue) ' The values came from bank UDR capture result.
            theexec.Flow.TestLimit resultVal:=GlbUtility.Hex2Dbl(field.DsscValue), lowVal:=value, hiVal:=value, Tname:=fieldStr, PinName:=""
        Next
    Next
    opbank.PrintReadFuseDatalog False, False

Exit Function
errHandler:
    Call Print_Error_Message(Error_Info, "VBT_ZeFuseRead", "Bank_Udrcmp")
    If AbortTest Then Exit Function Else Resume Next
End Function

'It auto-detected the bank supported double bits or not, if not then throw error.
Public Function Bank_SingleDoubleBitsCheck(opbank As eFuseBdfBank)
On Error GoTo errHandler
Dim result As New SiteLong
Dim nBits As Long
Dim ParallelWave As New DSPWave
Dim out1 As New DSPWave, out2 As New DSPWave
Dim site As Variant 'Carter, 20240304

    ''202004xx for ap
    'If Not opbank.DoubleBits Then GoTo errHandler: 'Not double bits, can't do the check!!
    If opbank.pgmMode = pgm_DAA Then
        Set result = DoubleBitChecks(opbank.DaaCapturedWave, out1, out2)
    Else
        ''202004xx for ap
        nBits = IIf(opbank.DoubleBits, 2, 1)
        'If nBits Then
        If nBits = 2 Then
            '2-bit case(APB)
            ' Convert it to parallel then check it like the DAA mode.
            Ze_DspwaveToParallelMsb opbank.JtagCapturedWave, 32, ParallelWave
            Set result = DoubleBitChecks(ParallelWave, out1, out2)
        Else
            '1-bit case
            'no duplicate bit to comapre
            result = 0
        End If
    End If
    theexec.Flow.TestLimit resultVal:=result, lowVal:=0, hiVal:=0, Tname:="SingleDoubleBits", PinName:="Value"

    If gB_efuse_DebugPrint_SingleDoubleBits_Flag Then
        For Each site In theexec.sites
            If result(site) <> 0 Then Call DebugPrintSingleDoubleBitCheck(out1, out2, site)
        Next
    End If
Exit Function
errHandler:
    Call Print_Error_Message(Error_Info, "VBT_ZeFuseRead", "Bank_SingleDoubleBitsCheck")
    If AbortTest Then Exit Function Else Resume Next
End Function

Public Function Bank_IsBlank()
On Error GoTo errHandler 'Blank = 1, non-Blank = 0
Dim opbank As eFuseBdfBank, bankstr As Variant

    For Each bankstr In BdfDataBase.Banks.Keys
        If Not GlbUtility.IsStrMatch(CStr(bankstr), "cmp") Then
        Set opbank = BdfDataBase.Banks(CStr(bankstr))
        theexec.Flow.TestLimit resultVal:=opbank.IsBlank, lowVal:=0, hiVal:=0, Tname:=bankstr & "_IsBlank", PinName:="Value"
        End If
    Next

Exit Function
errHandler:
    Call Print_Error_Message(Error_Info, "VBT_ZeFuseRead", "Bank_IsBlank")
    If AbortTest Then Exit Function Else Resume Next
End Function

'''The syntaxes checker are limited at the very last to save the TT.
Public Function Bank_SyntaxCheck(bank As String, _
                                 earlyfuse As Boolean, _
                                 ecid As Boolean, _
                                 checkAll As Boolean, _
                                 compareWR As Boolean, _
                                 Optional RvOnly As Boolean, _
                                 Optional printdecode As Boolean = False, _
                                 Optional PrintDspWave As Boolean, Optional MultitypJob As String = vbNullString, _
                                 Optional Validating_ As Boolean)
On Error GoTo errHandler
RvOnly = False
Dim funcName As String: funcName = "Bank_SyntaxCheck"
Dim opbank As eFuseBdfBank, m_EQNum As New SiteLong, maxEQ As New SiteLong
Dim m_unitType As UnitType, dicTrimmed As Dictionary, vChecks As Variant
Dim fieldStr As Variant, field As eFuseBdfField, capWave As DSPWave
Dim oValue() As New SiteDouble, oHlimit() As Double, oLimit() As Double, asigned As Boolean
Dim oName() As String, oUnit() As UnitType, oScaleType() As tlScaleType, iItem As Long, preiItem As Long, subItem As Long
Dim oCmpLlimit() As New SiteDouble, oCmpHlimit() As New SiteDouble
Dim m_ScaleType As tlScaleType
Dim fieldStage As String
Dim B_Hlimit() As New SiteDouble
Dim B_Llimit() As New SiteDouble
Dim LoopLimitPrint() As Boolean
Dim LLim As Double
Dim HLim As Double
Dim arrSize As Long
Dim Arraytmp As Variant
Dim Strtmp As Variant
Dim i As Long ''20210118
Dim instNameKey As String: instNameKey = vbNullString
Dim m_value As New SiteLong
Dim site As Variant
Dim m_WLFTEcidInfoStr As String

    If Validating_ Then Validating_ = False: Exit Function         ' Exit after validation

GlbUtility.TimerStarter
    Set opbank = GetBdfBank(bank)

    If opbank.IsExistMultiTypeKey(theexec.DataManager.instancename, instNameKey) Or MultitypJob <> "" Then
        If Not opbank.IsMultiTypInst(instNameKey, MultitypJob) Then
            m_value = 1
            theexec.Flow.TestLimit resultVal:=m_value, lowVal:=0, hiVal:=0, Tname:="Multi-type_SettingError"
            Exit Function
        End If
    End If

    Bank_TrimmedUpdate bank, earlyfuse, ecid, RvOnly, MultitypJob
    Set dicTrimmed = ObtainCatDictionary(opbank, ecid, earlyfuse, False, RvOnly, MultitypJob:=MultitypJob)
    Bank_Decode_Print opbank, printdecode, earlyfuse, IIf(MultitypJob <> "", dicTrimmed, Nothing)
    PutData2RegKeySave opbank, earlyfuse  'RegKeySave put at here
    
    ''20220524, modify for syntax check only check current job fields
    If gB_eFuse_Disable_SyntaxCheckAll_Flag = True Then checkAll = False

    ''202005xx for ap
    'If opbank.cmpFlag = True Then opbank.GetFieldHLlimit
    'PrintDspWave = False
    '202004xx for ap
    If PseudoFuseEnable And Not MixPseudoFuseEnable Then
        Dim tmpWave As New DSPWave
        If opbank.HadVddBinFuse Then
            Call ReadEfuseDataFromBinCut
        ElseIf opbank.HadIdsFuse Then
            Call GetIdsValues
        End If

        printdecode = True
        If opbank.pgmMode = pgm_DAA Then
            Set opbank.DaaCapWaveSerial = opbank.PseudoFuseDspWave
            Ze_XExpandDblBit opbank.PseudoFuseDspWave, tmpWave
            
            Set opbank.DaaCapWaveSerialFL = tmpWave
            For Each site In theexec.sites
                opbank.DaaCapturedWave = tmpWave.ConvertDataTypeTo(DspLong).ConvertStreamTo(tldspParallel, 32, 0, Bit0IsMsb)
            Next
        Else
            Set opbank.JtagCapturedSerial = opbank.PseudoFuseDspWave
            
            If opbank.DoubleBits = True Then
                Ze_XExpandDblBit opbank.PseudoFuseDspWave, opbank.JtagCapturedWave
            Else
                Set opbank.JtagCapturedWave = opbank.PseudoFuseDspWave
            End If
            For Each site In theexec.sites
                opbank.JtagCapturedWave = opbank.JtagCapturedWave.ConvertDataTypeTo(DspLong)
            Next
        End If
        If opbank.DoubleBits = True Then Bank_SingleDoubleBitsCheck opbank
    Else
        '20211214, Modify for ECID JTAG Read decode
        If False = opbank.NeedJTAGRead Then
            If opbank.DoubleBits = True Then Bank_SingleDoubleBitsCheck opbank
        End If
    End If
    
'    If compareWR Then
'        Bank_CompareWRData bank, earlyFuse
'        Bank_SingleDoubleBitsCheck bank
'    End If
    'If IsInstMatch("_Verify") Then Bank_SingleDoubleBitsCheck bank
    
    iItem = 0: subItem = 0: preiItem = 0
    vChecks = IIf(checkAll Or dicTrimmed.Count = 0, opbank.Fields.Keys, dicTrimmed.Keys)
    If MultitypJob <> "" Then
        vChecks = dicTrimmed.Keys
    End If

    arrSize = UBound(vChecks) + 5
    ReDim oValue(arrSize)
    ReDim oHlimit(arrSize)
    ReDim oLimit(arrSize)
    ReDim oName(arrSize)
    ReDim oUnit(arrSize)
    ReDim oScaleType(arrSize)
    ReDim oCmpLlimit(arrSize)
    ReDim oCmpHlimit(arrSize)
    ReDim B_Hlimit(arrSize)
    ReDim B_Llimit(arrSize)
    ReDim LoopLimitPrint(arrSize)

    If gB_eFuse_Disable_ChkLMT_Flag Then
        Dim Disable_chk_arr() As Long
        'ReDim Disable_chk_arr(UBound(vChecks) + 1)
        ReDim Disable_chk_arr(0)
    End If
            For Each fieldStr In vChecks
                Dim values As New SiteDouble, lotStr As String
                Set field = opbank.Fields(fieldStr)
                iItem = iItem + subItem: asigned = False
                
                LoopLimitPrint(iItem) = False
                
                If gB_eFuse_Disable_ChkLMT_Flag Then
                    ReDim Preserve Disable_chk_arr(iItem)
                    Disable_chk_arr(iItem) = 0
                    fieldStage = Replace(UCase(field.BlowLocation), "_EARLY", "")
                    
                    'Handle extral subitems parts, 20210318
                    If iItem - preiItem > 1 Then
                        For i = preiItem + 1 To iItem - 1
                            Disable_chk_arr(i) = Disable_chk_arr(preiItem)
                        Next i
                    End If
                    
                    If fieldStage = theexec.CurrentJob Or GlbUtility.testedStages.Exists(field.BlowLocation) Then
                        Disable_chk_arr(iItem) = 0
                    Else
                        Disable_chk_arr(iItem) = 1
                    End If

                    preiItem = iItem
                End If
                
                    If field.ReadOnly And field.size < 32 And field.Algorithm <> alg_base Then 'base needs to convert it to mV
                                subItem = 0: m_unitType = unitVolt
                                If field.Algorithm = alg_vddbin Then
                                                    oValue(iItem) = field.FuseMeasureValue.Multiply(0.001)
                                                    If Not asigned Then
                                                            oHlimit(iItem) = field.Hlimit: oLimit(iItem) = field.Llimit
                                                            oName(iItem) = fieldStr:    oUnit(iItem) = m_unitType:   oScaleType(iItem) = scaleMilli
                                                    End If
                                Else
                                               m_unitType = unitNone
                                               oValue(iItem) = field.DsscDecValue:
                                               oHlimit(iItem) = field.Hlimit:  oLimit(iItem) = field.Llimit
                                               oName(iItem) = fieldStr:     oUnit(iItem) = m_unitType:   oScaleType(iItem) = scaleNone
                                End If
                                subItem = subItem + 1
                    Else
                                For Each site In theexec.sites
                                subItem = 0:
                                    Select Case field.Algorithm
                                    Dim mytname As String
                                    ''202004xx for ap
'                                    Case alg_cond:
'                                            m_unitType = unitNone
'                                            Dim i32 As Long, dssc As String, atetrim As String, fourBytes As Long, mytname As String
'                                            Dim oridssc As String, oriatetrim As String, last4BytesLen As Long, getLen As Long, getStPos As Long, maxPos As Long
'                                            oridssc = Replace(field.DsscValue, "0x", "")
'                                            If (bLumpStages Or CfgDataBaseAct.EarlyStage) And Not earlyFuse Then
'                                                    oriatetrim = Replace(CfgDataBaseAct.CfgCmpValueLumpStages(field.Name), "0x", "") ' IIf(field.Trimmed, Replace(CfgDataBaseAct.CfgCmpValueLumpStages(field.Name), "0x", ""), Replace(field.TrimAteValue, "0x", ""))
'                                            Else
'                                                    oriatetrim = Replace(CfgDataBaseAct.CfgCmpValue(field.Name), "0x", "") 'IIf(field.Trimmed, Replace(CfgDataBaseAct.CfgCmpValue(field.Name), "0x", ""), Replace(field.TrimAteValue, "0x", ""))
'                                            End If
'                                            fourBytes = (field.hBytes - 1) \ 8 + 1
'                                            last4BytesLen = IIf(field.hBytes Mod 8 = 0, 8, field.hBytes Mod 8)
'                                            For i32 = 0 To fourBytes - 1 'CFG_Cond[0031:0000]_A00_0xA050C030
'                                                    getLen = IIf(i32 <> fourBytes - 1, 8, last4BytesLen)
'                                                    maxPos = IIf(field.msb < (i32 + 1) * 32 - 1, field.msb, (i32 + 1) * 32 - 1 + field.lsb)
'                                                    getStPos = IIf(field.hBytes - (i32 + 1) * 8 + 1 <= 0, 1, field.hBytes - (i32 + 1) * 8 + 1)
'                                                    dssc = Mid(oridssc, getStPos, getLen)
'                                                    atetrim = Mid(oriatetrim, getStPos, getLen)
'                                                    mytname = fieldStr & "[" & GlbUtility.TxtFmt(maxPos, 4, "0") & ":" & GlbUtility.TxtFmt(i32 * 32 + field.lsb, 4, "0") & "]_" & SelectedCfg & "_0x" & atetrim & " ... " & "0x" & dssc
'                                                    mytname = IIf(CFG_SVM, mytname & " <SVM>", mytname)
'                                                    values = IIf(dssc = atetrim, 1, 0)
'                                                    oValue(iItem + subItem) = values:
'                                                    If values = 0 Then GlbUtility.WriteDlg "Site" & Site & " " & mytname
'                                                    If Not asigned Or values = 1 Then
'                                                        oHlimit(iItem + subItem) = 1:  oLimit(iItem + subItem) = 1
'                                                        oName(iItem + subItem) = mytname:  oUnit(iItem + subItem) = m_unitType:  oScaleType(iItem + subItem) = scaleNone
'                                                    End If
'                                                    subItem = subItem + 1
'                                            Next
                                    Case alg_uid:
                                            m_unitType = unitNone
                                            If opbank.DicRandomFused.Exists(field.name) Then
                                            Dim stBit As Long: stBit = IIf(field.LSB < field.msb, field.LSB, field.msb)
                                                    If opbank.pgmMode = pgm_DAA And False = opbank.NeedJTAGRead Then
                                                            Set capWave = opbank.DaaCapWaveSerial.Select(stBit, 1, field.size).Copy
                                                    Else
                                                            Set capWave = opbank.JtagCapturedSerial.Select(stBit, 1, field.size).Copy
                                                    End If
                                                    values = capWave.CountElements(EqualTo, 1) / field.size
                                             Else
                                                    values = GlbUtility.Hex2Dbl(field.DsscValue)
                                             End If
                                                     oValue(iItem) = values:
                                                     If Not asigned Then
                                                            oHlimit(iItem) = field.Hlimit:  oLimit(iItem) = field.Llimit
                                                            oName(iItem) = fieldStr:   oUnit(iItem) = m_unitType:   oScaleType(iItem) = scaleNone
                                                    End If
                                                    subItem = subItem + 1
                                    Case Algorithm.alg_base:
                                                    m_unitType = unitVolt
                                                    values = field.BaseValue_V
                                                    oValue(iItem) = values:
                                                    If Not asigned Then
                                                        oHlimit(iItem) = field.Hlimit * 0.001: oLimit(iItem) = field.Llimit * 0.001
                                                        oName(iItem) = fieldStr:    oUnit(iItem) = m_unitType:   oScaleType(iItem) = scaleMilli
                                                    End If
                                                    subItem = subItem + 1
                                    Case Algorithm.alg_ids:
                                                    m_unitType = unitAmp
                                                    values = field.IdsValue_A
                                                    oValue(iItem) = values:
                                                    ''20210118----------------------------------
                                                    If field.name = "ids_vdd_gpu_10" Or _
                                                       field.name = "ids_vdd_sram_gpu_10" Or _
                                                       field.name = "ids_vdd_gpu_105_10" Or _
                                                       field.name = "ids_vdd_sram_gpu_105_10" Then
                                                        oLimit(iItem) = 0
                                                    Else
                                                        oLimit(iItem) = field.Llimit
                                                    End If
                                                    
                                                    If Not asigned Then
                                                        oHlimit(iItem) = field.Hlimit
                                                        oName(iItem) = fieldStr:    oUnit(iItem) = m_unitType:   oScaleType(iItem) = scaleMilli
                                                    End If
                                                    subItem = subItem + 1
                                    Case Algorithm.alg_vddbin
                                       If field.BlowLocation = GlbUtility.currStage Or GlbUtility.testedStages.Exists(field.BlowLocation) Then
                                                m_unitType = unitVolt
'
                                                If LCase(field.BlowLocation) = LCase(BincutAdditionalSheetName) Then
                                                    m_EQNum = GetVddBinEqu(field, CurrentPassBinCutNum_additional, maxEQ, values, True, False)
                                                Else
                                                    m_EQNum = GetVddBinEqu(field, CurrentPassBinCutNum_normal, maxEQ, values, , False)
                                                End If
                                                
                                                oValue(iItem) = values
                                                B_Hlimit(iItem) = field.BVHLimit * 0.001: B_Llimit(iItem) = field.BVLLimit * 0.001
                                                oName(iItem) = fieldStr:    oUnit(iItem) = m_unitType:   oScaleType(iItem) = scaleMilli

                                                LoopLimitPrint(iItem) = True
                                        Else
                                            m_unitType = unitVolt
                                            values = field.FuseMeasureValue * 0.001
                                            oValue(iItem) = values
                                            oName(iItem) = fieldStr:    oUnit(iItem) = m_unitType:   oScaleType(iItem) = scaleMilli
                                        End If
                                                 subItem = subItem + 1
                                     Case Algorithm.alg_crc
                                                m_unitType = unitNone
                                                ''202004xx for ap
                                                If ((Not EFUSE_ALWAYS_CHECK_CRC) And GlbUtility.testedStages.Exists(field.BlowLocation)) Then
                                                'If (field.TrimAteValue = "0x0") And (field.BlowLocation = GlbUtility.currStage Or bLumpStages) Or GlbUtility.testedStages.Exists(field.BlowLocation) Then
                                                    oValue(iItem) = 0:
                                                    If Not asigned Then
                                                            oHlimit(iItem) = 0:  oLimit(iItem) = 0
                                                            oName(iItem) = fieldStr:     oUnit(iItem) = m_unitType:   oScaleType(iItem) = scaleNone
                                                    End If
                                                Else
                                                    If (field.BlowLocation = GlbUtility.currStage Or bLumpStages) Or GlbUtility.testedStages.Exists(field.BlowLocation) Then
                                                        If field.crc_type = crc_normal Then
                                                            Dim myCrc As String
                                                            If opbank.pgmMode = pgm_DAA And False = opbank.NeedJTAGRead Then
                                                                Set capWave = opbank.DaaCapWaveSerial
                                                            Else
                                                                Set capWave = opbank.JtagCapturedSerial
                                                            End If
                                                            myCrc = opbank.CalculateCRC(field.name, capWave, getValOnly:=True)
                                                            If myCrc = field.DsscValue Then
                                                                values = -1
                                                                GlbUtility.WriteDlg "Site" & site & " Crc checked PASS~ [" & field.name & "] = " & field.DsscValue
                                                            Else
                                                                values = GlbUtility.Hex2Dbl(field.DsscValue)
                                                                GlbUtility.WriteDlg "Site" & site & " Crc checked FAIL! [" & field.name & "]," & " DSSC is = " & field.DsscValue & ", but calculate result is " & myCrc
                                                            End If
                                                        ElseIf field.crc_type = crc_onecomp Then
                                                            Dim myCrcComp As String
                                                            myCrcComp = opbank.CalculateCRCComp(field.name, getValOnly:=True)
                                                            If myCrcComp = field.DsscValue Then
                                                                values = -1
                                                                GlbUtility.WriteDlg "Site" & site & " CrcComp checked PASS~ [" & field.name & "] = " & field.DsscValue
                                                            Else
                                                                values = GlbUtility.Hex2Dbl(field.DsscValue)
                                                                GlbUtility.WriteDlg "Site" & site & " CrcComp checked FAIL! [" & field.name & "]," & " DSSC is = " & field.DsscValue & ", but calculate result is " & myCrcComp
                                                            End If
                                                        Else
                                                            values = GlbUtility.Hex2Dbl(field.DsscValue)
                                                        End If
                                                        
                                                        field.Llimit = -1: field.Hlimit = -1
                                                    Else
                                                        values = GlbUtility.Hex2Dbl(field.DsscValue)
                                                    End If
                                                    oValue(iItem) = values:
                                                    If Not asigned Then
                                                            oHlimit(iItem) = field.Hlimit: oLimit(iItem) = field.Llimit
                                                            oName(iItem) = fieldStr:    oUnit(iItem) = m_unitType:
                                                    End If
                                                End If
                                                 subItem = subItem + 1
                                     Case Algorithm.alg_lotid
                                             Dim iChr As Long, ascii As Long: lotStr = opbank.Decode(field, True)
                                             Dim tmpBin As String, tmpCnt As Long
                                             fieldstage = BdfDataBase.GetRealStage(field.BlowLocation)
                                             fieldstage = Replace(UCase(fieldstage), "_EARLY", "")
                                        
                                        If (fieldstage = GlbUtility.currStage) Or GlbUtility.testedStages.Exists(fieldstage) Then
                                            mytname = "LotID_1st_Char"
                                            tmpBin = GlbUtility.GetLotIdBin(mid(lotStr, 1, 1))
                                            values = 0
                                            For i = 1 To Len(tmpBin)
                                                values = values + CLng(mid(tmpBin, i, 1))
                                            Next i
                                            oValue(iItem + subItem) = values:
                                            If Not asigned Then
                                                oHlimit(iItem + subItem) = -999:  oLimit(iItem + subItem) = 1
                                                oName(iItem + subItem) = mytname:  oUnit(iItem + subItem) = m_unitType:  oScaleType(iItem + subItem) = scaleNone
                                            End If
                                            subItem = subItem + 1
                                            mytname = "LotID_2to6_Char"
                                            
                                            values = 0
                                            tmpBin = GlbUtility.GetLotIdBin(mid(lotStr, 2, 5))
                                            For i = 1 To Len(tmpBin)
                                                values = values + CLng(mid(tmpBin, i, 1))
                                            Next i
                                            oValue(iItem + subItem) = values
                                            
                                            If Not asigned Then
                                                oHlimit(iItem + subItem) = -999: oLimit(iItem + subItem) = 1
                                                oName(iItem + subItem) = mytname:  oUnit(iItem + subItem) = m_unitType:  oScaleType(iItem + subItem) = scaleNone
                                            End If
                                            subItem = subItem + 1
                                                   
                                            ''20210118----------------
                                            If GlbUtility.currStage = fieldstage Then
                                               mytname = "Prober_" + UCase(lotId) + "_vs_DUT_" + UCase(lotStr)
                                               If (UCase(lotId) = UCase(lotStr)) Then
                                                       ''''Pass
                                                       oValue(iItem + subItem) = 1
                                               Else
                                                       ''''Fail
                                                       oValue(iItem + subItem) = 0
                                               End If
                                               If Not asigned Then
                                                       oHlimit(iItem + subItem) = 1:  oLimit(iItem + subItem) = 1
                                                       oName(iItem + subItem) = mytname:  oUnit(iItem + subItem) = m_unitType:  oScaleType(iItem + subItem) = scaleNone
                                               End If
                                               subItem = subItem + 1
                                            End If
                                        Else
                                            mytname = fieldstr
                                            tmpBin = GlbUtility.GetLotIdBin(mid(lotStr, 1, 6))
                                            values = 0
                                            For i = 1 To Len(tmpBin)
                                                values = values + CLng(mid(tmpBin, i, 1))
                                            Next i
                                            oValue(iItem + subItem) = values:
                                            If Not asigned Then
                                               oHlimit(iItem + subItem) = 0:  oLimit(iItem + subItem) = 0
                                               oName(iItem + subItem) = mytname:  oUnit(iItem + subItem) = m_unitType:  oScaleType(iItem + subItem) = scaleNone
                                            End If
                                            subItem = subItem + 1
                                        End If
                                            
                                     Case Else
                                             fieldstage = BdfDataBase.GetRealStage(field.BlowLocation)
                                             fieldstage = Replace(UCase(fieldstage), "_EARLY", "")
                                             If field.size < 32 Then
                                                    values = field.DsscDecValue
                                                    If opbank.cmpFlag = True Then
                                                        Call field.GetFieldHLlimit(bank)
                                                        oCmpLlimit(iItem) = field.BVLLimit
                                                        oCmpHlimit(iItem) = field.BVLLimit
                                                    ElseIf field.SetWriteByBincut Then
                                                        values = field.VddBinValue_V * 0.001
                                                        oValue(iItem) = values
                                                        B_Hlimit(iItem) = field.BVHLimit * 0.001: B_Llimit(iItem) = field.BVLLimit * 0.001
        
                                                        LoopLimitPrint(iItem) = True
                                                    ElseIf field.Algorithm = alg_numeric And (Not GlbUtility.currStage Like "FT*") Then
                                                        If GlbUtility.currStage Like "CP*" And fieldstage Like "WLFT*" Then
                                                            B_Llimit(iItem) = 0
                                                            B_Hlimit(iItem) = 0
                                                        Else
                                                            If field.name Like "*x_coordinate" Then
                                                                If GlbUtility.currStage Like fieldstage Then
                                                                    B_Llimit(iItem) = XCoord
                                                                    B_Hlimit(iItem) = XCoord
                                                                Else
                                                                    B_Llimit(iItem) = CDbl(opbank.DsscXcoor)
                                                                    B_Hlimit(iItem) = CDbl(opbank.DsscXcoor)
                                                                End If
                                                                LoopLimitPrint(iItem) = True
                                                            ElseIf field.name Like "*y_coordinate" Then
                                                                If GlbUtility.currStage Like fieldstage Then
                                                                    B_Llimit(iItem) = YCoord
                                                                    B_Hlimit(iItem) = YCoord
                                                                Else
                                                                    B_Llimit(iItem) = CDbl(opbank.DsscYcoor)
                                                                    B_Hlimit(iItem) = CDbl(opbank.DsscYcoor)
                                                                End If
                                                                LoopLimitPrint(iItem) = True
                                                            ElseIf field.name Like "*wafer_id" Then
                                                                If GlbUtility.currStage Like fieldstage Then
                                                                    B_Llimit(iItem) = WaferID
                                                                    B_Hlimit(iItem) = WaferID
                                                                Else
                                                                    B_Llimit(iItem) = CDbl(opbank.DsscWfrStr)
                                                                    B_Hlimit(iItem) = CDbl(opbank.DsscWfrStr)
                                                                End If
                                                                LoopLimitPrint(iItem) = True
                                                            End If
                                                            
                                                            If LCase(field.BlowLocation) Like "wlft*" Then
                                                                m_WLFTEcidInfoStr = StrReverse(GlbUtility.Dec2Bin(field.DsscDecValue, field.size))
                                                                values = GlbUtility.Bin2Dec(m_WLFTEcidInfoStr)
                                                            End If
                                                        End If
                                                    ElseIf ((field.name = "bkm_package") Or (field.name = "bkm_process")) And (BdfDataBase.ReadRealBKMdone = True) Then
                                                        B_Llimit(iItem) = CInt(gS_BKM_IEDA)
                                                        B_Hlimit(iItem) = CInt(gS_BKM_IEDA)
                                                        LoopLimitPrint(iItem) = True
                                                    End If
                                             Else
                                                    'Compare Default Value is not 0 or Real Value in over 32 bits istuation
                                                    If (field.Llimit = field.Hlimit And field.Llimit <> 0) _
                                                        Or (field.Llimit <> field.Hlimit And field.Hlimit <> 0) Then
                                                            values = 0: field.Llimit = -1: field.Hlimit = -1
                                                            If GlbUtility.xHexCompare(field.DsscValue, field.xLlimit, ChkGreaterEqualThan) And _
                                                            GlbUtility.xHexCompare(field.DsscValue, field.xHlimit, ChkLessEqualThan) Then
                                                                values = -1
'                                                            Else
'                                                                    If GlbUtility.Hex2Dec(field.DsscValue) = 0 Then field.xLlimit = -1: field.xHlimit = -1: values = -1
                                                            End If
                                                    Else    'Compare Default Value is 0 in over 32 bits situation
                                                            values = 0
                                                            Strtmp = Replace(field.DsscValue, "0x", "")
                                                            Arraytmp = Split(Strtmp, "0")
                                                            If UBound(Arraytmp) <> Len(field.DsscValue) - 2 Then
                                                                values = -1
                                                            End If
                                                            
                                                            'values = field.DsscValue
                                                            'values = GlbUtility.Hex2Dbl(field.DsscValue)
                                                    End If
                                             End If
                                             oValue(iItem) = values
                                             m_unitType = unitNone
                                             m_ScaleType = scaleNone
                                             If field.SetWriteByBincut Then m_unitType = unitVolt: m_ScaleType = scaleMilli

                                             If Not asigned Then
                                                oHlimit(iItem) = field.Hlimit:  oLimit(iItem) = field.Llimit
                                                oName(iItem) = fieldstr:     oUnit(iItem) = m_unitType:   oScaleType(iItem) = m_ScaleType
                                             End If
                                             subItem = subItem + 1

                                             If (fieldstage = GlbUtility.currStage) Or GlbUtility.testedStages.Exists(fieldstage) Then
                                                 If (field.name Like "*y_coordinate") Then
                                                     mytname = "x_coordinate_plus_y_coordinate"
                                                     values = CDbl(opbank.DsscXcoor) + CDbl(opbank.DsscYcoor)
                                                     oValue(iItem + subItem) = values
                                                     If Not asigned Then
                                                         oHlimit(iItem + subItem) = -999:  oLimit(iItem + subItem) = 1
                                                         oName(iItem + subItem) = mytname:  oUnit(iItem + subItem) = m_unitType:  oScaleType(iItem + subItem) = scaleNone
                                                     End If
                                                     subItem = subItem + 1
                                                 End If
                                             End If
                                    End Select
                                    asigned = True
                            Next 'TheExec.sites
                End If
    Next

    iItem = iItem + subItem ' add the last item count
    If gB_eFuse_Disable_ChkLMT_Flag And subItem > 1 Then
        ReDim Preserve Disable_chk_arr(iItem - 1)
        'Handle extral subitems parts, 20210318
        If iItem - preiItem > 1 Then
            For i = preiItem + 1 To iItem - 1
                Disable_chk_arr(i) = Disable_chk_arr(preiItem)
            Next i
        End If
    End If
    If opbank.cmpFlag = True Then
        For i = 0 To iItem - 1 ' Limit all sites together to have better TT
            For Each site In theexec.sites
                theexec.Flow.TestLimit resultVal:=oValue(i), lowVal:=oCmpLlimit(i), hiVal:=oCmpHlimit(i), Tname:=oName(i), unit:=oUnit(i), scaletype:=oScaleType(i)
            Next
        Next
    Else
        For i = 0 To iItem - 1 ' Limit all sites together to have better TT
            If gB_eFuse_Disable_ChkLMT_Flag Then
                If Disable_chk_arr(i) = 0 Then
                    If LoopLimitPrint(i) = True Then
                        For Each site In theexec.sites
                            LLim = B_Llimit(i)
                            HLim = B_Hlimit(i)
                            theexec.Flow.TestLimit resultVal:=oValue(i), lowVal:=LLim, hiVal:=HLim, Tname:=oName(i), unit:=oUnit(i), scaletype:=oScaleType(i)
                        Next
                    ElseIf (oHlimit(i) = -999) Then
                        theexec.Flow.TestLimit resultVal:=oValue(i), lowVal:=oLimit(i), Tname:=oName(i), unit:=oUnit(i), scaletype:=oScaleType(i)
                    Else
                        theexec.Flow.TestLimit resultVal:=oValue(i), lowVal:=oLimit(i), hiVal:=oHlimit(i), Tname:=oName(i), unit:=oUnit(i), scaletype:=oScaleType(i)
                    End If
                Else
                    'if disable chk arr is 1,set HL limit value is ovalue(forcePass),20210318
                    For Each site In theexec.sites
                        LLim = oValue(i)
                        HLim = oValue(i)
                        theexec.Flow.TestLimit resultVal:=oValue(i), lowVal:=LLim, hiVal:=HLim, Tname:=oName(i), unit:=oUnit(i), scaletype:=oScaleType(i)
                    Next
                End If
            Else
                If LoopLimitPrint(i) = True Then
                    For Each site In theexec.sites
                        LLim = B_Llimit(i)
                        HLim = B_Hlimit(i)
                        theexec.Flow.TestLimit resultVal:=oValue(i), lowVal:=LLim, hiVal:=HLim, Tname:=oName(i), unit:=oUnit(i), scaletype:=oScaleType(i)
                    Next
                ElseIf (oHlimit(i) = -999) Then
                    theexec.Flow.TestLimit resultVal:=oValue(i), lowVal:=oLimit(i), Tname:=oName(i), unit:=oUnit(i), scaletype:=oScaleType(i)
                Else
                    theexec.Flow.TestLimit resultVal:=oValue(i), lowVal:=oLimit(i), hiVal:=oHlimit(i), Tname:=oName(i), unit:=oUnit(i), scaletype:=oScaleType(i)
                End If
            End If
        Next
    End If
skip:
    If False = opbank.NeedJTAGRead Then
        If ((Not earlyfuse And MultitypJob = "") Or IsInstMatch(NonDEID_Inst_Key)) Then
                If opbank.pgmMode = pgm_DAA Then
                       opbank.DumpDspWave4Stdf IIf(EFUSE_PRINT_DOUBLE_HEXMAP, opbank.DaaCapWaveSerialFL, opbank.DaaCapWaveSerial)
                Else
                       opbank.DumpDspWave4Stdf IIf(EFUSE_PRINT_DOUBLE_HEXMAP, opbank.JtagCapturedWave, opbank.JtagCapturedSerial)
                End If
        End If
    Else
        opbank.DumpDspWave4Stdf IIf(EFUSE_PRINT_DOUBLE_HEXMAP, opbank.JtagCapturedWave, opbank.JtagCapturedSerial)
    End If
GlbUtility.PersonalTimerLog "Limits Test"
    ''202004xx for ap
    'opbank.PrintReadFuseDatalog PrintDspWave, False, PrintDecode:=PrintDecode
GlbUtility.PersonalTimerLog "Print Out Cost"

Exit Function
errHandler:
    Call Print_Error_Message(Error_Info, "VBT_ZeFuseRead", "Bank_SyntaxCheck")
    If AbortTest Then Exit Function Else Resume Next
End Function

Public Function Bank_ReadWaferData(bankstr As String, Optional Validating_ As Boolean) As Long
On Error GoTo errHandler
Dim LotTmp As String, m_tmpwfid As String, tmpBinary As String, Loc_dash As Integer, iChr As Long, ascii As Long
Dim lotStrChk As Integer, lotidpf As New SiteBoolean, wfridpf As New SiteBoolean, xcoorpf As New SiteBoolean, ycoorpf As New SiteBoolean
Dim opbank As eFuseBdfBank, field As eFuseBdfField
Dim fusebankarr As Variant
Dim i As Long
Dim values As New SiteDouble
Dim site As Variant
Dim fieldstr As Variant
Dim dicLotInfo As New Dictionary

    'JackChou 202404 To allow Stage_FT with ChannelMap_CP case [Requester-TSMC-KoPei]
    'Take off the enableword of ReadWaferData in Flow_Table_Main_Init_Flows, so all stage would execute this function
    If (UCase(TheExec.CurrentChanMap) Like "CHANNELMAP_CP*" Or UCase(TheExec.CurrentChanMap) Like "CHANNELMAP_WL*") Then
        If UCase(TheExec.CurrentJob) Like "FT*" Then
            Call Print_Error_Message(Warning_Info, "VBT_ZeFuseRead", "Bank_ReadWaferData", "The Currrnt Job is " & TheExec.CurrentJob & " and the Current Channel Map is " & TheExec.CurrentChanMap & ". Please Check!!!")
            TheExec.AddOutput "<Warning> The Currrnt Job is " & TheExec.CurrentJob & " and the Current Channel Map is " & TheExec.CurrentChanMap & ". Please Check!!", vbBlue, True
        End If
        
    If Validating_ Then Validating_ = False: Exit Function         ' Exit after validation

    If bankstr = "" Then bankstr = "ECID"
    LotTmp = Trim(UCase(theexec.Datalog.Setup.LotSetup.lotid))
    m_tmpwfid = Trim(CStr(theexec.Datalog.Setup.WaferSetup.ID))

    If (LotTmp = "") Then
        LotTmp = "000000" ''''to avoid runtime VBT error stop
        theexec.Datalog.WriteComment "[WARNING] Input LotID is Empty, Set it to (000000). "
    End If
    
    Loc_dash = InStr(1, LotTmp, "-")
    If Loc_dash <> 0 Then LotTmp = mid(LotTmp, 1, Loc_dash - 1)
    
    'Checking for LotID string
    lotidpf = False
    If LotTmp = "000000" Then
        values = 0
    Else
        values = Len(LotTmp)
    End If
    lotid = LotTmp
     
    'Checking for Wafer ID string
    wfridpf = False
    fusebankarr = Split(bankstr, ",")
    For i = 0 To UBound(fusebankarr)
        xcoorpf = True: ycoorpf = True
        dicLotInfo.RemoveAll
        Set opbank = GetBdfBank(CStr(fusebankarr(i)))
        For Each fieldstr In opbank.Fields.keys
            Set field = opbank.Fields(fieldstr)
            If field.Algorithm = alg_lotid Then
                dicLotInfo.Add "lotid", fieldstr
            ElseIf field.Algorithm = alg_numeric Then
                If fieldstr Like "*wafer_id" Then
                    dicLotInfo.Add "waferid", fieldstr
                ElseIf LCase(fieldstr) Like "*x_coordinate" Then
                    dicLotInfo.Add "xcoordinate", fieldstr
                ElseIf LCase(fieldstr) Like "*y_coordinate" Then
                    dicLotInfo.Add "ycoordinate", fieldstr
                End If
            End If
        Next fieldstr
        Set field = opbank.Fields(dicLotInfo("lotid"))
        If (values = field.size / LotIdCharBits) Then lotidpf = True
        TheExec.Flow.TestLimit values, field.size / LotIdCharBits, field.size / LotIdCharBits, Tname:="Prober_LotID", PinName:=CStr(lotId)
        

        Set field = opbank.Fields(dicLotInfo("waferid"))
            If m_tmpwfid <> "" Then
                If (IsNumeric(m_tmpwfid) = True) Then
                    WaferID = CLng(m_tmpwfid)
                    If (field.Llimit <= WaferID And field.Hlimit >= WaferID) Then
                        wfridpf = True
                    Else
                        theexec.Datalog.WriteComment "Prober WaferID (" + CStr(WaferID) + ") is out of the range."
                    End If
                Else
                    theexec.Datalog.WriteComment "Prober WaferID (" + m_tmpwfid + ") is NOT numeric, set it to 0."
                    WaferID = 0
                End If
            Else
                WaferID = 0
            End If
            TheExec.Flow.TestLimit WaferID, field.Llimit, field.Hlimit, Tname:="Prober_WaferID", PinName:=CStr(WaferID)

        ''''---------- For CharZ Datalog --------
        HramLotId = lotId
        HramWaferId = WaferID
        ''''-------------------------------------
        Set field = opbank.Fields(dicLotInfo("xcoordinate"))
        For Each site In theexec.sites
            XCoord(site) = theexec.Datalog.Setup.WaferSetup.GetXCoord(site)
            YCoord(site) = theexec.Datalog.Setup.WaferSetup.GetYCoord(site)
            If (XCoord(site) < field.Llimit Or XCoord(site) > field.Hlimit) Then
                GlbUtility.WriteErrDlg "Prober X_Coordinate (" + CStr(XCoord(site)) + ") is out of the range [" + CStr(field.Llimit) + "..." + CStr(field.Hlimit) + "]."
                xcoorpf = False
            End If
        Next site
        
        If UCase(opbank.name) = "ECID" Then TheExec.Flow.TestLimit XCoord, field.Llimit, field.Hlimit, Tname:="Prober_X"
        
        Set field = opbank.Fields(dicLotInfo("ycoordinate"))
        For Each site In TheExec.sites
            If (YCoord(site) < field.Llimit Or YCoord(site) > field.Hlimit) Then
                GlbUtility.WriteErrDlg "Prober Y_Coordinate (" + CStr(YCoord(site)) + ") is out of the range [" + CStr(field.Llimit) + "..." + CStr(field.Hlimit) + "]."
                ycoorpf = False
            End If
        Next site
        
        If UCase(opbank.name) = "ECID" Then TheExec.Flow.TestLimit YCoord, field.Llimit, field.Hlimit, Tname:="Prober_Y"
        
        ''20210705, Modify for CP series to Set PRR
        If opbank.HadDeidFuse Then
            GenerateEcid opbank, lotId, lotidpf, WaferID, wfridpf, XCoord, xcoorpf, YCoord, ycoorpf, dicLotInfo
            
            If UCase(opbank.name) = "ECID" And (Not TheExec.CurrentJob Like "WLFT*") Then
                TheExec.Datalog.WriteComment ""
                For Each site In TheExec.sites
                        TheExec.Datalog.WriteComment "Site(" & site & ") Prober Hex code = " + opbank.EcidHexStr
                        Call TheExec.Datalog.Setup.SetPRRPartInfo(tl_SelectSite, site, , opbank.EcidHexStr)
                Next
                StdfPRR = True
            End If
        End If
    Next i

    ElseIf (UCase(TheExec.CurrentChanMap) Like "CHANNELMAP_FT*") And (UCase(TheExec.CurrentJob) Like "CP*" Or UCase(TheExec.CurrentJob) Like "WLFT*") Then '20240326
        Call Print_Error_Message(Warning_Info, "VBT_ZeFuseRead", "Bank_ReadWaferData", "The Currrnt Job is " & UCase(TheExec.CurrentJob) & " and the Current Channel Map is " & TheExec.CurrentChanMap & ".")
        TheExec.AddOutput "<Warning> The Currrnt Job is " & TheExec.CurrentJob & " and the Current Channel Map is " & TheExec.CurrentChanMap & ". Please Check!!", vbBlue, True
    Else
        If UCase(TheExec.CurrentChanMap) Like "CHANNELMAP_FT*" And UCase(TheExec.CurrentJob) Like "FT*" Then
            ' CurrentChanMap = FT* and CurrentJob = FT*
            Call Print_Error_Message(Warning_Info, "VBT_ZeFuseRead", "Bank_ReadWaferData", "The Currrnt Job is " & UCase(TheExec.CurrentJob) & " and the Current Channel Map is " & TheExec.CurrentChanMap & ".")
            TheExec.Flow.TestLimit 1, 1, 1, Tname:="Bank_ReadWaferData"
        Else
            Call Print_Error_Message(Error_Info, "VBT_ZeFuseRead", "Bank_ReadWaferData", "The Currrnt Job is " & UCase(TheExec.CurrentJob) & " and the Current Channel Map is " & TheExec.CurrentChanMap & ".")
            TheExec.Flow.TestLimit 0, 1, 1, Tname:="Bank_ReadWaferData"
        End If
    End If
    
Exit Function
errHandler:
    Call Print_Error_Message(Error_Info, "VBT_ZeFuseRead", "Bank_ReadWaferData")
    If AbortTest Then Exit Function Else Resume Next
End Function

'Check flow bin pass or fail
Public Function PassBinCheck()
On Error GoTo errHandler
Dim SiteVarValue As New SiteLong

    SiteVarValue = GlbUtility.BinCheck
    theexec.Flow.TestLimit resultVal:=SiteVarValue, lowVal:=1, hiVal:=1, Tname:="eFuse_is_BinOne?", PinName:="Value"

Exit Function
errHandler:
    Call Print_Error_Message(Error_Info, "VBT_ZeFuseRead", "PassBinCheck")
    If AbortTest Then Exit Function Else Resume Next
End Function

Private Function ObtainCatDictionary(opbank As eFuseBdfBank, ecid As Boolean, earlyfuse As Boolean, blankCheckAll As Boolean, _
            Optional RvOnly As Boolean = False, Optional DvOnly As Boolean = False, Optional MultitypJob As String = vbNullString) As Dictionary
On Error GoTo errHandler

    If blankCheckAll Then Set ObtainCatDictionary = Nothing:   Exit Function
    If opbank.DicForceFields.Count > 0 Then 'And Not earlyFuse Then
        Set ObtainCatDictionary = New Dictionary
        DicCloned ObtainCatDictionary, opbank.DicForceFields
        If earlyfuse Then
            DicSameKept ObtainCatDictionary, opbank.DicEarlyFused
        Else
            DicSameKept ObtainCatDictionary, opbank.DicOthers
        End If
        GlbUtility.WriteDlg "*** WARNING! The BDF contains the [force] algorithm~ Please check and confirm it, THANKS! ***"
        Exit Function
    End If

    If MultitypJob <> "" And opbank.DicMultiTypFields.Exists(MultitypJob) Then
        Set ObtainCatDictionary = opbank.DicMultiTypFields(MultitypJob)
        Exit Function
    End If
    If RvOnly Then Set ObtainCatDictionary = opbank.DicNonDefaultNonEarly:            Exit Function
    If DvOnly Then Set ObtainCatDictionary = opbank.DicDefaultIncludeEarly:           Exit Function
    Set ObtainCatDictionary = IIf(earlyfuse, opbank.DicEarlyFused, opbank.DicOthers): Exit Function

Exit Function
errHandler:
    Call Print_Error_Message(Error_Info, "VBT_ZeFuseRead", "ObtainCatDictionary")
    If AbortTest Then Exit Function Else Resume Next
End Function

Private Function OfflineDspWave(opbank As eFuseBdfBank, ecid As Boolean, earlyfuse As Boolean, blankCheck As Boolean, blankCheckAll As Boolean) As DSPWave
On Error GoTo errHandler
Dim mywave As New DSPWave
Dim sampleSize As Long

    If blankCheckAll Then
        Set mywave = FakeBankDsscResult(opbank, AteTrimData, ForceDoubleBit:=True, earlyfuse:=earlyfuse):
        sampleSize = GlbUtility.GetSampleSize(mywave)
        Set mywave = New DSPWave
        mywave.CreateConstant 0, sampleSize ' the above is used for sample size information only, set to 0 as blank
        Set OfflineDspWave = mywave
        Exit Function
    End If
    If ecid And Not blankCheck Then Set OfflineDspWave = FakeBankDsscResult(opbank, AteTrimData, True, ForceDoubleBit:=True, earlyfuse:=earlyfuse):                      Exit Function
    If blankCheck Then Set OfflineDspWave = FakeBankDsscResult(opbank, TrimmedData, ForceDoubleBit:=True, earlyfuse:=earlyfuse):                Exit Function
    Set OfflineDspWave = FakeBankDsscResult(opbank, AteTrimData, True, includeEarlyFuse:=True, ForceDoubleBit:=True, earlyfuse:=earlyfuse)   ' for read only

Exit Function
errHandler:
    Call Print_Error_Message(Error_Info, "VBT_ZeFuseRead", "OfflineDspWave")
    If AbortTest Then Exit Function Else Resume Next
End Function

Private Sub EarlyCheck4MultipleStage(opbank As eFuseBdfBank, Optional fieldStr As String = vbNullString)
On Error GoTo errHandler
Dim dicCfgTable As New Dictionary
Dim fieldstr_ As Variant

        If (LCase(opbank.name) = "config" Or LCase(opbank.name) = LCase(CfgBankName)) And opbank.HadEarlyFuse Then
            If fieldStr <> "" Then
                dicCfgTable.Add fieldStr, opbank.name
            Else
                Set dicCfgTable = opbank.DicCondTables
            End If
            If bLumpStages Then
                For Each fieldstr_ In dicCfgTable.Keys
                    If Not opbank.DicOthers.Exists(fieldstr_) Then opbank.DicOthers.Add fieldstr_, opbank.name
                Next
            End If
         End If

Exit Sub
errHandler:
    Call Print_Error_Message(Error_Info, "VBT_ZeFuseRead", "EarlyCheck4MultipleStage")
    If AbortTest Then Exit Sub Else Resume Next
End Sub

Private Function GetTName(earlyfuse As Boolean, ecid As Boolean) As String
On Error GoTo errHandler
Dim instance As String: instance = theexec.DataManager.instancename

    GetTName = "Unknow"
    If ecid Then
        GetTName = IIf(IsInstMatch(NonDEID_Inst_Key), NonDEID_Inst_Key, DEID_Inst_Key)
    Else
        GetTName = IIf(earlyfuse, "Early", "Post")
    End If

Exit Function
errHandler:
    Call Print_Error_Message(Error_Info, "VBT_ZeFuseRead", "GetTName")
    If AbortTest Then Exit Function Else Resume Next
End Function

''202004xx for ap
Public Function Bank_CompareWRData(bank As String, earlyfuse As Boolean, Optional RvOnly As Boolean, Optional MultitypJob As String = vbNullString, Optional Validating_ As Boolean)
On Error GoTo errHandler
Dim result As New SiteLong
Dim opbank As eFuseBdfBank
Dim m_pgmWave As New DSPWave
Dim m_ReadWave As New DSPWave
Dim m_SerialMode As Boolean
Dim m_MaskWave As New DSPWave
Dim instNameKey As String: instNameKey = vbNullString
Dim m_value As New SiteLong

    If Validating_ Then Validating_ = False: Exit Function   ' Exit after validation
    Set opbank = GetBdfBank(bank)

    If opbank.IsExistMultiTypeKey(theexec.DataManager.instancename, instNameKey) Or MultitypJob <> "" Then
        If Not opbank.IsMultiTypInst(instNameKey, MultitypJob) Then
            m_value = 1
            theexec.Flow.TestLimit resultVal:=m_value, lowVal:=0, hiVal:=0, Tname:="Multi-type_SettingError"
            Exit Function
        End If
    End If

    If earlyfuse Then
        m_pgmWave = opbank.DsscWave_Eff
        m_MaskWave = opbank.EarlyStageMask
    Else
        m_pgmWave = opbank.DsscWave_Eff
        If MultitypJob <> "" And opbank.StageMaskMultiTyp.Exists(MultitypJob) Then
            m_MaskWave = opbank.StageMaskMultiTyp(MultitypJob)
        Else
            m_MaskWave = opbank.StageMask
        End If
    End If

    If RvOnly Then
        If MultitypJob <> "" And opbank.MultiTyepRVReadWave.Exists(MultitypJob) Then
            m_ReadWave = opbank.MultiTyepRVReadWave(MultitypJob)
        Else
            m_ReadWave = opbank.RVReadWave
        End If
    Else
        If False = opbank.NeedJTAGRead And opbank.pgmMode = pgm_DAA Then
            m_ReadWave = opbank.DaaCapWaveSerial
        Else
            m_ReadWave = opbank.JtagCapturedSerial
        End If
    End If
    rundsp.CompareWRData m_pgmWave, m_ReadWave, m_MaskWave, result, RvOnly

    theexec.Flow.TestLimit resultVal:=result, lowVal:=0, hiVal:=0, Tname:="Read / Write Compare"
    
    Set m_MaskWave = Nothing
    Set m_ReadWave = Nothing
    Set m_pgmWave = Nothing
    If g_Rvenable Then g_Rvenable = False

Exit Function
errHandler:
    Call Print_Error_Message(Error_Info, "VBT_ZeFuseRead", "Bank_CompareWRData")
    If AbortTest Then Exit Function Else Resume Next
End Function

'''202006xx dsp modify
'Private Function DigCapWaveProcess(blankChk As Boolean, _
'                                   bank As eFuseBdfBank, _
'                                   ByRef capWave As DSPWave, _
'                                   result As SiteLong, _
'                                   Optional ByVal Stage As Boolean = True, _
'                                   Optional ByVal dicCategory As dictionary, _
'                                   Optional ForceDoubleBit As Boolean = False, _
'                                   Optional includeEarlyFuse As Boolean = False, _
'                                   Optional chkNonEarly As Boolean = False)
'On Error GoTo errHandler
'    Dim funcName As String: funcName = "DigCapWaveProcess"
'
'    Dim nBits As Long
'    Dim dicCared As New DSPWave
'    Dim stBit As Long, spBit As Long, iBit As Long
'    Dim fieldStr As Variant, field As eFuseBdfField
'    Dim dicCatCheck As dictionary
'    Dim paraWave As New DSPWave
'    Dim seriWave As New DSPWave
'    Dim seriWave_FL As New DSPWave
'    Dim name As String
'
'    nBits = IIf(bank.DoubleBits And ForceDoubleBit, 2, 1)
'    dicCared.CreateConstant 0, bank.FullSize
'
'    If blankChk Then
'        If Not dicCategory Is Nothing Then
'            Set dicCatCheck = New dictionary
'                For Each fieldStr In dicCategory.Keys
'                      dicCatCheck.Add fieldStr, False
'                Next
'        End If
'        'dicCared.CreateConstant 0, FullSize
'        For Each fieldStr In bank.Fields.Keys
'            Set field = bank.Fields(fieldStr)
'            bank.GetBitStSp field, stBit, spBit
'
'            If (Stage And (GlbUtility.currStage = field.BlowLocation Or bLumpStages)) Or (includeEarlyFuse And bank.DicEarlyFused.Exists(fieldStr)) Then
'                If dicCategory Is Nothing Then
'                            For iBit = stBit To spBit
'                                dicCared.Element(iBit) = 1
'                            Next
'                ElseIf dicCategory.Exists(fieldStr) Then
'                       If dicCategory(fieldStr) = bank.name Then
'                            dicCatCheck(fieldStr) = True
'                            For iBit = stBit To spBit
'                                dicCared.Element(iBit) = 1
'                            Next
'                      End If
'                End If
'            ElseIf Not dicCategory Is Nothing Then
'                If dicCategory.Exists(fieldStr) Then
'                        If dicCategory(fieldStr) = bank.name Then
'                            dicCatCheck(fieldStr) = True
'                            For iBit = stBit To spBit
'                                dicCared.Element(iBit) = 1
'                            Next
'                       End If
'                End If
'            Else
'                GlbUtility.MessageBox "Conditons is ambigious... Please confirm!"
'                GoTo errHandler
'            End If
'        Next
'        'inWave(0).Plot
'        If Not dicCatCheck Is Nothing Then
'        For Each fieldStr In bank.DicCondTables.Keys
'                If dicCategory.Exists(fieldStr) Then 'CfgEarlyTable
'                        Set field = bank.Fields(fieldStr)
'                        bank.GetBitStSp field, stBit, spBit
'                        For iBit = stBit To spBit
'                                If chkNonEarly And bLumpStages Then
'                                        dicCared.Element(iBit) = bank.BlankMaskDspWaveFutureStages.Element(iBit - stBit)
'                                Else
'                                        dicCared.Element(iBit) = bank.BlankMaskDspWave.Element(iBit - stBit)
'                                End If
'                        Next
'                End If
'        Next
'                For Each fieldStr In dicCategory.Keys
'                    If Not dicCatCheck.Exists(fieldStr) And dicCategory(fieldStr) = bank.name And fieldStr <> "" Then
'                        GlbUtility.MessageBox "Category specified didn't found! " & fieldStr
'                    End If
'                Next
'        End If
'    End If
'
'    ''------------------------------------------------------------------------------------
'    ''-Argument fill in--
'    ''- CapWave
'    ''- StageMaskWave
'    ''- Parallel bit size of Direct access mode
'    ''- Direct access mode / JTag mod
'    ''- 2-bit / 1-bit
'    ''-------------------------------------------------------------------------------------
'    Call rundsp.ProcessCapWaveAndCheck(capWave, dicCared, DaaParaBits, bank.pgmMode, nBits, bank.JtagMsbFirst, paraWave, seriWave, seriWave_FL, result)
'    'TheExec.Flow.TestLimit resultVal:=result, LowVal:=0, HiVal:=2, Tname:="test"
'    If bank.pgmMode = pgm_DAA Then
'        Set bank.DaaCapturedWave = paraWave
'        Set bank.DaaCapWaveSerial = seriWave
'        Set bank.DaaCapWaveSerialFL = seriWave_FL
'    Else
'        Set bank.JtagCapturedWave = capWave
'        Set bank.JtagCapturedSerial = seriWave
'    End If
'
'Exit Function
'errHandler:
'    GlbUtility.WriteDlg "<Error> " + funcName + ":: please check it out."
'    If AbortTest Then Exit Function Else Resume Next
'End Function

Public Function SplitDspWave(ReadPatSet As Pattern, PinRead As PinList, bank As String, sampleSize As Long, Optional InitPinsHi As PinList, Optional InitPinsLo As PinList, Optional InitPinsHiZ As PinList, Optional Validating_ As Boolean)
On Error GoTo errHandler
'Debug.Print vbCrLf & "<" & TheExec.DataManager.InstanceName & ">, SiteCnt = " & TheExec.sites.Selected.Count
Dim SignalName As String
Dim nBits As Long
Dim opbank As eFuseBdfBank, ReadPatt As String, PattArray() As String, PatCount As Long, capWave As New DSPWave

    If Validating_ Then
        If ReadPatSet.value <> "" Then Call PrLoadPattern(ReadPatSet.value)
        Exit Function    ' Exit after validation
    End If
    ''202005xx for ap
    Call RunDspSet
    
    Call PATT_GetPatListFromPatternSet(ReadPatSet.value, PattArray, PatCount)
    ReadPatt = PattArray(0)
    
    TheHdw.Digital.ApplyLevelsTiming True, True, True, tlPowered, InitPinsHi, InitPinsLo, InitPinsHiZ

    Set opbank = GetBdfBank(bank)
    Set opbank.CapturePin = PinRead
    SignalName = theexec.DataManager.instancename
    nBits = IIf(opbank.DoubleBits, 2, 1)
    sampleSize = sampleSize * nBits
    

    TheHdw.DSSC.Pins(opbank.CapturePin).Pattern(ReadPatt).Capture.Signals.Add (SignalName)
    With TheHdw.DSSC.Pins(opbank.CapturePin).Pattern(ReadPatt).Capture.Signals.item(SignalName)
            .sampleSize = sampleSize
            .LoadSettings
    End With

    capWave = TheHdw.DSSC.Pins(opbank.CapturePin).Pattern(ReadPatt).Capture.Signals(SignalName).DSPWave
    Call TheHdw.Patterns(ReadPatt).test(pfAlways, 0, tlResultModeDomain)
    
    Dic_SplitDspWave.Add SignalName, capWave
    opbank.SplitReadDspWaveFlag = True

GlbUtility.PersonalTimerLog "Print Out Cost"
Exit Function
SkipInstance:
    GlbUtility.WriteDlg theexec.DataManager.instancename & " was skipped!"
    Exit Function
errHandler:
    Call Print_Error_Message(Error_Info, "VBT_ZeFuseRead", "SplitDspWave")
    If AbortTest Then Exit Function Else Resume Next
End Function

Public Function Bank_Decode_Print(opbank As eFuseBdfBank, Optional printdecode As Boolean, Optional isEarly As Boolean, Optional dicTrimmed As Dictionary)
On Error GoTo errHandler
Dim funcName As String: funcName = "SplitDspWave"
Dim putWave2Db As New SiteBoolean:    putWave2Db = True
Dim site As Variant

    If PseudoFuseEnable And Not MixPseudoFuseEnable Then
        opbank.PutDsscRtToDb opbank.PseudoFuseDspWave, putWave2Db, dicTrimmed, isEarly:=isEarly
        If opbank.HadVddBinFuse Then
            Call ReadEfuseDataFromBinCut   '20210812, Add for TTR decode put dsscValue to MeasureValue
        ElseIf opbank.HadIdsFuse Then
            Call GetIdsValues
        End If
        opbank.PrintReadFuseDatalog False, False, printdecode:=True, dicItem4Print:=dicTrimmed, isEarly:=isEarly
    Else
        If False = opbank.NeedJTAGRead Then
            If opbank.pgmMode = pgm_DAA Then
                opbank.PutDsscRtToDb opbank.DaaCapWaveSerial, putWave2Db, dicTrimmed, isEarly:=isEarly
            Else
                opbank.PutDsscRtToDb opbank.JtagCapturedSerial, putWave2Db, dicTrimmed, isEarly:=isEarly
            End If
            If opbank.HadVddBinFuse Then
                Call ReadEfuseDataFromBinCut   '20210812, Add for TTR decode put dsscValue to MeasureValue
            ElseIf opbank.HadIdsFuse Then
                Call GetIdsValues
            End If
            opbank.PrintReadFuseDatalog False, False, printdecode:=printdecode, dicItem4Print:=dicTrimmed, isEarly:=isEarly
        End If
    End If
         
    If StdfPRR = False And IsInstMatch("ECID") Then
    'If StdfPRR = False And opbank.HadDeidFuse And Not IsInstMatch(NonDEID_Inst_Key) Then
       theexec.Datalog.WriteComment ""
           For Each site In theexec.sites
               theexec.Datalog.WriteComment "Site(" & site & ") Prober Hex code = " + opbank.EcidHexStr
               Call theexec.Datalog.Setup.SetPRRPartInfo(tl_SelectSite, site, , opbank.EcidHexStr)
               
               theexec.Datalog.WriteComment "DEVICE_CODE : " & HramLotId & "_W" _
                                                             & HramWaferId & "_X" _
                                                             & HramXCoord & "_Y" _
                                                             & HramYCoord & "_S" _
                                                             & site
           Next
           StdfPRR = True
    End If

Exit Function
errHandler:
    Call Print_Error_Message(Error_Info, "VBT_ZeFuseRead", "Bank_Decode_Print")
    If AbortTest Then Exit Function Else Resume Next
End Function
