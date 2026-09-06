Attribute VB_Name = "VBT_ZeFuseWrite"
Option Explicit

'' either DAA or JTAG mode, all data sources are in serial mode, so it doesn't matter to use the same dssc wave
'''''if run ecid, Your instance name should be put "_DEID" or "_nonDeid" to be differentiated, not case senstive.
Public Function Bank_Write(WritePatSet As Pattern, PinWrite As PinList, bank As String, earlyfuse As Boolean, ecid As Boolean, _
                    PwrPin As String, vpwr As Double, Optional RvOnly As Boolean, Optional printdecode As Boolean, Optional PrintDspWave As Boolean, Optional InitPinsHi As PinList, Optional InitPinsLo As PinList, Optional InitPinsHiZ As PinList, _
                    Optional PrePatSet As Pattern, Optional DigSource As String = vbNullString, Optional MultitypJob As String = vbNullString, Optional Validating_ As Boolean) 'RvOnly =>DefaultOrReal = Real, set these fields' bits to dssc source only.
On Error GoTo errHandler
Dim dicTrimmed As New Dictionary
Dim opbank As eFuseBdfBank, cmpbank As eFuseBdfBank, SiteVarValue As New SiteLong, SignalName As String, cmpStr As String
'20210714, Modify for Complement Fuse
Dim cmpfield As Variant
Dim stBit As Long, spBit As Long, iBit As Long
Dim instNameKey As String: instNameKey = vbNullString
Dim m_value As New SiteLong
Dim NoTrimMissingFlag As New SiteBoolean: NoTrimMissingFlag = False
Dim StoreSiteStatus As New SiteBoolean: StoreSiteStatus = False
Dim WritePat As String, PattArray() As String, PatCount As Long, TrimMissingCnt As New SiteLong, verRptCnt As Long, dicRowNum As DSPWave
Dim PrePatResult As New SiteBoolean
Dim srcWave As New DSPWave
Dim tmpDSP As New DSPWave
Dim tmpDSP2 As New DSPWave
Dim BySite As Boolean
Dim site As Variant
Dim initSiteStatus As New SiteBoolean
    
    If Validating_ Then
        If WritePatSet.value <> "" Then Call PrLoadPattern(WritePatSet.value)
        If PrePatSet.value <> "" Then Call PrLoadPattern(PrePatSet.value)
        Exit Function    ' Exit after validation
    End If
'     '******Obtaining the Read Pattern Name and eFuse Bank******

    '20210812,Add for enable word control printing
    If gB_eFuse_Disable_DecodeDataPrint_Flag = True Then printdecode = False
    If gB_eFuse_Disable_DSPwavePrint_Flag = True Then PrintDspWave = False
    
    ''202005xx for ap
    Call RunDspSet
    
    If (PrePatSet.value <> "") And ((Not PseudoFuseEnable) Or MixPseudoFuseEnable) Then
        TheHdw.Digital.ApplyLevelsTiming True, True, True, tlPowered, InitPinsHi, InitPinsLo, InitPinsHiZ
        initSiteStatus = TheExec.sites.Active
        PrePatResult = EfuseExecInitPattern(PrePatSet.value, DigSource)
        If Not PrePatResult.Any(True) Then
            Exit Function
        ElseIf PrePatResult.Any(False) Then
            TheExec.sites.Selected = PrePatResult
        End If
    End If
    Call PATT_GetPatListFromPatternSet(WritePatSet.value, PattArray, PatCount)
    
    If RvOnly Then g_Rvenable = True
    'PrintDspWave = False
    
    ''202004xx for ap
    WritePat = PattArray(0)
    Set opbank = GetBdfBank(bank)

    If opbank.IsExistMultiTypeKey(TheExec.DataManager.instancename, instNameKey) Or MultitypJob <> "" Then
        If Not opbank.IsMultiTypInst(instNameKey, MultitypJob) Then
            m_value = 1
            TheExec.flow.TestLimit resultVal:=m_value, lowVal:=0, hiVal:=0, Tname:="Multi-type_SettingError"
            Exit Function
        End If
    End If
    '20210428 Clear EcidBinStr when writeDSSC happened(BankRead might read 0 value)
    If opbank.HadDeidFuse Then
        opbank.EcidBinStr = vbNullString
    End If
    
    '20210910, Add for cmplfuse create realfields' mask and do setefuse
    If (gB_CmpFuseEnable) = True Then
        cmplfuse_createrealfielmask_setfuse opbank
    End If
    
    
GlbUtility.TimerStarter
    Set dicTrimmed = ObtainCatDictionary(opbank, ecid, earlyfuse, RvOnly, MultitypJob)

'    '****** Check early fused status******
    If earlyfuse Then
        If Not opbank.HadEarlyFuse Then GoTo errHandler:
        If opbank.IsEarlyFused Then GlbUtility.MessageBox bank & " bank had already done early fused! Please confirmed!": GoTo errHandler:
        If LCase(bank) = "config" Or LCase(bank) = LCase(CfgBankName) Then
            Dim field As New eFuseBdfField, fieldStr As Variant
             
            For Each fieldStr In opbank.DicCondTables.Keys
                Set field = opbank.Fields(fieldStr)
                If field.IsEarlyFuse Then Call UpdateCfg2BdfEfuse(CfgDataBaseAct.CfgValue(fieldStr), AteTrimData, field.name)
            Next
                        
        End If
    Else
        Call SetCustomizeFuse(opbank, dicTrimmed)
        
        ''202004xx for ap
        If LCase(bank) <> "ecid" Then Call EarlyCheck4MultipleStage(opbank)
        
    End If
GlbUtility.PersonalTimerLog "SetCustomizeFuse"
    If dicTrimmed.Count = 0 Then GoTo SkipInstance:
    
    If Not GlbUtility.OnlineMode Then
        opbank.AnyWriteRun = True
        If LCase(opbank.name) Like "udr*" Then
        cmpStr = Replace(LCase(opbank.name), "udr", "cmp") 'udr_p => cmp_p
        Set cmpbank = GetBdfBank(cmpStr)
        cmpbank.AnyWriteRun = True
        End If
    End If
    'If Not GlbUtility.IsOnline Then opbank.AnyWriteRun = True

    If Not GlbUtility.OnlineMode Then Call opbank.PseudoFusedFillup(dicTrimmed)
    
    If ecid And dicTrimmed.Exists("ecid_crc") Then
    'If ecid And IsInstMatch(NonDEID_Inst_Key) Then 'ecid & instance name like the key
        opbank.CalculateCRC "ECID_CRC"
    End If

    If Not (IsInstMatch(NonDEID_Inst_Key)) And Not earlyfuse And (IIf(MultitypJob <> "", DicTrimmedExistCrc(dicTrimmed, opbank.Crcs), True)) Then
GlbUtility.PersonalTimerLog "BeforeCalculateCRC"
     '''Put random and crc values
            'If Not GlbUtility.IsOnline Then Call opbank.PseudoFusedFillup(dicTrimmed)
            If opbank.HadRandomFuse Then
                Call opbank.PutRandomCodes
                
                opbank.CalculateCRC '"uid_crc" 'ecid_crc will also update @this step, if SeparateEarly=false
                
            ElseIf opbank.HadCrcFuse And Not ecid Then
                opbank.CalculateCRC
                
            End If
GlbUtility.PersonalTimerLog "CalculateCRC"
    End If
    
    If opbank.HadCrcCompFuse = True And Not earlyfuse And (IIf(MultitypJob <> "", DicTrimmedExistCrc(dicTrimmed, opbank.Crcs_Comp), True)) Then
        opbank.CalculateCRCComp
    End If

    '''Missing Trim Count Check
    StoreSiteStatus = TheExec.sites.Selected
    TrimMissingCnt = opbank.TrimedMissingCnt(dicTrimmed)
    
    TheExec.flow.TestLimit resultVal:=TrimMissingCnt, lowVal:=0, hiVal:=0, Tname:="TrimMissingCnt", PinName:="Value"
    verRptCnt = -1 '-1 will use default
GlbUtility.PersonalTimerLog "Pre fuse"
    BySite = True: SiteVarValue = 0
    For Each site In TheExec.sites
        If TrimMissingCnt(site) = 0 Then NoTrimMissingFlag(site) = True
    Next site

    TheExec.sites.Selected = NoTrimMissingFlag
    If NoTrimMissingFlag.Any(True) Then
    opbank.IsBlank = 0
                If earlyfuse And Not ecid Then
                    BySite = False ' false means that all sites source the same data
'                    If opbank.HadDeidFuse Then '20231220 Other bank had deidfuse
                    If opbank.HadDeidFuse Then '20231220 Other bank had deidfuse
                        For Each fieldStr In dicTrimmed
                            Set field = opbank.Fields(fieldStr)
                            If field.Algorithm = alg_lotid Or field.Algorithm = alg_numeric Then
                                TheExec.flow.TestLimit resultVal:=-1, lowVal:=0, hiVal:=0, Tname:=bank + "_HadDeidFuse"
                                Call Print_Error_Message(Error_Info, "VBT_ZeFuseWrite", "Bank_Write", "Please check this instance ecid argument.")
                                GoTo SkipInstance
                            End If
                        Next
                    End If
'                    End If
                    Set srcWave = opbank.EarlyDsscSourcedWave
                    
                    If PrintDspWave Then Set opbank.SingleBitWave = opbank.EarlySingleBitWave ' Just for print requirement, no use!
                    
                Else
                    If ecid Then
                        Set srcWave = opbank.GetDsscWave(dicTrimmed, useOthersDefault:=False)
                        
                    Else
                        Set srcWave = opbank.GetDsscWave(dicTrimmed, useOthersDefault:=True, RvOnly:=RvOnly, dicRowNum:=dicRowNum, MultiType:=MultitypJob)
                        
                    End If
                    Set opbank.DsscWave_Eff = srcWave
                    
                    '20210910, Add cmplfuse modify wave to avoid double fusing
                    If (gB_CmpFuseEnable) = True Then
                        cmplfuse_modifydspwave_avoiddoublefusing opbank, dicTrimmed
                    End If
                    
                    'Dim o As New DSPWave: Ze_GetPartialWave SrcWave, 0, 128, o  'if you had to split pattern then use this to obtain the dssc waveform, e.g. Cebu
                    ''202004xx for ap
                    If RvOnly Then verRptCnt = 1
                    'If RvOnly Then verRptCnt = 5
                    Set srcWave = opbank.GenDsscSrcDspWave(True, True, verRptCnt)
                    
                End If
GlbUtility.PersonalTimerLog "Retrieve Dssc SrcWave"
                If RvOnly Then PrintDspWave = False ' Print the rows which are real codes included, instead of print all rows
                If opbank.DoubleBits = True Then
                    opbank.PrintWriteFuseDatalog dicTrimmed, printdecode:=printdecode, PrintDoubleBitWave:=PrintDspWave, CompactPrint:=True, Early:=earlyfuse  ' write data pase and print out  the result
                Else
                
                    opbank.PrintWriteFuseDatalog dicTrimmed, printdecode:=printdecode, PrintSingleBitWave:=PrintDspWave, CompactPrint:=True, Early:=earlyfuse  ' write data pase and print out  the result
                End If
                
                If RvOnly And Not gB_eFuse_Disable_DSPwavePrint_Flag Then
                    Dim paraBit As Long: paraBit = IIf(opbank.DoubleBits, 16, 32)
                    opbank.DumpDspWave opbank.SingleBitWave, TheExec.DataManager.instancename, paraBit, dicRowNum
                    
                End If
                Call opbank.UpdateTrimmed(dicTrimmed, Not earlyfuse) ' To mark it's already trimed
                
                'BdfDataBase.DumpBdfData BankDumpPath 'Just for debug
GlbUtility.PersonalTimerLog "Print Out Cost"
'****************************READY TO BELOW THE FUSE****************************
                If PrePatSet.value = "" Then TheHdw.Digital.ApplyLevelsTiming True, True, True, tlPowered, InitPinsHi, InitPinsLo, InitPinsHiZ
                ''202005xx for ap
                If (Not PseudoFuseEnable) Or MixPseudoFuseEnable Then
                'vpwr = 0
                    If EFUSE_POWER_OFF_SETTING Then vpwr = 0
                    If (PwrPin <> "") Then Call TurnOnEfusePwrPins(PwrPin, vpwr)
                    'If (PwrPin <> "" And vpwr <> 0) Then Call TurnOnEfusePwrPins(PwrPin, vpwr)
    GlbUtility.PersonalTimerLog "TurnOnEfusePwrPins"
                    'If GlbUtility.OnlineMode Or earlyfuse Then
                    'If GlbUtility.IsOnline Or earlyfuse Then  'set earlyfuse/Wireless here is for offline trial test only
                            Set opbank.SourcePin = PinWrite
                            
                            ''202004xx for ap
                            Burn_With_MultiWritePat PattArray, opbank, srcWave, BySite, RvOnly
                            DebugPrintFunc Join(PattArray, ",")
                            'SignalName = "SignalSource"
                            'eFuse_DSSC_SrcSetup WritePat, opbank, SignalName, srcWave, BySite
    GlbUtility.PersonalTimerLog "eFuse_DSSC_SrcSetup"
                            'Call TheHdw.Patterns(WritePat).test(pfAlways, 0) 'run write pattern and source
                    'End If
                ''202004xx for ap
                If (PwrPin <> "") Then Call TurnOffEfusePwrPins(PwrPin, vpwr)
                End If
'****************************FINISHED*******************************************
GlbUtility.PersonalTimerLog "TurnOffEfusePwrPins"
            SiteVarValue = 1
    End If
    
    If PseudoFuseEnable And Not MixPseudoFuseEnable Then
        For Each site In TheExec.sites
        tmpDSP2 = opbank.DsscWave_Eff.ConvertDataTypeTo(DspLong)
        tmpDSP = opbank.PseudoFuseDspWave
        Next
        For Each site In TheExec.sites
            If tmpDSP2.SampleSize = tmpDSP.SampleSize Then
                tmpDSP = tmpDSP.BitwiseOr(tmpDSP2)
            End If
        Next

        Set opbank.PseudoFuseDspWave = tmpDSP
    End If

    TheExec.flow.TestLimit resultVal:=SiteVarValue, lowVal:=1, hiVal:=1, Tname:=bank + "_Fuse", PinName:="Value"
    TheExec.sites.Selected = StoreSiteStatus
Exit Function
SkipInstance:
    GlbUtility.WriteDlg TheExec.DataManager.instancename & " was skipped!"
    Exit Function
errHandler:
    Call Print_Error_Message(Error_Info, "VBT_ZeFuseWrite", "Bank_Write")
    If AbortTest Then Exit Function Else Resume Next
End Function

Public Function Bank_LoadDefaultValues(bank As String, Optional userSpecified As Boolean, Optional Validating_ As Boolean)
On Error GoTo errHandler
Dim opbank As eFuseBdfBank

    If Validating_ Then Exit Function
    Set opbank = GetBdfBank(bank)
    Call opbank.PutDefaultVauleToReal(userSpecified)
    TheExec.flow.TestLimit resultVal:=1, lowVal:=1, hiVal:=1, Tname:="LoadDefaultValues", PinName:="Value"

Exit Function
errHandler:
    Call Print_Error_Message(Error_Info, "VBT_ZeFuseWrite", "Bank_LoadDefaultValues")
    If AbortTest Then Exit Function Else Resume Next
End Function

Public Function Bank_TrimmedUpdate(bank As String, earlyfuse As Boolean, ecid As Boolean, Optional RvOnly As Boolean, Optional MultitypJob As String = vbNullString, Optional Validating_ As Boolean)
On Error GoTo errHandler
Dim dicTrimmed As New Dictionary
Dim opbank As eFuseBdfBank

    If Validating_ Then Exit Function
    Set opbank = GetBdfBank(bank)
    Set dicTrimmed = ObtainCatDictionary(opbank, ecid, earlyfuse, RvOnly, MultitypJob)
    Call opbank.UpdateTrimmed(dicTrimmed, Not earlyfuse) ' To mark it's already trimed

Exit Function
errHandler:
    Call Print_Error_Message(Error_Info, "VBT_ZeFuseWrite", "Bank_TrimmedUpdate")
    If AbortTest Then Exit Function Else Resume Next
End Function

Private Sub SetCustomizeFuse(opbank As eFuseBdfBank, dicTrimmed As Dictionary)
On Error GoTo errHandler
    
    If UCase(opbank.name) = "CFG" And dicTrimmed.Exists("efuse_database_revision") Then
        opbank.SetDefaultVauleToRealField "efuse_database_revision"
    End If
    
    '''''''''''''''''20231201 The PreWrite function is already set ids value'''''''''''''''''
'    If Not GlbUtility.OnlineMode And UCase(GlbUtility.currStage) <> "CP1" Then
'        opbank.PutIdsCodes
'    End If
    
Exit Sub
errHandler:
    Call Print_Error_Message(Error_Info, "VBT_ZeFuseWrite", "SetCustomizeFuse")
    If AbortTest Then Exit Sub Else Resume Next
End Sub

'This will need to do customize while early fuse checking is required
Private Function EarlyCheck4MultipleStage(opbank As eFuseBdfBank, Optional fieldStr As String = vbNullString) As Boolean
On Error GoTo errHandler
Dim dicCfgTable As New Dictionary 'CfgEarlyTable
Dim fieldBdfStr As Variant
EarlyCheck4MultipleStage = False
        If (LCase(opbank.name) = "config" Or LCase(opbank.name) = LCase(CfgBankName)) And (opbank.HadEarlyFuse Or Not opbank.SeparateEarly) Then
                If fieldStr <> "" Then
                    dicCfgTable.Add fieldStr, opbank.name
                Else
                    Set dicCfgTable = opbank.DicCondTables
                End If
                If CfgDataBaseAct.DicFuseStages.Exists(GlbUtility.currStage) Then ' And (GlbUtility.testedStages.Count > 0 Or bLumpStages) Then
                
                Dim field As New eFuseBdfField
                        For Each fieldBdfStr In dicCfgTable.Keys
                            Set field = opbank.Fields(fieldBdfStr)
                            If bLumpStages And CfgDataBaseAct.EarlyStage Then
                                    Call UpdateCfg2BdfEfuse(CfgDataBaseAct.CfgValueFutureStages(fieldBdfStr), AteTrimData, field.name)
                            ElseIf bLumpStages And Not opbank.SeparateEarly Then
                                    Call UpdateCfg2BdfEfuse(CfgDataBaseAct.CfgCmpValueLumpStages(fieldBdfStr), AteTrimData, field.name)
                            ''20200519
                            ElseIf Not field.IsEarlyFuse And field.BlowLocation = GlbUtility.currStage Then

                            'Else
                                    Call UpdateCfg2BdfEfuse(CfgDataBaseAct.CfgValue(fieldBdfStr), AteTrimData, field.name)
                            End If
                            '20200519
                            'If bLumpStages Or Not (GlbUtility.testedStages.Count = 0) Then
                            If bLumpStages Or (Not field.IsEarlyFuse And field.BlowLocation = TheExec.CurrentJob) Then
                            'If bLumpStages Or Not (Not field.IsEarlyFuse And GlbUtility.testedStages.Count = 0) Then
                                If Not opbank.DicOthers.Exists(fieldBdfStr) Then opbank.DicOthers.Add fieldBdfStr, opbank.name
                                field.Trimmed = False
                            End If
                            EarlyCheck4MultipleStage = True
                        Next
                Else
                        For Each fieldBdfStr In dicCfgTable.Keys
                            If opbank.DicOthers.Exists(fieldBdfStr) Then opbank.DicOthers.Remove fieldBdfStr
                        Next
                End If
         End If
         
Exit Function
errHandler:
    Call Print_Error_Message(Error_Info, "VBT_ZeFuseWrite", "EarlyCheck4MultipleStage")
    
        If AbortTest Then Exit Function Else Resume Next

End Function

Private Sub OfflineFakeTrim(opbank As eFuseBdfBank, fieldStr As String, v As Double, pf As Boolean)
On Error GoTo errHandler
Dim trimRt As New SiteVariant, PatPF As New SiteBoolean
       PatPF = True
       trimRt = v
       Call opbank.SetEfuse(fieldStr, trimRt, PatPF)
       
Exit Sub
errHandler:
    Call Print_Error_Message(Error_Info, "VBT_ZeFuseWrite", "OfflineFakeTrim")
  
        If AbortTest Then Exit Sub Else Resume Next

End Sub

Public Function ObtainCatDictionary(opbank As eFuseBdfBank, ecid As Boolean, earlyfuse As Boolean, RvOnly As Boolean, Optional MultitypJob As String = vbNullString) As Dictionary
On Error GoTo errHandler
    
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
    If RvOnly Then Set ObtainCatDictionary = opbank.DicNonDefaultNonEarly():          Exit Function
    Set ObtainCatDictionary = IIf(earlyfuse, opbank.DicEarlyFused, opbank.DicOthers): Exit Function

Exit Function
errHandler:
    Call Print_Error_Message(Error_Info, "VBT_ZeFuseWrite", "ObtainCatDictionary")
    If AbortTest Then Exit Function Else Resume Next
End Function

Private Function TurnOnEfusePwrPins(powerPin As String, _
                                   Optional v As Double = 1.8, _
                                   Optional i_rng As Double = 0.2, _
                                   Optional wait_before_gate As Double = 0.001, _
                                   Optional wait_after_gate As Double = 0.002, _
                                   Optional Steps As Integer = 10, _
                                   Optional RiseTime As Double = 0.001)

On Error GoTo errHandler
Dim m_currVolt As Double
Dim stringTypes() As String
Dim typesCount As Long
Dim stringType As String
Dim m_PowerPinArr() As String
    
    ''202004xx for ap
    m_PowerPinArr = Split(powerPin, ",")
    
    Call TheExec.DataManager.GetChannelTypes(m_PowerPinArr(0), typesCount, stringTypes())
    stringType = Trim(UCase(stringTypes(0)))
    If (stringType Like "DCVS*") Then 'DCVS
        eFuse_pwr_Ramp_meter_DCVS True, powerPin, v, i_rng, wait_before_gate, wait_after_gate, Steps, RiseTime ', True
        m_currVolt = TheHdw.DCVS.pins(powerPin).Voltage.Main.value
    ElseIf (stringType Like "DCVI*") Then 'DCVI
        eFuse_pwr_Ramp_meter_DCVI True, powerPin, v, i_rng, wait_before_gate, wait_after_gate, Steps, RiseTime ', True
        m_currVolt = TheHdw.DCVI.pins(powerPin).Voltage
    Else
        Call Print_Error_Message(Error_Info, "VBT_ZeFuseWrite", "TurnOnEfusePwrPins", "Please check pin (" + powerPin + ") type is DCVI or DCVS")
        GoTo errHandler
    End If
    
    TheExec.Datalog.WriteComment "TurnOnEfusePwrPins :: " + UCase(powerPin) + " = " + Format(m_currVolt, "0.000")
Exit Function
errHandler:
    Call Print_Error_Message(Error_Info, "VBT_ZeFuseWrite", "TurnOnEfusePwrPins")
   
        If AbortTest Then Exit Function Else Resume Next

End Function

Private Function TurnOffEfusePwrPins(powerPin As String, _
                                   Optional v As Double = 1.8, _
                                   Optional i_rng As Double = 0.2, _
                                   Optional wait_before_gate As Double = 0.001, _
                                   Optional wait_after_gate As Double = 0.002, _
                                   Optional Steps As Integer = 10, _
                                   Optional RiseTime As Double = 0.001)

On Error GoTo errHandler
Dim m_currVolt As Double
Dim stringTypes() As String
Dim typesCount As Long
Dim stringType As String
Dim m_PowerPinArr() As String

    ''202004xx for ap
    m_PowerPinArr = Split(powerPin, ",")
    
    Call TheExec.DataManager.GetChannelTypes(m_PowerPinArr(0), typesCount, stringTypes())
    stringType = Trim(UCase(stringTypes(0)))
    If (stringType Like "DCVS*") Then 'DCVS
        eFuse_pwr_Ramp_meter_DCVS False, powerPin, v, i_rng, wait_before_gate, wait_after_gate, Steps, RiseTime ', True
        m_currVolt = TheHdw.DCVS.pins(powerPin).Voltage.Main.value
    ElseIf (stringType Like "DCVI*") Then 'DCVI
        eFuse_pwr_Ramp_meter_DCVI False, powerPin, v, i_rng, wait_before_gate, wait_after_gate, Steps, RiseTime ', True
        m_currVolt = TheHdw.DCVI.pins(powerPin).Voltage
    Else
        Call Print_Error_Message(Error_Info, "VBT_ZeFuseWrite", "TurnOffEfusePwrPins", "Please check pin (" + powerPin + ") type is DCVI or DCVS")
        GoTo errHandler
    End If
    
    TheExec.Datalog.WriteComment "TurnOffEfusePwrPins :: " + UCase(powerPin) + " = " + Format(m_currVolt, "0.000")
Exit Function
errHandler:
    Call Print_Error_Message(Error_Info, "VBT_ZeFuseWrite", "TurnOffEfusePwrPins")
    
        If AbortTest Then Exit Function Else Resume Next

End Function

Private Function eFuse_pwr_Ramp_meter_DCVS(PwrOn As Boolean, pin As String, v As Double, i_rng As Double, _
                                               wait_before_gate As Double, wait_after_gate As Double, _
                                               Steps As Integer, RiseTime As Double, _
                                               Optional showPrint As Boolean = False)

On Error GoTo errHandler
Dim i_meter_rng As Double, setV As Double, StepV As Double, Vstart As Double, Vstring As String
Dim i As Integer, stepT As Double
    'set Force voltage and Current/Meter Range'set Force voltage and Current/Meter Range
    ''========================================''=================================================
    ''Description: __                __       ''Description
    ''                __|                     ''__                              __
    ''            __|                         ''  |__
    ''        __|                             ''     |_>|  |<--stepT             v
    ''    __| >|   |<--stepT  __        v     ''        |__          __
    ''__|                   __ stepV   __     ''           |__       __ stepV   __
    ''|<-- steps -->                          ''|<-- steps -->
    ''|<-FallTime ->                          ''|<-FallTime ->
    ''========================================''=================================================

    i_meter_rng = i_rng

    If PwrOn Then
        StepV = v / Steps
        Vstart = 0
        Vstring = " Pwr Up Voltage "
    Else
        StepV = -v / Steps
        Vstart = v
        Vstring = " Pwr Down Voltage "
    End If

    stepT = RiseTime / Steps

    With TheHdw.DCVS.pins(pin)
        .Connect
        .mode = tlDCVSModeVoltage
        .Voltage.Main = Vstart
        .SetCurrentRanges i_rng, i_meter_rng
'        .CurrentLimit.Source.FoldLimit.Level = i_rng
        .Meter.mode = tlDCVSMeterCurrent
        .CurrentRange.value = i_rng
        .CurrentLimit.Source.FoldLimit.level.value = i_rng
        If glb_TesterType = "Jaguar" Then .Meter.CurrentRange = i_rng
        TheHdw.Wait wait_before_gate   'wait for relay connect
        .Gate = True
    End With
    
    ''Pwr On/Off Ramp up slew-rate control============================
    For i = 1 To Steps
        setV = Vstart + i * StepV
        TheHdw.DCVS.pins(pin).Voltage.Main = setV
        If showPrint = True Then
            TheExec.Datalog.WriteComment "  Curr_" & pin & ", " & Vstring & "(" & CStr(i) & ") : " & CStr(setV) & " V"
        End If
        TheHdw.Wait stepT
    Next i
    ''============================================================
    TheHdw.Wait wait_after_gate
    
Exit Function
errHandler:
    Call Print_Error_Message(Error_Info, "VBT_ZeFuseWrite", "eFuse_pwr_Ramp_meter_DCVS")
    
        If AbortTest Then Exit Function Else Resume Next

End Function

Private Function eFuse_pwr_Ramp_meter_DCVI(PwrOn As Boolean, pin As String, v As Double, i_rng As Double, _
                                               wait_before_gate As Double, wait_after_gate As Double, _
                                               Steps As Integer, RiseTime As Double, _
                                               Optional showPrint As Boolean = False)

On Error GoTo errHandler
Dim i_meter_rng As Double, setV As Double, StepV As Double, Vstart As Double, Vstring As String
Dim i As Integer, stepT As Double
    'set Force voltage and Current/Meter Range'set Force voltage and Current/Meter Range
    ''========================================''=================================================
    ''Description: __                __       ''Description
    ''                __|                     ''__                              __
    ''            __|                         ''  |__
    ''        __|                             ''     |_>|  |<--stepT             v
    ''    __| >|   |<--stepT  __        v     ''        |__          __
    ''__|                   __ stepV   __     ''           |__       __ stepV   __
    ''|<-- steps -->                          ''|<-- steps -->
    ''|<-FallTime ->                          ''|<-FallTime ->
    ''========================================''=================================================

    i_meter_rng = i_rng

    If PwrOn Then
        StepV = v / Steps
        Vstart = 0
        Vstring = " Pwr Up Voltage "
    Else
        StepV = -v / Steps
        Vstart = v
        Vstring = " Pwr Down Voltage "
    End If

    stepT = RiseTime / Steps

    With TheHdw.DCVI.pins(pin)
        .Connect
        .mode = tlDCVIModeVoltage
        .Voltage = Vstart
        .SetCurrentAndRange i_rng, i_meter_rng
        .Meter.mode = tlDCVIMeterCurrent
        .CurrentRange.value = i_rng
        .FoldCurrentLimit.TimeOut = 0.01
        .FoldCurrentLimit.Behavior = tlDCVIFoldCurrentLimitBehaviorGateOff
        If glb_TesterType = "Jaguar" Then .Meter.CurrentRange = i_rng
        .ComplianceRange(tlDCVICompliancePositive).value = 7
        .ComplianceRange(tlDCVIComplianceNegative).value = -2
        
        TheHdw.Wait wait_before_gate   'wait for relay connect
        .Gate = True
    End With
    
    ''Pwr On/Off Ramp up slew-rate control============================
    For i = 1 To Steps
        setV = Vstart + i * StepV
        TheHdw.DCVI.pins(pin).Voltage = setV
        If showPrint = True Then
            TheExec.Datalog.WriteComment "  Curr_" & pin & ", " & Vstring & "(" & CStr(i) & ") : " & CStr(setV) & " V"
        End If
        TheHdw.Wait stepT
    Next i
    ''============================================================
    TheHdw.Wait wait_after_gate
    
Exit Function
errHandler:
    Call Print_Error_Message(Error_Info, "VBT_ZeFuseWrite", "eFuse_pwr_Ramp_meter_DCVI")
    
        If AbortTest Then Exit Function Else Resume Next

End Function

'20210629, Modify for pattern should expand dssc src wave
Private Sub Burn_With_MultiWritePat(patArr() As String, _
                                    bank As eFuseBdfBank, _
                                    srcWave As DSPWave, _
                                    bySiteFlag As Boolean, RvOnly As Boolean)
On Error GoTo errHandler
    Dim m_PatLoop As Long, i As Long
    Dim m_SplitDigSrcWave As New DSPWave
    Dim m_SplitSampleSize As Long
Dim m_CheckPatRun As New SiteBoolean
Dim m_patCnt As Long: m_patCnt = UBound(patArr) + 1
Dim DigSrcSignalName() As String:: ReDim DigSrcSignalName(m_patCnt)
Dim tempVarArray As Variant
Dim fEnableExpandDSSC As Boolean
Dim fAddMonReservedDSSC As Boolean
Dim m_MonReservedDigSrcWave As New DSPWave
Dim orderlist As Variant, tmpstr As String
Dim site As Variant
Dim start_bit As Long

    fEnableExpandDSSC = False   '20210629,if pattern exist "repeat send" vector, should turn on flag to true
    
    If bank.pgmMode = pgm_JTAG Then '20210629,JTAG Banks always don't need to expand
        fEnableExpandDSSC = False
    End If
    
    If LCase(bank.name) = "mon" Then '20210629,BitDef Mon Banks size didn't match pattern, should turn to true
        If bank.FullSize = MONFullSize And m_patCnt = 1 Then
            fAddMonReservedDSSC = False
        Else
            fAddMonReservedDSSC = True
        End If
    End If

    '20211227, Modify for multi config write patterns order control
    If bank.name = "CFG" Then
        If m_patCnt = 1 Then
            orderlist = Split(ConfigWritePatOrder_F, ",")
        ElseIf m_patCnt = 16 Then
            orderlist = Split(ConfigWritePatOrder_H, ",")
        ElseIf m_patCnt = 8 Then
            orderlist = Split(ConfigWritePatOrder_O, ",")
        ElseIf m_patCnt = 4 Then
            orderlist = Split(ConfigWritePatOrder_Q, ",")
        End If
    Else
        tmpstr = vbNullString
        For i = m_patCnt To 1 Step -1
            If i = 1 Then
                tmpstr = tmpstr + CStr(i)
            Else
                tmpstr = tmpstr + CStr(i) + ","
            End If
        Next
        orderlist = Split(tmpstr, ",")
    End If
   '20220215,Modify for Multi RV Pat
    If RvOnly = True And Flag_CFG_Multi_RV_Enable = True And LCase(bank.name) = "cfg" Then
        start_bit = CFG_Multi_RV_PatDsscTotalCnt
        For m_PatLoop = m_patCnt - 1 To 0 Step -1
            tempVarArray = TheHdw.DSSC.pins(bank.SourcePin).Pattern(patArr(m_PatLoop)).Source.Labels.list
            If tempVarArray(0) = "" Then
                DigSrcSignalName(m_PatLoop) = "DigSrcSignal_" & CStr(m_PatLoop)
            Else
                DigSrcSignalName(m_PatLoop) = tempVarArray(0)
            End If
        
            start_bit = start_bit - CFG_Multi_RV_PatDsscCnt(m_PatLoop)
            
            For Each site In TheExec.sites
                m_SplitDigSrcWave.CreateConstant 0, CFG_Multi_RV_PatDsscCnt(m_PatLoop), DspLong
                m_SplitDigSrcWave = srcWave(site).Select(start_bit, 1, CFG_Multi_RV_PatDsscCnt(m_PatLoop)).COPY
                
                If (m_SplitDigSrcWave(site).CalcSum <> 0) Then
                    m_CheckPatRun = True
                End If
            Next site
            
            If (m_CheckPatRun.Any(True)) Then
                If Not bySiteFlag Then
                    ''''if it's same values on all Sites to save TT and improve PTE
                    Call eFuse_DSSC_SetupDigSrcWave_allSites(patArr(m_PatLoop), bank.SourcePin, DigSrcSignalName(m_PatLoop), m_SplitDigSrcWave)
                Else
                    Call eFuse_DSSC_SetupDigSrcWave(patArr(m_PatLoop), bank.SourcePin, DigSrcSignalName(m_PatLoop), m_SplitDigSrcWave)
                End If

                Call TheHdw.patterns(patArr(m_PatLoop)).test(pfAlways, 0)
            End If
            m_CheckPatRun = False
            Set m_SplitDigSrcWave = Nothing
        Next
    Else
        If RvOnly And UBound(patArr) > 0 Then
            Call Print_Error_Message(Error_Info, "VBT_ZeFuseWrite", "Burn_With_MultiWritePat", "Multiple RV write pattern, please set global flag (Flag_CFG_Multi_RV_Enable) to true")
            TheExec.flow.TestLimit resultVal:=0, lowVal:=1, hiVal:=1, Tname:=bank.name + "_writeSetUpFail"
        Else
    For m_PatLoop = 0 To UBound(patArr)
        '20210621, Modify for pattern start with label
        tempVarArray = TheHdw.DSSC.pins(bank.SourcePin).Pattern(patArr(CLng(orderlist(m_PatLoop)) - 1)).Source.Labels.list
        If tempVarArray(0) = "" Then
            DigSrcSignalName(CLng(orderlist(m_PatLoop)) - 1) = "DigSrcSignal_" & orderlist(m_PatLoop)
        Else
            DigSrcSignalName(CLng(orderlist(m_PatLoop)) - 1) = tempVarArray(0)
        End If
        
        For Each site In TheExec.sites
            m_SplitSampleSize = srcWave(site).SampleSize / m_patCnt
            
            If fAddMonReservedDSSC = True Then  '20210629,Add Reserved 0 DSSC to m_SplitDigSrcWave
                m_MonReservedDigSrcWave.CreateConstant 0, MONFullSize * 2 - m_SplitSampleSize, DspDouble
                m_SplitDigSrcWave = srcWave(site).Select((CLng(orderlist(m_PatLoop)) - 1) * m_SplitSampleSize, 1, m_SplitSampleSize).COPY
                m_SplitDigSrcWave = m_SplitDigSrcWave.Concatenate(m_MonReservedDigSrcWave)
            Else
                m_SplitDigSrcWave.CreateConstant 0, m_SplitSampleSize, DspLong
                m_SplitDigSrcWave = srcWave(site).Select((CLng(orderlist(m_PatLoop)) - 1) * m_SplitSampleSize, 1, m_SplitSampleSize).COPY
            End If
            
            If (m_SplitDigSrcWave(site).CalcSum <> 0) Then
                m_CheckPatRun = True
            End If
        Next site
        
        If (m_CheckPatRun.Any(True)) Then
            'eFuse_DSSC_SrcSetup patArr(m_PatLoop), bank, DigSrcSignalName(m_PatLoop), m_SplitDigSrcWave, bySiteFlag
            If fEnableExpandDSSC = True Then
                If Not bySiteFlag Then
                    '20210629, Default set 1 bit expand to 120 bits
                    ''''if it's same values on all Sites to save TT and improve PTE
                    Call eFuse_DSSC_SetupDigSrcWave_allSites(patArr(CLng(orderlist(m_PatLoop)) - 1), bank.SourcePin, DigSrcSignalName(CLng(orderlist(m_PatLoop)) - 1), m_SplitDigSrcWave, True, 120)
                Else
                    Call eFuse_DSSC_SetupDigSrcWave(patArr(CLng(orderlist(m_PatLoop)) - 1), bank.SourcePin, DigSrcSignalName(CLng(orderlist(m_PatLoop)) - 1), m_SplitDigSrcWave, True, 120)
                End If
            Else
                If Not bySiteFlag Then
                    ''''if it's same values on all Sites to save TT and improve PTE
                    Call eFuse_DSSC_SetupDigSrcWave_allSites(patArr(CLng(orderlist(m_PatLoop)) - 1), bank.SourcePin, DigSrcSignalName(CLng(orderlist(m_PatLoop)) - 1), m_SplitDigSrcWave)
                Else
                    Call eFuse_DSSC_SetupDigSrcWave(patArr(CLng(orderlist(m_PatLoop)) - 1), bank.SourcePin, DigSrcSignalName(CLng(orderlist(m_PatLoop)) - 1), m_SplitDigSrcWave)
                End If
            End If
            Call TheHdw.patterns(patArr(CLng(orderlist(m_PatLoop)) - 1)).test(pfAlways, 0) 'Write ECID
        End If
        m_CheckPatRun = False
        Set m_SplitDigSrcWave = Nothing
    Next
        End If
    End If
GlbUtility.PersonalTimerLog "RunPattern"

Exit Sub
errHandler:
    Call Print_Error_Message(Error_Info, "VBT_ZeFuseWrite", "Burn_With_MultiWritePat")
    If AbortTest Then Exit Sub Else Resume Next
End Sub

Public Function DicTrimmedExistCrc(trimmedDic As Dictionary, Crcs As Dictionary) As Boolean
On Error GoTo errHandler
Dim crcItem As Variant

    For Each crcItem In Crcs
        If trimmedDic.Exists(crcItem) Then
            DicTrimmedExistCrc = True
            Exit Function
        End If
    Next crcItem
    
    DicTrimmedExistCrc = False

Exit Function
errHandler:
    Call Print_Error_Message(Error_Info, "VBT_ZeFuseWrite", "DicTrimmedExistCrc")
    
        If AbortTest Then Exit Function Else Resume Next

End Function
