Attribute VB_Name = "VBT_LIB_eFuse_Func"
Option Explicit
''''20160106 Update
Public Function auto_CleanRegData_New()
On Error GoTo errHandler
Dim k As Long
Dim siteNCnt As Long
Dim tmpstr As String
Dim site As Variant

    siteNCnt = theexec.sites.Existing.Count

    tmpstr = vbNullString
    If siteNCnt > 1 Then
        For k = 1 To siteNCnt - 1
            tmpstr = tmpstr + ","
        Next k
    End If
    
    Call RegKeySave("eFuseLotNumber", tmpstr)
    Call RegKeySave("eFuseWaferID", tmpstr)
    Call RegKeySave("eFuseDieX", tmpstr)
    Call RegKeySave("eFuseDieY", tmpstr)
    Call RegKeySave("eFuseIDSSOC", tmpstr)
    Call RegKeySave("eFuseIDSCPU", tmpstr)
    Call RegKeySave("eFuseIDSFIXED", tmpstr)
    Call RegKeySave("eFuseIDSGPU", tmpstr)
    Call RegKeySave("IDSSRAMSOC", tmpstr)
    Call RegKeySave("IDSSRAMCPU1", tmpstr)
    Call RegKeySave("IDSSRAMCPU2", tmpstr)
    Call RegKeySave("eFuseSpareParameter1", tmpstr)
    Call RegKeySave("eFuseSpareParameter2", tmpstr)
    Call RegKeySave("eFuseSpareParameter3", tmpstr)
    Call RegKeySave("eFuseSpareParameter4", tmpstr)
    Call RegKeySave("Hram_ECID_53bit", tmpstr)
    Call RegKeySave("Hram_DVFM_64bit", tmpstr)
    Call RegKeySave("Hram_BinData_46bit", tmpstr)
    Call RegKeySave("Hram_IDS_37bit", tmpstr)
    Call RegKeySave("eFuseSOCTRIM1", tmpstr)
    Call RegKeySave("eFuseSOCTRIM2", tmpstr)
    Call RegKeySave("eFuseBINFIXEDMD1", tmpstr)
    Call RegKeySave("eFuseBINGPUMD1", tmpstr)
    Call RegKeySave("eFuseBINGPUMD2", tmpstr)
    Call RegKeySave("eFuseBINGPUMD3", tmpstr)
    Call RegKeySave("eFuseBINGPUMD4", tmpstr)
    Call RegKeySave("eFuseBINSOCMD1", tmpstr)
    Call RegKeySave("eFuseBINSOCMD2", tmpstr)
    Call RegKeySave("eFuseIDSSRAMSOC", tmpstr)
    Call RegKeySave("eFuseIDSSRAMCPU1", tmpstr)
    Call RegKeySave("eFuseIDSSRAMCPU2", tmpstr)
    Call RegKeySave("HardBinName", "") ''tmpStr
    Call RegKeySave("SoftBinName", "") ''tmpStr
    
    Call RegKeySave("tmps0", tmpstr)
    Call RegKeySave("tmps1", tmpstr)
    Call RegKeySave("tmps2", tmpstr)
    Call RegKeySave("tmps3", tmpstr)
    Call RegKeySave("tmps4", tmpstr)
    
    Call RegKeySave("tmps0_trim", tmpstr)
    Call RegKeySave("tmps1_trim", tmpstr)
    Call RegKeySave("tmps2_trim", tmpstr)
    Call RegKeySave("tmps3_trim", tmpstr)
    Call RegKeySave("tmps4_trim", tmpstr)
    
    Call RegKeySave("DRAM_vendor", tmpstr)
    Call RegKeySave("BKM", tmpstr)
    Call RegKeySave("BKM_Fuse", tmpstr)

    'If (gB_findCFGCondTable_flag = True) Then
    Call RegKeySave("SVM_CFuse_288Bits", tmpstr) ''''20171103 add SVM_CFuse_288Bits
    'ElseIf (gB_findCFGTable_flag = True) Then
        'Call RegKeySave("CFG_First_64Bits", tmpStr)
    'End If
        
    Call UpdateDLogColumns(30)

    ''''if the TP does NOT run the TestInstance "eFuse_Initialize"
    If (Trim(GlbUtility.currStage) = "") Then
        GlbUtility.currStage = LCase(theexec.CurrentJob)
    End If

    ''''For all NOT CP/WLFT Jobs, reset datalog X,Y Coordinates to N/A (-32768).
'    If ((GlbUtility.currStage Like "*CP*") Or (GlbUtility.currStage Like "*WLFT*")) Then
    If UCase(TheExec.CurrentChanMap) Like "CHANNELMAP_CP*" Or UCase(TheExec.CurrentChanMap) Like "CHANNELMAP_WL*" Then  'JackChou 202404 To allow Stage_FT with ChannelMap_CP case [Requester-TSMC-KoPei]
        theexec.Flow.TestLimit 1, 1, 1, Tname:="CleanRegData", PinName:=UCase(GlbUtility.currStage)
    Else
        If theexec.Flow.enableWord("FT_SIM") = True Then
            ''''it's Allen's functioality, doing the FT simulation on CP environment (engineer).
            'Do not Reset X,Y to -32768(N/A)
            theexec.Flow.TestLimit 1, 1, 1, Tname:="CleanRegData", PinName:=UCase(GlbUtility.currStage)
        Else
            ''multisite for FT, initialize XY to N/A(-32768) for all sites
            ''Write X and Y coordinates to FT STDF file
            ''After running ECID, it will have the correct XY information.

            If (theexec.TesterMode = testModeOffline) Then ''''20160202 for the simulation
                For Each site In theexec.sites
                    XCoord(site) = theexec.Datalog.Setup.WaferSetup.GetXCoord(site)
                    YCoord(site) = theexec.Datalog.Setup.WaferSetup.GetYCoord(site)
                    If (XCoord(site) = -32768 Or YCoord(site) = -32768) Then
                        Call setXY(5, 6) ''''set a pseudo XY coordinate
                        theexec.Datalog.WriteComment vbTab & "Call setXY(5, 6) (pseudo XY_Coordinate)"
                        XCoord(site) = theexec.Datalog.Setup.WaferSetup.GetXCoord(site)
                        YCoord(site) = theexec.Datalog.Setup.WaferSetup.GetYCoord(site)
                    End If
                Next site
            Else
                ''''<MUST> Very Important
                If gB_FT_Wafer_PsudoFuse = False Then
                    For Each site In theexec.sites.Existing
                        Call theexec.Datalog.Setup.WaferSetup.SetXCoord(site, "-32768")
                        Call theexec.Datalog.Setup.WaferSetup.SetYCoord(site, "-32768")
                    Next site
                End If
            End If

            theexec.Flow.TestLimit 1, 1, 1, Tname:="CleanRegData_XY", PinName:=UCase(GlbUtility.currStage)
        End If
    End If

    Call UpdateDLogColumns__False

Exit Function
errHandler:
    Call Print_Error_Message(Error_Info, "VBT_LIB_eFuse_Func", "auto_CleanRegData_New")
    If AbortTest Then Exit Function Else Resume Next
End Function

''''20171211 update
Public Function auto_isBinaryString(ByVal Inputstr As String) As Boolean
On Error GoTo errHandler
Dim i As Long
Dim m_len As Long
Dim m_char As String
Dim m_match_flag As Boolean

    Inputstr = UCase(Trim(Inputstr))

    If (Inputstr Like "B*") Then
        m_match_flag = True ''''<MUST> initialize
        Inputstr = Replace(Inputstr, "B", "", 1, 1) ''''remove the first "B" character
        Inputstr = Replace(Inputstr, "_", "")       ''''for case like as b0001_0101

        ''''do the advanced analysis
        m_len = Len(Inputstr)
        For i = 1 To m_len
            m_char = mid(Inputstr, i, 1)
            If (m_char <> "0" And m_char <> "1") Then
                m_match_flag = False
                Exit For
            End If
        Next i
    Else
        m_match_flag = False
    End If

    auto_isBinaryString = m_match_flag

Exit Function
errHandler:
    Call Print_Error_Message(Error_Info, "VBT_LIB_eFuse_Func", "auto_isBinaryString")
    If AbortTest Then Exit Function Else Resume Next
End Function
''''2015-07-02
''''Here it's used to disable the column length in the datalog.
Public Function UpdateDLogColumns__False()
On Error GoTo errHandler
    ''''20170217 update
    If (gB_newDlog_Flag) Then Exit Function

    theexec.Datalog.Setup.Shared.ascii.Columns.EnableCustomWidths = False
    theexec.Datalog.ApplySetup

Exit Function
errHandler:
    Call Print_Error_Message(Error_Info, "VBT_LIB_eFuse_Func", "UpdateDLogColumns__False")
    If AbortTest Then Exit Function Else Resume Next
End Function

Public Function auto_eFuse_CheckPatternFlag(Optional bankstr As String = vbNullString)
On Error GoTo errHandler
Dim m_site As Variant
Dim m_value As New SiteVariant
Dim field As eFuseBdfField
Dim fieldStage As String
Dim opbank As eFuseBdfBank
Dim patFailCntRt As New SiteLong
Dim bank As Variant
Dim TrimMissingCnt As New SiteLong
Dim checkRV As Boolean

    For Each m_site In theexec.sites
        m_value(m_site) = theexec.sites.item(m_site).SortNumber
        If gB_Fuse_Skip(m_site) = True Then
            m_value(m_site) = 0
            Call Print_Error_Message(Warning_Info, "VBT_LIB_eFuse_Func", "auto_eFuse_CheckPatternFlag", "Site(" & CStr(m_site) & ") is overflow and skip fusing.")
            theexec.Flow.TestLimit resultVal:=m_value, lowVal:=-1, hiVal:=-1, Tname:="OverFlow"
        End If
    Next m_site
    
    theexec.Flow.TestLimit resultVal:=m_value, lowVal:=-1, hiVal:=-1, Tname:="SortNumber"
    
    For Each bank In BdfDataBase.Banks
        checkRV = False
        If bankstr <> "" And InStr(1, UCase(bankstr), UCase(bank)) < 1 Then
            GoTo NextBank
        End If

        Set opbank = BdfDataBase.Banks(bank)
        Set patFailCntRt = PatFailCnt(opbank)
        
        If opbank.Fields.Exists("fuse_revision") Then
            Set field = opbank.Fields("fuse_revision")
            If GlbUtility.testedStages.Exists(field.BlowLocation) Then
                checkRV = True
            End If
        End If
        
        For Each m_site In theexec.sites
            m_value(m_site) = 0                   '210702 forcrete
            m_value = patFailCntRt + m_value
            
            If checkRV Then
                If field.DsscValue <> field.Default Then
                    theexec.Flow.TestLimit resultVal:=0, lowVal:=-1, hiVal:=-1, Tname:="fuse_revision Check"
                Else
                    theexec.Flow.TestLimit resultVal:=-1, lowVal:=-1, hiVal:=-1, Tname:="fuse_revision Check"
                End If
            End If
        Next
        theexec.Flow.TestLimit resultVal:=m_value, lowVal:=0, hiVal:=0, Tname:=CStr(opbank.name) + "_CheckPatternFlag"

        If GlbUtility.OnlineMode And (Not UCase(opbank.name) Like "CMP*") Then
            If theexec.enableWord("CFG_Partial") And bankstr = "CFG" Then
                TrimMissingCnt = opbank.TrimedMissingCnt(opbank.DicNonDefaultNonEarly, TrimMissingFilter_Flag:=True)
            Else
                TrimMissingCnt = opbank.TrimedMissingCnt(opbank.DicOthers, TrimMissingFilter_Flag:=True)
            End If
            theexec.Flow.TestLimit resultVal:=TrimMissingCnt, lowVal:=0, hiVal:=0, Tname:=CStr(opbank.name) & "_TrimMissingCnt", PinName:="Value"
        End If
NextBank:
    Next

Exit Function
errHandler:
    Call Print_Error_Message(Error_Info, "VBT_LIB_eFuse_Func", "auto_eFuse_CheckPatternFlag")
    If AbortTest Then Exit Function Else Resume Next
End Function
    
Public Function EFUSE_Resistance(patset As Pattern, PwrPin As String, vpwr As Double, Optional PrePatSet As Pattern, Optional DigSource As String = vbNullString, Optional Validating_ As Boolean) As Long
On Error GoTo errHandler
Dim result As New SiteDouble
Dim current_data As New PinListData
Dim site As Variant
Dim power_pin_arr() As String
Dim power_pin_number As Long
Dim i As Long
Dim temp_power_pin As String
Dim PrePatResult As New SiteBoolean
Dim initSiteStatus As New SiteBoolean
    
    If Validating_ Then
        If patset.value <> "" Then Call PrLoadPattern(patset.value)
        If PrePatSet.value <> "" Then Call PrLoadPattern(PrePatSet.value)
        Exit Function    ' Exit after validation
    End If
    
    TheHdw.Digital.ApplyLevelsTiming True, True, True, tlPowered
    
    If PrePatSet.value <> "" Then
        initSiteStatus = TheExec.sites.Active
        PrePatResult = EfuseExecInitPattern(PrePatSet.value, DigSource)
        If Not PrePatResult.Any(True) Then
            Exit Function
        ElseIf PrePatResult.Any(False) Then
            TheExec.sites.Selected = PrePatResult
        End If
    End If
    
    Call TheExec.DataManager.DecomposePinList(PwrPin, power_pin_arr, power_pin_number)
            
    TheExec.Datalog.WriteComment "Setting: " + CStr(vpwr) + " V"
 
    Call TurnOnEfusePwrPins(PwrPin, vpwr)
    TheHdw.Wait 0.001

    Call HardIP_InitialSetupForPatgen
    Call TheHdw.Patterns(patset).start
    TheHdw.DCVS.Pins(PwrPin).SetCurrentRanges 0.02, 0.02
    Call TheHdw.Digital.Patgen.FlagWait(cpuA, 0) 'Meas during CPUA loop
    
    TheHdw.Wait 0.01
    
    current_data = TheHdw.DCVS.Pins(PwrPin).Meter.Read.Math.Multiply(1000)
    
    For i = 0 To power_pin_number - 1
        temp_power_pin = power_pin_arr(i)
        For Each site In TheExec.sites
        
            If (current_data.Pins(temp_power_pin).value < 0.0000001) Then
                current_data.Pins(temp_power_pin).value = 0.0000001
            End If
            
            result = vpwr / current_data.Pins(temp_power_pin).value
            '''theexec.Datalog.WriteComment "Site" + CStr(site) + ": " + temp_power_pin + " " + FormatNumber(CStr(current_data.Pins(temp_power_pin).Value), 3) + " mA" + " " + FormatNumber(CStr(result), 3) + " Kohm"
        Next
        
        theexec.Flow.TestLimit resultVal:=result.Multiply(1000), lowVal:=140, hiVal:=330, Tname:=temp_power_pin
    Next
    
    Call TheHdw.Digital.Patgen.Continue(0, cpuA + cpuB + cpuC + cpuD)
    TheHdw.Digital.Patgen.HaltWait ' Haltwait at patten end
    Call TurnOffEfusePwrPins(PwrPin, vpwr)
    
    TheExec.Datalog.WriteComment ""
    If (PrePatSet.value <> "") Then
        TheExec.sites.Selected = initSiteStatus
    End If
Exit Function
errHandler:
    Call Print_Error_Message(Error_Info, "VBT_LIB_eFuse_Func", "EFUSE_Resistance")
    If AbortTest Then Exit Function Else Resume Next
End Function

Public Function auto_UDR_UFP(UFP_pat As Pattern, PwrPin As String, vpwr As Double, Optional PrePatSet As Pattern, Optional DigSource As String = vbNullString, Optional Validating_ As Boolean) As Long
On Error GoTo errHandler
Dim patt As String
Dim PattArray() As String, PatCount As Long
Dim PrePatResult As New SiteBoolean
Dim initSiteStatus As New SiteBoolean

    ''''----------------------------------------------------------------------------------------------------
    ''''<Important>
    ''''Must be put before all implicit array variables, otherwise the validation will be error.
    '==================================
    '=  Validate/Load patterns        =
    '==================================
    ''''20161114 update to Validate/load pattern
    ''20210121 modify the pattern validation. use the same function with HIP.
    
    If Validating_ Then
        If UFP_pat.value <> "" Then Call PrLoadPattern(UFP_pat.value)
        If PrePatSet.value <> "" Then Call PrLoadPattern(PrePatSet.value)
        Exit Function    ' Exit after validation
    End If
    TheHdw.Digital.ApplyLevelsTiming True, True, True, tlPowered
    If PrePatSet.value <> "" Then
        initSiteStatus = TheExec.sites.Active
        PrePatResult = EfuseExecInitPattern(PrePatSet.value, DigSource)
        If Not PrePatResult.Any(True) Then
            Exit Function
        ElseIf PrePatResult.Any(False) Then
            TheExec.sites.Selected = PrePatResult
        End If
    End If
    Call PATT_GetPatListFromPatternSet(UFP_pat.value, PattArray, PatCount)
    patt = PattArray(0)
    
    'If (auto_eFuse_PatSetToPat_Validation(UFP_pat, patt, Validating_) = True) Then Exit Function
    ''''----------------------------------------------------------------------------------------------------

    ''TheHdw.Patterns(UFP_pat).Load
    
    If EFUSE_POWER_OFF_SETTING Then vpwr = 0
    
    Call TurnOnEfusePwrPins(PwrPin, vpwr)

    Call TheHdw.Patterns(patt).test(pfAlways, 0) ''''was UFP_pat
    DebugPrintFunc UFP_pat.value

    Call TurnOffEfusePwrPins(PwrPin, vpwr)
    If (PrePatSet.value <> "") Then
        TheExec.sites.Selected = initSiteStatus
    End If
Exit Function
errHandler:
    Call Print_Error_Message(Error_Info, "VBT_LIB_eFuse_Func", "auto_UDR_UFP")
    If AbortTest Then Exit Function Else Resume Next
End Function

Public Function auto_UDR_UFR(UFR_pat As Pattern, Optional PrePatSet As Pattern, Optional DigSource As String = vbNullString, Optional Validating_ As Boolean) As Long
On Error GoTo errHandler
Dim patt As String
Dim PattArray() As String, PatCount As Long
Dim PrePatResult As New SiteBoolean
Dim initSiteStatus As New SiteBoolean

    ''''----------------------------------------------------------------------------------------------------
    ''''<Important>
    ''''Must be put before all implicit array variables, otherwise the validation will be error.
    '==================================
    '=  Validate/Load patterns        =
    '==================================
    ''''20161114 update to Validate/load pattern
    ''20210121 modify the pattern validation. use the same function with HIP.
    
    If Validating_ Then
        If UFR_pat.value <> "" Then Call PrLoadPattern(UFR_pat.value)
        If PrePatSet.value <> "" Then Call PrLoadPattern(PrePatSet.value)
        Exit Function    ' Exit after validation
    End If
    
    TheHdw.Digital.ApplyLevelsTiming True, True, True, tlPowered
    If PrePatSet.value <> "" Then
        initSiteStatus = TheExec.sites.Active
        PrePatResult = EfuseExecInitPattern(PrePatSet.value, DigSource)
        If Not PrePatResult.Any(True) Then
            Exit Function
        ElseIf PrePatResult.Any(False) Then
            TheExec.sites.Selected = PrePatResult
        End If
    End If

    Call PATT_GetPatListFromPatternSet(UFR_pat.value, PattArray, PatCount)
    patt = PattArray(0)
    
    'If (auto_eFuse_PatSetToPat_Validation(UFR_pat, patt, Validating_) = True) Then Exit Function
    ''''----------------------------------------------------------------------------------------------------

    ''TheHdw.Patterns(UFR_pat).Load

    Call TheHdw.Patterns(patt).test(pfAlways, 0) '''' was UFR_pat
    DebugPrintFunc UFR_pat.value
    If (PrePatSet.value <> "") Then
        TheExec.sites.Selected = initSiteStatus
    End If
Exit Function
errHandler:
    Call Print_Error_Message(Error_Info, "VBT_LIB_eFuse_Func", "auto_UDR_UFR")
    If AbortTest Then Exit Function Else Resume Next
End Function

Public Function auto_ShowECIDData()
On Error GoTo errHandler
Dim opbank As eFuseBdfBank
Dim site As Variant

    theexec.Datalog.WriteComment ""
    Set opbank = GetBdfBank("ECID")
    'update by Jason's request to fixed Galaxy multi-site format, 140411

    For Each site In theexec.sites
        theexec.Datalog.WriteComment "<@efuse_lot_ID=" & site & "|" & HramLotId(site) & ">"
        theexec.Datalog.WriteComment "<@efuse_wafer_ID=" & site & "|" & HramLotId(site) & "." & Format(CStr(HramWaferId(site)), "00") & ">"
        theexec.Datalog.WriteComment "<@Prober_Hex_code=" & site & "|" & opbank.EcidHexStr & ">"
    Next site
    theexec.Datalog.WriteComment ""

Exit Function
errHandler:
    Call Print_Error_Message(Error_Info, "VBT_LIB_eFuse_Func", "auto_ShowECIDData")
    If AbortTest Then Exit Function Else Resume Next
End Function

''''20161121 Add for gereral Function Test with the (.pat/.pat.gz) in datalog
Public Function auto_Function_Test(patset As Pattern, Optional PrePatSet As Pattern, Optional DigSource As String = vbNullString, Optional Validating_ As Boolean) As Long
On Error GoTo errHandler
Dim patt As String
Dim PattArray() As String, PatCount As Long
Dim PrePatResult As New SiteBoolean
Dim initSiteStatus As New SiteBoolean

    ''''----------------------------------------------------------------------------------------------------
    ''''<Important>
    ''''Must be put before all implicit array variables, otherwise the validation will be error.
    '==================================
    '=  Validate/Load patterns        =
    '==================================
    ''''20161114 update to Validate/load pattern

    If Validating_ Then
        If patset.value <> "" Then Call PrLoadPattern(patset.value)
        If PrePatSet.value <> "" Then Call PrLoadPattern(PrePatSet.value)
        Exit Function    ' Exit after validation
    End If
    
    TheHdw.Digital.ApplyLevelsTiming True, True, True, tlPowered
    If PrePatSet.value <> "" Then
        initSiteStatus = TheExec.sites.Active
        PrePatResult = EfuseExecInitPattern(PrePatSet.value, DigSource)
        If Not PrePatResult.Any(True) Then
            Exit Function
        ElseIf PrePatResult.Any(False) Then
            TheExec.sites.Selected = PrePatResult
        End If
    End If

    Call PATT_GetPatListFromPatternSet(patset.value, PattArray, PatCount)
    patt = PattArray(0)
    
    'If (auto_eFuse_PatSetToPat_Validation(patset, patt, Validating_) = True) Then Exit Function
    ''''----------------------------------------------------------------------------------------------------

    Call TheHdw.Patterns(patt).test(pfAlways, 0)
    DebugPrintFunc patset.value

    If (PrePatSet.value <> "") Then
        TheExec.sites.Selected = initSiteStatus
    End If
Exit Function
errHandler:
    Call Print_Error_Message(Error_Info, "VBT_LIB_eFuse_Func", "auto_Function_Test")
    If AbortTest Then Exit Function Else Resume Next
End Function

Public Function auto_ConfigWrite_CFG_DV(CFG_DV_pat As Pattern, PwrPin As String, vpwr As Double, Optional Validating_ As Boolean) As Long
On Error GoTo errHandler
Dim patt As String
Dim PattArray() As String, PatCount As Long

    ''''----------------------------------------------------------------------------------------------------
    ''''<Important>
    ''''Must be put before all implicit array variables, otherwise the validation will be error.
    '==================================
    '=  Validate/Load patterns        =
    '==================================
    ''''20161114 update to Validate/load pattern
    
    If Validating_ Then Call PrLoadPattern(CFG_DV_pat.value):  Exit Function    ' Exit after validation
    Call PATT_GetPatListFromPatternSet(CFG_DV_pat.value, PattArray, PatCount)
    patt = PattArray(0)
    
    'If (auto_eFuse_PatSetToPat_Validation(CFG_DV_pat, patt, Validating_) = True) Then Exit Function
    ''''----------------------------------------------------------------------------------------------------

    ''TheHdw.Patterns(CFG_DV_pat).Load
    TheHdw.Digital.ApplyLevelsTiming True, True, True, tlPowered
    
    If EFUSE_POWER_OFF_SETTING Then vpwr = 0

    Call TurnOnEfusePwrPins(PwrPin, vpwr)

    Call TheHdw.Patterns(patt).test(pfAlways, 0)
    DebugPrintFunc CFG_DV_pat.value

    Call TurnOffEfusePwrPins(PwrPin, vpwr)

Exit Function
errHandler:
    Call Print_Error_Message(Error_Info, "VBT_LIB_eFuse_Func", "auto_ConfigWrite_CFG_DV")
    If AbortTest Then Exit Function Else Resume Next
End Function

Public Function DRAM_eFuse_Write(FuseType As String)
On Error GoTo errHandler
Dim m_value As New SiteVariant
Dim opbank As eFuseBdfBank
Dim field As eFuseBdfField
Dim m_Loop As Long
Dim m_catename As String
Dim enablewrdname As String
Dim dramType As Variant
Dim enableWdCnt As Integer
Dim dramSiteControl As Boolean
Dim enableWd As Variant
Dim msg As String: msg = vbNullString
Dim sitectrlArray As Variant
Dim listEnableWd As Collection
Dim i As Variant
Dim site As Variant 'Carter, 20240304

    Set listEnableWd = New Collection
    Set opbank = GetBdfBank(FuseType)

    For Each dramType In BdfDataBase.DicDramMap
        If UCase(dramType) Like "DRAM_*" And theexec.Flow.enableWord(dramType) Then
            GlbUtility.WriteDlg "The Dram Enable = " & dramType & ""
            enableWdCnt = enableWdCnt + 1
            listEnableWd.Add dramType, CStr(enableWdCnt)   'add enable word to list
            If BdfDataBase.DicDramMap(dramType)("sitecontrol") <> Empty Then
                dramSiteControl = True
                If (enableWdCnt > 2) Then
                    msg = "Site Control mode: more than Two Dram EnableWord selected!!!"
                    GoTo skip
                End If
            Else
                If (enableWdCnt > 1) Then
                    msg = "more than One Dram EnableWord selected!!!"
                    GoTo skip
                End If
            End If
        End If
    Next dramType

    If enableWdCnt = 0 Then
        msg = "The Dram EnableWord isn't selected!!!"
        GoTo skip
    ElseIf dramSiteControl And (enableWdCnt > 2 Or enableWdCnt < 2) Then
        msg = "Site Control mode: The Dram EnableWord count = " + CStr(enableWdCnt) + " !!!"
        GoTo skip
    End If
    
    For m_Loop = 0 To BdfDataBase.DramFieldSize - 1
        m_catename = BdfDataBase.DicDramMap.item("fieldname")(m_Loop)
        For Each enableWd In listEnableWd
            If dramSiteControl Then
                sitectrlArray = Split(BdfDataBase.DicDramMap.item(enableWd)("sitecontrol"), ",")
                For i = 0 To UBound(sitectrlArray)
                    For Each site In theexec.sites
                        If site = CInt(sitectrlArray(i)) Then
                            m_value(site) = CStr(BdfDataBase.DicDramMap.item(enableWd)("fuseData")(m_Loop))
                            If LCase(m_value(site)) Like "x*" Then
                                m_value(site) = GlbUtility.Hex2Dbl(CStr(m_value(site)))
                            ElseIf LCase(m_value(site)) Like "b*" Then
                                m_value(site) = GlbUtility.Bin2Dec(CStr(m_value(site)))
                            Else
                                m_value(site) = CStr(m_value(site))
                            End If
                            Exit For
                        End If
                    Next site
                Next
            Else
                For Each site In theexec.sites
                    m_value(site) = CStr(BdfDataBase.DicDramMap.item(enableWd)("fuseData")(m_Loop))
                    If LCase(m_value(site)) Like "x*" Then
                        m_value(site) = GlbUtility.Hex2Dbl(CStr(m_value(site)))
                    ElseIf LCase(m_value(site)) Like "b*" Then
                        m_value(site) = GlbUtility.Bin2Dec(CStr(m_value(site)))
                    Else
                        m_value(site) = CStr(m_value(site))
                    End If
                Next site
            End If
        Next enableWd
        
        Set field = opbank.Fields(m_catename)
        If field.ReadOnly Or field.DefaultOrReal = dr_default Then
            Call Print_Error_Message(Warning_Info, "VBT_LIB_eFuse_Func", "DRAM_eFuse_Write", "[" + m_catename + "] - Category has default value, bypass SetEfuse")
            theexec.Flow.TestLimit resultVal:=m_value, lowVal:=field.Llimit, hiVal:=field.Hlimit, Tname:="DRAM_eFuse_Write_" & field.name
        Else
            opbank.SetEfuse field.name, m_value, , , , , True
        End If
    Next
    
Exit Function
skip:
    Call Print_Error_Message(Error_Info, "VBT_LIB_eFuse_Func", "DRAM_eFuse_Write", msg)
    theexec.Flow.TestLimit resultVal:=999, lowVal:=1, hiVal:=1, Tname:="DRAM_eFuse_Write"
    Exit Function
errHandler:
    Call Print_Error_Message(Error_Info, "VBT_LIB_eFuse_Func", "DRAM_eFuse_Write")
    If AbortTest Then Exit Function Else Resume Next
End Function

Public Function BinXYCheck()
On Error GoTo errHandler
Dim opbank As eFuseBdfBank
Dim field As eFuseBdfField
Dim m_value As New SiteVariant
Dim site As Variant 'Carter, 20240304

    For Each site In theexec.sites
        Set field = BdfDataBase.Bank_Cfg.Fields("Product_Identifier")
        m_value = field.DsscDecValue
        theexec.Flow.TestLimit resultVal:=m_value, lowVal:=0, hiVal:=0, Tname:="BinXYCheck"
    Next site

Exit Function
errHandler:
    Call Print_Error_Message(Error_Info, "VBT_LIB_eFuse_Func", "BinXYCheck")
    If AbortTest Then Exit Function Else Resume Next
End Function

Public Function CPM_eFuse_Write()
On Error GoTo errHandler
Dim opbank As eFuseBdfBank
Dim field As eFuseBdfField
Dim mSD_value As New SiteDouble
Dim keyword As String
Dim i As Long, j As Long
Dim m_value As New SiteVariant
Dim BankArry As Variant
Dim FieldArry As Variant
Dim FlagArry As Variant
Dim ContentArry As Variant
Dim site As Variant 'Carter, 20240304

    BankArry = BdfDataBase.DicCPMMap("bank")
    FieldArry = BdfDataBase.DicCPMMap("field")
    FlagArry = BdfDataBase.DicCPMMap("flag")
    
    mSD_value = 1
    
    For i = 0 To UBound(FlagArry)
        For Each site In theexec.sites
            theexec.Datalog.WriteComment "Site " & site & "   " & FlagArry(i) & ":" & theexec.sites.item(site).FlagState(FlagArry(i))
        Next site
    Next i
    
    For j = 0 To UBound(FieldArry)
        For Each site In theexec.sites
            keyword = vbNullString
            For i = 0 To UBound(FlagArry)
                If theexec.sites.item(site).FlagState(FlagArry(i)) = logicTrue Then
                    keyword = keyword + "true,"
                Else
                    keyword = keyword + "false,"
                End If
            Next i
            If Not BdfDataBase.DicCPMMap.Exists(keyword) Then
                theexec.Datalog.WriteComment "Site " & site & " flag is mixed up."
                mSD_value = 0
                m_value = 0
            Else
                ContentArry = BdfDataBase.DicCPMMap(keyword)
                If LCase(ContentArry(j)) Like "x*" Then
                    m_value = GlbUtility.Hex2Dbl(CStr(ContentArry(j)))
                ElseIf LCase(ContentArry(j)) Like "b*" Then
                    m_value = GlbUtility.Bin2Dec(CStr(ContentArry(j)))
                Else
                    m_value = CStr(ContentArry(j))
                End If
            End If
        Next site
        Set opbank = GetBdfBank(CStr(BankArry(j)))
        Set field = opbank.Fields(FieldArry(j))
        If field.BlowLocation = GlbUtility.currStage Then
            opbank.SetEfuse field.name, m_value, , , , , True
        End If
    Next j
    
    theexec.Flow.TestLimit resultVal:=mSD_value, lowVal:=1, hiVal:=1, Tname:="CPM_eFuse_Write"
Exit Function
errHandler:
    Call Print_Error_Message(Error_Info, "VBT_LIB_eFuse_Func", "CPM_eFuse_Write")
    If AbortTest Then Exit Function Else Resume Next
End Function

Public Function CPM_eFuse_Read()
On Error GoTo errHandler
Dim opbank As eFuseBdfBank
Dim field As eFuseBdfField
Dim mSD_value As New SiteDouble
Dim m_flag_status() As New SiteBoolean
Dim keyword As String
Dim i As Long, j As Long
Dim key As Variant
Dim str() As String
Dim findkey As New SiteBoolean
Dim findkeyname As New SiteVariant
Dim m_value As New SiteVariant
Dim BankArry As Variant
Dim FieldArry As Variant
Dim FlagArry As Variant
Dim ContentArry As Variant
Dim ContentTmp As Variant
Dim stage As String: stage = vbNullString
Dim site As Variant 'Carter, 20240304

    BankArry = BdfDataBase.DicCPMMap("bank")
    FieldArry = BdfDataBase.DicCPMMap("field")
    FlagArry = BdfDataBase.DicCPMMap("flag")
    
    findkey = False
    ReDim m_flag_status(UBound(FlagArry))
    For i = 0 To UBound(FlagArry)
        m_flag_status(i) = False
    Next i
    
    mSD_value = 1
    
    For j = 0 To UBound(FieldArry)
        Set opbank = GetBdfBank(CStr(BankArry(j)))
        Set field = opbank.Fields(FieldArry(j))
        m_value = field.DsscDecValue
        stage = Replace(UCase(BdfDataBase.GetRealStage(field.BlowLocation)), "_EARLY", "")
        For Each site In theexec.sites
            If GlbUtility.testedStages.Exists(stage) Or (UCase(GlbUtility.currStage) = stage And m_value <> 0) Then
                If findkey = False And mSD_value = 1 Then
                    For Each key In BdfDataBase.DicCPMMap
                        If key <> "bank" And key <> "field" And key <> "flag" Then
                            ContentArry = BdfDataBase.DicCPMMap(key)
                            If LCase(ContentArry(j)) Like "x*" Then
                                ContentTmp = GlbUtility.Hex2Dbl(CStr(ContentArry(j)))
                            ElseIf LCase(ContentArry(j)) Like "b*" Then
                                ContentTmp = GlbUtility.Bin2Dec(CStr(ContentArry(j)))
                            Else
                                ContentTmp = CStr(ContentArry(j))
                            End If
                            If m_value = ContentTmp Then
                                str = Split(key, ",")
                                For i = 0 To UBound(str) - 1
                                    If str(i) = "true" Then
                                        m_flag_status(i) = True
                                    Else
                                        m_flag_status(i) = False
                                    End If
                                Next i
                                findkey = True
                                findkeyname = key
                                Exit For
                            End If
                        End If
                    Next key
                    If findkey = False Then
                        Call Print_Error_Message(Error_Info, "VBT_LIB_eFuse_Func", "CPM_eFuse_Read", "Site " & site & "   " & FieldArry(j) & ":" & CStr(m_value) & ", Value didn't match.")
                        mSD_value = 0
                    End If
                ElseIf findkey = True Then
                    ContentArry = BdfDataBase.DicCPMMap(CStr(findkeyname))
                    If LCase(ContentArry(j)) Like "x*" Then
                        ContentTmp = GlbUtility.Hex2Dbl(CStr(ContentArry(j)))
                    ElseIf LCase(ContentArry(j)) Like "b*" Then
                        ContentTmp = GlbUtility.Bin2Dec(CStr(ContentArry(j)))
                    Else
                        ContentTmp = CStr(ContentArry(j))
                    End If
                    If m_value <> ContentTmp Then
                        Call Print_Error_Message(Error_Info, "VBT_LIB_eFuse_Func", "CPM_eFuse_Read", "Site " & site & "   " & FieldArry(j) & ":" & CStr(m_value) & ", Value didn't match.")
                        mSD_value = 0
                    End If
                End If
            End If
        Next site
    Next j
    
    For i = 0 To UBound(FlagArry)
        For Each site In theexec.sites
            If findkey Then
                If mSD_value = 1 Then
                    If m_flag_status(i) = True Then
                        theexec.sites.item(site).FlagState(FlagArry(i)) = logicTrue
                    Else
                        theexec.sites.item(site).FlagState(FlagArry(i)) = logicFalse
                    End If
                    theexec.Datalog.WriteComment "Site " & site & "   " & FlagArry(i) & ":" & theexec.sites.item(site).FlagState(FlagArry(i))
                End If
            End If
        Next site
    Next i
    
    theexec.Flow.TestLimit resultVal:=mSD_value, lowVal:=1, hiVal:=1, Tname:="CPM_eFuse_Read"

Exit Function
errHandler:
    Call Print_Error_Message(Error_Info, "VBT_LIB_eFuse_Func", "CPM_eFuse_Read")
    If AbortTest Then Exit Function Else Resume Next
End Function

Public Function ECID_ReadChk()
On Error GoTo errHandler
Dim opbank As eFuseBdfBank
Dim field As eFuseBdfField
Dim m_value As New SiteVariant
Dim item As Variant
Dim keyword As String
Dim m_Loop As Long
Dim m_catename As String
Dim enablewrdname As String
Dim dramType As Variant
    
    Set opbank = GetBdfBank("ecid")
    If opbank.HadDeidFuse = False Then GlbUtility.WriteDlg "Not ECID Bank, skip check!": Exit Function
    
    For Each item In opbank.Fields.Keys
        If GlbUtility.IsStrMatch(CStr(item), "x.*coor") Then
            Set field = opbank.Fields(item)
            theexec.Flow.TestLimit resultVal:=XCoord, lowVal:=field.Llimit, hiVal:=field.Hlimit, Tname:=field.name
        ElseIf GlbUtility.IsStrMatch(CStr(item), "y.*coor") Then
            Set field = opbank.Fields(item)
            theexec.Flow.TestLimit resultVal:=YCoord, lowVal:=field.Llimit, hiVal:=field.Hlimit, Tname:=field.name
        End If
    Next
    
Exit Function
errHandler:
    Call Print_Error_Message(Error_Info, "VBT_LIB_eFuse_Func", "ECID_ReadChk")
    If AbortTest Then Exit Function Else Resume Next
End Function

Public Function BankVarChk(bank As String)
On Error GoTo errHandler
Dim bankarr As Variant
Dim m_siteVar As String
Dim i As Integer
Dim site As Variant 'Carter, 20240304

    bankarr = Split(bank, ",")
    For i = 0 To UBound(bankarr)
        For Each site In theexec.sites
            m_siteVar = bankarr(i) + "Chk_Var"
            theexec.Flow.TestLimit resultVal:=theexec.sites(site).SiteVariableValue(m_siteVar), lowVal:=-1, hiVal:=1, Tname:=m_siteVar
        Next site
    Next i
    
Exit Function
errHandler:
    Call Print_Error_Message(Error_Info, "VBT_LIB_eFuse_Func", "BankVarChk")
    If AbortTest Then Exit Function Else Resume Next
End Function

Public Function Efuse_bincheck(Optional preWrite As Boolean = False)
On Error GoTo errHandler
Dim m_site As Variant
Dim flagDic As New Dictionary
Dim tmpFlagDic As Dictionary
Dim Flag As Variant, bank As Variant
Dim key As Variant
Dim m_Index As Long, m_Loop As Long, m_ruleIdx As Long, i As Long
Dim m_catename As String, m_bankname As String, m_stage As String, operator As String
Dim opbank As eFuseBdfBank
Dim bkmfield As eFuseBdfField, field As eFuseBdfField
Dim m_tmpDSP As DSPWave
Dim result As New SiteBoolean
Dim m_continueFor As New SiteBoolean
Dim m_FusedValue As New SiteVariant
Dim m_skipTest As New SiteBoolean
Dim oValue() As Double, oHiLimit() As Double, oLoLimit() As Double
Dim oCataName() As String
Dim FieldFail As Boolean, hasFlag As Boolean: hasFlag = False
Dim isCurrentAndTestedJob As Boolean: isCurrentAndTestedJob = False
Dim dicPreWriteWave As New Dictionary
Dim dicTrimmed As New Dictionary
Dim newWave As New DSPWave
Dim conditionSkip As Boolean
Dim item As Variant, tmp As Variant

    operator = vbNullString
    Call UpdateDLogColumns(50)
    dicPreWriteWave.RemoveAll
    
    For Each bank In DictBinBankMapping.Keys
        Set opbank = GetBdfBank(CStr(bank))
        If Not GlbUtility.OnlineMode Then
            Set dicTrimmed = ObtainCatDictionary(opbank, False, False, False, "")
            opbank.PseudoFusedFillup dicTrimmed
        End If
        Set newWave = opbank.GetBinCheckDsscWave
        dicPreWriteWave.Add opbank.name, newWave
        Set newWave = Nothing
    Next bank
    For Each m_site In theexec.sites
        m_continueFor = False
        hasFlag = False
        '1. Check fail flag to get which limit column.
        GetFlagDicByFlagStatus m_site, hasFlag, flagDic

        If Not hasFlag Then
            Call Print_Error_Message(Error_Info, "VBT_LIB_eFuse_Func", "Efuse_bincheck", "Can't find any BIN flag is true in site(" & CStr(m_site) & "), please check it !!!")
            theexec.Flow.TestLimit resultVal:=999, lowVal:=0, hiVal:=0, Tname:="Efuse_bincheck"
            m_continueFor = True
        ElseIf flagDic.Count > 1 Then
            Call Print_Error_Message(Error_Info, "VBT_LIB_eFuse_Func", "Efuse_bincheck", "More than one BIN flag is true in site(" & CStr(m_site) & "), please check it !!!")
            theexec.Flow.TestLimit resultVal:=999, lowVal:=0, hiVal:=0, Tname:="Efuse_bincheck"
            m_continueFor = True
        End If
        If Not m_continueFor Then
            Set opbank = GetBdfBank("CFG")
            If opbank.Fields.Exists("bkm_package") Then
                Set bkmfield = opbank.Fields("bkm_package")
            ElseIf opbank.Fields.Exists("bkm_process") Then
                Set bkmfield = opbank.Fields("bkm_process")
            Else
                Call Print_Error_Message(Error_Info, "VBT_LIB_eFuse_Func", "Efuse_bincheck", "Please check bkm category name!!!")
                theexec.Flow.TestLimit resultVal:=999, lowVal:=0, hiVal:=0, Tname:="PassBinNotMatch"
                m_continueFor = True
            End If
        End If
        
        If Not m_continueFor Then
            For Each Flag In flagDic
                If flagDic(Flag).Count > 1 Then
                    GlbUtility.WriteDlg "Site(" & CStr(m_site) & ") have " + CStr(flagDic(Flag).Count) + " column for Bin flag " + Flag + " check."
                End If
                Set tmpFlagDic = New Dictionary: tmpFlagDic.compareMode = TextCompare
                Set tmpFlagDic = flagDic(Flag)
                For Each key In tmpFlagDic
                    '2.Select BinLite limit column
                    If DictBinNameMapping.Exists(key) Then
                         m_Index = DictBinNameMapping.item(key)
                    Else
                        Call Print_Error_Message(Error_Info, "VBT_LIB_eFuse_Func", "Efuse_bincheck", "The BIN flag (" + key + ") not exist in DictBinNameMapping in site(" & CStr(m_site) & "), please check it !!!")
                        theexec.Flow.TestLimit resultVal:=999, lowVal:=0, hiVal:=0, Tname:="Efuse_bincheck"
                        m_continueFor = True
                        Exit For
                    End If
                    '3. Check the "bkm_package" category has fused or not
                    If Not m_continueFor Then
                        If bkmfield.DsscDecValue(m_site) <> 0 Or preWrite Then
                            FieldFail = False
                            Erase oValue(): Erase oHiLimit(): Erase oLoLimit(): Erase oCataName()
                            For m_Loop = 0 To UBound(BinCheckData(m_Index).CateArr)
                                If BinCheckData(m_Index).CateArr(m_loop).skipCheck Then
                                    GoTo fieldSkipCheck
                                End If
                                ReDim Preserve oValue(m_Loop): ReDim Preserve oHiLimit(m_Loop): ReDim Preserve oLoLimit(m_Loop): ReDim Preserve oCataName(m_Loop)
                                m_catename = BinCheckData(m_Index).CateArr(m_Loop).cateName
                                m_bankname = BinCheckData(m_Index).CateArr(m_Loop).bankName
                                
                                '4. Check which the category's bank name.
                                Set opbank = GetBdfBank(m_bankname)
                                
                                If BinCheckData(m_Index).CateArr(m_Loop).conditionCheck Then
                                    tmp = Split(BinCheckData(m_Index).CateArr(m_Loop).conditionInfo, "=")
                                    
                                    Set field = opbank.Fields(tmp(0))
                                    If GlbUtility.testedStages.Exists(field.BlowLocation) Then
                                        If field.DsscValue(m_site) = GlbUtility.String2Hex(tmp(1), hBytes:=field.hBytes) Then
                                            BinCheckData(m_Index).CateArr(m_Loop).conditionResult = True
                                        Else
                                            BinCheckData(m_Index).CateArr(m_Loop).conditionResult = False
                                        End If
                                    Else
                                        If field.TrimAteValue(m_site) = GlbUtility.String2Hex(tmp(1), hBytes:=field.hBytes) Then
                                            BinCheckData(m_Index).CateArr(m_Loop).conditionResult = True
                                        Else
                                            
                                            BinCheckData(m_Index).CateArr(m_Loop).conditionResult = False
                                        End If
                                    End If
                                End If
                                
                                If BinCheckData(m_Index).CateArr(m_Loop).MergeBitCheck = False And Not BinCheckData(m_Index).CateArr(m_Loop).CombineFieldCheck Then
                                    Set field = opbank.Fields(m_catename)
                                    If preWrite Then
                                        m_FusedValue = GetFieldValue(field)
                                    Else
                                        m_FusedValue = field.DsscDecValue
                                    End If
                                End If
                                
                                If BinCheckData(m_Index).CateArr(m_Loop).MergeBitCheck Or BinCheckData(m_Index).CateArr(m_Loop).SplitBitCheck Then
                                    m_catename = BinCheckData(m_Index).CateArr(m_Loop).cateNameOri
                                End If
                                
                                If preWrite Then
                                    Set m_tmpDSP = dicPreWriteWave(m_bankname)
                                Else
                                    If opbank.pgmMode = pgm_DAA Then
                                        Set m_tmpDSP = opbank.DaaCapWaveSerial
                                    ElseIf opbank.pgmMode = pgm_JTAG Then
                                        Set m_tmpDSP = opbank.JtagCapturedSerial
                                    End If
                                End If
                                
                                If BinCheckData(m_Index).CateArr(m_Loop).CombineFieldCheck Then
                                    m_stage = Replace(UCase(BdfDataBase.GetRealStage(BinCheckData(m_Index).CateArr(m_Loop).CombineFieldJob)), "_EARLY", "")
                                Else
                                    m_stage = Replace(UCase(BdfDataBase.GetRealStage(field.BlowLocation)), "_EARLY", "")
                                End If
                                '5. Limit Type
                                '  a. IDS
                                '  b. specific value
                                '  b. walking
                                '  c. two-1
                                '  d. group
                                '  e. rnage
                                operator = vbNullString
                                m_skipTest(m_site) = False
                                
                                For m_ruleIdx = 0 To UBound(BinCheckData(m_Index).CateArr(m_Loop).checkRules)
                                    If BinCheckData(m_Index).CateArr(m_Loop).conditionCheck Then
                                        If BinCheckData(m_Index).CateArr(m_Loop).conditionResult <> BinCheckData(m_Index).CateArr(m_Loop).checkRules(m_ruleIdx).ConditionChkTrue Then
                                            GoTo skip
                                        End If
                                    End If
                                    
                                    If BinCheckData(m_Index).CateArr(m_Loop).checkRules(m_ruleIdx).SkipTest Or m_skipTest(m_site) Then
                                        m_skipTest(m_site) = True
                                    Else
                                        If BinCheckData(m_Index).CateArr(m_Loop).checkRules(m_ruleIdx).siteInfo.Count = 0 Or _
                                           BinCheckData(m_Index).CateArr(m_Loop).checkRules(m_ruleIdx).siteInfo.Exists(CStr(m_site)) Then
                                            If BinCheckData(m_Index).CateArr(m_Loop).checkRules(m_ruleIdx).isOperator Then
                                                operator = BinCheckData(m_Index).CateArr(m_Loop).checkRules(m_ruleIdx).Rule
                                            Else
                                                If operator <> "" Then
                                                    If operator Like "|" Then
                                                        result(m_site) = result(m_site) Or BinCheckResult(m_tmpDSP, opbank, field, BinCheckData(m_Index).CateArr(m_Loop), m_ruleIdx, preWrite)
                                                    ElseIf operator Like "&" Then
                                                        result(m_site) = result(m_site) And BinCheckResult(m_tmpDSP, opbank, field, BinCheckData(m_Index).CateArr(m_Loop), m_ruleIdx, preWrite)
                                                    Else
                                                        Call Print_Error_Message(Error_Info, "VBT_LIB_eFuse_Func", "Efuse_bincheck", "The BKM category has not been fused, please check it !!!")
                                                        theexec.Flow.TestLimit resultVal:=999, lowVal:=0, hiVal:=0, Tname:="Efuse_bincheck"
                                                    End If
                                                    operator = vbNullString
                                                Else
                                                    result(m_site) = BinCheckResult(m_tmpDSP, opbank, field, BinCheckData(m_Index).CateArr(m_Loop), m_ruleIdx, preWrite)
                                                End If
                                            End If
                                        End If
                                    End If
skip:
                                Next m_ruleIdx
                                
                                oHiLimit(m_Loop) = 1: oLoLimit(m_Loop) = 1
                                oCataName(m_Loop) = m_catename
                                isCurrentAndTestedJob = IIf(UCase(m_stage) = UCase(theexec.CurrentJob) Or GlbUtility.testedStages.Exists(m_stage), True, False)
                                If gB_eFuse_Disable_ChkLMT_Flag And (Not isCurrentAndTestedJob) Then
                                    oValue(m_Loop) = 1
                                Else
                                    If m_skipTest(m_site) Then
                                        oValue(m_loop) = 1
                                    ElseIf isCurrentAndTestedJob Then
                                        If result(m_site) Then
                                            oValue(m_loop) = 1
                                        Else
                                            FieldFail = True
                                            oValue(m_loop) = 0
                                            If key <> tmpFlagDic.Keys(tmpFlagDic.Count - 1) Then Exit For
                                        End If
                                    Else
                                        oValue(m_loop) = CDbl(m_FusedValue(m_site))
                                        oHiLimit(m_loop) = 0: oLoLimit(m_loop) = 0
                                        If oValue(m_loop) <> 0 Then
                                            FieldFail = True
                                            If key <> tmpFlagDic.Keys(tmpFlagDic.Count - 1) Then Exit For
                                        End If
                                    End If
                                End If
fieldSkipCheck:
                            Next m_loop
                            
                            If Not FieldFail Or key = tmpFlagDic.Keys(tmpFlagDic.Count - 1) Then
                                GlbUtility.WriteDlg "Site(" & CStr(m_site) & ") Choose => " & key & " column from FuseBinCheck Table."
                                For i = 0 To UBound(oValue)
                                    theexec.Flow.TestLimit resultVal:=oValue(i), lowVal:=oLoLimit(i), hiVal:=oHiLimit(i), Tname:=oCataName(i)
                                Next i
                                Exit For
                            End If
                        Else
                            Call Print_Error_Message(Error_Info, "VBT_LIB_eFuse_Func", "Efuse_bincheck", "The BKM category has not been fused, please check it !!!")
                            theexec.Flow.TestLimit resultVal:=999, lowVal:=0, hiVal:=0, Tname:="Efuse_bincheck"
                        End If
                    End If
                Next key
            Next Flag
        End If
    Next m_site

Exit Function
errHandler:
    Call Print_Error_Message(Error_Info, "VBT_LIB_eFuse_Func", "Efuse_bincheck")
    If AbortTest Then Exit Function Else Resume Next
End Function

Private Function GetFieldValue(field As eFuseBdfField, Optional hexValue As Boolean = False) As SiteVariant
On Error GoTo errHandler

    If GlbUtility.testedStages.Exists(field.BlowLocation) Then
        If hexValue Then
            Set GetFieldValue = field.DsscValue
        Else
            Set GetFieldValue = field.DsscDecValue
        End If
    Else
        If hexValue Then
           Set GetFieldValue = field.TrimAteValue
        Else
           Set GetFieldValue = field.TrimAteDecValue
        End If
    End If

Exit Function
errHandler:
    Call Print_Error_Message(Error_Info, "VBT_LIB_eFuse_Func", "GetFieldValue")
    If AbortTest Then Exit Function Else Resume Next
End Function


Public Function Temp_eFuse_setWrite(FuseType As String, Srcfields As String, Tarfields As String)
On Error GoTo errHandler
Dim m_Loop As Integer
Dim m_value As New SiteVariant
Dim opbank As eFuseBdfBank
Dim sfield As eFuseBdfField, tfield As eFuseBdfField
Dim PassFlag As Long
Dim fieldsrcarr As Variant, fieldtararr As Variant

    PassFlag = 1

    fieldsrcarr = Split(CStr(Srcfields), ",")
    fieldtararr = Split(CStr(Tarfields), ",")
    
    If UBound(fieldsrcarr) <> UBound(fieldtararr) Then
        theexec.Datalog.WriteComment "Input arguments num are different!!"
        PassFlag = 0
        GoTo judgeresult
    End If
    
    Set opbank = GetBdfBank(FuseType)
    
    For m_Loop = 0 To UBound(fieldsrcarr)
        If opbank.Fields.Exists(fieldsrcarr(m_Loop)) Then
            Set sfield = opbank.Fields(fieldsrcarr(m_Loop))
    
            If sfield.BlowLocation = GlbUtility.currStage Then
                m_value = sfield.TrimAteDecValue
            Else
                m_value = sfield.DsscDecValue
            End If
            theexec.Datalog.WriteComment "Src Field : " & sfield.name
        Else
            theexec.Datalog.WriteComment "Src Field : " & sfield.name & " doesn't exist"
            PassFlag = 0
            GoTo judgeresult
        End If
        
        If opbank.Fields.Exists(fieldtararr(m_Loop)) Then
            Set tfield = opbank.Fields(fieldtararr(m_Loop))
            If tfield.BlowLocation = GlbUtility.currStage Then
                opbank.SetEfuse tfield.name, m_value, , , , , True
            End If
        Else
            theexec.Datalog.WriteComment "Target Field : " & tfield.name & " doesn't exist"
            PassFlag = 0
            GoTo judgeresult
        End If
    Next m_Loop
    
judgeresult:
    theexec.Flow.TestLimit resultVal:=PassFlag, lowVal:=1, hiVal:=1, Tname:="Temp_eFuse_setWrite"
    
Exit Function
errHandler:
    Call Print_Error_Message(Error_Info, "VBT_LIB_eFuse_Func", "Temp_eFuse_setWrite")
    If AbortTest Then Exit Function Else Resume Next
End Function

Public Function Temp_eFuse_setWriteZero(FuseType As String, Fields As String)
On Error GoTo errHandler
Dim fieldarr As Variant
Dim m_Loop As Integer
Dim m_value As New SiteVariant
Dim opbank As eFuseBdfBank
Dim field As eFuseBdfField
Dim PassFlag As Long
    
    PassFlag = 1
    
    fieldarr = Split(CStr(Fields), ",")
    
    Set opbank = GetBdfBank(FuseType)
    
    For m_Loop = 0 To UBound(fieldarr)
        If opbank.Fields.Exists(fieldarr(m_Loop)) Then
            m_value = 0
            Set field = opbank.Fields(fieldarr(m_Loop))
            If field.BlowLocation = GlbUtility.currStage Then
                opbank.SetEfuse field.name, m_value, , , , , True
            End If
        Else
            theexec.Datalog.WriteComment fieldarr(m_Loop) & " doesn't exist"
            PassFlag = 0
        End If
    Next

    theexec.Flow.TestLimit resultVal:=PassFlag, lowVal:=1, hiVal:=1, Tname:="Temp_eFuse_setWriteZero"
    
Exit Function
errHandler:
    Call Print_Error_Message(Error_Info, "VBT_LIB_eFuse_Func", "Temp_eFuse_setWriteZero")
    If AbortTest Then Exit Function Else Resume Next
End Function

'20190503: Created for the independent instance
Public Function ForceSepvmFusing_mod(performance_mode As String, FuseName As String)
On Error GoTo errHandler
Dim p_mode As Integer
Dim vdd_val As New SiteDouble
Dim EnableFuse As Boolean
Dim Sepvm_table As SEPVM_SHEET_Type
Dim site As Variant
Dim i As Integer

    p_mode = VddBinStr2Enum(performance_mode)
    
'    If AllBinCut(p_mode).Used = True Then
''        TheExec.Datalog.WriteComment "ForceSepvmFusing:" & VddBinName(p_mode)
    For Each site In theexec.sites
        vdd_val = BinCut(p_mode, CurrentPassBinCutNum(site)).OTHER_CP_Vmax(p_mode) + BinCut(p_mode, CurrentPassBinCutNum(site)).OTHER_CP1_GB(p_mode)
    Next site
    EnableFuse = True
    
    For i = 0 To UBound(Sepvm_table_Arr)
        If UCase(Sepvm_table_Arr(i).FuseName) Like UCase(FuseName) Then
            Sepvm_table = Sepvm_table_Arr(i)
            Exit For
        End If
    Next i
    
    Call SEPVM_Fuse_mod_AllCaseSupport2(p_mode, vdd_val, Sepvm_table, EnableFuse) 'for Crete
'    Else
'        TheExec.ErrorLogMessage VddBinName(p_mode) & " isn't used or tested in BinCut. ForceSepvmFusing has the error. Error!!"
'        TheExec.Flow.TestLimit resultVal:=999, lowVal:=1, hiVal:=1, Tname:=VddBinName(p_mode) & " isn't used or tested in BinCut. ForceSepvmFusing has the error. Error!!"
'    End If
Exit Function
errHandler:
    Call Print_Error_Message(Error_Info, "VBT_LIB_eFuse_Func", "ForceSepvmFusing_mod")
    If AbortTest Then Exit Function Else Resume Next
End Function

'20211109, Replace setWrite customfunction part to set bincut efuse and check
Public Function BinCut_PreSetAndCheck() As Long
On Error GoTo errHandler
Dim funcName As String: funcName = "BinCut_PreSetAndCheck"
Dim m_EQNum As New SiteLong, maxEQ As New SiteLong, GRADEVDD As New SiteDouble
Dim bankstr As Variant, bank As eFuseBdfBank, fieldStr As Variant, field As eFuseBdfField, fieldstg As String
Dim field_normal As eFuseBdfField, field_addition As eFuseBdfField
Dim need_calculate As New SiteBoolean: need_calculate = True
Dim need_calculate_addition As New SiteBoolean: need_calculate_addition = True
Dim site As Variant 'Carter, 20240304

    If BdfDataBase.Banks("CFG").Fields.Exists("product_identifier") Then
        Set field_normal = BdfDataBase.Banks("CFG").Fields("product_identifier")
        For Each site In theexec.sites
            If GlbUtility.testedStages.Exists(field_normal.BlowLocation) Then
                CurrentPassBinCutNum_normal(site) = field_normal.DsscDecValue + 1
                If CurrentPassBinCutNum_normal(site) > Total_Bincut_Num Then
                    need_calculate(site) = False
                    GlbUtility.WriteDlg "site:" & site & ", product_identifier " & CurrentPassBinCutNum_normal(site) & " > Total_Bincut_Num " & Total_Bincut_Num & " , Error!!!"
                    theexec.Flow.TestLimit resultVal:=0, lowVal:=-1, hiVal:=-1, Tname:="product_identifier Error"
                End If
            Else
                CurrentPassBinCutNum_normal(site) = CurrentPassBinCutNum(site)
            End If
        Next site
        CurrentPassBinCutNum_additional = CurrentPassBinCutNum_normal
    Else
        CurrentPassBinCutNum_normal = 1
    End If
    
    '''notice. if bdf don't have product_identifier_shadow for wlft2, it will become use cp1 fused product_identify, not use CurrentPassBinCutNum
    If BdfDataBase.Banks("CFG").Fields.Exists("product_identifier_shadow") Then
        Set field_addition = BdfDataBase.Banks("CFG").Fields("product_identifier_shadow")
        If field_addition.DefaultOrReal = dr_real Then
            For Each site In theexec.sites
                If GlbUtility.testedStages.Exists(field_addition.BlowLocation) Then
                    CurrentPassBinCutNum_additional(site) = field_addition.DsscDecValue + 1
                    If CurrentPassBinCutNum_additional(site) > Total_Bincut_Num Then
                        need_calculate_addition(site) = False
                        GlbUtility.WriteDlg "site:" & site & ", product_identifier_shadow " & CurrentPassBinCutNum_additional(site) & " > Total_Bincut_Num " & Total_Bincut_Num & " , Error!!!"
                        theexec.Flow.TestLimit resultVal:=0, lowVal:=-1, hiVal:=-1, Tname:="product_identifier_shadow Error"
                    End If
                Else
                    CurrentPassBinCutNum_additional(site) = CurrentPassBinCutNum(site)
                End If
            Next site
        End If
    End If

    For Each bankstr In BdfDataBase.Banks.Keys
        If bankstr <> Empty Then
            Set bank = BdfDataBase.Banks(bankstr)
            
            If Not GlbUtility.OnlineMode Then bank.PutIdsCodes 'offline

            If bank.HadVddBinFuse Then
                bank.PutBinCutCodes
                
                For Each fieldStr In bank.Fields.Keys
                    Set field = bank.Fields(fieldStr)
                    fieldstg = UCase(BdfDataBase.GetRealStage(field.BlowLocation))
                    If field.Algorithm = alg_vddbin And (fieldstg = theexec.CurrentJob Or GlbUtility.testedStages.Exists(fieldstg)) Then
                        If Not BdfDataBase.DicVddBinIdsFieldMap.Exists(Replace(LCase(fieldStr), "_shadow", "")) Then
                            theexec.Flow.TestLimit resultVal:=0, lowVal:=1, hiVal:=1, Tname:=funcName
                            Call Print_Error_Message(Error_Info, "VBT_LIB_eFuse_Func", "BinCut_PreSetAndCheck", "Can't get ids mapping field! " & bankstr & ">>" & fieldStr)
                        Else
                            For Each site In theexec.sites
                                If field.BlowLocation = GlbUtility.currStage Then
                                    If LCase(field.BlowLocation) = LCase(BincutAdditionalSheetName) Then
                                        If need_calculate_addition(site) Then
                                            m_EQNum = GetVddBinEqu(field, CurrentPassBinCutNum_additional, maxEQ, GRADEVDD, True, , False)
                                            theexec.Flow.TestLimit resultVal:=m_EQNum, lowVal:=1, hiVal:=maxEQ, Tname:=field.name & "_EQ"
                                        End If
                                    Else
                                        If need_calculate(site) Then
                                            m_EQNum = GetVddBinEqu(field, CurrentPassBinCutNum_normal, maxEQ, GRADEVDD, , , False)
                                            theexec.Flow.TestLimit resultVal:=m_EQNum, lowVal:=1, hiVal:=maxEQ, Tname:=field.name & "_EQ"
                                        End If
                                    End If
                                ElseIf GlbUtility.testedStages.Exists(field.BlowLocation) Then
                                    If LCase(field.BlowLocation) = LCase(BincutAdditionalSheetName) Then
                                        If need_calculate_addition(site) Then
                                            m_EQNum = GetVddBinEqu(field, CurrentPassBinCutNum_additional, maxEQ, GRADEVDD, True)
                                            theexec.Flow.TestLimit resultVal:=m_EQNum, lowVal:=1, hiVal:=maxEQ, Tname:=field.name & "_EQ"
                                        End If
                                    Else
                                        If need_calculate(site) Then
                                            m_EQNum = GetVddBinEqu(field, CurrentPassBinCutNum_normal, maxEQ, GRADEVDD)
                                            theexec.Flow.TestLimit resultVal:=m_EQNum, lowVal:=1, hiVal:=maxEQ, Tname:=field.name & "_EQ"
                                        End If
                                    End If
                                End If
                            Next
                        End If
                    End If
                Next
            End If
        End If
    Next
    
Exit Function
errHandler:
    Call Print_Error_Message(Error_Info, "VBT_LIB_eFuse_Func", "BinCut_PreSetAndCheck")
    If AbortTest Then Exit Function Else Resume Next
End Function

Public Function EFUSE_Flat_Pattern_Check(Flat_Pattern As Pattern, bank As String, Optional PrePatSet As Pattern, Optional DigSource As String = vbNullString, Optional Validating_ As Boolean)
On Error GoTo errHandler
Dim instance_name As String:: instance_name = theexec.DataManager.instancename
Dim SiteVarValue As New SiteLong, m_siteVar As String
Dim opbank As eFuseBdfBank
Dim PrePatResult As New SiteBoolean
Dim tempVarValue As New SiteVariant
Dim site As Variant 'Carter, 20240304
Dim initSiteStatus As New SiteBoolean

    If Validating_ Then
        If Flat_Pattern.value <> "" Then Call PrLoadPattern(Flat_Pattern.value)
        If PrePatSet.value <> "" Then Call PrLoadPattern(PrePatSet.value)
        Exit Function
    End If
    
    Set opbank = GetBdfBank(bank)
    m_siteVar = bank + "Chk_Var"
    
    '20230314, T-Col request from C651
    For Each site In theexec.sites
        tempVarValue = theexec.sites(site).SiteVariableValue(m_siteVar)
    Next site
    opbank.BlankCheckCount = opbank.BlankCheckCount + 1
    GlbUtility.IniSiteVar m_siteVar, -1

    TheHdw.Digital.ApplyLevelsTiming True, True, True, tlPowered
    If PrePatSet.value <> "" Then
        initSiteStatus = TheExec.sites.Active
        PrePatResult = EfuseExecInitPattern(PrePatSet.value, DigSource)
        If Not PrePatResult.Any(True) Then
            Exit Function
        ElseIf PrePatResult.Any(False) Then
            TheExec.sites.Selected = PrePatResult
        End If
    End If
    Call TheHdw.Patterns(Flat_Pattern).test(pfAlways, 0)

    For Each site In theexec.sites
        If (TheHdw.Digital.Patgen.PatternBurstPassedPerSite(site) = False) Then 'If not blank
            opbank.IsBlank = 0
            SiteVarValue = 2
        Else ' It's means this Fuse area is all balnk.
            SiteVarValue = 1
        End If
        
        If EFUSE_REFUSE_FOR_PTE And EFUSE_POWER_OFF_SETTING Then
            SiteVarValue = 1
        End If
        If Not GlbUtility.OnlineMode Then SiteVarValue = 1 'And Site = 1 Then SiteVarValue = 2
        theexec.sites(site).SiteVariableValue(m_siteVar) = CLng(SiteVarValue)
        
        If tempVarValue > 0 And tempVarValue <> SiteVarValue And opbank.BlankCheckCount = 2 Then
            opbank.IsSameCheckVar = 0
        End If
    Next site

    theexec.Flow.TestLimit resultVal:=SiteVarValue, lowVal:=1, hiVal:=2, Tname:=m_siteVar
    If opbank.BlankCheckCount = 2 Then
        theexec.Flow.TestLimit resultVal:=opbank.IsSameCheckVar, lowVal:=1, hiVal:=1, Tname:=bank + "_FlatCheckSummary"
        For Each site In theexec.sites
            If opbank.IsSameCheckVar(site) = 0 Then
                theexec.sites(site).FlagState("F_FlatCheck_Summary") = logicTrue
            End If
        Next site
        Call opbank.InitialMultiCheckVarVariable
    End If

    DebugPrintFunc Flat_Pattern.value
    If (PrePatSet.value <> "") Then
        TheExec.sites.Selected = initSiteStatus
    End If
Exit Function
errHandler:
    Call Print_Error_Message(Error_Info, "VBT_LIB_eFuse_Func", "EFUSE_Flat_Pattern_Check")
    If AbortTest Then Exit Function Else Resume Next
End Function

Public Function Bincut_eFuse_OtherSetWrite(FuseType As String, Srcfields As String, Tarfields As String, offset As Long)
On Error GoTo errHandler
Dim funcName As String:: funcName = "BinCut_OtherSet"
Dim opbank As eFuseBdfBank
Dim m_Loop As Integer
Dim m_value As New SiteVariant
Dim m_fusevalue As New SiteVariant
Dim p_mode As Long, hiV As Double, lowV As Double
Dim sfield As eFuseBdfField, tfield As eFuseBdfField, field_identifier As eFuseBdfField
Dim fieldsrcarr As Variant, fieldtararr As Variant
Dim hasAdditionBinCut As Boolean: hasAdditionBinCut = False
Dim PassFlag As Long: PassFlag = 1
Dim vddFuseValue As New SiteVariant
Dim site As Variant 'Carter, 20240304

    fieldsrcarr = Split(CStr(Srcfields), ",")
    fieldtararr = Split(CStr(Tarfields), ",")
    
    If UBound(fieldsrcarr) <> UBound(fieldtararr) Then
        theexec.Datalog.WriteComment "Input arguments num are different!!"
        PassFlag = 0
        GoTo judgeresult
    End If

    Set opbank = GetBdfBank(FuseType)

    For m_Loop = 0 To UBound(fieldsrcarr)
        If opbank.Fields.Exists(fieldsrcarr(m_Loop)) And opbank.Fields.Exists(fieldtararr(m_Loop)) Then
            Set sfield = opbank.Fields(fieldsrcarr(m_Loop))
            Set tfield = opbank.Fields(fieldtararr(m_Loop))
            If LCase(sfield.BlowLocation) = LCase(BincutAdditionalSheetName) Then hasAdditionBinCut = True

            p_mode = IIf(Not GlbUtility.OnlineMode, BdfDataBase.DicVddBinPmodeMap(Replace(LCase(sfield.name), "_shadow", "")), VddBinStr2Enum(Replace(LCase(sfield.name), "_shadow", "")))
            If tfield.BlowLocation = GlbUtility.currStage Or (GlbUtility.testedStages.Exists(tfield.BlowLocation)) Then
                tfield.Resolution = sfield.Resolution
                tfield.SetWriteByBincut = True
                For Each site In theexec.sites
                    If sfield.BlowLocation = GlbUtility.currStage Then
                        If sfield.DsscDecValue <> 0 Then  'for retest case
                            vddFuseValue = sfield.DsscValue
                        Else
                            vddFuseValue = sfield.TrimAteValue
                        End If
                        m_value = GlbUtility.Hex2Dbl(CStr(vddFuseValue)) * sfield.Resolution + BdfDataBase.BaseVoltage
                    Else
                        m_value = sfield.FuseMeasureValue
                    End If
                    theexec.Datalog.WriteComment "site" & CStr(site) & "Src Field : " & sfield.name & " Bincut voltage = " & CStr(m_value)
                    m_fusevalue = GlbUtility.CeilingValue((m_value - offset - BdfDataBase.BaseVoltage) / sfield.Resolution, 1)
                    theexec.Datalog.WriteComment "site" & CStr(site) & "Target Field : " & tfield.name & " Bincut voltage - Offset" & "(" & CStr(offset) & ")" & " = " & CStr(m_fusevalue)
                    tfield.BVHLimit = m_value - offset
                    tfield.BVLLimit = m_value - offset
                Next site
                If tfield.BlowLocation = GlbUtility.currStage Then opbank.SetEfuse tfield.name, m_fusevalue, , , , , True
            End If
        Else
            theexec.Datalog.WriteComment "Src Field : " & sfield.name & " or " & "Target Field : " & tfield.name & " doesn't exist"
            PassFlag = 0
            GoTo judgeresult
        End If
    Next m_Loop
    
judgeresult:
    theexec.Flow.TestLimit resultVal:=PassFlag, lowVal:=1, hiVal:=1, Tname:="BinCut_OtherSet"

Exit Function
errHandler:
    Call Print_Error_Message(Error_Info, "VBT_LIB_eFuse_Func", "Bincut_eFuse_OtherSetWrite")
    If AbortTest Then Exit Function Else Resume Next
End Function

Public Function EfuseExecInitPattern(prePattern As String, DigSource As String) As SiteBoolean
On Error GoTo errHandler
Dim instance_name As String
Dim site As Variant, PrePat As Variant
Dim PrePatArray() As String, PatCount As Long
Dim PrePatResultBySite As New SiteBoolean: PrePatResultBySite = True

    If prePattern <> "" Then
        instance_name = LCase(theexec.DataManager.instancename)
        PrePatArray = theexec.DataManager.Raw.GetPatternsInSet(prePattern, PatCount)
        For Each PrePat In PrePatArray
            If UCase(CStr(PrePat)) Like "*_DSRMDSSC_*" Or UCase(CStr(PrePat)) Like "*_DSRMDSSC*" Or UCase(CStr(PrePat)) Like "*_SRMDSSC*" Then
                If DigSource <> "" Then
                    If CStr(PrePat) Like ".\*" Then
                        Call Sub_SourceEMA_SelSRM(prePattern, instance_name, DigSource)
                    Else
                        Call Sub_SourceEMA_SelSRM(CStr(PrePat), instance_name, DigSource)
                    End If
                End If
            End If
            Call TheHdw.Patterns(PrePat).test(pfAlways, 0)
            Call update_Pattern_result_to_PattPass(TheHdw.Digital.Patgen.PatternBurstPassedPerSite, PrePatResultBySite)
        Next PrePat

    End If
    
    Set EfuseExecInitPattern = PrePatResultBySite
    
Exit Function
errHandler:
    Call Print_Error_Message(Error_Info, "VBT_LIB_eFuse_Func", "EfuseExecInitPattern")
    If AbortTest Then Exit Function Else Resume Next
End Function

Private Function BinCheckResult(wave As DSPWave, opbank As eFuseBdfBank, field As eFuseBdfField, checkData As EFuseBinCheckData, ruleIdx As Long, preWrite As Boolean) As Boolean
On Error GoTo errHandler
Dim idsField As eFuseBdfField
Dim m_fusevalueHex As String
Dim m_fusevalueDec As Double
Dim m_startBin As Long, m_endBit As Long, Count As Long, groupCnt As Long, i As Long
Dim groupDic As New Dictionary
Dim key As Variant
Dim isFail As Boolean: isFail = False
Dim binaryStr As String: binaryStr = vbNullString
Dim bitwidth As Long, bytes As Long
Dim over32Bits As Boolean: over32Bits = False
Dim item As Variant
Dim tmpDspWave As New DSPWave
Dim tmpCnt As Long

    If checkData.CombineFieldCheck Then
        tmpCnt = 0
        tmpDspWave.CreateConstant 0, checkData.Width
        For Each item In checkData.listCombineFieldsGroup
            'Set field = opbank.Fields(item.name)
            For i = item("LSB") To item("MSB")
                binaryStr = CStr(wave.Element(i)) + binaryStr
                tmpDspWave.Element(tmpCnt) = wave.Element(i)
                tmpCnt = tmpCnt + 1
            Next i
        Next item

        If checkData.Width > 32 Then
            over32Bits = True
            m_fusevalueHex = GlbUtility.Bin2HexStr(binaryStr, Ceiling(checkData.Width / 4))
        Else
            m_fusevalueDec = GlbUtility.Bin2Dec(binaryStr)
        End If
        
        m_startBin = 0
        m_endBit = checkData.Width - 1
    ElseIf checkData.SplitBitCheck Or checkData.MergeBitCheck Then
        m_startBin = checkData.Bit_LSB
        m_endBit = checkData.Bit_MSB
        
        If checkData.SplitBitCheck Then
            m_startBin = field.LSB + m_startBin
            m_endBit = field.LSB + m_endBit
        End If
        bitwidth = m_endBit - m_startBin + 1
        With checkData.checkRules(ruleIdx)
            If .IDSCheck Or .ValueCheck Or .RangeCheck Then
                For i = m_startBin To m_endBit
                    binaryStr = CStr(wave.Element(i)) + binaryStr
                Next i
                If bitwidth > 32 Then
                    over32Bits = True
                    bytes = Ceiling(bitwidth / 4)
                    m_fusevalueHex = GlbUtility.Bin2HexStr(binaryStr, bytes)
                Else
                    m_fusevalueDec = GlbUtility.Bin2Dec(binaryStr)
                End If
            End If
        End With
    Else
        bitwidth = field.size
        If bitwidth > 32 Then
            over32Bits = True
            m_fusevalueHex = IIf(preWrite, UCase(CStr(GetFieldValue(field, True))), UCase(CStr(field.DsscValue)))
        Else
            m_fusevalueDec = IIf(preWrite, GetFieldValue(field), field.DsscDecValue)
        End If
        m_startBin = field.LSB
        m_endBit = field.msb
    End If
    
    If checkData.checkRules(ruleIdx).IDSCheck Then
        idsField = opbank.Fields(checkData.checkRules(ruleIdx).Rule)
        If over32Bits Then
            If preWrite And GlbUtility.xHexCompare(CStr(GetFieldValue(idsField, True)), m_fusevalueHex, ChkEqualTo) Then
                BinCheckResult = True
            ElseIf GlbUtility.xHexCompare(CStr(idsField.DsscValue), m_fusevalueHex, ChkEqualTo) Then
                BinCheckResult = True
            Else
                BinCheckResult = False
            End If
        Else
            If preWrite And m_fusevalueDec = GetFieldValue(idsField) Then
                BinCheckResult = True
            ElseIf m_fusevalueDec = idsField.DsscDecValue Then
                BinCheckResult = True
            Else
                BinCheckResult = False
            End If
        End If
    ElseIf checkData.checkRules(ruleIdx).ValueCheck Then
        If over32Bits Then
            If GlbUtility.xHexCompare(checkData.checkRules(ruleIdx).specificHexValue, m_fusevalueHex, ChkEqualTo) Then
                BinCheckResult = True
            Else
                BinCheckResult = False
            End If
        Else
            If checkData.checkRules(ruleIdx).specificDecValue = m_fusevalueDec Then
                BinCheckResult = True
            Else
                BinCheckResult = False
            End If
        End If
    ElseIf checkData.checkRules(ruleIdx).WalkingCheck Then
        Count = 0
        If checkData.CombineFieldCheck Then
            Count = tmpDspWave.CountElements(EqualTo, 1)
        Else
            For i = m_startBin To m_endBit
                If wave.Element(i) = 1 Then Count = Count + 1
            Next i
        End If
        If Count = checkData.checkRules(ruleIdx).Walkingvalue Then
            BinCheckResult = True
        Else
            BinCheckResult = False
        End If
    ElseIf checkData.checkRules(ruleIdx).GroupCheck Or checkData.checkRules(ruleIdx).TwoCheck Then
        Count = 0
        groupCnt = 1

        For i = m_startBin To m_endBit
            If (i - m_startBin) < checkData.checkRules(ruleIdx).GroupBits * groupCnt Then
                If checkData.CombineFieldCheck Then
                    If tmpDspWave.Element(i) = 1 Then Count = Count + 1
                Else
                    If wave.Element(i) = 1 Then Count = Count + 1
                End If
                
            Else
                groupDic.Add groupCnt, Count
                groupCnt = groupCnt + 1
                Count = 0
                If checkData.CombineFieldCheck Then
                    If tmpDspWave.Element(i) = 1 Then Count = Count + 1
                Else
                    If wave.Element(i) = 1 Then Count = Count + 1
                End If
            End If
        Next i
        groupDic.Add groupCnt, Count
        
        Count = 0
        For Each key In groupDic
            If groupDic.item(key) = checkData.checkRules(ruleIdx).FailBit Then
                Count = Count + 1
            ElseIf groupDic.item(key) > checkData.checkRules(ruleIdx).FailBit Then
                isFail = True
                Exit For
            End If
        Next key
        
        If isFail Or (Count > checkData.checkRules(ruleIdx).GroupPick) Or (checkData.checkRules(ruleIdx).TwoCheck And Count <> 2) Then
            BinCheckResult = False
        Else
            BinCheckResult = True
        End If
    ElseIf checkData.checkRules(ruleIdx).RangeCheck Then
        If bitwidth > 32 Then
            If GlbUtility.xHexCompare(m_fusevalueHex, checkData.checkRules(ruleIdx).LowHexValue, ChkGreaterEqualThan) And GlbUtility.xHexCompare(m_fusevalueHex, checkData.checkRules(ruleIdx).HighHexValue, ChkLessEqualThan) Then
                BinCheckResult = True
            Else
                BinCheckResult = False
            End If
        Else
            If (m_fusevalueDec >= checkData.checkRules(ruleIdx).LowDecValue) And (m_fusevalueDec <= checkData.checkRules(ruleIdx).HighDecValue) Then
                BinCheckResult = True
            Else
                BinCheckResult = False
            End If
        End If
    End If

Exit Function
errHandler:
    Call Print_Error_Message(Error_Info, "VBT_LIB_eFuse_Func", "BinCheckResult")
    If AbortTest Then Exit Function Else Resume Next
End Function

Private Function GetFlagDicByFlagStatus(m_site As Variant, hasFlag As Boolean, flagDic As Dictionary)
On Error GoTo errHandler
Dim binstr As New Dictionary: binstr.compareMode = TextCompare
Dim key As Variant, binStrArr As Variant
Dim tmpKey As String

    Set flagDic = New Dictionary: flagDic.compareMode = TextCompare
    For Each key In DictBinNameMapping
        If key <> "" Then
            If theexec.sites.item(m_site).FlagState(key) = logicTrue Then
                If Not flagDic.Exists(key) Then
                    binstr.Add key, True
                    flagDic.Add key, binstr
                    hasFlag = True
                End If
            ElseIf theexec.sites.item(m_site).FlagState("F_" + key) = logicTrue Then
                Set binstr = Nothing
                hasFlag = True
                If Not flagDic.Exists(key) Then
                    binstr.Add key, True
                    flagDic.Add "F_" + key, binstr
                Else
                    Set binstr = flagDic(key)
                    binstr.Add key, True
                    Set flagDic("F_" + key) = binstr
                End If
            ElseIf key Like "*_*" Then
                Set binstr = Nothing
                binStrArr = Split(key, "_")
                tmpKey = binStrArr(0)
                If theexec.sites.item(m_site).FlagState(tmpKey) = logicTrue Then
                    hasFlag = True
                    If Not flagDic.Exists(tmpKey) Then
                        binstr.Add key, True
                        flagDic.Add tmpKey, binstr
                    Else
                        Set binstr = flagDic(tmpKey)
                        binstr.Add key, True
                        Set flagDic(tmpKey) = binstr
                    End If
                End If
            End If
        End If
    Next key
    
Exit Function
errHandler:
    Call Print_Error_Message(Error_Info, "VBT_LIB_eFuse_Func", "GetFlagDicByFlagStatus")
    If AbortTest Then Exit Function Else Resume Next
End Function
