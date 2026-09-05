Attribute VB_Name = "VBT_LIB_Common_AP"
Option Explicit
Function VBT_IEDA_Registry(RegistryName As String, Optional OnOff As Boolean = True, Optional DebugPrint As Boolean = True)

On Error GoTo errHandler
    Dim funcName As String:: funcName = "VBT_IEDA_Registry"
    Dim InputStr As String
    If OnOff Then
        Call IEDA_Initialize(InputStr)  'clean up strings
        Call IEDA_GetString(InputStr, RegistryName)  'compose ieda string
        Call IEDA_AutoCheck_Print(InputStr, RegistryName, DebugPrint)   'show log
        Call IEDA_SaveRegistry(InputStr, RegistryName)  'save to registry
    End If

Exit Function

errHandler:
    Call Print_Error_Message(Error_Info, "VBT_LIB_Common_AP", "VBT_IEDA_Registry") 'Add ErrHandler 2023/08/18
     If AbortTest Then Exit Function Else Resume Next

End Function

Public Function ECID_DTS() 'VBT function
    On Error GoTo errHandler
    Dim funcName As String:: funcName = "ECID_DTS"
    '20200529 Oliver
    Dim OCR_String_32Sites As String, OCR_StringBySite() As String, SiteCount As Integer
    Dim OCR_String As String, Condition_1 As String, Condition_2 As String
    Dim i As Integer
    OCR_String_32Sites = vbNullString
    OCR_String_32Sites = RegKeyRead("HandlerBarCodeString") 'Read from regedit
    OCR_StringBySite = Split(OCR_String_32Sites, ",")
    Dim Flag_OCR_function As Boolean

   

    '20200907 Check OCR function start
    Flag_OCR_function = False
    For i = UBound(OCR_StringBySite) To 0 Step -1
        If OCR_StringBySite(i) = "0" Then
            Flag_OCR_function = False
        ElseIf OCR_StringBySite(i) = Chr(32) Then
            Flag_OCR_function = False
        Else
            Flag_OCR_function = True
            Exit For
        End If
    Next i
    '20200907 Check OCR function End
'    OCR_String = Replace(OCR_String_32Sites, ",", "")
'    SiteCount = UBound(OCR_StringBySite)                    'To judge how manys sites automaticall
'    If SiteCount = -1 Then
'        Condition_1 = String(0, "0")                        'all "0" status
'        Condition_2 = String(0, ",")                        'all "," status
'    Else
'        Condition_1 = String(SiteCount + 1, "0")
'        Condition_2 = String(SiteCount, ",")
'    End If
            
'    If Flag_DTS_function = True Then
'        If OCR_String = Condition_1 Or OCR_String_32Sites = Condition_2 Or OCR_String_32Sites = "" Then
'            Call GetStoredData_ECIDOCR_Compare
'        Else
'            CloseFunction
'        End If
'    Else
'        CloseFunction
'    End If
    
    '200907 update DTS function and OCR function conditions
    If Flag_DTS_function = True Then
        Call DTS_GetStoredData_Compare
    ElseIf Flag_OCR_function = True Then
        CloseFunction
    Else
        TheExec.Datalog.WriteComment ("WarNing: DTS system and OCR Handler system do not run.")
    End If
'''    If Flag_DTS_function = True And Flag_OCR_function = False Then
'''        Call DTS_GetStoredData_Compare
'''    ElseIf Flag_DTS_function = False And Flag_OCR_function = True Then
'''        CloseFunction
'''    ElseIf Flag_DTS_function = True And Flag_OCR_function = True Then
'''        CloseFunction
'''    ElseIf Flag_DTS_function = False And Flag_OCR_function = False Then
'''        CloseFunction
'''    End If
    
    Exit Function

errHandler:
    Call Print_Error_Message(Error_Info, "VBT_LIB_Common_AP", "ECID_DTS") 'Add ErrHandler 2023/08/18
    If AbortTest Then Exit Function Else Resume Next
End Function

Public Function Harvest_eFuse_Write(ByVal FuseBlockName As String, ByVal FuseCategoryName As String, ByVal FailFlagOrValue As String) As Long
    On Error GoTo errHandler
    Dim funcName As String:: funcName = "Harvest_eFuse_Write"

''''FuseBlockName = CFG;CFG;CFG;CFG
''''FuseCategoryName = gfx_fstp_harvest_te_misc;CFG_Condition_[47:44];ecpu_harvest;harvesting_bin
''''FailFlagOrValue = F_GFX_HARV[4:0];F_ECPU_CORE[3:0];0;0
''''------Special Case: Use the() Function to indicate which value want to Fuse------
''''------BitPos2Dec(F_GFX_HARV[4:0]) or BitPos2Dec(b00100)------
    'Dim vSite As Variant
    Dim i As Long
    Dim j As Long
    Dim site As Variant
    
    Dim FuseBlockNameArr() As String
    Dim FuseCategoryNameArr() As String
    Dim FailFlagOrValueArr() As String
    Dim tempStr As String
    Dim TempStrArr() As String
    Dim tempFuseCategoryNameArr() As String
    Dim tempFailFlagOrValueArr() As String
'    Dim Slng_FuseValue As New SiteLong
    Dim Sdbl_FuseValue As New SiteDouble
    Dim tempSlng As New SiteLong
'    Dim Svar_ReturnDec() As New SiteVariant
    
'    Dim Svar_FuseValue As New SiteVariant
    Dim sInstName As String
    Dim opbank As eFuseBdfBank
    Dim field As eFuseBdfField
    Dim HarvCustomFunc As String
    
         
    FuseBlockNameArr = Split(FuseBlockName, ";")
    FuseCategoryNameArr = Split(FuseCategoryName, ";")
    FailFlagOrValueArr = Split(FailFlagOrValue, ";")
    
    sInstName = TheExec.DataManager.instanceName
    TheExec.Datalog.WriteComment "<" & sInstName & ">"

    If (UBound(FuseBlockNameArr) = UBound(FuseCategoryNameArr)) And (UBound(FuseBlockNameArr) = UBound(FailFlagOrValueArr)) Then
        For i = 0 To UBound(FuseBlockNameArr)
            HarvCustomFunc = vbNullString
            tempStr = Harvest_StrExpand(FuseCategoryNameArr(i))
            tempFuseCategoryNameArr = Split(tempStr, ",")
            tempStr = Harvest_StrExpand(FailFlagOrValueArr(i), , , , , HarvCustomFunc)
            tempFailFlagOrValueArr = Split(tempStr, ",")
            
            Set opbank = GetBdfBank(FuseBlockNameArr(i)) 'Set FuseType
            Set field = opbank.Fields(FuseCategoryNameArr(i)) 'Set FuseName
            
            If (UBound(tempFuseCategoryNameArr) = 0) And (UBound(tempFailFlagOrValueArr) = 0) Then
                If auto_isBinaryString(tempFailFlagOrValueArr(0)) Then
                    tempFailFlagOrValueArr(0) = Replace(UCase(tempFailFlagOrValueArr(0)), "B", "")
                    tempFailFlagOrValueArr(0) = Replace(UCase(tempFailFlagOrValueArr(0)), "_", "")
                    tempFailFlagOrValueArr(0) = CStr(Bin2Dec(tempFailFlagOrValueArr(0)))
                ElseIf auto_isHexString(tempFailFlagOrValueArr(0)) Then
                    tempFailFlagOrValueArr(0) = CStr(auto_HexStr2Value(tempFailFlagOrValueArr(0)))
                ElseIf UCase(tempFailFlagOrValueArr(0)) = "NA" Then
                    
                End If
                If IsNumeric(tempFailFlagOrValueArr(0)) Then
                    'FuseCategoryName = harvesting_bin
                    'FailFlagOrValue = 0
                    Sdbl_FuseValue = CDbl(tempFailFlagOrValueArr(0))
'                    Call auto_eFuse_SetWriteVariable_SiteAware(FuseBlockNameArr(i), FuseCategoryNameArr(i), Sdbl_FuseValue, True)
                    opbank.SetEfuse field.name, Sdbl_FuseValue, , , , , True
''''''                    For Each site In theexec.sites.Active
''''''                        opbank.SetEfuseBySite field.name, Sdbl_FuseValue
''''''                        opbank.SetEfuseBySitePf field.name, True
''''''                    Next site
                    
                ElseIf HarvCustomFunc <> "" Then
                    Call JudgeCustomFun(CVar(tempFailFlagOrValueArr(0)), HarvCustomFunc, Sdbl_FuseValue)
'                    Call auto_eFuse_SetWriteVariable_SiteAware(FuseBlockNameArr(i), FuseCategoryNameArr(i), Sdbl_FuseValue, True)
                    opbank.SetEfuse field.name, Sdbl_FuseValue, , , , , True
''''''                    For Each site In theexec.sites.Active
''''''                        opbank.SetEfuseBySite field.name, Sdbl_FuseValue
''''''                        opbank.SetEfuseBySitePf field.name, True
''''''                    Next site
                ElseIf InStr(1, tempStr, "&") > 0 Then '231120 QJ: Add for 3 cases power binning. F_PWRBIN_OTHER = 3; F_PWRBIN_HIGH =2; F_PWRBIN_LOW =1
                    'TempStr = F_PWRBIN_OTHER&F_PWRBIN_HIGH&F_PWRBIN_LOW
                    TempStrArr = Split(tempStr, "&")
                    For Each site In TheExec.sites.Active
                        Sdbl_FuseValue = UBound(TempStrArr) + 1
                        For j = 0 To UBound(TempStrArr)
                            If TheExec.sites.item(site).FlagState(TempStrArr(j)) = logicTrue Then
                                opbank.SetEfuse field.name, Sdbl_FuseValue, , , , , True
                            Else
                                Sdbl_FuseValue = Sdbl_FuseValue - 1
                            End If
                        Next j
                        If Sdbl_FuseValue = 0 Then
                            TheExec.Datalog.WriteComment "<Error> Harvest_eFuse_Write: " & tempStr & " are all FALSE. One of them should be TRUE."
                        End If
                    Next site
                Else
                    'FuseCategoryName = CFG_Condition_47
                    'FailFlagOrValue = F_ECPU_CORE3
                    If UCase(tempFailFlagOrValueArr(0)) = "NA" Then
                        TheExec.Datalog.WriteComment "Harvest_eFuse_Write:" & tempFuseCategoryNameArr(0) & ":Did not need to fuse in current stage."
                    Else
                        Sdbl_FuseValue = Harvest_GetAllSiteFlagState(tempFailFlagOrValueArr(0), 1, 0, 1)
    '                    Call auto_eFuse_SetWriteVariable_SiteAware(FuseBlockNameArr(i), FuseCategoryNameArr(i), Sdbl_FuseValue, True)
                        opbank.SetEfuse field.name, Sdbl_FuseValue, , , , , True
''''''                        For Each site In theexec.sites.Active
''''''                            opbank.SetEfuseBySite field.name, Sdbl_FuseValue
''''''                            opbank.SetEfuseBySitePf field.name, True
''''''                        Next site
                    End If
                End If
''''                Call auto_eFuse_SetWriteVariable_SiteAware(FuseBlockNameArr(i), FuseCategoryNameArr(i), Slng_FuseValue, True)
                
            ElseIf UBound(tempFuseCategoryNameArr) = UBound(tempFailFlagOrValueArr) Then
                'FuseCategoryName = CFG_Condition_[47:44]
                'FailFlagOrValue = F_ECPU_CORE[3:0]
                For j = 0 To UBound(tempFuseCategoryNameArr)
                    Sdbl_FuseValue = Harvest_GetAllSiteFlagState(tempFailFlagOrValueArr(j), 1, 0, 1)
'                    Call auto_eFuse_SetWriteVariable_SiteAware(FuseBlockNameArr(i), tempFuseCategoryNameArr(j), Sdbl_FuseValue, True)
                    opbank.SetEfuse field.name, Sdbl_FuseValue, , , , , True
''''''                    For Each site In theexec.sites.Active
''''''                        opbank.SetEfuseBySite field.name, Sdbl_FuseValue
''''''                        opbank.SetEfuseBySitePf field.name, True
''''''                    Next site
                Next j
                
            ElseIf (UBound(tempFuseCategoryNameArr) = 0) And (UBound(tempFailFlagOrValueArr) > 0) Then
                'FuseCategoryName = gfx_fstp_harvest_te_misc
                'FailFlagOrValue = F_GFX_HARV[4:0]
                Sdbl_FuseValue = 0
                For Each site In TheExec.sites.Active
                    For j = 0 To UBound(tempFailFlagOrValueArr)
                        tempSlng = Harvest_GetAllSiteFlagState(tempFailFlagOrValueArr(j), 1, 0, 1)
                        If tempSlng(site) = 1 Then
                            Sdbl_FuseValue = Sdbl_FuseValue.Add(Application.WorksheetFunction.power(2, UBound(tempFailFlagOrValueArr) - j))
                        End If
                    Next j
                Next site
'                Call auto_eFuse_SetWriteVariable_SiteAware(FuseBlockNameArr(i), FuseCategoryNameArr(i), Sdbl_FuseValue, True)
                opbank.SetEfuse field.name, Sdbl_FuseValue, , , , , True
''''''                For Each site In theexec.sites.Active
''''''                    opbank.SetEfuseBySite field.name, Sdbl_FuseValue
''''''                    opbank.SetEfuseBySitePf field.name, True
''''''                Next site

            ElseIf (UBound(tempFuseCategoryNameArr) > 0) And (UBound(tempFailFlagOrValueArr) = 0) Then
                'FuseCategoryName = CFG_Condition_47, CFG_Condition_46, CFG_Condition_45, CFG_Condition_44
                'FailFlagOrValue = b1000
                Sdbl_FuseValue = 0
                If auto_isBinaryString(tempFailFlagOrValueArr(0)) Then
                    tempFailFlagOrValueArr(0) = Replace(UCase(tempFailFlagOrValueArr(0)), "B", "")
                ElseIf auto_isHexString(tempFailFlagOrValueArr(0)) Then
                    tempFailFlagOrValueArr(0) = CStr(auto_HexStr2Value(tempFailFlagOrValueArr(0)))
                End If
                For j = 0 To UBound(tempFuseCategoryNameArr)
                    Sdbl_FuseValue = mid(tempFailFlagOrValueArr(0), j + 1, 1)
'                    Call auto_eFuse_SetWriteVariable_SiteAware(FuseBlockNameArr(i), tempFuseCategoryNameArr(j), Sdbl_FuseValue, True)
                    opbank.SetEfuse field.name, Sdbl_FuseValue, , , , , True
''''''                    For Each site In theexec.sites.Active
''''''                        opbank.SetEfuseBySite field.name, Sdbl_FuseValue
''''''                        opbank.SetEfuseBySitePf field.name, True
''''''                    Next site
                Next j
            Else
                TheExec.Datalog.WriteComment "<Error> Harvest_eFuse_Write: did not support this case yet."
            End If

        Next i
    Else
        TheExec.Datalog.WriteComment "<Error> Harvest_eFuse_Write: input is not match."
    End If

Exit Function

errHandler:
    Call Print_Error_Message(Error_Info, "VBT_LIB_Common_AP", "Harvest_eFuse_Write") 'Add ErrHandler 2023/08/18
    If AbortTest Then Exit Function Else Resume Next
End Function
Public Function Harvest_eFuse_Read(ByVal FuseBlockName As String, ByVal FuseCategoryName As String, ByVal FailFlagOrValue As String) As Long
    On Error GoTo errHandler
    Dim funcName As String:: funcName = "Harvest_eFuse_Read"

''''FuseBlockName = CFG;CFG
''''FuseCategoryName = gfx_fstp_harvest_te_misc;CFG_Condition_[47:44]
''''FailFlagOrValue = F_GFX_HARV[4:0];F_ECPU_CORE[3:0]

    Dim vsite As Variant
    Dim i As Long
    Dim j As Long
    
    Dim FuseBlockNameArr() As String
    Dim FuseCategoryNameArr() As String
    Dim FailFlagOrValueArr() As String
    Dim tempStr As String
    Dim tempFuseCategoryNameArr() As String
    Dim tempFailFlagOrValueArr() As String
    Dim Slng_ReadFuseValue As New SiteLong
    Dim SDbl_ReadFuseValue As New SiteDouble
    Dim Svar_ReadFuseValue As New SiteVariant
    Dim tempSlng As New SiteLong
    
    Dim sInstName As String
    Dim opbank As eFuseBdfBank
    Dim field As eFuseBdfField
    Dim HarvCustomFunc As String
    
    Dim Slng_TempAry() As New SiteLong
    Dim Svar_Temp As New SiteVariant
    glb_Harvest_Fail_Flag_From_eFuseRead_Dict.RemoveAll 'Initial Dictionary for record the Harvest Fail Flag from previous stage
    FuseBlockNameArr = Split(FuseBlockName, ";")
    FuseCategoryNameArr = Split(FuseCategoryName, ";")
    FailFlagOrValueArr = Split(FailFlagOrValue, ";")
    
    sInstName = TheExec.DataManager.instanceName
    TheExec.Datalog.WriteComment "<" & sInstName & ">"
    
    If (UBound(FuseBlockNameArr) = UBound(FuseCategoryNameArr)) And (UBound(FuseBlockNameArr) = UBound(FailFlagOrValueArr)) Then
        For i = 0 To UBound(FuseBlockNameArr)
            tempStr = Harvest_StrExpand(FuseCategoryNameArr(i))
            tempFuseCategoryNameArr = Split(tempStr, ",")
            tempStr = Harvest_StrExpand(FailFlagOrValueArr(i), , , , , HarvCustomFunc)
            tempFailFlagOrValueArr = Split(tempStr, ",")
            Set opbank = GetBdfBank(FuseBlockNameArr(i)) 'Set FuseType
            Set field = opbank.Fields(FuseCategoryNameArr(i)) 'Set FuseName
             
            If (UBound(tempFuseCategoryNameArr) = 0) And (UBound(tempFailFlagOrValueArr) = 0) Then
                If IsNumeric(tempFailFlagOrValueArr(0)) Then
                    TheExec.Datalog.WriteComment "<Error> Harvest_eFuse_Read: do not support this case."
                ElseIf HarvCustomFunc <> "" Then
                    Dim TempFlagArr() As String
                    For Each vsite In TheExec.sites.Selected
'                        Sdbl_ReadFuseValue(vSite) = auto_eFuse_GetWriteDecimal(FuseBlockNameArr(i), FuseCategoryNameArr(i), True)
                        SDbl_ReadFuseValue(vsite) = field.DsscDecValue(vsite)
                    Next vsite
                    tempFailFlagOrValueArr(0) = Harvest_StrExpand(tempFailFlagOrValueArr(0))
                    TempFlagArr = Split(tempFailFlagOrValueArr(0), ",")
                    Call JudgeCustomFun(tempFailFlagOrValueArr(0), HarvCustomFunc, SDbl_ReadFuseValue, Svar_ReadFuseValue)
                    For Each vsite In TheExec.sites.Selected
                        For j = 0 To Len(Svar_ReadFuseValue) - 1
                            tempSlng = mid(Svar_ReadFuseValue, 1 + j, 1)
                            Call Harvest_SetAllSiteFlagState(TempFlagArr(j), tempSlng)
                        Next j
                    Next vsite
                ElseIf UCase(FuseCategoryNameArr(i)) Like UCase("te_misc_evs_cp1") Then
                    For Each vsite In TheExec.sites.Selected
                        SDbl_ReadFuseValue(vsite) = field.DsscDecValue(vsite)
                        TheExec.Datalog.WriteComment "Site" & vsite & " => " & UCase(FuseCategoryNameArr(i)) & " = " & SDbl_ReadFuseValue(vsite)
                        If SDbl_ReadFuseValue(vsite) = 0 Then
                            TheExec.sites.item(vsite).FlagState("F_EVS_Alarm_CP1") = logicClear
                            'TheExec.sites.item(vsite).FlagState("F_EVS_Conduct") = logicFalse
                        ElseIf SDbl_ReadFuseValue(vsite) = 1 Then
                            TheExec.sites.item(vsite).FlagState("F_EVS_Alarm_CP1") = logicFalse
                            'TheExec.sites.item(vsite).FlagState("F_EVS_Conduct") = logicTrue
                        ElseIf SDbl_ReadFuseValue(vsite) = 2 Then
                            TheExec.sites.item(vsite).FlagState("F_EVS_Alarm_CP1") = logicTrue
                            'TheExec.sites.item(vsite).FlagState("F_EVS_Conduct") = logicTrue
                        Else
                            TheExec.sites.item(vsite).FlagState("F_EVS_Alarm_CP1") = logicClear
                            TheExec.Datalog.WriteComment "<Error> Please check te_misc_evs_cp1 Value of." & "Site" & vsite
                        End If
                    Next vsite
                ElseIf UCase(FuseCategoryNameArr(i)) Like UCase("te_misc_evs_cp2") Then
                    For Each vsite In TheExec.sites.Selected
                        SDbl_ReadFuseValue(vsite) = field.DsscDecValue(vsite)
                        TheExec.Datalog.WriteComment "Site" & vsite & " => " & UCase(FuseCategoryNameArr(i)) & " = " & SDbl_ReadFuseValue(vsite)
                        If SDbl_ReadFuseValue(vsite) = 0 Then
                            TheExec.sites.item(vsite).FlagState("F_EVS_Alarm_CP2") = logicClear
                            'TheExec.sites.item(vsite).FlagState("F_EVS_Conduct") = logicFalse
                        ElseIf SDbl_ReadFuseValue(vsite) = 1 Then
                            TheExec.sites.item(vsite).FlagState("F_EVS_Alarm_CP2") = logicFalse
                            'TheExec.sites.item(vsite).FlagState("F_EVS_Conduct") = logicTrue
                        ElseIf SDbl_ReadFuseValue(vsite) = 2 Then
                            TheExec.sites.item(vsite).FlagState("F_EVS_Alarm_CP2") = logicTrue
                            'TheExec.sites.item(vsite).FlagState("F_EVS_Conduct") = logicTrue
                Else
                            TheExec.sites.item(vsite).FlagState("F_EVS_Alarm_CP2") = logicClear
                            TheExec.Datalog.WriteComment "<Error> Please check te_misc_evs_cp2 Value of." & "Site" & vsite
                        End If
                    Next vsite
                Else
                    'FuseCategoryName = CFG_Condition_47
                    'FailFlagOrValue = F_ECPU_CORE3
                    For Each vsite In TheExec.sites.Selected
                        SDbl_ReadFuseValue(vsite) = field.DsscDecValue(vsite)
''''                        Slng_ReadFuseValue(vSite) = auto_eFuse_GetReadDecimal(FuseBlockNameArr(i), FuseCategoryNameArr(i), True)
                    Next vsite
                    If InStr(1, FailFlagOrValueArr(i), "[") <> 0 And InStr(1, FailFlagOrValueArr(i), "]") <> 0 Then
                        FailFlagOrValueArr(i) = tempStr
                    End If
                    
                    Call Harvest_SetAllSiteFlagState(FailFlagOrValueArr(i), SDbl_ReadFuseValue)
                End If
                
            ElseIf UBound(tempFuseCategoryNameArr) = UBound(tempFailFlagOrValueArr) Then
                'FuseCategoryName = CFG_Condition_[47:44]
                'FailFlagOrValue = F_ECPU_CORE[3:0]
                For j = 0 To UBound(tempFuseCategoryNameArr)
                    For Each vsite In TheExec.sites.Selected
                        Slng_ReadFuseValue(vsite) = CDbl(field.DsscDecValue(vsite))
''''                        Slng_ReadFuseValue(vSite) = auto_eFuse_GetReadDecimal(FuseBlockNameArr(i), tempFuseCategoryNameArr(j), True)
                    Next vsite
                    Call Harvest_SetAllSiteFlagState(tempFailFlagOrValueArr(j), Slng_ReadFuseValue)
                Next j
                
            ElseIf (UBound(tempFuseCategoryNameArr) = 0) And (UBound(tempFailFlagOrValueArr) > 0) Then
                'FuseCategoryName = gfx_fstp_harvest_te_misc
                'FailFlagOrValue = F_GFX_HARV[4:0]
                 'FuseCategoryName = gfx_fstp_harvest_te_misc
                'FailFlagOrValue = F_GFX_HARV[4:0]
                For Each vsite In TheExec.sites.Selected
                    If UBound(tempFailFlagOrValueArr) > 32 Then
                        Svar_Temp(vsite) = field.DsscValue(vsite)  'auto_eFuse_GetWriteDecimal(FuseBlockNameArr(i), FuseCategoryNameArr(i), True)
                    Else
                        SDbl_ReadFuseValue(vsite) = field.DsscDecValue(vsite)
                        Svar_Temp(vsite) = GlbUtility.Dec2HexStr(SDbl_ReadFuseValue(vsite))
                    End If
                Next vsite
                Slng_TempAry = GlbUtility.Hex2BinArray(Svar_Temp, UBound(tempFailFlagOrValueArr) + 1)
                
                For j = 0 To UBound(tempFailFlagOrValueArr)
                    Call Harvest_SetAllSiteFlagState(tempFailFlagOrValueArr(j), Slng_TempAry(UBound(tempFailFlagOrValueArr) - j))
                Next j

'                ElseIf (UBound(tempFuseCategoryNameArr) = 0) And (UBound(tempFailFlagOrValueArr) = -1) Then
'                    For Each vSite In TheExec.sites.Selected
'                        SDbl_ReadFuseValue(vSite) = field.DsscDecValue(vSite)
'                    Next vSite
            ElseIf UCase(FuseCategoryNameArr(i)) Like UCase("*bin_fuse*") Then
                For Each vsite In TheExec.sites.Selected
                    SDbl_ReadFuseValue(vsite) = field.DsscDecValue(vsite)
                    TheExec.Datalog.WriteComment "Site" & vsite & " => " & "Bin_Fuse = " & SDbl_ReadFuseValue(vsite)
'''''''                    If SDbl_ReadFuseValue(vsite) = 6 Then
'''''''                        theexec.sites.item(vsite).FlagState("F_IDS_PCPU_HIGH_LEAKAGE") = logicTrue
'''''''                    Else
'''''''                        theexec.sites.item(vsite).FlagState("F_IDS_PCPU_HIGH_LEAKAGE") = logicFalse
'''''''                    End If
                Next vsite
            Else
                TheExec.Datalog.WriteComment "<Error> Harvest_eFuse_Read: did not support this case yet."
            End If
        Next i
    Else
        TheExec.Datalog.WriteComment "<Error> Harvest_eFuse_Read: input is not match."
    End If
    
    
Exit Function

errHandler:
    Call Print_Error_Message(Error_Info, "VBT_LIB_Common_AP", "Harvest_eFuse_Read") 'Add ErrHandler 2023/08/18
    If AbortTest Then Exit Function Else Resume Next
End Function


Public Function Harvest_SUBFLOW()
On Error GoTo errHandler 'Add ErrHandler 2023/08/18
    
    Dim funcName As String:: funcName = "Harvest_SUBFLOW"
    
    
    Dim site As Variant
    
    TheExec.sites.item(1).SiteVariableValue("HARVEST_PASS") = 1
    TheExec.sites.item(2).SiteVariableValue("HARVEST_PASS") = 2
    
    Exit Function

errHandler:
    Call Print_Error_Message(Error_Info, "VBT_LIB_Common_AP", "Harvest_SUBFLOW") 'Add ErrHandler 2023/08/18
    If AbortTest Then Exit Function Else Resume Next
End Function



Public Function Harvest_FailedCoreCount(ByVal BlockType As String, ByVal Harv_GlobalFailFlag As String, _
                                        ByVal Harv_FailCoreSum As String, ByVal Harv_FailCoreSumFlag As String, _
                                        Optional ByVal Harv_AllCorePassFlag As String, Optional ByVal Harv_LocalFailFlag As String) As Long
    On Error GoTo errHandler
    Dim funcName As String:: funcName = "Harvest_FailedCoreCount"
''''BlockType = GFXSA
''''Harv_GlobalFailFlag = F_GFX_HARV[0:4]
''''Harv_FailCoreNum = 0;1
''''Harv_FailCoreSumFlag = F_GFX_Sum_0;F_GFX_Sum_1
''''Harv_AllCorePassFlag = F_GFX_ALLCORE_PASS
''''Harv_LocalFailFlag = F_GFX_SA_HARV[0:4]

    Dim vsite As Variant
    Dim sInstName As String
    Dim i As Long
    
    Dim FailCoreSumArr() As String
    Dim FailCoreSumFlagArr() As String
    Dim lng_MaxFailCoreSum As Long
    Dim FailFlagDefaultName As String
    Dim FailFlagStartIdx As Long
    Dim FailFlagEndIdx As Long
    Dim FailCoreAry_TempBySite() As String
    Dim EachCoreResultAry_TempBySite() As String
    Dim FailCoreCnt_Slng As New SiteLong
    Dim FirstFailCore_Slng As New SiteLong
    
    sInstName = TheExec.DataManager.instanceName
    
    If Harv_FailCoreSum <> "" And Harv_FailCoreSumFlag <> "" Then
        FailCoreSumArr = Split(Harv_FailCoreSum, ";")
        FailCoreSumFlagArr = Split(Harv_FailCoreSumFlag, ";")
        If UBound(FailCoreSumArr) = UBound(FailCoreSumFlagArr) Then
            lng_MaxFailCoreSum = -1
            For i = 0 To UBound(FailCoreSumArr)
                lng_MaxFailCoreSum = Application.WorksheetFunction.Max(lng_MaxFailCoreSum, FailCoreSumArr(i))
            Next i
        Else
            TheExec.Datalog.WriteComment "Warning: Harv_FailCoreSum and Harv_FailCoreSumFlag did not match."
            TheExec.Flow.TestLimit resultVal:=1, lowVal:=0, hiVal:=0, Tname:=sInstName, ForceResults:=tlForceNone
            Exit Function
        End If
    Else
        TheExec.Datalog.WriteComment "Warning: Harv_FailCoreSum or Harv_FailCoreSumFlag did not have any information."
        TheExec.Flow.TestLimit resultVal:=1, lowVal:=0, hiVal:=0, Tname:=sInstName, ForceResults:=tlForceNone
        Exit Function
    End If
    
    If Harv_GlobalFailFlag <> "" Then
        If InStr(Harv_GlobalFailFlag, "[") <> 0 And InStr(Harv_GlobalFailFlag, "]") <> 0 And InStr(Harv_GlobalFailFlag, ":") <> 0 Then
            FailFlagDefaultName = mid(Harv_GlobalFailFlag, 1, InStr(Harv_GlobalFailFlag, "[") - 1)
            FailFlagStartIdx = CLng(mid(Harv_GlobalFailFlag, InStr(Harv_GlobalFailFlag, "[") + 1, InStr(Harv_GlobalFailFlag, ":") - InStr(Harv_GlobalFailFlag, "[") - 1))
            FailFlagEndIdx = CLng(mid(Harv_GlobalFailFlag, InStr(Harv_GlobalFailFlag, ":") + 1, InStr(Harv_GlobalFailFlag, "]") - InStr(Harv_GlobalFailFlag, ":") - 1))
            
            'only consider [0:xxx]
            FailCoreCnt_Slng = 0
            FirstFailCore_Slng = -1
            For Each vsite In TheExec.sites.Selected

                ReDim EachCoreResultAry_TempBySite(FailFlagEndIdx) '(FailFlagEndIdx - FailFlagStartIdx + 1)
                ReDim FailCoreAry_TempBySite(0)
                FailCoreAry_TempBySite(FailFlagStartIdx) = "N/A"
                For i = FailFlagStartIdx To FailFlagEndIdx
                    If TheExec.sites.item(vsite).FlagState(FailFlagDefaultName & CStr(i)) = logicTrue Then
                        FailCoreCnt_Slng = FailCoreCnt_Slng + 1
                        If FirstFailCore_Slng = -1 Then
                            FirstFailCore_Slng = i
                        End If
                        ReDim Preserve FailCoreAry_TempBySite(FailCoreCnt_Slng - 1)
                        FailCoreAry_TempBySite(FailCoreCnt_Slng - 1) = CStr(i)
                        EachCoreResultAry_TempBySite(i) = "F" '(i - FailFlagStartIdx)
                    ElseIf TheExec.sites.item(vsite).FlagState(FailFlagDefaultName & CStr(i)) = logicFalse Then
                        EachCoreResultAry_TempBySite(i) = "P"
                    Else
                        FailCoreCnt_Slng = FailCoreCnt_Slng + 1
                        ReDim Preserve FailCoreAry_TempBySite(FailCoreCnt_Slng - 1)
                        FailCoreAry_TempBySite(FailCoreCnt_Slng - 1) = CStr(i)
                        EachCoreResultAry_TempBySite(i) = "X"
                    End If
                Next i
                
                '----datalog printing----
                'ex: GfxSa_Harvested_Core,Site0,N/A
                TheExec.Datalog.WriteComment BlockType & "_Harvested_Core," + "Site" + CStr(vsite) + "," + Join(FailCoreAry_TempBySite, ",")
                'ex: GfxSa_Harvesting_Result,Site0,p,p,p,p,p
                TheExec.Datalog.WriteComment BlockType & "_Harvesting_Result," + "Site" + CStr(vsite) + "," + Join(EachCoreResultAry_TempBySite, ",")

                TheExec.sites.item(vsite).FlagState(Harv_AllCorePassFlag) = logicFalse
                If FailCoreCnt_Slng > lng_MaxFailCoreSum Then
                    FirstFailCore_Slng = 999
                    For i = 0 To lng_MaxFailCoreSum
                        TheExec.sites.item(vsite).FlagState(FailCoreSumFlagArr(i)) = logicClear
                    Next i
                ElseIf FailCoreCnt_Slng = 0 Then
                    'all core pass
                    TheExec.sites.item(vsite).FlagState(Harv_AllCorePassFlag) = logicTrue
                    TheExec.sites.item(vsite).FlagState(FailCoreSumFlagArr(FailCoreCnt_Slng)) = logicTrue
                Else
                    'FailCoreCnt_Slng = 1 to lng_MaxFailCoreSum
                    For i = 0 To lng_MaxFailCoreSum - 1
                        TheExec.sites.item(vsite).FlagState(FailCoreSumFlagArr(i)) = logicClear
                    Next i
                    TheExec.sites.item(vsite).FlagState(FailCoreSumFlagArr(FailCoreCnt_Slng)) = logicTrue
                End If
                
                'for T-ELLIS correlation
                If FailCoreCnt_Slng = 1 Then
                    TheExec.sites.item(vsite).FlagState("F_Gfx_HARV") = logicTrue
                End If
                
            Next vsite
            
            TheExec.Flow.TestLimit resultVal:=FirstFailCore_Slng, lowVal:=-1, hiVal:=FailFlagEndIdx, Tname:=BlockType + "_Harvested_Core", ForceResults:=tlForceNone
            TheExec.Flow.TestLimit resultVal:=FailCoreCnt_Slng, lowVal:=0, hiVal:=lng_MaxFailCoreSum, Tname:=BlockType + "_Fail_Core_Count", ForceResults:=tlForceNone

        ElseIf InStr(Harv_GlobalFailFlag, ",") <> 0 Then
        End If
    End If
    
Exit Function

errHandler:
    Call Print_Error_Message(Error_Info, "VBT_LIB_Common_AP", "Harvest_FailedCoreCount") 'Add ErrHandler 2023/08/18
    If AbortTest Then Exit Function Else Resume Next
End Function

Public Function Harvest_Summary(ByVal Harv_AllFlag As String, ByVal Harv_GlobalFailFlag As String, _
                                ByVal Harv_FailCoreSumFlag As String, Optional ByVal Harv_AllCorePassFlag As String, _
                                Optional ByVal Harv_BinOutFailFlag As String, _
                                Optional CustHarv_GlobalFailFlag As String, Optional CustHarv_FailCoreSumFlag As String, _
                                Optional CustHarv_BinOutFailFlag As String) As Long
    On Error GoTo errHandler
    Dim funcName As String:: funcName = "Harvest_Summary"

''''Harv_AllFlag = F_SOC_GroupE+F_SOC_GroupI+F_SOC_ALL+F_ECPU_CORE3+F_ECPU_CORE2+F_ECPU_CORE1+F_ECPU_CORE0+F_ECPU_HARV+F_ECPU_ALL+F_ECPU_CORE_SUM_0+F_ECPU_CORE_SUM_1+F_Gfx_HARV4+F_Gfx_HARV3+F_Gfx_HARV2+F_Gfx_HARV1+F_Gfx_HARV0+F_Gfx_HARV+F_GFX_HAV_SUM_0+F_GFX_HAV_SUM_1+F_Bin3_Pass+F_Bin4X_Pass
''''Harv_GlobalFailFlag = F_GFX_HARV[0:4]|F_ECPU_CORE[0:3] / F_ECPU_CORE0+F_ECPU_CORE1+F_ECPU_CORE2+F_ECPU_CORE3|F_GFX_HARV0+F_GFX_HARV1+F_GFX_HARV2+F_GFX_HARV3+F_GFX_HARV4
''''Harv_FailCoreSumFlag = F_GFX_HARV_Sum_0;F_GFX_HARV_Sum_1|F_ECPU_CORE_Sum_0;F_ECPU_CORE_Sum_1 / F_GFX_HARV_Sum_0+F_GFX_HARV_Sum_1|F_ECPU_CORE_Sum_0+F_ECPU_CORE_Sum_1
''''Harv_AllCorePassFlag = F_GFX_ALL_CORE_PASS|F_ECPU_ALL_CORE_PASS
''''Harv_BinOutFailFlag = F_GFX_HARV_SUM|F_ECPU_CORE_SUM|F_Bin3_Pass|F_Bin4X_Pass
    Dim vsite As Variant
    Dim sInstName As String
    Dim i As Long
''''Harv_GlobalFailFlag = "F_GFX_HARV0+F_GFX_HARV1+F_GFX_HARV2+F_GFX_HARV3+F_GFX_HARV4|F_ECPU_CORE0+F_ECPU_CORE1+F_ECPU_CORE2+F_ECPU_CORE3"
    Dim FailCoreSumArr() As String
    Dim FailCoreSumFlagArr() As String
    Dim FailCoreSumFlagArr_Split() As String
    Dim lng_MaxFailCoreSum As Long
    Dim FailFlagDefaultName As String
    Dim FailFlagStartIdx As Long
    Dim FailFlagEndIdx As Long
    Dim FailCoreAry_TempBySite() As String
    Dim EachCoreResultAry_TempBySite() As String
    Dim FailCoreCnt_Slng As New SiteLong
    Dim FirstFailCore_Slng As New SiteLong
    
    
''''New format for harv range
    Dim TempStoreSumRange() As String
    Dim TempBoolSumRange As Boolean: TempBoolSumRange = False
    Dim TempRangeLB As Long: TempRangeLB = 0
    Dim TempRangeUB As Long: TempRangeUB = 0
    
    
''    If UCase(theexec.DataManager.instanceName) = UCase("Harvest_Summary_After_Harvesting_Descision") Then Stop
    With TheExec.Datalog
        .Setup.Shared.ascii.Columns.EnableCustomWidths = True
        .Setup.Shared.ascii.Columns.Parametric.TestName.Width = 75
        .Setup.Shared.ascii.Columns.Parametric.measured.Width = 75
'        .Setup.Shared.ascii.Columns.Functional.TestName.Width = 75
'        .Setup.Shared.ascii.Columns.Functional.Pattern.Width = 75
        .Setup.Shared.ascii.Columns.Parametric.measured.Width = 10
        .ApplySetup
    End With

    sInstName = TheExec.DataManager.instanceName
    
    Dim Block_Idx As Long
''    Dim Block_FailCoreCnt_Slng() As Long
    Dim Block_BinOutFailFlagArr() As String
    Dim Block_GlobalFailFlagArr() As String
    Dim Block_FailCoreSumFlagArr() As String
    Dim Block_AllCorePassFlagArr() As String
    Dim Block_GlobalFailFlagArr_Split() As String
    
    Dim Block_FailCoreCnt_Slng() As New SiteLong
    
    Block_AllCorePassFlagArr = Split(Harv_AllCorePassFlag, "|")
    If Harv_FailCoreSumFlag <> "" And Harv_GlobalFailFlag <> "" And Harv_BinOutFailFlag <> "" Then
        Block_BinOutFailFlagArr = Split(Harv_BinOutFailFlag, "|")
        Block_GlobalFailFlagArr = Split(Harv_GlobalFailFlag, "|")
        Block_FailCoreSumFlagArr = Split(Harv_FailCoreSumFlag, "|")
    Else
        TheExec.Datalog.WriteComment "Warning: Harv_FailCoreSumFlag/Harv_GlobalFailFlag/Harv_BinOutFailFlag did not have any information."
        TheExec.Flow.TestLimit resultVal:=1, lowVal:=0, hiVal:=0, Tname:=sInstName, ForceResults:=tlForceNone
        Exit Function
    End If
    
    ReDim Block_FailCoreCnt_Slng(UBound(Block_BinOutFailFlagArr))
    
    Call initHarvSumFlag(Harv_FailCoreSumFlag, Harv_AllCorePassFlag, Harv_BinOutFailFlag, CustHarv_FailCoreSumFlag, CustHarv_BinOutFailFlag)
    
    For Block_Idx = 0 To UBound(Block_GlobalFailFlagArr)
    
        FailCoreSumFlagArr = Split(Block_FailCoreSumFlagArr(Block_Idx), "+") ''FailCoreSumFlagArr = Split(Block_FailCoreSumFlagArr(Block_Idx), ";")
        ReDim FailCoreSumArr(UBound(FailCoreSumFlagArr))
        For i = 0 To UBound(FailCoreSumFlagArr)
            If InStr(FailCoreSumFlagArr(i), "_") <> 0 Then
                FailCoreSumFlagArr_Split = Split(FailCoreSumFlagArr(i), "_")
                FailCoreSumArr(i) = FailCoreSumFlagArr_Split(UBound(FailCoreSumFlagArr_Split))
            End If
        Next i
        
        lng_MaxFailCoreSum = -1
        For i = 0 To UBound(FailCoreSumArr)
            If InStr(1, UCase(FailCoreSumArr(i)), UCase("to")) <> 0 Then
                TempStoreSumRange = Split(UCase(FailCoreSumArr(i)), "TO")
'''                If TempStoreSumRange(0) <> 0 Then
'''                    TempRangeLB = TempStoreSumRange(0)
'''                End If
                lng_MaxFailCoreSum = Application.WorksheetFunction.Max(lng_MaxFailCoreSum, TempStoreSumRange(UBound(TempStoreSumRange)))
'''                If lng_MaxFailCoreSum > TempRangeUB Then
'''                    TempRangeUB = lng_MaxFailCoreSum
'''                End If
                TempBoolSumRange = True
            Else
                lng_MaxFailCoreSum = Application.WorksheetFunction.Max(lng_MaxFailCoreSum, FailCoreSumArr(i))
            End If
        Next i
        
        
        If InStr(Block_GlobalFailFlagArr(Block_Idx), "[") <> 0 And InStr(Block_GlobalFailFlagArr(Block_Idx), "]") <> 0 And InStr(Block_GlobalFailFlagArr(Block_Idx), ":") <> 0 Then
            FailFlagDefaultName = mid(Block_GlobalFailFlagArr(Block_Idx), 1, InStr(Block_GlobalFailFlagArr(Block_Idx), "[") - 1)
            FailFlagStartIdx = CLng(mid(Block_GlobalFailFlagArr(Block_Idx), InStr(Block_GlobalFailFlagArr(Block_Idx), "[") + 1, InStr(Block_GlobalFailFlagArr(Block_Idx), ":") - InStr(Block_GlobalFailFlagArr(Block_Idx), "[") - 1))
            FailFlagEndIdx = CLng(mid(Block_GlobalFailFlagArr(Block_Idx), InStr(Block_GlobalFailFlagArr(Block_Idx), ":") + 1, InStr(Block_GlobalFailFlagArr(Block_Idx), "]") - InStr(Block_GlobalFailFlagArr(Block_Idx), ":") - 1))
            
            'only consider [0:xxx]
            FailCoreCnt_Slng = 0
            FirstFailCore_Slng = -1
            
            For Each vsite In TheExec.sites.Selected
                Block_FailCoreCnt_Slng(Block_Idx)(vsite) = 0
                ReDim EachCoreResultAry_TempBySite(FailFlagEndIdx) '(FailFlagEndIdx - FailFlagStartIdx + 1)
                ReDim FailCoreAry_TempBySite(0)
                FailCoreAry_TempBySite(FailFlagStartIdx) = "N/A"
                For i = FailFlagStartIdx To FailFlagEndIdx
                    If TheExec.sites.item(vsite).FlagState(FailFlagDefaultName & CStr(i)) = logicTrue Then
                        FailCoreCnt_Slng = FailCoreCnt_Slng + 1
                        Block_FailCoreCnt_Slng(Block_Idx)(vsite) = Block_FailCoreCnt_Slng(Block_Idx)(vsite) + 1
                        If FirstFailCore_Slng = -1 Then
                            FirstFailCore_Slng = i
                        End If
                        ReDim Preserve FailCoreAry_TempBySite(FailCoreCnt_Slng - 1)
                        FailCoreAry_TempBySite(FailCoreCnt_Slng - 1) = CStr(i)
                        EachCoreResultAry_TempBySite(i) = "F" '(i - FailFlagStartIdx)
                    ElseIf TheExec.sites.item(vsite).FlagState(FailFlagDefaultName & CStr(i)) = logicFalse Then
                        EachCoreResultAry_TempBySite(i) = "P"
                    Else
                        FailCoreCnt_Slng = FailCoreCnt_Slng + 1
                        Block_FailCoreCnt_Slng(Block_Idx)(vsite) = Block_FailCoreCnt_Slng(Block_Idx)(vsite) + 1
                        ReDim Preserve FailCoreAry_TempBySite(FailCoreCnt_Slng - 1)
                        FailCoreAry_TempBySite(FailCoreCnt_Slng - 1) = CStr(i)
                        EachCoreResultAry_TempBySite(i) = "X"
                    End If
                Next i

                If FailCoreCnt_Slng > lng_MaxFailCoreSum Then
                    FirstFailCore_Slng = 999
                    For i = 0 To lng_MaxFailCoreSum - 1
                        TheExec.sites.item(vsite).FlagState(FailCoreSumFlagArr(i)) = logicClear
                    Next i
                    TheExec.sites.item(vsite).FlagState(Block_BinOutFailFlagArr(Block_Idx)) = logicTrue
                    TheExec.sites.item(vsite).FlagState(Block_AllCorePassFlagArr(Block_Idx)) = logicFalse
                ElseIf FailCoreCnt_Slng = 0 Then
                    'all core pass
                    TheExec.sites.item(vsite).FlagState(Block_BinOutFailFlagArr(Block_Idx)) = logicFalse

                    TheExec.sites.item(vsite).FlagState(Block_AllCorePassFlagArr(Block_Idx)) = logicTrue
                    TheExec.sites.item(vsite).FlagState(FailCoreSumFlagArr(FailCoreCnt_Slng)) = logicTrue
                Else
                    'FailCoreCnt_Slng = 1 to lng_MaxFailCoreSum
                    For i = 0 To lng_MaxFailCoreSum - 1
                        TheExec.sites.item(vsite).FlagState(FailCoreSumFlagArr(i)) = logicClear
                    Next i
                    TheExec.sites.item(vsite).FlagState(FailCoreSumFlagArr(FailCoreCnt_Slng)) = logicTrue
                    TheExec.sites.item(vsite).FlagState(Block_BinOutFailFlagArr(Block_Idx)) = logicFalse
                    TheExec.sites.item(vsite).FlagState(Block_AllCorePassFlagArr(Block_Idx)) = logicFalse
                End If
               
            Next vsite

        ElseIf InStr(Block_GlobalFailFlagArr(Block_Idx), "+") <> 0 Then
            
            Block_GlobalFailFlagArr_Split = Split(Block_GlobalFailFlagArr(Block_Idx), "+") ''F_ECPU_CORE0+F_ECPU_CORE1+F_ECPU_CORE2+F_ECPU_CORE3
            
            FailCoreCnt_Slng = 0
            For Each vsite In TheExec.sites.Selected
                Block_FailCoreCnt_Slng(Block_Idx)(vsite) = 0
                For i = 0 To UBound(Block_GlobalFailFlagArr_Split)
                    If TheExec.sites.item(vsite).FlagState(Block_GlobalFailFlagArr_Split(i)) = logicTrue Then
                        FailCoreCnt_Slng = FailCoreCnt_Slng + 1
                        Block_FailCoreCnt_Slng(Block_Idx)(vsite) = Block_FailCoreCnt_Slng(Block_Idx)(vsite) + 1
                        If FirstFailCore_Slng = -1 Then
                            FirstFailCore_Slng = i
                        End If
                    ElseIf TheExec.sites.item(vsite).FlagState(Block_GlobalFailFlagArr_Split(i)) = logicFalse Then
                    Else
                        FailCoreCnt_Slng = FailCoreCnt_Slng + 1
                        Block_FailCoreCnt_Slng(Block_Idx)(vsite) = Block_FailCoreCnt_Slng(Block_Idx)(vsite) + 1
                    End If
                Next i
                
                If FailCoreCnt_Slng > lng_MaxFailCoreSum Then
                    FirstFailCore_Slng = 999
                    Dim TempFailCoreSum As Long
                    If TempBoolSumRange = True Then
                        TempFailCoreSum = UBound(FailCoreSumFlagArr) + 1
                    Else
                        TempFailCoreSum = lng_MaxFailCoreSum
                    End If
                    For i = 0 To TempFailCoreSum - 1
                        TheExec.sites.item(vsite).FlagState(FailCoreSumFlagArr(i)) = logicClear
                    Next i
                    TheExec.sites.item(vsite).FlagState(Block_BinOutFailFlagArr(Block_Idx)) = logicTrue
                    TheExec.sites.item(vsite).FlagState(Block_AllCorePassFlagArr(Block_Idx)) = logicFalse

                ElseIf FailCoreCnt_Slng = 0 Then
                'all core pass
                    TheExec.sites.item(vsite).FlagState(Block_BinOutFailFlagArr(Block_Idx)) = logicFalse
                    TheExec.sites.item(vsite).FlagState(Block_AllCorePassFlagArr(Block_Idx)) = logicTrue
                    If TempBoolSumRange = True Then
                        For i = 0 To UBound(FailCoreSumArr)
                            If InStr(1, UCase(FailCoreSumArr(i)), UCase("0to")) <> 0 Then
                                TheExec.sites.item(vsite).FlagState(FailCoreSumFlagArr(i)) = logicTrue
                            End If
                        Next i
'''''                        theexec.sites.item(vSite).FlagState(Block_AllCorePassFlagArr(Block_Idx)) = logicTrue
                    Else
                        TheExec.sites.item(vsite).FlagState(FailCoreSumFlagArr(FailCoreCnt_Slng)) = logicTrue
                    End If
                ElseIf TempBoolSumRange = True Then
                    For i = 0 To UBound(FailCoreSumArr)
                        If InStr(1, UCase(FailCoreSumArr(i)), UCase("to")) <> 0 Then
                            TempStoreSumRange = Split(UCase(FailCoreSumArr(i)), UCase("to"))
                            TempRangeLB = TempStoreSumRange(0)
                            TempRangeUB = TempStoreSumRange(1)
                            If TempRangeLB <= FailCoreCnt_Slng And FailCoreCnt_Slng <= TempRangeUB Then
                                TheExec.sites.item(vsite).FlagState(FailCoreSumFlagArr(i)) = logicTrue
                            End If
                        ElseIf FailCoreCnt_Slng = CInt(FailCoreSumArr(i)) Then
                            TheExec.sites.item(vsite).FlagState(FailCoreSumFlagArr(i)) = logicTrue
                        End If
                    Next i

                    TheExec.sites.item(vsite).FlagState(Block_BinOutFailFlagArr(Block_Idx)) = logicFalse
                    TheExec.sites.item(vsite).FlagState(Block_AllCorePassFlagArr(Block_Idx)) = logicFalse
                    
                Else
                    'FailCoreCnt_Slng = 1 to lng_MaxFailCoreSum
                    For i = 0 To lng_MaxFailCoreSum - 1
                        TheExec.sites.item(vsite).FlagState(FailCoreSumFlagArr(i)) = logicClear
                    Next i
                    TheExec.sites.item(vsite).FlagState(FailCoreSumFlagArr(FailCoreCnt_Slng)) = logicTrue
                    TheExec.sites.item(vsite).FlagState(Block_BinOutFailFlagArr(Block_Idx)) = logicFalse
                    TheExec.sites.item(vsite).FlagState(Block_AllCorePassFlagArr(Block_Idx)) = logicFalse

                End If
            Next vsite
        End If
    Next Block_Idx
    
    ''''' New request for special case ''''''
    
    If CustHarv_GlobalFailFlag <> "" And CustHarv_FailCoreSumFlag <> "" And CustHarv_BinOutFailFlag <> "" Then
        Call Harvest_CustomFlagJudge(CustHarv_GlobalFailFlag, CustHarv_FailCoreSumFlag, CustHarv_BinOutFailFlag, lng_MaxFailCoreSum)
    Else
        TheExec.Datalog.WriteComment "Custom Context Mismatch, Please Check "
    End If
    
    '''''''''''''''''''''''''''''''''''''''''
    
    Dim FlagCnt As Long
    Dim lTestResultFail As Long
    
    Dim HarvestFlagArr() As String
    Dim TempHarvestStr() As String
    
    Dim Tname_str As String
    Dim HarvestFlagArr_Num As String
    
    Dim sLSumResult  As New SiteLong

    
    
    HarvestFlagArr = Split(Harv_AllFlag, "+")
    HarvestFlagArr_Num = UBound(HarvestFlagArr)
    ReDim TempHarvestStr(HarvestFlagArr_Num)
    lTestResultFail = 2 ''Fail value
    For Each vsite In TheExec.sites
        If Harv_BinOutFailFlag <> "" Then
            For FlagCnt = 0 To HarvestFlagArr_Num
                Tname_str = HarvestFlagArr(FlagCnt)
                ''''Harv_BinOutFailFlag = F_GFX_HARV_SUM|F_ECPU_CORE_SUM
                If InStr(Harv_BinOutFailFlag, HarvestFlagArr(FlagCnt)) <> 0 And UCase(HarvestFlagArr(FlagCnt)) <> "F_SOC_BV_AN" Then
                    For i = 0 To UBound(Block_BinOutFailFlagArr)
                        If UCase(HarvestFlagArr(FlagCnt)) = UCase(Block_BinOutFailFlagArr(i)) Then
                            If TheExec.sites.item(vsite).FlagState(HarvestFlagArr(FlagCnt)) = logicTrue Then
                                TempHarvestStr(FlagCnt) = 1
                            Else
                                TempHarvestStr(FlagCnt) = 0
                            End If
'                                If InStr(HarvestFlagArr(FlagCnt), "SUM_0") = 0 Then
'                                    TheExec.Flow.TestLimit resultVal:=Block_FailCoreCnt_Slng(i)(vsite), lowVal:=0, hiVal:=0, Tname:=Tname_str, ForceResults:=tlForceNone
'                                Else
'                                    TheExec.Flow.TestLimit resultVal:=Block_FailCoreCnt_Slng(i)(vsite), lowVal:=1, hiVal:=1, Tname:=Tname_str, ForceResults:=tlForceNone
'                                End If
    ''                            TheExec.Flow.TestLimit resultVal:=TempHarvestStr(FlagCnt), lowVal:=0, hiVal:=0, Tname:=Tname_Str, ForceResults:=tlForceNone
                        End If
                    Next i
                    TheExec.Flow.TestLimit resultVal:=TempHarvestStr(FlagCnt), lowVal:=0, hiVal:=0, Tname:=Tname_str, ForceResults:=tlForceNone
                ''''Harv_AllCorePassFlag = F_GFX_ALL_CORE_PASS|F_ECPU_ALL_CORE_PASS
                ElseIf InStr(Harv_AllCorePassFlag, HarvestFlagArr(FlagCnt)) <> 0 And UCase(HarvestFlagArr(FlagCnt)) <> "F_SOC_BV_AN" Then
                    For i = 0 To UBound(Block_AllCorePassFlagArr)
                        If UCase(HarvestFlagArr(FlagCnt)) = UCase(Block_AllCorePassFlagArr(i)) Then
''                            If UCase(Block_AllCorePassFlagArr(i)) = "F_ECPU_ALL_CORE_PASS" Then Stop
                            If TheExec.sites.item(vsite).FlagState(HarvestFlagArr(FlagCnt)) = logicTrue Then
                                TempHarvestStr(FlagCnt) = 1
                            Else
                                TempHarvestStr(FlagCnt) = 0
                            End If
''                            TheExec.Flow.TestLimit resultVal:=TempHarvestStr(FlagCnt), Tname:=Tname_Str, ForceResults:=tlForceNone
                            If InStr(HarvestFlagArr(FlagCnt), "SUM_0") = 0 Then
                                TheExec.Flow.TestLimit resultVal:=TempHarvestStr(FlagCnt), lowVal:=0, hiVal:=0, Tname:=Tname_str, ForceResults:=tlForceNone
                            Else '230727 QJ: set high limit to 1 if contain "SUM_0". Bin1 will be no (F)
                                TheExec.Flow.TestLimit resultVal:=TempHarvestStr(FlagCnt), lowVal:=1, hiVal:=1, Tname:=Tname_str, ForceResults:=tlForceNone
                            End If
                        End If
                    Next i
                Else
                    If TheExec.sites.item(vsite).FlagState(HarvestFlagArr(FlagCnt)) = logicTrue Then
                        TempHarvestStr(FlagCnt) = 1
                    Else
                        TempHarvestStr(FlagCnt) = 0
                    End If
''                    TheExec.Flow.TestLimit resultVal:=TempHarvestStr(FlagCnt), Tname:=Tname_Str, ForceResults:=tlForceNone
                    '230727 QJ: set high limit to 1 if contain "SUM_0", F_PWRBIN_LOW, Bin1. Bin1 die will be no (F)
                    If InStr(HarvestFlagArr(FlagCnt), "SUM_0") > 0 Or HarvestFlagArr(FlagCnt) Like "F_EVS_Conduct" Then
                        TheExec.Flow.TestLimit resultVal:=TempHarvestStr(FlagCnt), lowVal:=1, hiVal:=1, Tname:=Tname_str, ForceResults:=tlForceNone
                    ElseIf HarvestFlagArr(FlagCnt) Like "F_PWRBIN_LOW" Then
                        TheExec.Flow.TestLimit resultVal:=TempHarvestStr(FlagCnt), lowVal:=0, hiVal:=1, Tname:=Tname_str, ForceResults:=tlForceNone
                    ElseIf HarvestFlagArr(FlagCnt) Like "Bin01" Or HarvestFlagArr(FlagCnt) Like "F_Bin01_LowPower" Then
                        TheExec.Flow.TestLimit resultVal:=TempHarvestStr(FlagCnt), lowVal:=0, hiVal:=1, Tname:=Tname_str, ForceResults:=tlForceNone '230727 QJ: allow flag Bin1 to be 0 and 1. To avoid (f) at the cases of others good bin
                    Else
                        TheExec.Flow.TestLimit resultVal:=TempHarvestStr(FlagCnt), lowVal:=0, hiVal:=0, Tname:=Tname_str, ForceResults:=tlForceNone
                    End If
                End If
            Next FlagCnt
        End If
        
    Next vsite
'
'    With TheExec.Datalog
'        .Setup.Shared.ascii.Columns.EnableCustomWidths = True
'        .Setup.Shared.ascii.Columns.Parametric.TestName.Width = 150
'        .Setup.Shared.ascii.Columns.Parametric.measured.Width = 16
'        .Setup.Shared.ascii.Columns.Functional.TestName.Width = 150
'        .Setup.Shared.ascii.Columns.Functional.Pattern.Width = 100
'        .ApplySetup
'    End With

Exit Function

errHandler:
    Call Print_Error_Message(Error_Info, "VBT_LIB_Common_AP", "Harvest_Summary") 'Add ErrHandler 2023/08/18
    If AbortTest Then Exit Function Else Resume Next
End Function
Public Function Harvest_ReadValue(ByVal PatternName As String, ByVal HarvSrcKey As String) As Long
    On Error GoTo errHandler
    Dim funcName As String:: funcName = "Harvest_eFuse_Read"

''''PatternName = PatA;PatB
''''HarvSrcKey = HarvSrcKeyA;HarvSrcKeyB

    Dim vsite As Variant
    Dim i As Long
    Dim j As Long
    
    Dim PatternNameArr() As String
    Dim FuseBlockNameArr() As String
    Dim FuseCategoryNameArr() As String
    Dim FailFlagOrValueArr() As String
    Dim HarvSrcKeyArr() As String
    Dim tempStr As String
    Dim tempFuseCategoryNameArr() As String
    Dim tempFailFlagOrValueArr() As String
    Dim Slng_ReadFuseValue As New SiteLong
    Dim tempSlng As New SiteLong
    
    Dim sInstName As String
    
    PatternNameArr = Split(PatternName, ";")
    HarvSrcKeyArr = Split(HarvSrcKey, ";")
    
    sInstName = TheExec.DataManager.instanceName
    TheExec.Datalog.WriteComment "<" & sInstName & ">"
    
    
    Dim temp_Harv_DSSC_DSP As New DSPWave
    If UBound(PatternNameArr) = UBound(HarvSrcKeyArr) Then
        tempStr = vbNullString
        For i = 0 To UBound(PatternNameArr)
''            Call Harvest_DigSrc(PatternNameArr(i), HarvSrcKeyArr(i), tempStr, temp_Harv_DSSC_DSP)
''            Call Harvest_CreateDigSrc(PatternNameArr(i), HarvSrcKeyArr(i))
            
            Dim tempHarvDsp As New DSPWave
            Dim tempHarvDsp_SampleSize As Long
            Call Harvest_CreateDigSrc(PatternNameArr(i), HarvSrcKeyArr(i), , tempHarvDsp, tempHarvDsp_SampleSize)

        Next i
    Else
        TheExec.Datalog.WriteComment "<Error> Harvest_eFuse_Read: input is not match."
    End If
    
    
Exit Function

errHandler:
    Call Print_Error_Message(Error_Info, "VBT_LIB_Common_AP", "Harvest_ReadValue") 'Add ErrHandler 2023/08/18
    If AbortTest Then Exit Function Else Resume Next
End Function


Public Function Harvest_Print_Flag_Status(HarvFlag As String, Optional HarvMaxCore As Integer = -1) As Long
    On Error GoTo errHandler
    Dim funcName As String:: funcName = "Harvest_Print_Flag_Status"
    
    Dim vsite As Variant
    Dim HarvFlagArr() As String
    Dim FlagCnt As Integer
    Dim TempHarvestStr As Integer
    Dim sInstName As String
    Dim Tname_str As String
    Dim FlagTrueCnt As New SiteLong
    

    TheExec.Datalog.Setup.Shared.ascii.Columns.Parametric.TestName.Width = 75
    TheExec.Datalog.ApplySetup
 
    HarvFlagArr = Split(HarvFlag, ",")
    
    sInstName = TheExec.DataManager.instanceName
    'TheExec.Datalog.WriteComment "<" & sInstName & ">"
    
    For Each vsite In TheExec.sites
        FlagTrueCnt(vsite) = 0
        For FlagCnt = 0 To UBound(HarvFlagArr)
            Tname_str = HarvFlagArr(FlagCnt)
            If TheExec.sites.item(vsite).FlagState(HarvFlagArr(FlagCnt)) = logicTrue Then
                TempHarvestStr = 1
                FlagTrueCnt(vsite) = FlagTrueCnt(vsite) + 1
            ElseIf TheExec.sites.item(vsite).FlagState(HarvFlagArr(FlagCnt)) = logicFalse Then
                TempHarvestStr = 0
            Else
                TempHarvestStr = -1 'HarvFlagArr(FlagCnt) is Clear or doesn't exist.
            End If
            TheExec.Flow.TestLimit resultVal:=TempHarvestStr, lowVal:=-1, hiVal:=1, Tname:=Tname_str
            
        Next FlagCnt
    Next vsite
    If HarvMaxCore > 0 Then
        For Each vsite In TheExec.sites
            TheExec.Flow.TestLimit resultVal:=FlagTrueCnt(vsite), lowVal:=0, hiVal:=HarvMaxCore, Tname:=sInstName & "_Core_Over_Check"
        Next vsite
    End If
Exit Function

errHandler:
    TheExec.Datalog.WriteComment "<Error> " + funcName + ":: please check it out."
    If AbortTest Then Exit Function Else Resume Next
End Function

'Public Function Harvest_Simulation(HarvFlag As String)
'    Dim HarvFlagArr() As String
'    Dim FlagCount As Integer
'    Dim Vsite As Variant
'    Dim tmpScenario As Long
'    Dim lastDigit As Integer
'
'    HarvFlagArr = Split(HarvFlag, ",")
'    FlagCount = UBound(HarvFlagArr) + 1
'    For Each Vsite In TheExec.sites
'        TheExec.Datalog.WriteComment "HarvScenario: " & HarvScenario
'        tmpScenario = HarvScenario
'        While tmpScenario > 0
'            lastDigit = tmpScenario Mod (FlagCount + 1)
'            If lastDigit > 0 Then
'                TheExec.sites.item(Vsite).FlagState(HarvFlagArr(lastDigit - 1)) = logicTrue
'            End If
'            tmpScenario = tmpScenario \ (FlagCount + 1)
'        Wend
'        HarvScenario = HarvScenario + 1
'    Next Vsite
'End Function
Public Function Set_Flag_By_Sheet(sheetName As String)
    On Error GoTo errHandler
    Dim funcName As String:: funcName = "Set_Flag_By_Sheet"
    'setting flag clear/ false/ true if (X,Y) and condition are match
    
    Dim ScenarioTable As Worksheet
    Dim vsite As Variant
    Dim MaxRow As Long
    Dim rowCount As Long
    Dim colXY, colCondition, colRepeat, colFlagClear, colFlagFalse, colFlagTrue, colNote As Integer
    Dim targetX, targetY As Integer
    Dim ForceSetFlagFinishCasesArr() As String
    Dim tmpCaseStr As String
    Dim conditionArr() As String
    Dim tmpFlagArr() As String
    Dim i As Integer
    Dim siteDone As Boolean
    
    Dim HarvFlagArr() As String
    Dim FlagCount As Integer
    Dim tmpScenario As Long
    Dim lastDigit As Integer
    Set ScenarioTable = Sheets(sheetName)
    MaxRow = ScenarioTable.UsedRange.Rows.Count
    
    'colnum number for each header
    colXY = 1
    colCondition = 2
    colRepeat = 3
    colFlagClear = 4
    colFlagFalse = 5
    colFlagTrue = 6
    colNote = 7
    
    'Check Header. The input sheet need to have below header.
    'X,Y    condition   repeat   flag-clear  flag-false  flag-true   note
    If UCase(ScenarioTable.Cells(1, colXY).value) Like UCase("X,Y") And _
    UCase(ScenarioTable.Cells(1, colCondition).value) Like UCase("condition") And _
    UCase(ScenarioTable.Cells(1, colRepeat).value) Like UCase("repeat*") And _
    UCase(ScenarioTable.Cells(1, colFlagClear).value) Like UCase("flag-clear") And _
    UCase(ScenarioTable.Cells(1, colFlagFalse).value) Like UCase("flag-false") And _
    UCase(ScenarioTable.Cells(1, colFlagTrue).value) Like UCase("flag-true") And _
    UCase(ScenarioTable.Cells(1, colNote).value) Like UCase("note") Then
    'Header correct. Continue
    Else
        TheExec.Datalog.WriteComment "The header for the sheet" & sheetName & "is incorrect. Should be: X,Y condition   repeat   flag-clear  flag-false  flag-true   note"
        GoTo errHandler
        Exit Function
    End If
    
    
    For Each vsite In TheExec.sites
        siteDone = False 'the flag to record if this site match any case
        rowCount = 2
        While Not ScenarioTable.Cells(rowCount, colXY) Like "end" And rowCount <= MaxRow And siteDone = False
            'get XY from the sheet
            targetX = -999
            targetY = -999
            If ScenarioTable.Cells(rowCount, colXY) <> "" Then
                targetX = CInt(Split(ScenarioTable.Cells(rowCount, colXY), ",")(0))
                targetY = CInt(Split(ScenarioTable.Cells(rowCount, colXY), ",")(1))
            End If
            'get condition from the sheet
            Erase conditionArr
            If ScenarioTable.Cells(rowCount, colCondition) <> "" Then
                conditionArr = Split(ScenarioTable.Cells(rowCount, colCondition), ",")
            End If
            'Check this row has been executed or not
            'ForecSetFlagFinishCases = "SheetName1:ROW1;SheetName1:ROW2;SheetName2:ROW1;SheetName1:ROW3"
            ForceSetFlagFinishCasesArr = Split(ForceSetFlagFinishCases, ";") 'the global variable that record executed cases.
            tmpCaseStr = sheetName & ":" & rowCount 'the string to represent this case.
            If Not IsExistInArr(ForceSetFlagFinishCasesArr, tmpCaseStr) Then
                'check XY
                If ScenarioTable.Cells(rowCount, colXY) = "" Or (targetX = XCoord And targetY = YCoord) Then
                    'check condition
                    If ScenarioTable.Cells(rowCount, colCondition) = "" Or CheckConditionBySite(conditionArr, vsite) Then
                        'start to set flag
                        TheExec.Datalog.WriteComment "****************************Condition Match. Start to set flag****************************"
                        TheExec.Datalog.WriteComment "Force Set " & lotId & "-" & WaferID & "_S" & vsite & "X" & XCoord & "Y" & YCoord & " as " & tmpCaseStr & "    Note: " & ScenarioTable.Cells(rowCount, colNote)
                        'Flag-clear
                        tmpFlagArr = Split(ScenarioTable.Cells(rowCount, colFlagClear), ",")
                        For i = 0 To UBound(tmpFlagArr)
                            TheExec.Datalog.WriteComment tmpFlagArr(i) & ": " & FlagStatusNum2Str(TheExec.sites.item(vsite).FlagState(tmpFlagArr(i))) & " -> Clear"
                            TheExec.sites.item(vsite).FlagState(tmpFlagArr(i)) = logicClear
                        Next i
                        'Flag-false
                        tmpFlagArr = Split(ScenarioTable.Cells(rowCount, colFlagFalse), ",")
                        For i = 0 To UBound(tmpFlagArr)
                            TheExec.Datalog.WriteComment tmpFlagArr(i) & ": " & FlagStatusNum2Str(TheExec.sites.item(vsite).FlagState(tmpFlagArr(i))) & " -> False"
                            TheExec.sites.item(vsite).FlagState(tmpFlagArr(i)) = logicFalse
                        Next i
                        'Flag-true
                        tmpFlagArr = Split(ScenarioTable.Cells(rowCount, colFlagTrue), ",")
                        For i = 0 To UBound(tmpFlagArr)
                            TheExec.Datalog.WriteComment tmpFlagArr(i) & ": " & FlagStatusNum2Str(TheExec.sites.item(vsite).FlagState(tmpFlagArr(i))) & " -> True"
                            TheExec.sites.item(vsite).FlagState(tmpFlagArr(i)) = logicTrue
                        Next i
                        'if REPEAT is not "YES", record this case in global variable
                        If Not UCase(ScenarioTable.Cells(rowCount, colRepeat).value) Like "YES" Then
                            ForceSetFlagFinishCases = ForceSetFlagFinishCases & tmpCaseStr & ";"
                        End If
                        'mark this site match a case then end the while loop and go to next site.
                        siteDone = True
                    End If
                End If
            End If
            rowCount = rowCount + 1
        Wend
    Next vsite
Exit Function
errHandler:
    TheExec.Datalog.WriteComment "<Error> " + funcName + ":: please check it out."
    If AbortTest Then Exit Function Else Resume Next
End Function

Public Function IsExistInArr(arr As Variant, target As Variant) As Boolean
    Dim i As Double
    IsExistInArr = False
    For i = 0 To UBound(arr)
        If arr(i) = target Then
            IsExistInArr = True
            Exit Function
        End If
    Next i
End Function

Public Function CheckConditionBySite(arr() As String, site As Variant) As Boolean
    Dim i As Double
    CheckConditionBySite = True
    If Not Not arr Then
        For i = 0 To UBound(arr)
            If arr(i) <> "" Then
                If InStr(1, arr(i), "!", vbTextCompare) = 0 Then
                    If TheExec.sites.item(site).FlagState(arr(i)) <> logicTrue Then
                        CheckConditionBySite = False
                        Exit Function
                    End If
                Else
                    If TheExec.sites.item(site).FlagState(Replace(arr(i), "!", "")) <> logicFalse Then
                        CheckConditionBySite = False
                        Exit Function
                    End If
                End If
            End If
        Next i
    End If
End Function

Public Function FlagStatusNum2Str(Flag As Variant) As String
    If Flag = logicTrue Then
        FlagStatusNum2Str = "True"
    ElseIf Flag = logicFalse Then
        FlagStatusNum2Str = "False"
    Else
        FlagStatusNum2Str = "Clear"
    End If
End Function

Public Function Harvest_Postcheck(ByVal FuseBlockName As String, ByVal FuseCategoryName As String, ByVal FailFlagOrValue As String)
    On Error GoTo errHandler
    Dim funcName As String:: funcName = "Harvest_Postcheck"

    'For CP1 <=> FT1 fuse check, put before down bin.
    'This is for FT1 test binswap, since the new method, the test result will influence the efuse
    'If CP1 & FT1 fuse is different and not from down bin, then set BinOut for this device
    
    Dim i As Long
    Dim j As Long
    Dim k As Long
    
    Dim FuseBlockNameArr() As String
    Dim FuseCategoryNameArr() As String
    Dim FailFlagOrValueArr() As String
    Dim tempFailFlagOrValueArr() As String
    Dim vsite As Variant
    Dim Slng_ReadFuseValue As New SiteDouble
    Dim Svar_HexToBin As New SiteVariant
    
    
    Dim sInstName As String
    Dim opbank As eFuseBdfBank
    Dim field As eFuseBdfField
    Dim tempStr As String
    Dim tempSlng As New SiteLong
    Dim tempFuseSlng As New SiteLong
    
    Dim Slng_FuseValue As New SiteVariant
    Dim Svar_Temp As New SiteVariant
    Dim SDbl_ReadFuseValue As New SiteDouble
    Dim HarvCustomFunc As String
    

    Dim DefaultStr As String
    Dim OutputStrArr() As String
    Dim TempHarvStr As String
    Dim StarIdx, EndIdx As Long
    Dim StepOrder As Long
    
    Dim coreIdx As Long
    
    
''''    Dim FinalHarvStr As String: FinalHarvStr = ""
    Dim StoreFlagIdx As Long: StoreFlagIdx = 0

    
    FuseBlockNameArr = Split(FuseBlockName, ";")
    FuseCategoryNameArr = Split(FuseCategoryName, ";")
    FailFlagOrValueArr = Split(FailFlagOrValue, ";")

    sInstName = TheExec.DataManager.instanceName
    TheExec.Datalog.WriteComment "<" & sInstName & ">"
    
    If (UBound(FuseBlockNameArr) = UBound(FuseCategoryNameArr)) And (UBound(FuseBlockNameArr) = UBound(FailFlagOrValueArr)) Then
        For i = 0 To UBound(FuseBlockNameArr)
            HarvCustomFunc = vbNullString
            tempStr = vbNullString
            tempStr = Harvest_StrExpand(FailFlagOrValueArr(i), , , , , HarvCustomFunc)
            If InStr(1, tempStr, "&") <> 0 Then
                tempFailFlagOrValueArr = Split(tempStr, "&")
            Else
                tempFailFlagOrValueArr = Split(tempStr, ",")
            End If
            Set opbank = GetBdfBank(FuseBlockNameArr(i)) 'Set FuseType
            Set field = opbank.Fields(FuseCategoryNameArr(i)) 'Set FuseName
            'Slng_FuseValue => Convert the value from Flags
            'Slng_ReadFuseValue => Convert the value from fuse
            For Each vsite In TheExec.sites.Selected
                tempStr = vbNullString
                Slng_FuseValue(vsite) = 0 'Iintial Site variable for each site
                Svar_Temp(vsite) = 0
                SDbl_ReadFuseValue(vsite) = 0
                Svar_HexToBin(vsite) = 0
                tempFuseSlng(vsite) = 0
'*******************After flow execution, read the flag result and convert to decimal => [Slng_FuseValue]'*******************
                If InStr(1, FailFlagOrValueArr(i), "]") <> 0 And InStr(1, FailFlagOrValueArr(i), "[") <> 0 And InStr(1, FailFlagOrValueArr(i), "&") <> 0 Then
                    For j = 0 To UBound(tempFailFlagOrValueArr)
                        If IsNumeric(tempFailFlagOrValueArr(j)) Then
                            For k = 0 To Len(tempFailFlagOrValueArr(j)) - 1
                                ReDim Preserve OutputStrArr(UBound(OutputStrArr) + 1)
                                OutputStrArr(UBound(OutputStrArr)) = mid(tempFailFlagOrValueArr(j), k + 1, 1)
                            Next k
                        Else
                            DefaultStr = mid(tempFailFlagOrValueArr(j), 1, InStr(tempFailFlagOrValueArr(j), "[") - 1)
                            StarIdx = CLng(mid(tempFailFlagOrValueArr(j), InStr(tempFailFlagOrValueArr(j), "[") + 1, InStr(tempFailFlagOrValueArr(j), ":") - InStr(tempFailFlagOrValueArr(j), "[") - 1))
                            EndIdx = CLng(mid(tempFailFlagOrValueArr(j), InStr(tempFailFlagOrValueArr(j), ":") + 1, InStr(tempFailFlagOrValueArr(j), "]") - InStr(tempFailFlagOrValueArr(j), ":") - 1))
                            If StarIdx > EndIdx Then
                                StepOrder = -1
                            Else
                                StepOrder = 1
                            End If
                            ReDim Preserve OutputStrArr(Abs(StarIdx - EndIdx + StoreFlagIdx))
                            For k = StoreFlagIdx To UBound(OutputStrArr)
                                OutputStrArr(k) = DefaultStr & CStr(StarIdx + k * StepOrder + StoreFlagIdx)
                            Next k
                            StoreFlagIdx = UBound(OutputStrArr) + 1
                        End If
                    Next j
                    
                    If UBound(OutputStrArr) > 32 Then
                        Svar_Temp(vsite) = field.DsscValue(vsite)  'auto_eFuse_GetWriteDecimal(FuseBlockNameArr(i), FuseCategoryNameArr(i), True)
                        Svar_HexToBin(vsite) = GlbUtility.Hex2BinStr(Svar_Temp, field.size)
                    Else
                        SDbl_ReadFuseValue(vsite) = field.DsscDecValue(vsite)
                        Svar_Temp(vsite) = GlbUtility.Dec2Bin(SDbl_ReadFuseValue(vsite), field.size)
                    End If
                    
                    
                    For j = 0 To UBound(OutputStrArr)
                        If UBound(OutputStrArr) > 32 Then
                            tempFuseSlng = CLng(mid(Svar_HexToBin(vsite), j + 1, 1))
                        Else
                            tempFuseSlng = CLng(mid(Svar_Temp(vsite), j + 1, 1))
                        End If
                        tempSlng = Harvest_GetAllSiteFlagState(OutputStrArr(j), 1, 0, 2)
                        If tempFuseSlng = tempSlng Then
                            TheExec.Datalog.WriteComment OutputStrArr(j) & " site" & vsite & " => " & "Current Flag status is mapping Fuse Flag status"
                        ElseIf tempFuseSlng <> tempSlng Then
                            TheExec.sites.item(vsite).FlagState("F_Harvest_PostCheck") = logicTrue
                            TheExec.Datalog.WriteComment "Harvest Binswap, BinOut!!!"
                        End If
                        
                    Next j
                    ReDim OutputStrArr(UBound(OutputStrArr) + 1)
                    StoreFlagIdx = 0
                    
                ElseIf InStr(1, FailFlagOrValueArr(i), "]") <> 0 And InStr(1, FailFlagOrValueArr(i), "[") <> 0 And InStr(1, FailFlagOrValueArr(i), "&") = 0 And InStr(1, FailFlagOrValueArr(i), ":") <> 0 Then
                    
                    DefaultStr = mid(FailFlagOrValueArr(i), 1, InStr(FailFlagOrValueArr(i), "[") - 1)
                    StarIdx = CLng(mid(FailFlagOrValueArr(i), InStr(FailFlagOrValueArr(i), "[") + 1, InStr(FailFlagOrValueArr(i), ":") - InStr(FailFlagOrValueArr(i), "[") - 1))
                    EndIdx = CLng(mid(FailFlagOrValueArr(i), InStr(FailFlagOrValueArr(i), ":") + 1, InStr(FailFlagOrValueArr(i), "]") - InStr(FailFlagOrValueArr(i), ":") - 1))
                    
                    If UBound(tempFailFlagOrValueArr) > 32 Then
                        Svar_Temp(vsite) = field.DsscValue(vsite)  'auto_eFuse_GetWriteDecimal(FuseBlockNameArr(i), FuseCategoryNameArr(i), True)
                        Svar_HexToBin(vsite) = GlbUtility.Hex2BinStr(Svar_Temp, field.size)
                    Else
                        SDbl_ReadFuseValue(vsite) = field.DsscDecValue(vsite)
                        Svar_Temp(vsite) = GlbUtility.Dec2Bin(SDbl_ReadFuseValue(vsite), field.size)
                    End If
  
                    For j = 0 To UBound(tempFailFlagOrValueArr)
                        If StarIdx > EndIdx Then
                            If UBound(tempFailFlagOrValueArr) > 32 Then
                                tempFuseSlng = CLng(mid(Svar_HexToBin(vsite), j + 1, 1))
                            Else
                                tempFuseSlng = CLng(mid(Svar_Temp(vsite), j + 1, 1))
                            End If
                            tempSlng = Harvest_GetAllSiteFlagState(tempFailFlagOrValueArr(j), 1, 0, 2)
                            If tempFuseSlng = tempSlng Then
                                TheExec.Datalog.WriteComment tempFailFlagOrValueArr(j) & " site" & vsite & " => " & "Current Flag status is mapping Fuse Flag status"
                            ElseIf tempFuseSlng <> tempSlng Then
                                TheExec.sites.item(vsite).FlagState("F_Harvest_PostCheck") = logicTrue
                                TheExec.Datalog.WriteComment "Harvest Binswap, BinOut!!!"
                            End If
                            
                        ElseIf StarIdx < EndIdx Then
                            If UBound(tempFailFlagOrValueArr) > 32 Then
                                tempFuseSlng = CLng(mid(Svar_HexToBin(vsite), j + 1, 1))
                            Else
                                tempFuseSlng = CLng(mid(Svar_Temp(vsite), j + 1, 1))
                            End If
                            tempSlng = Harvest_GetAllSiteFlagState(tempFailFlagOrValueArr(j), 1, 0, 2)
                            If tempFuseSlng = tempSlng Then
                                TheExec.Datalog.WriteComment tempFailFlagOrValueArr(j) & " site" & vsite & " => " & "Current Flag status is mapping Fuse Flag status"
                            ElseIf tempFuseSlng <> tempSlng Then
                                TheExec.sites.item(vsite).FlagState("F_Harvest_PostCheck") = logicTrue
                                TheExec.Datalog.WriteComment "Harvest Binswap, BinOut!!!"
                            End If
                            
                        Else
                            TheExec.Datalog.WriteComment "Please Check The Flag Format"
                        End If
                        
                    Next j
                ElseIf InStr(1, FailFlagOrValueArr(i), "]") <> 0 And InStr(1, FailFlagOrValueArr(i), "[") <> 0 And InStr(1, FailFlagOrValueArr(i), "&") = 0 And InStr(1, FailFlagOrValueArr(i), ":") = 0 Then
                    ''''''F_Soc_DE[0] -> F_Soc_DE0
                    
                    
                    If field.size > 32 Then
                        Svar_Temp(vsite) = field.DsscValue(vsite)  'auto_eFuse_GetWriteDecimal(FuseBlockNameArr(i), FuseCategoryNameArr(i), True)
                        Svar_HexToBin(vsite) = GlbUtility.Hex2BinStr(Svar_Temp, field.size)
                    Else
                        SDbl_ReadFuseValue(vsite) = field.DsscDecValue(vsite)
                        Svar_Temp(vsite) = GlbUtility.Dec2Bin(SDbl_ReadFuseValue(vsite), field.size)
                    End If
                    
                    
                    For j = 0 To UBound(tempFailFlagOrValueArr)
                        If UBound(tempFailFlagOrValueArr) > 32 Then
                            tempFuseSlng = CLng(mid(Svar_HexToBin(vsite), j + 1, 1))
                        Else
                            tempFuseSlng = CLng(mid(Svar_Temp(vsite), j + 1, 1))
                        End If
                        tempSlng = Harvest_GetAllSiteFlagState(tempFailFlagOrValueArr(j), 1, 0, 2)
                        If tempFuseSlng = tempSlng Then
                            TheExec.Datalog.WriteComment tempFailFlagOrValueArr(j) & " site" & vsite & " => " & "Current Flag status is mapping Fuse Flag status"
                        ElseIf tempFuseSlng <> tempSlng Then
                            TheExec.sites.item(vsite).FlagState("F_Harvest_PostCheck") = logicTrue
                            TheExec.Datalog.WriteComment "Harvest Binswap, BinOut!!!"
                        End If
                    Next j

                End If
                
            Next vsite
        Next i
    End If

Exit Function

errHandler:
    Call Print_Error_Message(Error_Info, "VBT_LIB_Common_AP", "Harvest_Postcheck") 'Add ErrHandler 2023/08/18
    If AbortTest Then Exit Function Else Resume Next 'Add ErrHandler 2023/08/18
End Function
