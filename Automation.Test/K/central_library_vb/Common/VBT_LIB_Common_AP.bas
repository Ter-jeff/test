Attribute VB_Name = "VBT_LIB_Common_AP"
Option Explicit
Function VBT_IEDA_Registry(RegistryName As String, Optional OnOff As Boolean = True, Optional DebugPrint As Boolean = True)
On Error GoTo errHandler 'Add ErrHandler 2023/08/18

    Dim funcName As String:: funcName = "VBT_IEDA_Registry"
    Dim Inputstr As String
    If OnOff Then
        Call IEDA_Initialize(Inputstr)  'clean up strings
        Call IEDA_GetString(Inputstr, RegistryName)  'compose ieda string
        Call IEDA_AutoCheck_Print(Inputstr, RegistryName, DebugPrint)   'show log
        Call IEDA_SaveRegistry(Inputstr, RegistryName)  'save to registry
    End If

Exit Function 'Add ErrHandler 2023/08/18
errHandler: 'Add ErrHandler 2023/08/18
    Call Print_Error_Message(Error_Info, "VBT_LIB_Common_AP", "VBT_IEDA_Registry") 'Add ErrHandler 2023/08/18
    If AbortTest Then Exit Function Else Resume Next 'Add ErrHandler 2023/08/18
End Function

Public Function ECID_DTS() 'VBT function
On Error GoTo errHandler 'Add ErrHandler 2023/08/18
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
        theexec.Datalog.WriteComment ("WarNing: DTS system and OCR Handler system do not run.")
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
    
Exit Function 'Add ErrHandler 2023/08/18
errHandler: 'Add ErrHandler 2023/08/18
    Call Print_Error_Message(Error_Info, "VBT_LIB_Common_AP", "ECID_DTS") 'Add ErrHandler 2023/08/18
    If AbortTest Then Exit Function Else Resume Next 'Add ErrHandler 2023/08/18
End Function

Public Function Harvest_eFuse_Write(ByVal FuseBlockName As String, ByVal FuseCategoryName As String, ByVal FailFlagOrValue As String) As Long
On Error GoTo errHandler 'Add ErrHandler 2023/08/18
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
    
    Dim FlagAssign As Boolean
         
    FuseBlockNameArr = Split(FuseBlockName, ";")
    FuseCategoryNameArr = Split(FuseCategoryName, ";")
    FailFlagOrValueArr = Split(FailFlagOrValue, ";")
    
    sInstName = theexec.DataManager.instancename
    theexec.Datalog.WriteComment "<" & sInstName & ">"

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
                    tempFailFlagOrValueArr(0) = CStr(Bin2Dec(tempFailFlagOrValueArr(0)))
                ElseIf auto_isHexString(tempFailFlagOrValueArr(0)) Then
                    tempFailFlagOrValueArr(0) = CStr(auto_HexStr2Value(tempFailFlagOrValueArr(0)))
                ElseIf UCase(tempFailFlagOrValueArr(0)) = "NA" Then
                '20240923 michael
                ElseIf InStr(tempFailFlagOrValueArr(0), "[") = 0 And InStr(tempFailFlagOrValueArr(0), ":") Then
                    FlagAssign = True
                    
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
                ElseIf FlagAssign Then
                    Sdbl_FuseValue = GetFlagAssignValue(tempFailFlagOrValueArr(0))
                    opbank.SetEfuse field.name, Sdbl_FuseValue, , , , , True
                    
                ElseIf HarvCustomFunc <> "" Then
                    Call JudgeCustomFun(CVar(tempFailFlagOrValueArr(0)), HarvCustomFunc, Sdbl_FuseValue)
'                    Call auto_eFuse_SetWriteVariable_SiteAware(FuseBlockNameArr(i), FuseCategoryNameArr(i), Sdbl_FuseValue, True)
                    opbank.SetEfuse field.name, Sdbl_FuseValue, , , , , True
''''''                    For Each site In theexec.sites.Active
''''''                        opbank.SetEfuseBySite field.name, Sdbl_FuseValue
''''''                        opbank.SetEfuseBySitePf field.name, True
''''''                    Next site
                    
                Else
                    'FuseCategoryName = CFG_Condition_47
                    'FailFlagOrValue = F_ECPU_CORE3
                    If UCase(tempFailFlagOrValueArr(0)) = "NA" Then
                        theexec.Datalog.WriteComment "Harvest_eFuse_Write:" & tempFuseCategoryNameArr(0) & ":Did not need to fuse in current stage."
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
                For Each site In theexec.sites.Active
                    For j = 0 To UBound(tempFailFlagOrValueArr)
                        tempSlng = Harvest_GetAllSiteFlagState(tempFailFlagOrValueArr(j), 1, 0, 1)
                        If tempSlng(site) = 1 Then
                            Sdbl_FuseValue = Sdbl_FuseValue.Add(Application.WorksheetFunction.Power(2, UBound(tempFailFlagOrValueArr) - j))
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
                theexec.Datalog.WriteComment "<Error> Harvest_eFuse_Write: did not support this case yet."
            End If
            FlagAssign = False
        Next i
    Else
        theexec.Datalog.WriteComment "<Error> Harvest_eFuse_Write: input is not match."
    End If
Exit Function 'Add ErrHandler 2023/08/18
errHandler: 'Add ErrHandler 2023/08/18
    Call Print_Error_Message(Error_Info, "VBT_LIB_Common_AP", "Harvest_eFuse_Write") 'Add ErrHandler 2023/08/18
    If AbortTest Then Exit Function Else Resume Next 'Add ErrHandler 2023/08/18
End Function

Public Function Harvest_eFuse_Read(ByVal FuseBlockName As String, ByVal FuseCategoryName As String, ByVal FailFlagOrValue As String) As Long
On Error GoTo errHandler 'Add ErrHandler 2023/08/18
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
    FuseBlockNameArr = Split(FuseBlockName, ";")
    FuseCategoryNameArr = Split(FuseCategoryName, ";")
    FailFlagOrValueArr = Split(FailFlagOrValue, ";")
    
    sInstName = theexec.DataManager.instancename
    theexec.Datalog.WriteComment "<" & sInstName & ">"
    
    If (UBound(FuseBlockNameArr) = UBound(FuseCategoryNameArr)) And (UBound(FuseBlockNameArr) = UBound(FailFlagOrValueArr)) Then
        For i = 0 To UBound(FuseBlockNameArr)
            tempStr = Harvest_StrExpand(FuseCategoryNameArr(i))
            tempFuseCategoryNameArr = Split(tempStr, ",")
            tempStr = Harvest_StrExpand(FailFlagOrValueArr(i), , , , , HarvCustomFunc)
            tempFailFlagOrValueArr = Split(tempStr, ",")
            Set opbank = GetBdfBank(FuseBlockNameArr(i)) 'Set FuseType
            Set field = opbank.Fields(FuseCategoryNameArr(i)) 'Set FuseName
             
            If (UBound(tempFuseCategoryNameArr) = 0) And (UBound(tempFailFlagOrValueArr) = 0) Then
                If IsNumeric(tempFailFlagOrValueArr(0))  Or auto_isBinaryString(tempFailFlagOrValueArr(0)) Or auto_isHexString(tempFailFlagOrValueArr(0)) Then
                    theexec.Datalog.WriteComment "<Error> Harvest_eFuse_Read: do not support this case."
                ElseIf HarvCustomFunc <> "" Then
                    Dim TempFlagArr() As String
                    For Each vsite In theexec.sites.Selected
'                        Sdbl_ReadFuseValue(vSite) = auto_eFuse_GetWriteDecimal(FuseBlockNameArr(i), FuseCategoryNameArr(i), True)
                        SDbl_ReadFuseValue(vsite) = field.DsscDecValue(vsite)
                    Next vsite
                    tempFailFlagOrValueArr(0) = Harvest_StrExpand(tempFailFlagOrValueArr(0))
                    TempFlagArr = Split(tempFailFlagOrValueArr(0), ",")
                    Call JudgeCustomFun(tempFailFlagOrValueArr(0), HarvCustomFunc, SDbl_ReadFuseValue, Svar_ReadFuseValue)
                    For Each vsite In theexec.sites.Selected
                        For j = 0 To Len(Svar_ReadFuseValue) - 1
                            tempSlng = mid(Svar_ReadFuseValue, 1 + j, 1)
                            Call Harvest_SetAllSiteFlagState(TempFlagArr(j), tempSlng)
                        Next j
                    Next vsite
                ElseIf InStr(tempFailFlagOrValueArr(0), "&") AND InStr(tempFailFlagOrValueArr(0), ":")Then
                    For Each vsite In TheExec.sites.Selected
                        SDbl_ReadFuseValue(vsite) = field.DsscDecValue(vsite)
''''                        Slng_ReadFuseValue(vSite) = auto_eFuse_GetReadDecimal(FuseBlockNameArr(i), FuseCategoryNameArr(i), True)
                    Next vsite
                    Call Harvest_SetAllSiteFlagState(FailFlagOrValueArr(i), SDbl_ReadFuseValue)
                Else
                    'FuseCategoryName = CFG_Condition_47
                    'FailFlagOrValue = F_ECPU_CORE3
                    For Each vsite In theexec.sites.Selected
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
                    For Each vsite In theexec.sites.Selected
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
                For Each vsite In theexec.sites.Selected
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
                For Each vsite In theexec.sites.Selected
                    SDbl_ReadFuseValue(vsite) = field.DsscDecValue(vsite)
                    theexec.Datalog.WriteComment "Site" & vsite & " => " & "Bin_Fuse = " & SDbl_ReadFuseValue(vsite)
'''''''                    If SDbl_ReadFuseValue(vsite) = 6 Then
'''''''                        theexec.sites.item(vsite).FlagState("F_IDS_PCPU_HIGH_LEAKAGE") = logicTrue
'''''''                    Else
'''''''                        theexec.sites.item(vsite).FlagState("F_IDS_PCPU_HIGH_LEAKAGE") = logicFalse
'''''''                    End If
                Next vsite
            Else
                theexec.Datalog.WriteComment "<Error> Harvest_eFuse_Read: did not support this case yet."
            End If
        Next i
    Else
        theexec.Datalog.WriteComment "<Error> Harvest_eFuse_Read: input is not match."
    End If

Exit Function 'Add ErrHandler 2023/08/18
errHandler: 'Add ErrHandler 2023/08/18
    Call Print_Error_Message(Error_Info, "VBT_LIB_Common_AP", "Harvest_eFuse_Read") 'Add ErrHandler 2023/08/18
    If AbortTest Then Exit Function Else Resume Next 'Add ErrHandler 2023/08/18
End Function

Public Function Harvest_SUBFLOW()
On Error GoTo errHandler 'Add ErrHandler 2023/08/18
    
    Dim funcName As String:: funcName = "Harvest_SUBFLOW"
    
    
    Dim site As Variant
    
    theexec.sites.item(1).SiteVariableValue("HARVEST_PASS") = 1
    theexec.sites.item(2).SiteVariableValue("HARVEST_PASS") = 2
    
Exit Function 'Add ErrHandler 2023/08/18
errHandler: 'Add ErrHandler 2023/08/18
    Call Print_Error_Message(Error_Info, "VBT_LIB_Common_AP", "Harvest_SUBFLOW") 'Add ErrHandler 2023/08/18
    If AbortTest Then Exit Function Else Resume Next 'Add ErrHandler 2023/08/18
End Function



Public Function Harvest_FailedCoreCount(ByVal BlockType As String, ByVal Harv_GlobalFailFlag As String, _
                                        ByVal Harv_FailCoreSum As String, ByVal Harv_FailCoreSumFlag As String, _
                                        Optional ByVal Harv_AllCorePassFlag As String, Optional ByVal Harv_LocalFailFlag As String) As Long
On Error GoTo errHandler 'Add ErrHandler 2023/08/18
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
    
    sInstName = theexec.DataManager.instancename
    
    If Harv_FailCoreSum <> "" And Harv_FailCoreSumFlag <> "" Then
        FailCoreSumArr = Split(Harv_FailCoreSum, ";")
        FailCoreSumFlagArr = Split(Harv_FailCoreSumFlag, ";")
        If UBound(FailCoreSumArr) = UBound(FailCoreSumFlagArr) Then
            lng_MaxFailCoreSum = -1
            For i = 0 To UBound(FailCoreSumArr)
                lng_MaxFailCoreSum = Application.WorksheetFunction.max(lng_MaxFailCoreSum, FailCoreSumArr(i))
            Next i
        Else
            theexec.Datalog.WriteComment "Warning: Harv_FailCoreSum and Harv_FailCoreSumFlag did not match."
            theexec.Flow.TestLimit resultVal:=1, lowVal:=0, hiVal:=0, Tname:=sInstName, ForceResults:=tlForceNone
            Exit Function
        End If
    Else
        theexec.Datalog.WriteComment "Warning: Harv_FailCoreSum or Harv_FailCoreSumFlag did not have any information."
        theexec.Flow.TestLimit resultVal:=1, lowVal:=0, hiVal:=0, Tname:=sInstName, ForceResults:=tlForceNone
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
            For Each vsite In theexec.sites.Selected

                ReDim EachCoreResultAry_TempBySite(FailFlagEndIdx) '(FailFlagEndIdx - FailFlagStartIdx + 1)
                ReDim FailCoreAry_TempBySite(0)
                FailCoreAry_TempBySite(FailFlagStartIdx) = "N/A"
                For i = FailFlagStartIdx To FailFlagEndIdx
                    If theexec.sites.item(vsite).FlagState(FailFlagDefaultName & CStr(i)) = logicTrue Then
                        FailCoreCnt_Slng = FailCoreCnt_Slng + 1
                        If FirstFailCore_Slng = -1 Then
                            FirstFailCore_Slng = i
                        End If
                        ReDim Preserve FailCoreAry_TempBySite(FailCoreCnt_Slng - 1)
                        FailCoreAry_TempBySite(FailCoreCnt_Slng - 1) = CStr(i)
                        EachCoreResultAry_TempBySite(i) = "F" '(i - FailFlagStartIdx)
                    ElseIf theexec.sites.item(vsite).FlagState(FailFlagDefaultName & CStr(i)) = logicFalse Then
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
                theexec.Datalog.WriteComment BlockType & "_Harvested_Core," + "Site" + CStr(vsite) + "," + Join(FailCoreAry_TempBySite, ",")
                'ex: GfxSa_Harvesting_Result,Site0,p,p,p,p,p
                theexec.Datalog.WriteComment BlockType & "_Harvesting_Result," + "Site" + CStr(vsite) + "," + Join(EachCoreResultAry_TempBySite, ",")

                theexec.sites.item(vsite).FlagState(Harv_AllCorePassFlag) = logicFalse
                If FailCoreCnt_Slng > lng_MaxFailCoreSum Then
                    FirstFailCore_Slng = 999
                    For i = 0 To lng_MaxFailCoreSum
                        theexec.sites.item(vsite).FlagState(FailCoreSumFlagArr(i)) = logicClear
                    Next i
                ElseIf FailCoreCnt_Slng = 0 Then
                    'all core pass
                    theexec.sites.item(vsite).FlagState(Harv_AllCorePassFlag) = logicTrue
                    theexec.sites.item(vsite).FlagState(FailCoreSumFlagArr(FailCoreCnt_Slng)) = logicTrue
                Else
                    'FailCoreCnt_Slng = 1 to lng_MaxFailCoreSum
                    For i = 0 To lng_MaxFailCoreSum - 1
                        theexec.sites.item(vsite).FlagState(FailCoreSumFlagArr(i)) = logicClear
                    Next i
                    theexec.sites.item(vsite).FlagState(FailCoreSumFlagArr(FailCoreCnt_Slng)) = logicTrue
                End If
                
                'for T-ELLIS correlation
                If FailCoreCnt_Slng = 1 Then
                    theexec.sites.item(vsite).FlagState("F_Gfx_HARV") = logicTrue
                End If
                
            Next vsite
            
            theexec.Flow.TestLimit resultVal:=FirstFailCore_Slng, lowVal:=-1, hiVal:=FailFlagEndIdx, Tname:=BlockType + "_Harvested_Core", ForceResults:=tlForceNone
            theexec.Flow.TestLimit resultVal:=FailCoreCnt_Slng, lowVal:=0, hiVal:=lng_MaxFailCoreSum, Tname:=BlockType + "_Fail_Core_Count", ForceResults:=tlForceNone

        ElseIf InStr(Harv_GlobalFailFlag, ",") <> 0 Then
        End If
    End If
Exit Function 'Add ErrHandler 2023/08/18
errHandler: 'Add ErrHandler 2023/08/18
    Call Print_Error_Message(Error_Info, "VBT_LIB_Common_AP", "Harvest_FailedCoreCount") 'Add ErrHandler 2023/08/18
    If AbortTest Then Exit Function Else Resume Next 'Add ErrHandler 2023/08/18
End Function

Public Function Harvest_Summary(ByVal Harv_AllFlag As String, ByVal Harv_GlobalFailFlag As String, _
                                ByVal Harv_FailCoreSumFlag As String, Optional ByVal Harv_AllCorePassFlag As String, _
                                Optional ByVal Harv_BinOutFailFlag As String) As Long
On Error GoTo errHandler 'Add ErrHandler 2023/08/18
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
    With theexec.Datalog
        .Setup.Shared.ascii.Columns.EnableCustomWidths = True
        .Setup.Shared.ascii.Columns.Parametric.TestName.Width = 150
        .Setup.Shared.ascii.Columns.Parametric.Measured.Width = 16
        .Setup.Shared.ascii.Columns.Functional.TestName.Width = 150
        .Setup.Shared.ascii.Columns.Functional.Pattern.Width = 100
        .ApplySetup
    End With

    sInstName = theexec.DataManager.instancename
    
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
        theexec.Datalog.WriteComment "Warning: Harv_FailCoreSumFlag/Harv_GlobalFailFlag/Harv_BinOutFailFlag did not have any information."
        theexec.Flow.TestLimit resultVal:=1, lowVal:=0, hiVal:=0, Tname:=sInstName, ForceResults:=tlForceNone
        Exit Function
    End If
    
    ReDim Block_FailCoreCnt_Slng(UBound(Block_BinOutFailFlagArr))
    
    Call initHarvSumFlag(Harv_FailCoreSumFlag, Harv_AllCorePassFlag, Harv_BinOutFailFlag)
    
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
                lng_MaxFailCoreSum = Application.WorksheetFunction.max(lng_MaxFailCoreSum, TempStoreSumRange(UBound(TempStoreSumRange)))
'''                If lng_MaxFailCoreSum > TempRangeUB Then
'''                    TempRangeUB = lng_MaxFailCoreSum
'''                End If
                TempBoolSumRange = True
            Else
                lng_MaxFailCoreSum = Application.WorksheetFunction.max(lng_MaxFailCoreSum, FailCoreSumArr(i))
            End If
        Next i
        
        
        If InStr(Block_GlobalFailFlagArr(Block_Idx), "[") <> 0 And InStr(Block_GlobalFailFlagArr(Block_Idx), "]") <> 0 And InStr(Block_GlobalFailFlagArr(Block_Idx), ":") <> 0 Then
            FailFlagDefaultName = mid(Block_GlobalFailFlagArr(Block_Idx), 1, InStr(Block_GlobalFailFlagArr(Block_Idx), "[") - 1)
            FailFlagStartIdx = CLng(mid(Block_GlobalFailFlagArr(Block_Idx), InStr(Block_GlobalFailFlagArr(Block_Idx), "[") + 1, InStr(Block_GlobalFailFlagArr(Block_Idx), ":") - InStr(Block_GlobalFailFlagArr(Block_Idx), "[") - 1))
            FailFlagEndIdx = CLng(mid(Block_GlobalFailFlagArr(Block_Idx), InStr(Block_GlobalFailFlagArr(Block_Idx), ":") + 1, InStr(Block_GlobalFailFlagArr(Block_Idx), "]") - InStr(Block_GlobalFailFlagArr(Block_Idx), ":") - 1))
            
            'only consider [0:xxx]
            FailCoreCnt_Slng = 0
            FirstFailCore_Slng = -1
            
            For Each vsite In theexec.sites.Selected
                Block_FailCoreCnt_Slng(Block_Idx)(vsite) = 0
                ReDim EachCoreResultAry_TempBySite(FailFlagEndIdx) '(FailFlagEndIdx - FailFlagStartIdx + 1)
                ReDim FailCoreAry_TempBySite(0)
                FailCoreAry_TempBySite(FailFlagStartIdx) = "N/A"
                For i = FailFlagStartIdx To FailFlagEndIdx
                    If theexec.sites.item(vsite).FlagState(FailFlagDefaultName & CStr(i)) = logicTrue Then
                        FailCoreCnt_Slng = FailCoreCnt_Slng + 1
                        Block_FailCoreCnt_Slng(Block_Idx)(vsite) = Block_FailCoreCnt_Slng(Block_Idx)(vsite) + 1
                        If FirstFailCore_Slng = -1 Then
                            FirstFailCore_Slng = i
                        End If
                        ReDim Preserve FailCoreAry_TempBySite(FailCoreCnt_Slng - 1)
                        FailCoreAry_TempBySite(FailCoreCnt_Slng - 1) = CStr(i)
                        EachCoreResultAry_TempBySite(i) = "F" '(i - FailFlagStartIdx)
                    ElseIf theexec.sites.item(vsite).FlagState(FailFlagDefaultName & CStr(i)) = logicFalse Then
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
                        theexec.sites.item(vsite).FlagState(FailCoreSumFlagArr(i)) = logicClear
                    Next i
                    theexec.sites.item(vsite).FlagState(Block_BinOutFailFlagArr(Block_Idx)) = logicTrue
                    theexec.sites.item(vsite).FlagState(Block_AllCorePassFlagArr(Block_Idx)) = logicFalse
                ElseIf FailCoreCnt_Slng = 0 Then
                    'all core pass
                    theexec.sites.item(vsite).FlagState(Block_BinOutFailFlagArr(Block_Idx)) = logicFalse

                    theexec.sites.item(vsite).FlagState(Block_AllCorePassFlagArr(Block_Idx)) = logicTrue
                    theexec.sites.item(vsite).FlagState(FailCoreSumFlagArr(FailCoreCnt_Slng)) = logicTrue
                Else
                    'FailCoreCnt_Slng = 1 to lng_MaxFailCoreSum
                    For i = 0 To lng_MaxFailCoreSum - 1
                        theexec.sites.item(vsite).FlagState(FailCoreSumFlagArr(i)) = logicClear
                    Next i
                    theexec.sites.item(vsite).FlagState(FailCoreSumFlagArr(FailCoreCnt_Slng)) = logicTrue
                    theexec.sites.item(vsite).FlagState(Block_BinOutFailFlagArr(Block_Idx)) = logicFalse
                    theexec.sites.item(vsite).FlagState(Block_AllCorePassFlagArr(Block_Idx)) = logicFalse
                End If
               
            Next vsite

        ElseIf InStr(Block_GlobalFailFlagArr(Block_Idx), "+") <> 0 Then
            
            Block_GlobalFailFlagArr_Split = Split(Block_GlobalFailFlagArr(Block_Idx), "+") ''F_ECPU_CORE0+F_ECPU_CORE1+F_ECPU_CORE2+F_ECPU_CORE3
            
            FailCoreCnt_Slng = 0
            For Each vsite In theexec.sites.Selected
                Block_FailCoreCnt_Slng(Block_Idx)(vsite) = 0
                For i = 0 To UBound(Block_GlobalFailFlagArr_Split)
                    If theexec.sites.item(vsite).FlagState(Block_GlobalFailFlagArr_Split(i)) = logicTrue Then
                        FailCoreCnt_Slng = FailCoreCnt_Slng + 1
                        Block_FailCoreCnt_Slng(Block_Idx)(vsite) = Block_FailCoreCnt_Slng(Block_Idx)(vsite) + 1
                        If FirstFailCore_Slng = -1 Then
                            FirstFailCore_Slng = i
                        End If
                    ElseIf theexec.sites.item(vsite).FlagState(Block_GlobalFailFlagArr_Split(i)) = logicFalse Then
                    Else
                        FailCoreCnt_Slng = FailCoreCnt_Slng + 1
                        Block_FailCoreCnt_Slng(Block_Idx)(vsite) = Block_FailCoreCnt_Slng(Block_Idx)(vsite) + 1
                    End If
                Next i
                
                If FailCoreCnt_Slng > lng_MaxFailCoreSum Then
                    FirstFailCore_Slng = 999
                    Dim tempFailCoreSum As Long
                    If TempBoolSumRange = True Then
                        tempFailCoreSum = UBound(FailCoreSumFlagArr) + 1
                    Else
                        tempFailCoreSum = lng_MaxFailCoreSum
                    End If
                    For i = 0 To tempFailCoreSum - 1
                        theexec.sites.item(vsite).FlagState(FailCoreSumFlagArr(i)) = logicClear
                    Next i
                    theexec.sites.item(vsite).FlagState(Block_BinOutFailFlagArr(Block_Idx)) = logicTrue
                    theexec.sites.item(vsite).FlagState(Block_AllCorePassFlagArr(Block_Idx)) = logicFalse

                ElseIf FailCoreCnt_Slng = 0 Then
                'all core pass
                    theexec.sites.item(vsite).FlagState(Block_BinOutFailFlagArr(Block_Idx)) = logicFalse
                    theexec.sites.item(vsite).FlagState(Block_AllCorePassFlagArr(Block_Idx)) = logicTrue
                    If TempBoolSumRange = True Then
                        For i = 0 To UBound(FailCoreSumArr)
                            If InStr(1, UCase(FailCoreSumArr(i)), UCase("0to")) <> 0 Then
                                theexec.sites.item(vsite).FlagState(FailCoreSumFlagArr(i)) = logicTrue
                            End If
                        Next i
'''''                        theexec.sites.item(vSite).FlagState(Block_AllCorePassFlagArr(Block_Idx)) = logicTrue
                    Else
                        theexec.sites.item(vsite).FlagState(FailCoreSumFlagArr(FailCoreCnt_Slng)) = logicTrue
                    End If
                ElseIf TempBoolSumRange = True Then
                    For i = 0 To UBound(FailCoreSumArr)
                        If InStr(1, UCase(FailCoreSumArr(i)), UCase("to")) <> 0 Then
                            TempStoreSumRange = Split(UCase(FailCoreSumArr(i)), UCase("to"))
                            TempRangeLB = TempStoreSumRange(0)
                            TempRangeUB = TempStoreSumRange(1)
                            If TempRangeLB <= FailCoreCnt_Slng And FailCoreCnt_Slng <= TempRangeUB Then
                                theexec.sites.item(vsite).FlagState(FailCoreSumFlagArr(i)) = logicTrue
                            End If
                        ElseIf FailCoreCnt_Slng = CInt(FailCoreSumArr(i)) Then
                            theexec.sites.item(vsite).FlagState(FailCoreSumFlagArr(i)) = logicTrue
                        End If
                    Next i

                    theexec.sites.item(vsite).FlagState(Block_BinOutFailFlagArr(Block_Idx)) = logicFalse
                    theexec.sites.item(vsite).FlagState(Block_AllCorePassFlagArr(Block_Idx)) = logicFalse
                    
                Else
                    'FailCoreCnt_Slng = 1 to lng_MaxFailCoreSum
                    For i = 0 To lng_MaxFailCoreSum - 1
                        theexec.sites.item(vsite).FlagState(FailCoreSumFlagArr(i)) = logicClear
                    Next i
                    theexec.sites.item(vsite).FlagState(FailCoreSumFlagArr(FailCoreCnt_Slng)) = logicTrue
                    theexec.sites.item(vsite).FlagState(Block_BinOutFailFlagArr(Block_Idx)) = logicFalse
                    theexec.sites.item(vsite).FlagState(Block_AllCorePassFlagArr(Block_Idx)) = logicFalse

                End If
            Next vsite
        End If
    Next Block_Idx
    
    
    Dim FlagCnt As Long
    Dim lTestResultFail As Long
    
    Dim HarvestFlagArr() As String
    Dim TempHarvestStr() As String
    
    Dim Tname_Str As String
    Dim HarvestFlagArr_Num As String
    
    Dim sLSumResult  As New SiteLong
    
    HarvestFlagArr = Split(Harv_AllFlag, "+")
    HarvestFlagArr_Num = UBound(HarvestFlagArr)
    ReDim TempHarvestStr(HarvestFlagArr_Num)
    lTestResultFail = 2 ''Fail value
    For Each vsite In theexec.sites
        If Harv_BinOutFailFlag <> "" Then
            For FlagCnt = 0 To HarvestFlagArr_Num
                Tname_Str = HarvestFlagArr(FlagCnt)
                ''''Harv_BinOutFailFlag = F_GFX_HARV_SUM|F_ECPU_CORE_SUM
                If InStr(Harv_BinOutFailFlag, HarvestFlagArr(FlagCnt)) <> 0 Then
                    For i = 0 To UBound(Block_BinOutFailFlagArr)
                        If UCase(HarvestFlagArr(FlagCnt)) = UCase(Block_BinOutFailFlagArr(i)) Then
                            If theexec.sites.item(vsite).FlagState(HarvestFlagArr(FlagCnt)) = logicTrue Then
                                TempHarvestStr(FlagCnt) = 1
                            Else
                                TempHarvestStr(FlagCnt) = 0
                            End If
                            theexec.Flow.TestLimit resultVal:=Block_FailCoreCnt_Slng(i)(vsite), lowVal:=0, hiVal:=0, Tname:=Tname_Str, ForceResults:=tlForceNone
    ''                            TheExec.Flow.TestLimit resultVal:=TempHarvestStr(FlagCnt), lowVal:=0, hiVal:=0, Tname:=Tname_Str, ForceResults:=tlForceNone
                        End If
                    Next i
                    theexec.Flow.TestLimit resultVal:=TempHarvestStr(FlagCnt), lowVal:=0, hiVal:=0, Tname:=Tname_Str, ForceResults:=tlForceNone
                ''''Harv_AllCorePassFlag = F_GFX_ALL_CORE_PASS|F_ECPU_ALL_CORE_PASS
                ElseIf InStr(Harv_AllCorePassFlag, HarvestFlagArr(FlagCnt)) <> 0 Then
                    For i = 0 To UBound(Block_AllCorePassFlagArr)
                        If UCase(HarvestFlagArr(FlagCnt)) = UCase(Block_AllCorePassFlagArr(i)) Then
''                            If UCase(Block_AllCorePassFlagArr(i)) = "F_ECPU_ALL_CORE_PASS" Then Stop
                            If theexec.sites.item(vsite).FlagState(HarvestFlagArr(FlagCnt)) = logicTrue Then
                                TempHarvestStr(FlagCnt) = 1
                            Else
                                TempHarvestStr(FlagCnt) = 0
                            End If
''                            TheExec.Flow.TestLimit resultVal:=TempHarvestStr(FlagCnt), Tname:=Tname_Str, ForceResults:=tlForceNone
                            theexec.Flow.TestLimit resultVal:=TempHarvestStr(FlagCnt), lowVal:=0, hiVal:=1, Tname:=Tname_Str, ForceResults:=tlForceNone
                        End If
                    Next i
                Else
                    If theexec.sites.item(vsite).FlagState(HarvestFlagArr(FlagCnt)) = logicTrue Then
                        TempHarvestStr(FlagCnt) = 1
                    Else
                        TempHarvestStr(FlagCnt) = 0
                    End If
''                    TheExec.Flow.TestLimit resultVal:=TempHarvestStr(FlagCnt), Tname:=Tname_Str, ForceResults:=tlForceNone
                    theexec.Flow.TestLimit resultVal:=TempHarvestStr(FlagCnt), lowVal:=0, hiVal:=0, Tname:=Tname_Str, ForceResults:=tlForceNone
                    
                End If
            Next FlagCnt
        End If
    Next vsite
    
    With theexec.Datalog
        .Setup.Shared.ascii.Columns.EnableCustomWidths = True
        .Setup.Shared.ascii.Columns.Parametric.TestName.Width = 150
        .Setup.Shared.ascii.Columns.Parametric.Measured.Width = 16
        .Setup.Shared.ascii.Columns.Functional.TestName.Width = 150
        .Setup.Shared.ascii.Columns.Functional.Pattern.Width = 100
        .ApplySetup
    End With

Exit Function 'Add ErrHandler 2023/08/18
errHandler: 'Add ErrHandler 2023/08/18
    Call Print_Error_Message(Error_Info, "VBT_LIB_Common_AP", "Harvest_Summary") 'Add ErrHandler 2023/08/18
    If AbortTest Then Exit Function Else Resume Next 'Add ErrHandler 2023/08/18
End Function

'[20240503][All][Clyde] Function not used
Public Function Harvest_ReadValue(ByVal PatternName As String, ByVal HarvSrcKey As String) As Long
    On Error GoTo errHandler
    Dim funcName As String:: funcName = "Harvest_ReadValue"

''''PatternName = PatA;PatB
''''HarvSrcKey = HarvSrcKeyA;HarvSrcKeyB

'    Dim vsite As Variant
'    Dim i As Long
'    Dim j As Long
'
'    Dim PatternNameArr() As String
'    Dim FuseBlockNameArr() As String
'    Dim FuseCategoryNameArr() As String
'    Dim FailFlagOrValueArr() As String
'    Dim HarvSrcKeyArr() As String
'    Dim tempStr As String
'    Dim tempFuseCategoryNameArr() As String
'    Dim tempFailFlagOrValueArr() As String
'    Dim Slng_ReadFuseValue As New SiteLong
'    Dim tempSlng As New SiteLong
'
'    Dim sInstName As String
    
    TheExec.Datalog.WriteComment "Function not support!!"
    Exit Function
    
'    PatternNameArr = Split(PatternName, ";")
'    HarvSrcKeyArr = Split(HarvSrcKey, ";")
'
'    sInstName = TheExec.DataManager.instanceName
'    TheExec.Datalog.WriteComment "<" & sInstName & ">"
'
'
'    Dim temp_Harv_DSSC_DSP As New DSPWave
'    If UBound(PatternNameArr) = UBound(HarvSrcKeyArr) Then
'        tempStr = vbNullString
'        For i = 0 To UBound(PatternNameArr)
'''            Call Harvest_DigSrc(PatternNameArr(i), HarvSrcKeyArr(i), tempStr, temp_Harv_DSSC_DSP)
'''            Call Harvest_CreateDigSrc(PatternNameArr(i), HarvSrcKeyArr(i))
'
''            Dim tempHarvDsp As New DSPWave
''            Dim tempHarvDsp_SampleSize As Long
''            Call Harvest_CreateDigSrc(PatternNameArr(i), HarvSrcKeyArr(i), , tempHarvDsp, tempHarvDsp_SampleSize)
'
'        Next i
'    Else
'        TheExec.Datalog.WriteComment "<Error> Harvest_eFuse_Read: input is not match."
'    End If

errHandler:
    Call Print_Error_Message(Error_Info, "VBT_LIB_Common_AP", funcName) 'Add ErrHandler 2023/08/18
    If AbortTest Then Exit Function Else Resume Next
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
    
    Dim CoreIdx As Long
    
    
''''    Dim FinalHarvStr As String: FinalHarvStr = ""
    Dim StoreFlagIdx As Long: StoreFlagIdx = 0

    
    FuseBlockNameArr = Split(FuseBlockName, ";")
    FuseCategoryNameArr = Split(FuseCategoryName, ";")
    FailFlagOrValueArr = Split(FailFlagOrValue, ";")

    sInstName = theexec.DataManager.instancename
    theexec.Datalog.WriteComment "<" & sInstName & ">"
    
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
            For Each vsite In theexec.sites.Selected
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
                            theexec.Datalog.WriteComment OutputStrArr(j) & " site" & vsite & " => " & "Current Flag status is mapping Fuse Flag status"
                        ElseIf tempFuseSlng <> tempSlng Then
                            theexec.sites.item(vsite).FlagState("F_Harvest_PostCheck") = logicTrue
                            theexec.Datalog.WriteComment "Harvest Binswap, BinOut!!!"
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
                                theexec.Datalog.WriteComment tempFailFlagOrValueArr(j) & " site" & vsite & " => " & "Current Flag status is mapping Fuse Flag status"
                            ElseIf tempFuseSlng <> tempSlng Then
                                theexec.sites.item(vsite).FlagState("F_Harvest_PostCheck") = logicTrue
                                theexec.Datalog.WriteComment "Harvest Binswap, BinOut!!!"
                            End If
                            
                        ElseIf StarIdx < EndIdx Then
                            If UBound(tempFailFlagOrValueArr) > 32 Then
                                tempFuseSlng = CLng(mid(Svar_HexToBin(vsite), j + 1, 1))
                            Else
                                tempFuseSlng = CLng(mid(Svar_Temp(vsite), j + 1, 1))
                            End If
                            tempSlng = Harvest_GetAllSiteFlagState(tempFailFlagOrValueArr(j), 1, 0, 2)
                            If tempFuseSlng = tempSlng Then
                                theexec.Datalog.WriteComment tempFailFlagOrValueArr(j) & " site" & vsite & " => " & "Current Flag status is mapping Fuse Flag status"
                            ElseIf tempFuseSlng <> tempSlng Then
                                theexec.sites.item(vsite).FlagState("F_Harvest_PostCheck") = logicTrue
                                theexec.Datalog.WriteComment "Harvest Binswap, BinOut!!!"
                            End If
                            
                        Else
                            theexec.Datalog.WriteComment "Please Check The Flag Format"
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
                            theexec.Datalog.WriteComment tempFailFlagOrValueArr(j) & " site" & vsite & " => " & "Current Flag status is mapping Fuse Flag status"
                        ElseIf tempFuseSlng <> tempSlng Then
                            theexec.sites.item(vsite).FlagState("F_Harvest_PostCheck") = logicTrue
                            theexec.Datalog.WriteComment "Harvest Binswap, BinOut!!!"
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
