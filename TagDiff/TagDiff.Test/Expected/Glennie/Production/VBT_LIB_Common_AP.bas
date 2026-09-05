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

Exit Function 'Add ErrHandler 2023/08/18
errHandler: 'Add ErrHandler 2023/08/18
    Call Print_Error_Message(Error_Info, "VBT_LIB_Common_AP", "VBT_IEDA_Registry") 'Add ErrHandler 2023/08/18
    If AbortTest Then Exit Function Else Resume Next 'Add ErrHandler 2023/08/18
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
    
    sInstName = TheExec.DataManager.instancename
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
                    
                    If sInstName Like "*A1_Harvest_eFuse*" Then
                        If tempFailFlagOrValueArr(0) Like "*F_SOC_ANE_*" Then
                            If Sdbl_FuseValue = 2 Then
                                Sdbl_FuseValue = 4
                            ElseIf Sdbl_FuseValue = 4 Then
                                Sdbl_FuseValue = 2
                            ElseIf Sdbl_FuseValue = 32 Then
                                Sdbl_FuseValue = 64
                            ElseIf Sdbl_FuseValue = 64 Then
                                Sdbl_FuseValue = 32
                            ElseIf Sdbl_FuseValue = 512 Then
                                Sdbl_FuseValue = 1024
                            ElseIf Sdbl_FuseValue = 1024 Then
                                Sdbl_FuseValue = 512
                            ElseIf Sdbl_FuseValue = 8192 Then
                                Sdbl_FuseValue = 16384
                            ElseIf Sdbl_FuseValue = 16384 Then
                                Sdbl_FuseValue = 8192
                            End If
                        End If
                    End If
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
Exit Function 'Add ErrHandler 2023/08/18
errHandler: 'Add ErrHandler 2023/08/18
    Call Print_Error_Message(Error_Info, "VBT_LIB_Common_AP", "Harvest_eFuse_Write") 'Add ErrHandler 2023/08/18
    If AbortTest Then Exit Function Else Resume Next 'Add ErrHandler 2023/08/18
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
    
    sInstName = TheExec.DataManager.instancename
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
                Else
                    'FuseCategoryName = CFG_Condition_47
                    'FailFlagOrValue = F_ECPU_CORE3
                    For Each vsite In TheExec.sites.Selected
                        SDbl_ReadFuseValue(vsite) = field.DsscDecValue(vsite)
                        If UCase(FuseCategoryNameArr(i)) Like UCase("*harvesting_bin*") And SDbl_ReadFuseValue(vsite) = "4" Then
                            TheExec.sites.item(vsite).FlagState("F_IDS_CP1_HIGH_LEAKAGE") = logicTrue
                        End If
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
                        
                        If tempFailFlagOrValueArr(0) Like "*F_SOC_ANE_*" Then
                            If SDbl_ReadFuseValue = 2 Then
                                SDbl_ReadFuseValue = 4
                            ElseIf SDbl_ReadFuseValue = 4 Then
                                SDbl_ReadFuseValue = 2
                            ElseIf SDbl_ReadFuseValue = 32 Then
                                SDbl_ReadFuseValue = 64
                            ElseIf SDbl_ReadFuseValue = 64 Then
                                SDbl_ReadFuseValue = 32
                            ElseIf SDbl_ReadFuseValue = 512 Then
                                SDbl_ReadFuseValue = 1024
                            ElseIf SDbl_ReadFuseValue = 1024 Then
                                SDbl_ReadFuseValue = 512
                            ElseIf SDbl_ReadFuseValue = 8192 Then
                                SDbl_ReadFuseValue = 16384
                            ElseIf SDbl_ReadFuseValue = 16384 Then
                                SDbl_ReadFuseValue = 8192
                            End If
                        End If
                        
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

Exit Function 'Add ErrHandler 2023/08/18
errHandler: 'Add ErrHandler 2023/08/18
    Call Print_Error_Message(Error_Info, "VBT_LIB_Common_AP", "Harvest_eFuse_Read") 'Add ErrHandler 2023/08/18
    If AbortTest Then Exit Function Else Resume Next 'Add ErrHandler 2023/08/18
End Function


Public Function Harvest_SUBFLOW()
    
    Dim funcName As String:: funcName = "Harvest_SUBFLOW"
    
    On Error GoTo errHandler
    
    Dim site As Variant
    
    TheExec.sites.item(1).SiteVariableValue("HARVEST_PASS") = 1
    TheExec.sites.item(2).SiteVariableValue("HARVEST_PASS") = 2
    
Exit Function 'Add ErrHandler 2023/08/18
errHandler: 'Add ErrHandler 2023/08/18
    Call Print_Error_Message(Error_Info, "VBT_LIB_Common_AP", "Harvest_SUBFLOW") 'Add ErrHandler 2023/08/18
    If AbortTest Then Exit Function Else Resume Next 'Add ErrHandler 2023/08/18
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
    
    sInstName = TheExec.DataManager.instancename
    
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
            TheExec.flow.TestLimit resultVal:=1, lowVal:=0, hiVal:=0, Tname:=sInstName, ForceResults:=tlForceNone
            Exit Function
        End If
    Else
        TheExec.Datalog.WriteComment "Warning: Harv_FailCoreSum or Harv_FailCoreSumFlag did not have any information."
        TheExec.flow.TestLimit resultVal:=1, lowVal:=0, hiVal:=0, Tname:=sInstName, ForceResults:=tlForceNone
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
            
            TheExec.flow.TestLimit resultVal:=FirstFailCore_Slng, lowVal:=-1, hiVal:=FailFlagEndIdx, Tname:=BlockType + "_Harvested_Core", ForceResults:=tlForceNone
            TheExec.flow.TestLimit resultVal:=FailCoreCnt_Slng, lowVal:=0, hiVal:=lng_MaxFailCoreSum, Tname:=BlockType + "_Fail_Core_Count", ForceResults:=tlForceNone

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
                                Optional ByVal Harv_BinOutFailFlag As String, Optional CustHarv_GlobalFailFlag As String, _
                                Optional CustHarv_FailCoreSumFlag As String, Optional CustHarv_BinOutFailFlag As String) As Long
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
    
''' Create for multiple core group
    Dim StoreHavFailCoreDic As New Dictionary
    StoreHavFailCoreDic.RemoveAll
    Dim StoreHavAllCorePassDic As New Dictionary
    Dim multipleMaxFailCoreSum As Long: multipleMaxFailCoreSum = 0
    Dim multipleCodeGroupBool() As Boolean ': multipleCodeGroupBool() = False
    'New format For Harvest Sum Range Flag, ex: F_GUP_0To1
    Dim TempStoreSumRange() As String
    Dim TempBoolSumRange As Boolean: TempBoolSumRange = False
    Dim TempRangeLB As Long: TempRangeLB = 0
    Dim TempRangeUB As Long: TempRangeUB = 0
    
''    If UCase(theexec.DataManager.instanceName) = UCase("Harvest_Summary_After_Harvesting_Descision") Then Stop
    With TheExec.Datalog
        .Setup.Shared.ascii.Columns.EnableCustomWidths = True
        .Setup.Shared.ascii.Columns.Parametric.TestName.Width = 30
        .Setup.Shared.ascii.Columns.Parametric.measured.Width = 16
        .Setup.Shared.ascii.Columns.Functional.TestName.Width = 30
        .Setup.Shared.ascii.Columns.Functional.Pattern.Width = 100
        .ApplySetup
    End With

    sInstName = TheExec.DataManager.instancename
    
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
        TheExec.flow.TestLimit resultVal:=1, lowVal:=0, hiVal:=0, Tname:=sInstName, ForceResults:=tlForceNone
        Exit Function
    End If
    
    ReDim multipleCodeGroupBool(UBound(Block_GlobalFailFlagArr))
    ReDim Block_FailCoreCnt_Slng(UBound(Block_BinOutFailFlagArr))
    
    
    
    '''''Judge the multiple code group'''''
    Dim multipleCodeKeyWord() As String
    ReDim multipleCodeKeyWord(UBound(Block_FailCoreSumFlagArr))
    
    For i = 0 To UBound(Block_FailCoreSumFlagArr) - 1
        If InStr(Block_FailCoreSumFlagArr(i), "_") <> 0 Then
            multipleCodeKeyWord(i) = Split(Block_FailCoreSumFlagArr(i), "_")(1)
        End If
        If i = 0 Then
            multipleCodeGroupBool(i) = False
        Else
            If multipleCodeKeyWord(i) = multipleCodeKeyWord(i - 1) Then
                multipleCodeGroupBool(i) = True
                multipleCodeGroupBool(i - 1) = True
            Else
                multipleCodeGroupBool(i) = False
            End If
        End If
    Next i
    
    ''''''''''''''''''''''''''''''''''''''
    
    
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
                lng_MaxFailCoreSum = Application.WorksheetFunction.Max(lng_MaxFailCoreSum, TempStoreSumRange(UBound(TempStoreSumRange)))
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
        StoreHavFailCoreDic.RemoveAll
        StoreHavAllCorePassDic.RemoveAll
        If Harv_BinOutFailFlag <> "" Then
            For FlagCnt = 0 To HarvestFlagArr_Num
                Tname_str = HarvestFlagArr(FlagCnt)
                ''''Harv_BinOutFailFlag = F_GFX_HARV_SUM|F_ECPU_CORE_SUM
                If InStr(Harv_BinOutFailFlag, HarvestFlagArr(FlagCnt)) <> 0 Then
                    For i = 0 To UBound(Block_BinOutFailFlagArr)
                        If UCase(HarvestFlagArr(FlagCnt)) = UCase(Block_BinOutFailFlagArr(i)) Then
                            If multipleCodeGroupBool(i) = False Then
                                If TheExec.sites.item(vsite).FlagState(HarvestFlagArr(FlagCnt)) = logicTrue Then
                                    TempHarvestStr(FlagCnt) = 1
                                Else
                                    TempHarvestStr(FlagCnt) = 0
                                End If
                                TheExec.flow.TestLimit resultVal:=Block_FailCoreCnt_Slng(i)(vsite), lowVal:=0, hiVal:=0, Tname:=Tname_str, ForceResults:=tlForceNone
                                'glb_TestInstance = theexec.DataManager.instancename
                                'theexec.Datalog.WriteComment glb_TestInstance & ("---->Forcenonecase")
    ''                            TheExec.Flow.TestLimit resultVal:=TempHarvestStr(FlagCnt), lowVal:=0, hiVal:=0, Tname:=Tname_Str, ForceResults:=tlForceNone
                            Else
                                If StoreHavFailCoreDic.Exists(Block_BinOutFailFlagArr(i)) = False Then
                                    StoreHavFailCoreDic.Add (Block_BinOutFailFlagArr(i)), Block_FailCoreCnt_Slng(i)(vsite)
                                Else
                                    Dim tempCountFailCode As New SiteLong
                                    tempCountFailCode(vsite) = StoreHavFailCoreDic(Block_BinOutFailFlagArr(i)) + Block_FailCoreCnt_Slng(i)(vsite)
                                    StoreHavFailCoreDic.Remove (Block_BinOutFailFlagArr(i))
                                    StoreHavFailCoreDic.Add (Block_BinOutFailFlagArr(i)), tempCountFailCode(vsite)
                                End If
                            End If
                        End If
                    Next i
                    If StoreHavFailCoreDic.Exists(HarvestFlagArr(FlagCnt)) = True Then
                        TheExec.flow.TestLimit resultVal:=tempCountFailCode(vsite), lowVal:=0, hiVal:=0, Tname:=Tname_str, ForceResults:=tlForceNone
                        'glb_TestInstance = theexec.DataManager.instancename
                        'theexec.Datalog.WriteComment glb_TestInstance & ("---->Forcenonecase")
                    End If
                ''''Harv_AllCorePassFlag = F_GFX_ALL_CORE_PASS|F_ECPU_ALL_CORE_PASS
                ElseIf InStr(Harv_AllCorePassFlag, HarvestFlagArr(FlagCnt)) <> 0 Then
                    For i = 0 To UBound(Block_AllCorePassFlagArr)
                        If UCase(HarvestFlagArr(FlagCnt)) = UCase(Block_AllCorePassFlagArr(i)) Then
''                            If UCase(Block_AllCorePassFlagArr(i)) = "F_ECPU_ALL_CORE_PASS" Then Stop
                            If TheExec.sites.item(vsite).FlagState(HarvestFlagArr(FlagCnt)) = logicTrue Then
                                TempHarvestStr(FlagCnt) = 1
                            Else
                                TempHarvestStr(FlagCnt) = 0
                            End If
''                            TheExec.Flow.TestLimit resultVal:=TempHarvestStr(FlagCnt), Tname:=Tname_Str, ForceResults:=tlForceNone
                            If multipleCodeGroupBool(i) = False Then
                                TheExec.flow.TestLimit resultVal:=TempHarvestStr(FlagCnt), lowVal:=1, hiVal:=1, Tname:=Tname_str, ForceResults:=tlForceNone
                                'glb_TestInstance = theexec.DataManager.instancename
                                'theexec.Datalog.WriteComment glb_TestInstance & ("---->Forcenonecase")
                            Else
                                If StoreHavAllCorePassDic.Exists(Block_AllCorePassFlagArr(i)) = False Then
                                    StoreHavAllCorePassDic.Add Block_AllCorePassFlagArr(i), TempHarvestStr(FlagCnt)
                                Else
                                    
                                    TheExec.flow.TestLimit resultVal:=StoreHavAllCorePassDic(Block_AllCorePassFlagArr(i)), lowVal:=1, hiVal:=1, Tname:=Tname_str, ForceResults:=tlForceNone
                                    'glb_TestInstance = theexec.DataManager.instancename
                                    'theexec.Datalog.WriteComment glb_TestInstance & ("---->Forcenonecase")
                                End If
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
                    TheExec.flow.TestLimit resultVal:=TempHarvestStr(FlagCnt), lowVal:=0, hiVal:=0, Tname:=Tname_str, ForceResults:=tlForceNone
                    
                End If
            Next FlagCnt
        End If
    Next vsite
    
    With TheExec.Datalog
        .Setup.Shared.ascii.Columns.EnableCustomWidths = True
        .Setup.Shared.ascii.Columns.Parametric.TestName.Width = 110
        .Setup.Shared.ascii.Columns.Parametric.measured.Width = 16
        .Setup.Shared.ascii.Columns.Functional.TestName.Width = 110
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

Public Function HarvFlagPrinting() As Long

    Dim InstName As String
    Dim site As Variant
    Dim GfxFlagcnt As Long: GfxFlagcnt = 0
    Dim PCPUFlagcnt As Long: PCPUFlagcnt = 0
    Dim finalCnt() As Long
    ReDim finalCnt(1)
    Dim FlagArray() As String
    ReDim FlagArray(1)
    Dim i As Integer
    
    InstName = TheExec.DataManager.instancename
    FlagArray(0) = "F_GFX_HARV_8_0_SUM_1"
    FlagArray(1) = "F_GFX_HARV_17_9_SUM_1"
    
    For Each site In TheExec.sites

        If UCase(InstName) Like UCase("Gfx_Harvest_Fail") Then
            TheExec.flow.TestLimit resultVal:=TheExec.sites.item(site).FlagState("F_GFX_HARV_8_0_SUM"), lowVal:=0, hiVal:=0, Tname:="F_GFX_HARV_8_0_SUM", ForceResults:=tlForceNone
            TheExec.flow.TestLimit resultVal:=TheExec.sites.item(site).FlagState("F_GFX_HARV_17_9_SUM"), lowVal:=0, hiVal:=0, Tname:="F_GFX_HARV_17_9_SUM", ForceResults:=tlForceNone
           
        ElseIf UCase(InstName) Like UCase("Gfx_Harvest") Then
            TheExec.flow.TestLimit resultVal:=TheExec.sites.item(site).FlagState("F_GFX_HARV_8_0_SUM_1"), lowVal:=0, hiVal:=0, Tname:="F_GFX_HARV_8_0_SUM_1", ForceResults:=tlForceNone
            TheExec.flow.TestLimit resultVal:=TheExec.sites.item(site).FlagState("F_GFX_HARV_17_9_SUM_1"), lowVal:=0, hiVal:=0, Tname:="F_GFX_HARV_17_9_SUM_1", ForceResults:=tlForceNone
        
        ElseIf UCase(InstName) Like UCase("PCPU_Harvest_Fail") Then
            TheExec.flow.TestLimit resultVal:=TheExec.sites.item(site).FlagState("F_PCPU_CORE_5_0_SUM"), lowVal:=0, hiVal:=0, Tname:="F_PCPU_CORE_5_0_SUM", ForceResults:=tlForceNone
            TheExec.flow.TestLimit resultVal:=TheExec.sites.item(site).FlagState("F_PCPU_MAJ"), lowVal:=0, hiVal:=0, Tname:="F_PCPU_MAJ", ForceResults:=tlForceNone
            TheExec.flow.TestLimit resultVal:=TheExec.sites.item(site).FlagState("F_PCPU_CPM"), lowVal:=0, hiVal:=0, Tname:="F_PCPU_CPM", ForceResults:=tlForceNone
            TheExec.flow.TestLimit resultVal:=TheExec.sites.item(site).FlagState("F_PCPU_OTHER"), lowVal:=0, hiVal:=0, Tname:="F_PCPU_OTHER", ForceResults:=tlForceNone
    
        ElseIf UCase(InstName) Like UCase("PCPU_Harvest") Then
            TheExec.flow.TestLimit resultVal:=TheExec.sites.item(site).FlagState("F_PCPU_CORE_5_0_SUM_1"), lowVal:=0, hiVal:=0, Tname:="F_PCPU_CORE_5_0_SUM_1", ForceResults:=tlForceNone
    
        Else
            For i = 0 To 1
                If TheExec.sites.item(site).FlagState(FlagArray(i)) = logicTrue Then
                    GfxFlagcnt = 1
                Else
                    GfxFlagcnt = 0
                End If
                
                If TheExec.sites.item(site).FlagState("F_PCPU_CORE_5_0_SUM_1") = logicTrue Then
                    PCPUFlagcnt = 1
                Else
                    PCPUFlagcnt = 0
                End If
                finalCnt(i) = GfxFlagcnt + PCPUFlagcnt
            Next i
            
            TheExec.flow.TestLimit resultVal:=finalCnt(0), lowVal:=0, hiVal:=1, Tname:="F_GFX_HARV_8_0_SUM_1 + F_PCPU_CORE_5_0_SUM_1", ForceResults:=tlForceNone
            TheExec.flow.TestLimit resultVal:=finalCnt(1), lowVal:=0, hiVal:=1, Tname:="F_GFX_HARV_17_9_SUM_1 + F_PCPU_CORE_5_0_SUM_1", ForceResults:=tlForceNone
        
        
        End If
        
    Next site
    

End Function

Public Function Harvest_CustomFlagJudge(CustHarv_GlobalFailFlag As String, CustHarv_FailCoreSumFlag As String, CustHarv_BinOutFailFlag As String, lng_MaxFailCoreSum As Long) As Long

On Error GoTo errHandler

Dim funcName As String: funcName = "Harvest_CustomFlagJudge"

    Dim CustMultiGlbFlagAry() As String
    Dim CustGlbFailFlagAry() As String
    Dim CustCaseStr As String
    Dim TempGlbFailFlagAry() As String
    Dim CustFailFlagDefaultName() As String
    Dim CustFailFlagStartIdx As Long
    Dim CustFailFlagEndIdx As Long
    Dim CustFailCoreCnt_Slng As New SiteLong
    Dim CustHarv_FailCoreSumFlagAry() As String
    Dim CustHarv_BinOutFailFlagAry() As String
    
    Dim Order As Long
    Dim TempVal() As New SiteLong
    Dim ActualVal As Long
    Dim i, j, k As Long
    Dim vsite As Variant
        
    Dim FailCoreSumFlagAry() As String
            
       CustMultiGlbFlagAry = Split(CustHarv_GlobalFailFlag, "|")
       CustHarv_FailCoreSumFlagAry = Split(CustHarv_FailCoreSumFlag, "|")
       CustHarv_BinOutFailFlagAry = Split(CustHarv_BinOutFailFlag, "|")
       
For i = 0 To UBound(CustMultiGlbFlagAry)

    CustFailCoreCnt_Slng = 0
    If InStr(1, CustMultiGlbFlagAry(i), "SUM") <> 0 Then
        CustGlbFailFlagAry = Split(CustMultiGlbFlagAry(i), "(")
        CustCaseStr = CustGlbFailFlagAry(0) & CustGlbFailFlagAry(1)
        CustGlbFailFlagAry(UBound(CustGlbFailFlagAry)) = Replace(CustGlbFailFlagAry(UBound(CustGlbFailFlagAry)), "))", "")
        TempGlbFailFlagAry = Split(CustGlbFailFlagAry(UBound(CustGlbFailFlagAry)), ",")
    Else
        CustGlbFailFlagAry = Split(CustMultiGlbFlagAry(i), "(")
        CustCaseStr = CustGlbFailFlagAry(0)
        CustGlbFailFlagAry(1) = Replace(CustGlbFailFlagAry(1), ")", "")
        TempGlbFailFlagAry = Split(CustGlbFailFlagAry(1), ",")
    End If
    
    FailCoreSumFlagAry = Split(CustHarv_FailCoreSumFlagAry(i), "+")
    
    ReDim CustFailFlagDefaultName(UBound(TempGlbFailFlagAry))
    ReDim TempVal((UBound(TempGlbFailFlagAry)))
    
    Select Case UCase(CustCaseStr)
        
        Case "SUMOR":
        
            For Each vsite In TheExec.sites.Active
            
                CustFailFlagDefaultName(0) = mid(TempGlbFailFlagAry(0), 1, InStr(TempGlbFailFlagAry(0), "[") - 1)
                CustFailFlagDefaultName(1) = mid(TempGlbFailFlagAry(1), 1, InStr(TempGlbFailFlagAry(1), "[") - 1)
                CustFailFlagStartIdx = CLng(mid(TempGlbFailFlagAry(0), InStr(TempGlbFailFlagAry(0), "[") + 1, InStr(TempGlbFailFlagAry(0), ":") - InStr(TempGlbFailFlagAry(j), "[") - 1))
                CustFailFlagEndIdx = CLng(mid(TempGlbFailFlagAry(0), InStr(TempGlbFailFlagAry(0), ":") + 1, InStr(TempGlbFailFlagAry(0), "]") - InStr(TempGlbFailFlagAry(j), ":") - 1))
    
                If CustFailFlagStartIdx > CustFailFlagEndIdx Then
                    Order = -1
                Else
                    Order = 1
                End If
    
                For k = CustFailFlagStartIdx To CustFailFlagEndIdx Step Order
                    If TheExec.sites.item(vsite).FlagState(CustFailFlagDefaultName(0) & CStr(k)) = logicTrue Then
                        TempVal(0)(vsite) = 1
    
                    ElseIf TheExec.sites.item(vsite).FlagState(CustFailFlagDefaultName(0) & CStr(k)) = logicFalse Then
                        TempVal(0)(vsite) = 0
    
                    ElseIf TheExec.sites.item(vsite).FlagState(CustFailFlagDefaultName(0) & CStr(k)) = logicClear Then
                        TheExec.Datalog.WriteComment "Please confrim the flag status"
                    End If
                    
                    If TheExec.sites.item(vsite).FlagState(CustFailFlagDefaultName(1) & CStr(k)) = logicTrue Then
                        TempVal(1)(vsite) = 1
    
                    ElseIf TheExec.sites.item(vsite).FlagState(CustFailFlagDefaultName(1) & CStr(k)) = logicFalse Then
                        TempVal(1)(vsite) = 0
    
                    ElseIf TheExec.sites.item(vsite).FlagState(CustFailFlagDefaultName(1) & CStr(k)) = logicClear Then
                        TheExec.Datalog.WriteComment "Please confrim the flag status"
                    End If
                    
                    ActualVal = TempVal(0).BitwiseOr(TempVal(1))
                    
                    If ActualVal = 1 Then
                        CustFailCoreCnt_Slng(vsite) = CustFailCoreCnt_Slng(vsite) + 1
                    End If
                    
                Next k
                
                If CustFailCoreCnt_Slng > lng_MaxFailCoreSum Then
                    TheExec.sites.item(vsite).FlagState(CustHarv_BinOutFailFlagAry(i)) = logicTrue
                       
                ElseIf CustFailCoreCnt_Slng = 0 Then
                    For k = 0 To UBound(FailCoreSumFlagAry)
                        If right(FailCoreSumFlagAry(k), 1) = 0 Then
                            TheExec.sites.item(vsite).FlagState(FailCoreSumFlagAry(k)) = logicTrue
                        End If
                    Next k
                 
                Else
                    For k = 0 To UBound(FailCoreSumFlagAry)
                        If right(FailCoreSumFlagAry(k), 1) = CustFailCoreCnt_Slng Then
                            TheExec.sites.item(vsite).FlagState(FailCoreSumFlagAry(k)) = logicTrue
                        End If
                    Next k
                    
                End If
            Next vsite
            
    End Select
Next i
   
   
Exit Function
errHandler:
    TheExec.Datalog.WriteComment "<Error> " + funcName + ":: please check it out."
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
    
    HarvFlagArr = Split(HarvFlag, ",")
    
    sInstName = TheExec.DataManager.instancename
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
            TheExec.flow.TestLimit resultVal:=TempHarvestStr, lowVal:=-1, hiVal:=1, Tname:=Tname_str
            
        Next FlagCnt
    Next vsite
    If HarvMaxCore > 0 Then
        For Each vsite In TheExec.sites
            TheExec.flow.TestLimit resultVal:=FlagTrueCnt(vsite), lowVal:=0, hiVal:=HarvMaxCore, Tname:=sInstName & "_Core_Over_Check"
        Next vsite
    End If
Exit Function

errHandler:
    TheExec.Datalog.WriteComment "<Error> " + funcName + ":: please check it out."
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

    sInstName = TheExec.DataManager.instancename
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
