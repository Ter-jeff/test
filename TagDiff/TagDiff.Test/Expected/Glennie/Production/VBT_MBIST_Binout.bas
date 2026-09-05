Attribute VB_Name = "VBT_MBIST_Binout"
Option Explicit

Public Function SetFlag_PassFail_for_MbistBlocks(site As Variant, m_testName As String, m_Block As String, m_Pmode As String, _
                                                 m_groupName As String, bistType As String, testResult As String) As Long
'SWLIN 2017/01/18
On Error GoTo errHandler

Dim funcName As String:: funcName = "SetFlag_PassFail_for_MbistBlocks"
Dim m_BXX As String
Dim CurrentStatus_FailFlag As New SiteBoolean
Dim BinOutFlag As String
Dim tmpGroupName As String
Dim printflag As Boolean:: printflag = True
        'm_groupName = Replace(m_groupName, "_HV", "") 'remove HV/LV key word in group name
        'm_groupName = Replace(m_groupName, "_LV", "")
        
        If UCase(m_testName) Like "*_HV" Or UCase(m_testName) Like "*_MHV" Then
            'm_BXX = m_BXX & "_HV"
            tmpGroupName = m_groupName & "_HV"
        ElseIf UCase(m_testName) Like "*_LV" Or UCase(m_testName) Like "*_MLV" Then
            'm_BXX = m_BXX & "_LV"
            tmpGroupName = m_groupName & "_LV"
        ElseIf UCase(m_testName) Like "*_NV" Then
            tmpGroupName = m_groupName & "_NV"
        'ElseIf UCase(m_testName) Like "_NV" Then
'        ElseIf UCase(m_testName) Like "_MHV" Then
'            m_BXX = m_BXX & "_MHV"
'        ElseIf UCase(m_testName) Like "_MLV" Then
'            m_BXX = m_BXX & "_MLV"
        End If
        
        'Mbist Flag
        If UCase(m_testName) Like "*CPUMBIST*" Then
            BinOutFlag = "F_CPU_Mbist_" & m_Block & "_" & m_Pmode & "_" & tmpGroupName
        ElseIf UCase(m_testName) Like "*SOCMBIST*" Then
            BinOutFlag = "F_SOC_Mbist_" & m_Block & "_" & m_Pmode & "_" & tmpGroupName
        End If
        'BinOutFlag = "F_" & m_Pmode & "_" & m_BXX
        
        'CurrentStatus_FailFlag(Site) = TheExec.Sites.Item(Site).FlagState(m_BXX)
        
        'If UCase(m_testName) Like UCase("*Vmargin*") Then
            If UCase(testResult) = UCase("fail") Then 'if itme fail then set fail flag is true
                If TheExec.sites.item(site).FlagState(BinOutFlag) = logicClear Then TheExec.sites.item(site).FlagState(BinOutFlag) = logicTrue
                TheExec.sites.item(site).FlagState(BinOutFlag) = TheExec.sites.item(site).FlagState(BinOutFlag) Or logicTrue
                If printflag Then TheExec.Datalog.WriteComment "Site: " & site & ", Set BinOutFlag: " & CStr(BinOutFlag) & ", " & CStr(TheExec.sites.item(site).FlagState(BinOutFlag))
            ElseIf UCase(testResult) = UCase("pass") Then 'if itme pass then set fail flag is false
                If TheExec.sites.item(site).FlagState(BinOutFlag) = logicClear Then TheExec.sites.item(site).FlagState(BinOutFlag) = logicFalse
                TheExec.sites.item(site).FlagState(BinOutFlag) = TheExec.sites.item(site).FlagState(BinOutFlag) Or logicFalse
                If printflag Then TheExec.Datalog.WriteComment "Site: " & site & ", Set BinOutFlag: " & CStr(BinOutFlag) & ", " & CStr(TheExec.sites.item(site).FlagState(BinOutFlag))
            End If
        'End If
    Exit Function
errHandler:
    TheExec.Datalog.WriteComment "<Error> " + funcName + ":: please check it out."
    If AbortTest Then Exit Function Else Resume Next
    
End Function

Public Function CPU_MBist_BinOut() As Long  '''2017/1/12 scpan for osprey

On Error GoTo errHandler
    Dim funcName As String:: funcName = "CPU_MBist_BinOut"

    Dim i As Integer
    Dim str1 As String
    Dim m_BXX As String
    
    Dim Block_Fail_count As Integer
    Dim site As Variant
    
    
    For Each site In TheExec.sites
      
      Block_Fail_count = 0

      For i = 0 To 15 '''for 16 Block 2017/1/10 sean pan Osprey
            
        str1 = "00" & Trim(str(i))
        str1 = right(str1, 2)
              
        m_BXX = "CPU_MBIST_B" & str1
                
              
        If TheExec.sites.item(site).FlagState(m_BXX) = logicTrue Then
        
          Block_Fail_count = Block_Fail_count + 1
          
          TheExec.Datalog.WriteComment "Site: " + Trim(CStr(site)) + ", " + CStr(m_BXX) + ": " + "Fail"
      
        Else
          
          TheExec.Datalog.WriteComment "Site: " + Trim(CStr(site)) + ", " + CStr(m_BXX) + ": " + "Pass"
        
        End If
            
      Next
      
      Select Case Block_Fail_count
            Case 0
                TheExec.sites.item(site).FlagState("F_16CPU_pass") = logicTrue
            Case 1
                TheExec.sites.item(site).FlagState("F_15CPU_pass") = logicTrue
            Case 2
                TheExec.sites.item(site).FlagState("F_14CPU_pass") = logicTrue
            Case 3
                TheExec.sites.item(site).FlagState("F_13CPU_pass") = logicTrue
            Case 4
                TheExec.sites.item(site).FlagState("F_12CPU_pass") = logicTrue
            Case 5
                TheExec.sites.item(site).FlagState("F_11CPU_pass") = logicTrue
            Case 6
                TheExec.sites.item(site).FlagState("F_10CPU_pass") = logicTrue
            Case 7
                TheExec.sites.item(site).FlagState("F_09CPU_pass") = logicTrue
            Case 8
                TheExec.sites.item(site).FlagState("F_08CPU_pass") = logicTrue
            Case 9
                TheExec.sites.item(site).FlagState("F_07CPU_pass") = logicTrue
            Case 10
                TheExec.sites.item(site).FlagState("F_06CPU_pass") = logicTrue
            Case 11
                TheExec.sites.item(site).FlagState("F_05CPU_pass") = logicTrue
            Case 12
                TheExec.sites.item(site).FlagState("F_04CPU_pass") = logicTrue
            Case 13
                TheExec.sites.item(site).FlagState("F_03CPU_pass") = logicTrue
            Case 14
                TheExec.sites.item(site).FlagState("F_02CPU_pass") = logicTrue
            Case 15
                TheExec.sites.item(site).FlagState("F_01CPU_pass") = logicTrue
            Case 16
                TheExec.sites.item(site).FlagState("F_00CPU_pass") = logicTrue
            
      End Select
    
    Next site

    CPU_MBist_BinOut = 1

Exit Function

errHandler:
    TheExec.Datalog.WriteComment "<Error> " + funcName + ":: please check it out."
    If AbortTest Then Exit Function Else Resume Next
     
End Function



Public Function SOC_MBist_BinOut() As Long  '''2017/1/12 scpan for osprey

On Error GoTo errHandler
    Dim funcName As String:: funcName = "SOC_MBist_BinOut"

    Dim i As Integer
    Dim str1 As String
    Dim m_BXX As String
    Dim site As Variant
    
    Dim Block_Fail_count As Integer
    
    
    For Each site In TheExec.sites
      
      Block_Fail_count = 0

      For i = 0 To 15 '''for 16 Block 2017/1/10 sean pan Osprey
            
        str1 = "00" & Trim(str(i))
        str1 = right(str1, 2)
              
        m_BXX = "SOC_MBIST_B" & str1
                
              
        If TheExec.sites.item(site).FlagState(m_BXX) = logicTrue Then
        
          Block_Fail_count = Block_Fail_count + 1
          TheExec.Datalog.WriteComment "Site: " + Trim(CStr(site)) + ", " + CStr(m_BXX) + ": " + "Fail"
          
        Else
        
          TheExec.Datalog.WriteComment "Site: " + Trim(CStr(site)) + ", " + CStr(m_BXX) + ": " + "Pass"
          
        End If
            
      Next
      
      Select Case Block_Fail_count
            Case 0
                TheExec.sites.item(site).FlagState("F_16SOC_pass") = logicTrue
            Case 1
                TheExec.sites.item(site).FlagState("F_15SOC_pass") = logicTrue
            Case 2
                TheExec.sites.item(site).FlagState("F_14SOC_pass") = logicTrue
            Case 3
                TheExec.sites.item(site).FlagState("F_13SOC_pass") = logicTrue
            Case 4
                TheExec.sites.item(site).FlagState("F_12SOC_pass") = logicTrue
            Case 5
                TheExec.sites.item(site).FlagState("F_11SOC_pass") = logicTrue
            Case 6
                TheExec.sites.item(site).FlagState("F_10SOC_pass") = logicTrue
            Case 7
                TheExec.sites.item(site).FlagState("F_09SOC_pass") = logicTrue
            Case 8
                TheExec.sites.item(site).FlagState("F_08SOC_pass") = logicTrue
            Case 9
                TheExec.sites.item(site).FlagState("F_07SOC_pass") = logicTrue
            Case 10
                TheExec.sites.item(site).FlagState("F_06SOC_pass") = logicTrue
            Case 11
                TheExec.sites.item(site).FlagState("F_05SOC_pass") = logicTrue
            Case 12
                TheExec.sites.item(site).FlagState("F_04SOC_pass") = logicTrue
            Case 13
                TheExec.sites.item(site).FlagState("F_03SOC_pass") = logicTrue
            Case 14
                TheExec.sites.item(site).FlagState("F_02SOC_pass") = logicTrue
            Case 15
                TheExec.sites.item(site).FlagState("F_01SOC_pass") = logicTrue
            Case 16
                TheExec.sites.item(site).FlagState("F_00SOC_pass") = logicTrue
            
      End Select
    
    Next site

    SOC_MBist_BinOut = 1

Exit Function

errHandler:
    TheExec.Datalog.WriteComment "<Error> " + funcName + ":: please check it out."
    If AbortTest Then Exit Function Else Resume Next
     
End Function


Public Function MBist_BinOutFlag(CPU_BlockCount As Long, SOCX_BlockCount As Long, M9X_BlockCount As Long) As Long '''2017/1/18 SWLin for osprey

On Error GoTo errHandler
    Dim funcName As String:: funcName = "MBist_BinOutFlag"

    Dim i, j As Long
    Dim str1 As String
    Dim str2 As String
    Dim m_BXX As String
    Dim site As Variant
    
    Dim Block_Fail_count As Long
    Dim Block_Pass_count As Long
     
    Dim F_noTest_allBlocks As New SiteBoolean
    Dim F_noTest_temp As New SiteBoolean
    
    Dim MbistCateCount As Long:: MbistCateCount = 3
    Dim MbistCate As String
    Dim MbistCate_BlockName As String
    Dim MbistBlockCount As Long
    Dim strPassFailResult As String
    Dim varPassFailResult As New SiteDouble
    Dim BinOutFlag As String
    
    TheExec.Datalog.WriteComment "---------------------------------------------------------"
    TheExec.Datalog.WriteComment "Start to Print Out Pass/Fail Result for each Mbist Blocks"
    TheExec.Datalog.WriteComment "---------------------------------------------------------"
    
        For j = 0 To MbistCateCount - 1

            Select Case j
                Case 0
                    MbistCate = "CPU_MBist_B": MbistBlockCount = CPU_BlockCount: MbistCate_BlockName = "CPU"
                Case 1
                    MbistCate = "SOC_MBist_SOC": MbistBlockCount = SOCX_BlockCount: MbistCate_BlockName = "SOC"
                Case 2
                    MbistCate = "SOC_MBist_M9X": MbistBlockCount = M9X_BlockCount: MbistCate_BlockName = "M9X"
            End Select
            
            For Each site In TheExec.sites
                Block_Pass_count = 0
                BinOutFlag = ""
                For i = 0 To MbistBlockCount - 1
                    strPassFailResult = "N/A" ' initialize
                    varPassFailResult = 2
                    'compose string for block name
                    str1 = "00" & Trim(str(i))
                    If MbistCate Like "*CPU*" Then
                        str1 = right(str1, 2)
                    ElseIf MbistCate Like "*SOC*" Then
                        str1 = right(str1, 1)
                    End If
                    m_BXX = MbistCate & str1 'm_BXX = CPU_MBist_BXX / SOC_MBist_SOCX /SOC_MBist_M9XX
                    
                    If TheExec.sites.item(site).FlagState(m_BXX & "_HV") = logicTrue And TheExec.sites.item(site).FlagState(m_BXX & "_LV") = logicTrue Then
                        strPassFailResult = "Fail"
                        varPassFailResult = 2
                    ElseIf (TheExec.sites.item(site).FlagState(m_BXX & "_HV") = logicClear And TheExec.sites.item(site).FlagState(m_BXX & "_LV") = logicTrue _
                        Or TheExec.sites.item(site).FlagState(m_BXX & "_HV") = logicTrue And TheExec.sites.item(site).FlagState(m_BXX & "_LV") = logicClear) Then
                        strPassFailResult = "Fail"
                        varPassFailResult = 2
                    ElseIf TheExec.sites.item(site).FlagState(m_BXX & "_HV") = logicFalse Or TheExec.sites.item(site).FlagState(m_BXX & "_LV") = logicFalse Then
                        Block_Pass_count = Block_Pass_count + 1
                        strPassFailResult = "Pass"
                        varPassFailResult = 1
                    ElseIf TheExec.sites.item(site).FlagState(m_BXX & "_HV") = logicClear And TheExec.sites.item(site).FlagState(m_BXX & "_LV") = logicClear Then
                        strPassFailResult = "N/A"
                        varPassFailResult = 0
                    End If
                    
                    TheExec.Datalog.WriteComment "Site: " + Trim(CStr(site)) + ", " + CStr(m_BXX) + ": " + strPassFailResult
                    TheExec.flow.TestLimit varPassFailResult, 1, 1, Tname:=CStr(m_BXX)
                Next i

                    str2 = "00" & Trim(Block_Pass_count)
                    str2 = right(str2, 2)
                    BinOutFlag = "F_" & str2 & MbistCate_BlockName & "_pass" 'BinOutFlag = F_XXCPU_pass / F_XXSOC_pass / F_XXM9X_pass
                    TheExec.sites.item(site).FlagState(BinOutFlag) = logicTrue
                    TheExec.Datalog.WriteComment "Site: " + Trim(CStr(site)) + ", " + MbistCate_BlockName + ": Total Pass-Blocks: " + CStr(Block_Pass_count) + ", Turn on flag: " + BinOutFlag + ": " + CStr(TheExec.sites.item(site).FlagState(BinOutFlag))
                    TheExec.Datalog.WriteComment ""
            Next site
        Next j
        
    'MBist_BinOutFlag_PassRate = 1  'Delete, Leon Li, 20190629
    TheExec.Datalog.WriteComment "---------------------------------------------------------"
    TheExec.Datalog.WriteComment "---------------------------------------------------------"
    
Exit Function

errHandler:
    TheExec.Datalog.WriteComment "<Error> " + funcName + ":: please check it out."
    If AbortTest Then Exit Function Else Resume Next
     
End Function


Public Function BinOutFlag_BlockPassRate(CPU_BlockCount As Long, SOCX_BlockCount As Long, M9X_BlockCount As Long, BypassTestType As String) As Long '''2017/1/18 SWLin for osprey

On Error GoTo errHandler
    Dim funcName As String:: funcName = "BinOutFlag_BlockPassRate"

    Dim i, j, k As Long
    Dim str1 As String
    Dim str2 As String
    Dim str3 As String
    Dim m_BXX As String
    Dim site As Variant
    
    Dim Block_Fail_count As Long
    Dim Block_Pass_count As Long
     
    Dim F_TestTypePass As New SiteBoolean
    Dim F_BlockPass As New SiteBoolean
    Dim TestTypeCount As Long:: TestTypeCount = 3 '(TD/SA/Mbist)
    Dim TestTypeName As String
    Dim MbistCateCount As Long:: MbistCateCount = 3
    Dim MbistCate As String
    Dim MbistCate_BlockName As String
    Dim MbistBlockCount As Long
    Dim strPassFailResult As String
    Dim varPassFailResult As New SiteDouble
    Dim BinOutFlag As String
    Dim BlockResultFalg As String
    Dim BypassTestType_ary() As String
    Dim BypassType As Variant
    Dim BinOutCateName As String
    Dim SOC_Total_Block_Pass_count As New SiteLong
    Dim SOC_Total_Pass_Flag As String

    TheExec.Datalog.WriteComment "---------------------------------------------------------"
    TheExec.Datalog.WriteComment "  Start to Print Out Pass/Fail Result for each Blocks    "
    TheExec.Datalog.WriteComment "---------------------------------------------------------"
    
        BypassTestType_ary() = Split(BypassTestType, ",")
        
        For Each site In TheExec.sites
            SOC_Total_Block_Pass_count = 0
        Next site
    
        For j = 0 To MbistCateCount - 1
            Select Case j
                Case 0
                    MbistCate = "CPU": MbistBlockCount = CPU_BlockCount: MbistCate_BlockName = "B": BinOutCateName = "CPU_BXX"
                Case 1
                    MbistCate = "SOC": MbistBlockCount = SOCX_BlockCount: MbistCate_BlockName = "SOC": BinOutCateName = "SOC_SOCX"
                Case 2
                    MbistCate = "SOC": MbistBlockCount = M9X_BlockCount: MbistCate_BlockName = "M9X": BinOutCateName = "SOC_M9XX"
            End Select
            If MbistBlockCount <> 0 Then
                For Each site In TheExec.sites
                    Block_Pass_count = 0
                    BinOutFlag = ""
                    For i = 0 To MbistBlockCount - 1
                        F_BlockPass = True
                        For k = 0 To TestTypeCount - 1
                            F_TestTypePass = False ' initialize
                            strPassFailResult = "N/A"
                            varPassFailResult = 0
                            Select Case k
                                Case 0
                                    TestTypeName = "TD"
                                Case 1
                                    TestTypeName = "SA"
                                Case 2
                                    TestTypeName = "Mbist"
                            End Select
                            
                            For Each BypassType In BypassTestType_ary
                                If BypassType = TestTypeName Then
                                    GoTo NextType
                                End If
                            Next BypassType
                            
                            'compose string for block name
                            str1 = "00" & Trim(str(i))
                            If MbistCate Like "*CPU*" Then
                                str1 = right(str1, 2)
                            ElseIf MbistCate Like "*SOC*" Then
                                str1 = right(str1, 1)
                            End If
                            'm_BXX = MbistCate & str1 'm_BXX = CPU_MBist_BXX / SOC_MBist_SOCX /SOC_MBist_M9XX
                            'F_CPU_B02_TD
                            BlockResultFalg = "F_" & MbistCate & "_" & MbistCate_BlockName & str1 & "_" & TestTypeName
                            
                            If TheExec.sites.item(site).FlagState(BlockResultFalg) = logicFalse Then
                                strPassFailResult = "Fail"
                                varPassFailResult = 0
                                F_TestTypePass = False
                            ElseIf TheExec.sites.item(site).FlagState(BlockResultFalg) = logicClear Then
                                strPassFailResult = "N/A"
                                varPassFailResult = -1
                                F_TestTypePass = False
                            ElseIf TheExec.sites.item(site).FlagState(BlockResultFalg) = logicTrue Then
                                'Block_Pass_count = Block_Pass_count + 1
                                strPassFailResult = "Pass"
                                varPassFailResult = 1
                                F_TestTypePass = True
                            End If
                            
                            F_BlockPass = F_BlockPass And F_TestTypePass
                            
                            str3 = Replace(BlockResultFalg, "F_", "")
                            
                            TheExec.flow.TestLimit varPassFailResult, -1, 1, Tname:=CStr(str3)
                            TheExec.Datalog.WriteComment "Site: " + Trim(CStr(site)) + ", " + CStr(str3) + ": " + strPassFailResult
NextType:
                        Next k
                        If F_BlockPass = True Then Block_Pass_count = Block_Pass_count + 1 'this block pass
                    Next i
    
                        str2 = "00" & Trim(Block_Pass_count)
                        str2 = right(str2, 2)
                        BinOutFlag = "F_" & str2 & BinOutCateName & "_pass" 'BinOutFlag = F_XXCPU_pass / F_XXSOC_pass / F_XXM9X_pass
                        TheExec.sites.item(site).FlagState(BinOutFlag) = logicTrue
                        TheExec.flow.TestLimit Block_Pass_count, 0, MbistBlockCount, Tname:=BinOutCateName + ", Total Pass-Blocks"
                        TheExec.Datalog.WriteComment "Site: " + Trim(CStr(site)) + ", " + BinOutCateName + ":Total Pass-Blocks: " + CStr(Block_Pass_count) + ", Turn On Flag: " + BinOutFlag + ": " + CStr(TheExec.sites.item(site).FlagState(BinOutFlag))
                        TheExec.Datalog.WriteComment ""
                        
                        If MbistCate = "SOC" Then SOC_Total_Block_Pass_count(site) = SOC_Total_Block_Pass_count(site) + Block_Pass_count
                Next site
            End If
        Next j
        
        'swlinza, skua does not have M9 blocks
''        For Each Site In TheExec.sites
''            str2 = "00" & Trim(SOC_Total_Block_Pass_count(Site))
''            str2 = Right(str2, 2)
''            SOC_Total_Pass_Flag = "F_" & str2 & "_SOCX_M9XX_pass"
''            TheExec.sites.Item(Site).FlagState(SOC_Total_Pass_Flag) = logicTrue
''            TheExec.Datalog.WriteComment "Site: " + Trim(CStr(Site)) + ", " + "SOCX & M9XX " + ":Total Pass-Blocks: " + CStr(SOC_Total_Block_Pass_count(Site)) + ", Turn On Flag: " + SOC_Total_Pass_Flag + ": " + CStr(TheExec.sites.Item(Site).FlagState(SOC_Total_Pass_Flag))
''        Next Site
        
        
    BinOutFlag_BlockPassRate = 1
    TheExec.Datalog.WriteComment "---------------------------------------------------------"
    TheExec.Datalog.WriteComment "---------------------------------------------------------"
    
Exit Function

errHandler:
    TheExec.Datalog.WriteComment "<Error> " + funcName + ":: please check it out."
    If AbortTest Then Exit Function Else Resume Next
     
End Function

Public Function WriteRepairInfo_in_STDF(RepairDieCandidate As Boolean) As Long
'SWLIN 2017/01/18

'Delete start, Leon Li, 20190629
'On Error GoTo errHandler

'    Dim funcName As String:: funcName = "WriteRepairInfo_in_STDF"
'    Dim m_LoopIndex_DM As Long
'    Dim m_LoopIndex_PM As Long
'    Dim m_LoopIndex_BM As Long
'    Dim m_BMname As String
'    Dim m_testNumber As Long
'    Dim site As Variant
'
'    Init_Datalog_Setup
'
'    'get block name and blockSequence
'    Call auto_Mbist_GetLoopIndex_DM_PM_BM(m_LoopIndex_DM, m_LoopIndex_PM, m_LoopIndex_BM, False)
'
'    'get test number
'    For Each site In TheExec.sites
'        m_testNumber = TheExec.sites.Item(site).TestNumber
'        If TheExec.sites(site).SiteVariableValue("MbistBlockChk") = "M9" Then
'            m_LoopIndex_BM = m_LoopIndex_BM + SOC_count + 1
'            Exit For
'        End If
'    Next site
'
'    If UCase(TheExec.sites(site).SiteVariableValue("MbistBlockChk")) = "M9" Or UCase(TheExec.sites(site).SiteVariableValue("MbistBlockChk")) = "SOC" Then
'        m_BMname = UCase(SocMbist_Block.Category(m_LoopIndex_BM).name)
'    ElseIf UCase(TheExec.sites(site).SiteVariableValue("MbistBlockChk")) = "CPU" Then
'        m_BMname = UCase(CpuMbist_Block.Category(m_LoopIndex_BM).name)
'    End If
'
'    'Wirte comment and stdf
'    If RepairDieCandidate = True Then
'        TheExec.Datalog.WriteComment "-------------------------------"
'        TheExec.Datalog.WriteComment "  Repair Candidate :" & RepairDieCandidate
'        TheExec.Datalog.WriteComment "-------------------------------"
'        TheExec.Flow.TestLimit 1, 0, 1, TName:="Repair Candidate" & "_" & m_BMname, TNum:=m_testNumber + m_LoopIndex_BM
'    Else
'        TheExec.Datalog.WriteComment "-------------------------------"
'        TheExec.Datalog.WriteComment "  Repair Candidate :" & RepairDieCandidate
'        TheExec.Datalog.WriteComment "-------------------------------"
'        TheExec.Flow.TestLimit 0, 0, 1, TName:="Repair Candidate" & "_" & m_BMname, TNum:=m_testNumber + m_LoopIndex_BM
'    End If
'
'
'    Exit Function
'errHandler:
'    TheExec.Datalog.WriteComment "<Error> " + funcName + ":: please check it out."
'    If AbortTest Then Exit Function Else Resume Next
'Delete end, Leon Li, 20190629
    
End Function


