Attribute VB_Name = "LIB_VDD_BINNING"
Option Explicit
#Const IGXL_VER_1030 = False

Public Function decide_test_type(Test_Type As testType, inst_name As String)
On Error GoTo errHandler 'Add ErrHandler 2023/08/01
    Dim strTemp As String
    strTemp = LCase(inst_name)
    
    If strTemp Like "*elb*" Or strTemp Like "*ilb*" Or strTemp Like "*tmps*" _
    Or strTemp Like "*gfxtd*" Or strTemp Like "*gputd*" Or strTemp Like "*cputd*" Or strTemp Like "*soctd*" _
    Or strTemp Like "*gfxsa*" Or strTemp Like "*gpusa*" Or strTemp Like "*cpusa*" Or strTemp Like "*socsa*" Then
        Test_Type = testType.TD
    ElseIf strTemp Like "*cpu*bist*" Or strTemp Like "*gfx*bist*" Or strTemp Like "*soc*bist*" Or strTemp Like "*gpu*bist*" Then
        Test_Type = testType.Mbist
    ElseIf strTemp Like "*spi*" Then
        Test_Type = testType.SPI
    ElseIf strTemp Like "*rtos*" Then
        Test_Type = testType.RTOS
    Else
        Call Print_Error_Message(Error_Info, "LIB_VDD_BINNING", "decide_test_type", "Test instance:" & TheExec.DataManager.instancename & ", it doesn't have the correct keyword to decide TestType. Error!!!")
        'TheExec.ErrorLogMessage "Test instance:" & TheExec.DataManager.instanceName & ", it doesn't have the correct keyword to decide TestType. Error!!!"    ''20230807: Removed by sanity check
    End If
Exit Function
errHandler:
    Call Print_Error_Message(Error_Info, "LIB_VDD_BINNING", "decide_test_type") 'Add ErrHandler 2023/08/01
    If AbortTest Then Exit Function Else Resume Next
End Function

'20210831: Modified to print the info if Find_IDS_ZONE_per_site can't find the IDS_Zone for the performance mode.
'20210830: Modified to initialize the siteVariable find_ids_zone_flag.
'20210420: Modified to remove the unused vbt code of GradeSearchMethod. Once if C651 provided any new definition of GradeSearchMethod, we will revise the vbt code.
'20210331: Modified to print the full performance mode for Find_IDS_ZONE_per_site.
'20200914: Modified to merge the redundant branches.
'20200317: Modified for SearchByPmode.
'20191127: Modified for the revised InitVddBinTable.
'20190716: Modified to unify the unit for IDS. ids_current with unit mA.
'20190507: Modified to add "Cdec" for IDS to avoid double format accuracy issues.
'20190422: Modified to define the bin number for DUT with IDS on the IDS_limit.
Public Function Find_IDS_ZONE_per_site(ids_current As SiteDouble, p_mode As Integer)
On Error GoTo errHandler 'Add ErrHandler 2023/08/01
    Dim site As Variant
    Dim Test_Type As testType
    Dim ids_zone_num As Long
    Dim find_ids_zone_flag As New SiteBoolean
    Dim str_flag_BinOut As String
'''//==================================================================================================================================================================================//'''
'''//Note:
'''1. Compare each IDS Range to find out the IDS ZONE by site.
'''2. If the first step passbin of the IDS ZONE is greater than current passbin, use the passbin from the IDS ZONE to be current passbin.
'''3. It updates flags "F_IDS_Binx" and "F_IDS_Biny" for Bin_Table.
'''//==================================================================================================================================================================================//'''
    '''//init.
    '''//The default Testtype is TD.
    Test_Type = testType.TD
    '''20210830: Modified to initialize the siteVariable find_ids_zone_flag.
    find_ids_zone_flag = False
    
    For Each site In TheExec.sites
        For ids_zone_num = 0 To Max_IDS_Zone - 1
            If VBIN_IDS_ZONE(p_mode).Used = True Then
                '''//If ids_current >= ids_range, use the next zone.
                '''//IDS calculation uses the scale and the unit in "mA".
                If TheExec.enableWord("Vddbin_PTE_Debug") = True Then
                    VBIN_IDS_ZONE(p_mode).IDS_ZONE_NUMBER = ids_zone_num
                    find_ids_zone_flag(site) = True
                    Exit For
                Else
                    If CDec(ids_current(site)) >= CDec(VBIN_IDS_ZONE(p_mode).Ids_range(ids_zone_num, Test_Type)) _
                    And CDec(ids_current(site)) < CDec(VBIN_IDS_ZONE(p_mode).Ids_range(ids_zone_num + 1, Test_Type)) Then
                        VBIN_IDS_ZONE(p_mode).IDS_ZONE_NUMBER = ids_zone_num
                        find_ids_zone_flag(site) = True
                        
                        '''//PE asked us to add the flag to distinguish the BinX parts with IDS > Bin1E1 IDS limit from the BinX parts with IDS < Bin1E1 IDS limit.
                        If VBIN_IDS_ZONE(p_mode).IDS_ZONE_NUMBER <> 0 And CDec(ids_current(site)) > CDec(BinCut(p_mode, 1).IDS_CP_LIMIT(0)) Then
                            str_flag_BinOut = "F_" & UCase(VddBinName(p_mode)) & "_IDS"
                            TheExec.sites.item(site).FlagState(str_flag_BinOut) = logicTrue
                        End If
                        
                        '''//Check if CurrentPassBinCutNum is same as VBIN_IDS_ZONE(p_mode).passBinCut.
                        If CurrentPassBinCutNum(site) < VBIN_IDS_ZONE(p_mode).passBinCut(ids_zone_num, 0) Then
                            CurrentPassBinCutNum(site) = VBIN_IDS_ZONE(p_mode).passBinCut(ids_zone_num, 0)
                        End If
                        
                        '''//Update FlagState of "F_IDS_BinX" and "F_IDS_BinY" by CurrentPassBinCutNum(Site).
                        If Flag_BinX_Info_Parsed = True Then '20240131 Bin1 BinX merge IDS binning flag
                            If VBIN_IDS_ZONE(p_mode).IDS_ZONE_NUMBER = 1 Then
                                TheExec.sites.item(site).FlagState("F_IDS_BinX") = logicTrue
                            ElseIf VBIN_IDS_ZONE(p_mode).IDS_ZONE_NUMBER = 2 Then
                                TheExec.sites.item(site).FlagState("F_IDS_BinY") = logicTrue
                            Else
                                'BinX BinY else Bin flag
                            End If
                        ElseIf CurrentPassBinCutNum(site) = 2 Then
                            TheExec.sites.item(site).FlagState("F_IDS_BinX") = logicTrue '''for Binx IDS binning
                        ElseIf CurrentPassBinCutNum = 3 Then
                            TheExec.sites.item(site).FlagState("F_IDS_BinY") = logicTrue '''for Biny IDS binning
                        End If
                        
                        '''//Exit ids_zone-loop
                        Exit For
                    End If
                End If
            End If
        Next ids_zone_num
        
        '''//If it can't find any IDS zone...
        If find_ids_zone_flag(site) = False Then
            '''20210831: Modified to print the info if Find_IDS_ZONE_per_site can't find the IDS_Zone for the performance mode.
            Call Print_Error_Message(Error_Info, "LIB_VDD_BINNING", "Find_IDS_ZONE_per_site", "site:" & site & "," & VddBinName(p_mode) & ", Find_IDS_ZONE_per_site can't find the IDS_Zone for the performance mode. Error!!!")
        End If
    Next site
Exit Function
errHandler:
    Call Print_Error_Message(Error_Info, "LIB_VDD_BINNING", "Find_IDS_ZONE_per_site") 'Add ErrHandler 2023/08/01
    If AbortTest Then Exit Function Else Resume Next
End Function

'20210810: Modified to skip printing the info about the step-adjusted voltage for p_mode, requested by C651 Si and TSMC ZYLINI.
'20210726: Modified to add the argument "bincutNum As Long" to the vbt function Adjust_Multi_PassBinCut_Per_Site.
'20210722: Modified to use VBIN_IDS_ZONE(p_mode).Product_Voltage(Max_IDS_Zone, Max_IDS_Step) and DYNAMIC_VBIN_IDS_ZONE(p_mode).Product_Voltage(Max_IDS_Step) for GradeVDD.
'20210611: Modified to bin out the failed site in Adjust_Multi_PassBinCut_Per_Site.
'20210420: Modified to remove the unused vbt code of GradeSearchMethod. Once if C651 provided any new definition of GradeSearchMethod, we will revise the vbt code.
'20210319: Modified to print bincutNum(site).
'20210317: Modified to revise the format of string about the adjusted BinCut voltage for Adjust_Multi_PassBinCut_Per_Site.
'20210223: Modified to replace "Dim step As Long" with "Dim idx_step As Long".
'20210120: Modified to merge the branches.
'20200317: Modified for SearchByPmode.
'20191127: Modified for the revised InitVddBinTable.
'20181004: If bincutNum>VBIN_RESULT(P_mode).PASSBINCUT, we adjust the bin number.
Public Function Adjust_Multi_PassBinCut_Per_Site(p_mode As Integer, site As Variant, bincutNum As Long, Optional bool_SkipPrintingVoltage As Boolean = False, Optional align_instance As Boolean = False)
On Error GoTo errHandler 'Add ErrHandler 2023/08/01
    Dim idx_step As Long
    Dim find_out_flag As Boolean
'''//==================================================================================================================================================================================//'''
'''//Note:
'''If the CurrentPassBinCutNum is Bin2 but the result of current performance mode is still Bin1.
'''We will find the fisrt step which is Bin2 in the IDS zone to adjust the grade and gradevdd to Bin2.
'''Ex: original grade is Bin1 EQ4, adjust to Bin2
'''                  C                              EQ                             PassBinCut
'''     step0  step1  step2  step3      step0  step1  step2  step3        step0  step1  step2  step3
'''     700    720    780    800          4      3      2      1            1      1      2      2
'''                    V                                V                                 V
'''//==================================================================================================================================================================================//'''
    '''//init
    find_out_flag = False
    
    '''//If bincutNum>VBIN_RESULT(P_mode).PASSBINCUT, we adjust the bin number.
    If VBIN_RESULT(p_mode).passBinCut < bincutNum Then
        '''20230314: Only show error when Bin 1
        Dim fuse_bincutnum As Integer
        Dim field As eFuseBdfField
        Set field = BdfDataBase.Bank_Cfg.Fields("Product_Identifier")
        fuse_bincutnum = field.DsscDecValue(site)
        
        ''Non-CP1 not allow Bin1->BinX.
        ''But if we have some modes haven't been searched, we allow BinX die to sync those modes from Bin1->BinX
        If Not LCase(TheExec.CurrentJob) = "cp1" And align_instance = False And fuse_bincutnum = 0 Then
            Call Print_Error_Message(Error_Info, "LIB_VDD_BINNING", "Adjust_Multi_PassBinCut_Per_Site", "site:" & TheExec.sites.siteNumber & "," & VddBinName(p_mode) & "," & "bin=" & bincutNum & "," & _
                                            "It's not allowed to upgrade the product identifier in non-CP1 stage!!!")
            
            '''20210611: Modified to bin out the failed site in Adjust_Multi_PassBinCut_Per_Site.
            '''Warning!!!please check SortNumber and binNumber of Flag_Vddbinning_Fail_Stop in Bin_Table before using this.
            'Bin1 to BinX no show
                        TheExec.sites.item(site).FlagState(strGlb_Flag_Vddbinning_Fail_Stop) = logicTrue
            Call BinoutSite(site, 9806, 20, "F_Bin1_To_BinX_InWLFT1")
            '''//Shut down the failed site. As per discussion with Chihome, he suggested us to ensure that Sort Number/Bin Number/fail-stop should be updated before .result = tlResultFail.
            Exit Function
        End If
        
        For idx_step = 0 To DYNAMIC_VBIN_IDS_ZONE(p_mode).Max_Step - 1
            '''//Find out the step in Dynamic_IDS_ZONE to match the current BinCut number of DUT.
            '''//Then update step, BinCut voltage(Grade) and Efuse product voltage(GradeVDD).
            If DYNAMIC_VBIN_IDS_ZONE(p_mode).passBinCut(idx_step) = bincutNum Then
                VBIN_RESULT(p_mode).passBinCut = bincutNum
                VBIN_RESULT(p_mode).step_in_IDS_Zone = idx_step
                VBIN_RESULT(p_mode).step_in_BinCut = DYNAMIC_VBIN_IDS_ZONE(p_mode).EQ_Num(idx_step) - 1
                VBIN_RESULT(p_mode).GRADE = DYNAMIC_VBIN_IDS_ZONE(p_mode).Voltage(idx_step)
                VBIN_RESULT(p_mode).GRADEVDD = DYNAMIC_VBIN_IDS_ZONE(p_mode).Product_Voltage(idx_step)
                find_out_flag = True
                Exit For
            End If
        Next idx_step
        
        '''//Check if p_mode has the matched step in DYNAMIC_IDS_Zone for current PassBin.
        If find_out_flag = True Then
            If bool_SkipPrintingVoltage = False Then
                '''20210810: Modified to skip printing the info about the step-adjusted voltage for p_mode, requested by C651 Si and TSMC ZYLINI.
                TheExec.Datalog.WriteComment "site:" & TheExec.sites.siteNumber & "," & VddBinName(p_mode) & "," & "bin=" & bincutNum & "," & _
                                                "Adjust_Multi_PassBinCut_Per_Site changes " & AllBinCut(p_mode).powerPin & "=" & VBIN_RESULT(p_mode).GRADE
            End If
        Else '''If find_out_flag = False
            Call Print_Error_Message(Error_Info, "LIB_VDD_BINNING", "Adjust_Multi_PassBinCut_Per_Site", "site:" & TheExec.sites.siteNumber & "," & VddBinName(p_mode) & "," & "bin=" & bincutNum & "," & _
                                            "it can't find out the correct step for Adjust_Multi_PassBinCut_Per_Site. Error!!!")
            
            '''20210611: Modified to bin out the failed site in Adjust_Multi_PassBinCut_Per_Site.
            '''Warning!!!please check SortNumber and binNumber of Flag_Vddbinning_Fail_Stop in Bin_Table before using this.
            '''20240303: Use modular function to bin out site
            Call BinoutSite(site, VddbinFailStopSortNumber, VddbinFailStopBinNumber, strGlb_Flag_Vddbinning_Fail_Stop)
            
        End If
    End If
Exit Function
errHandler:
    Call Print_Error_Message(Error_Info, "LIB_VDD_BINNING", "Adjust_Multi_PassBinCut_Per_Site") 'Add ErrHandler 2023/08/01
    If AbortTest Then Exit Function Else Resume Next
End Function

'20210903: Modified to merge properties of "Public Type Instance_Step_Control" into Public Type Instance_Info "Public Type Instance_Info".
'20210901: Modified to rename "IndexLevelPerSite As New SiteLong" as "Step_Current As New SiteLong" for Public Type Instance_Info.
'20210812: Modified to rename the property "step_lowest As New SiteLong" as "step_inherit As New SiteLong".
'20210810: Modified to add the property "step_Lowest As New SiteLong" to Public Type DYNAMIC_VBIN_IDS_ZONE.
'20210726: Modified to add the argument "bincutNum As Long" to the vbt function Adjust_Multi_PassBinCut_Per_Site.
'20210722: Modified to use VBIN_IDS_ZONE(p_mode).Product_Voltage(Max_IDS_Zone, Max_IDS_Step) and DYNAMIC_VBIN_IDS_ZONE(p_mode).Product_Voltage(Max_IDS_Step) for GradeVDD.
'20210623: Modified to update step_Start for the vbt functon find_start_voltage if EnableWord "Vddbin_PTE_Debug" is enabled.
'20210611: Modified to bin out the failed site in find_start_voltage.
'20210526: Modified to remove Monotonicity_Offset check from find_start_voltage because C651 Si revised the check rules.
'20210518: Modified to update inst_info.is_Monotonicity_Offset_triggered(site).
'20210507: Modified to remove the redundant site-loop and use DYNAMIC_VBIN_IDS_ZONE(inst_info.p_mode).step_Mapping(PassBin,1) for "Vddbin_PTE_Debug".
'20210503: Modified to check if VBIN_RESULT(p_mode).GradeVDD < VBIN_RESULT(VBIN_RESULT(AllBinCut(inst_info.p_mode).PREVIOUS_Performance_Mode)).GradeVDD for Monotonicity_Offset.
'20210429: Modified to replace voltage_CalculatedFromIds with DYNAMIC_VBIN_IDS_ZONE(inst_info.p_mode).Voltage(step_Start).
'20210427: Modified to add "Monotonicity_Offset" for GradeVDD check of p_mode.
'20210420: C651 Si did internal syncup and confirmed that Montonicitiy Check should use product voltage(PV) only.
'20210420: Modified to remove the unused vbt code of GradeSearchMethod. Once if C651 provided any new definition of GradeSearchMethod, we will revise the vbt code.
'20210419: Modified replace "AllBinCut(inst_info.p_mode).Allow_Equal <> cntVddbinPmode + 1" with "AllBinCut(inst_info.p_mode).Allow_Equal <> 0".
'20210419: Modified to set step_stop = step_inherit because voltage heritance between p_mode and previous mode.
'20210408: Modified to overwrite step_inherit and VBIN_RESULT(inst_info.p_mode).step_in_IDS_Zone if p_mode is interpolated.
'20210407: Modified to revise the vbt code for the new Interpolation method proposed by C651 Toby.
'20210226: Modified to use step_Start and step_Stop to get startVoltage and StopVoltage.
'20201210: Modified to use the arguments "inst_info As Instance_Info" and "step_control As Instance_Step_Control" for find_start_voltage.
'20200502: Modified to replace variable name "IdsVoltage" with "voltage_CalculatedFromIds".
'20200423: Modified to replace "BinCut(p_mode, bincutNum(site)).tested = True" with "VBIN_RESULT(p_mode).tested=True".
'20200317: Modified for SearchByPmode.
'20200120: Modified to print the information about voltage adjustment.
'20191127: Modified for the revised InitVddBinTable.
'20190716: Modified to unify the unit for IDS.
'20190507: Modified to add "Cdec" to avoid double format accuracy issues.
Public Function find_start_voltage(inst_info As Instance_Info)
On Error GoTo errHandler 'Add ErrHandler 2023/08/01
    Dim site As Variant
    Dim step_Start As New SiteLong '''Bin (1[highest V]-6) decided by IDS_Distribution table
    Dim step_inherit As New SiteLong
    Dim EQ_Num As Long
    Dim exit_while_flag As Boolean
    Dim bincutNum As New SiteLong
    Dim i As Integer
    Dim field As eFuseBdfField
'''//==================================================================================================================================================================================//'''
'''//Note:
'''//Find the start voltage of searching for each test Instance, The consider condition as following:
'''1. If the Current Pass Bin Number has been Bin2, adjust the start Voltage to first Bin2 step in the IDS ZONE Number.
'''2. If the previous performance mode had been tested, we need to consider the inheritance by comparing the product value.
'''3. If the current performance mode had been tested, we need to consider the inheritance the step in IDS Zone.
'''4. If the current performance mode was not tested, we need to follow the IDS_START_EQ to find the start step in IDS ZONE.
'''20210420: C651 Si did internal syncup and confirmed that Montonicitiy Check should use product voltage(PV) only.
'''//==================================================================================================================================================================================//'''
'=============================================================================================================================================================
' A. The variable "step_inherit" is the step in the IDS Zone which inherit from the PREVIOUS Performance Mode and Current Performance Mode.
'    1. If the current performance mode did not been tested, we will use last EQ Number (step 0 in ids zone) to be the "step_inherit".
'    2. If the current performance mode had been tested, we will use the result(p_mode).step_in_ids_zone to inherit the previous test items to be the "step_inherit".
'    3. If the PREVIOUS Performance Mode had been tested, we will compare the efuse value (product value) to inherit the step in ids zone to be the "step_inherit".
'    4. If the current pass bincut number is Bin2, we will adjust the current performance mode to the Bin2 last EQ number in the ids zone.
'
' B. The variable "Step_start" is the step in the IDS Zone which come from the start EQ number in the IDS distribution Table.
'    1. If the current performance mode had not been tested,
'       ==> we need to consider the start EQ number in the IDS Distribution Table and base on start EQ number to find out the step in ids zone to be the "Step_start".
'
'    2. If the current performance mode had been tested but the test type is "SPI",
'       ==> we need to consider the start EQ number in the IDS Distribution Table and base on start EQ number to find out the step in ids zone to be the "Step_start".
'
'    3. If the current performance mode had been tested but the test type is not "SPI",
'       ==> we don't need to consider the IDS Distribution Table. And set the "Step_start" to 0(Just let the "step_inherit" to be the final step in ids zone).
'
' C. Compare the "step_inherit" and "Step_start" to decide the final step in the ids zone, And use the step to calculate the start voltage.
'=============================================================================================================================================================
    Set field = BdfDataBase.Bank_Cfg.Fields("Product_Identifier")
    For Each site In TheExec.sites
        '''//Get PassBin of the performance mode.
        bincutNum(site) = CurrentPassBinCutNum(site)
        
        '''//If the passbincut doesn't match CurrentPassBinCutNum, adjust the correct step to align passbincut of P_mode with CurrentPassBinCutNum, and update Grade/GradeVDD.
        If VBIN_RESULT(inst_info.p_mode).passBinCut <> bincutNum(site) Then
            If (LCase(TheExec.CurrentJob) = "wlft1" Or LCase(TheExec.CurrentJob) = "wlft2") And field.DsscDecValue > 0 Then
                Adjust_Multi_PassBinCut_Per_Site inst_info.p_mode, site, bincutNum(site), , True
            Else
                Adjust_Multi_PassBinCut_Per_Site inst_info.p_mode, site, bincutNum(site)
            End If
        End If
        
'''//Start of determining step_inherit//'''
        '''=============================================================================================================================================================
        ''' A. Just consider the step for inherit from previous performance mode and current performance mode to define the step in ids zone, step, Grade and GradeVDD,
        '''   ==> The final step needs to compare with "Step_start".
        '''=============================================================================================================================================================
        '''//Decide PassBin by considering previous test result of the same performance mode and previous lower performance mode (for voltage inheritance).
        If VBIN_RESULT(AllBinCut(inst_info.p_mode).PREVIOUS_Performance_Mode).tested = True Then  'if previous mode has been Tested, the flag is ture; if not, bypass.
            '''**************************************************************************************************************************************************************************'''
            '''Judge whether current performance mode has been Tested or not, or ECID fail, because of ECID fail will become 0,
            '''If the grade had adjusted to Bin2, the grade is not 0. (if the result had become 0 and the current performance mode has been tested, we can not assign the value to it.)
            '''**************************************************************************************************************************************************************************'''
            If VBIN_RESULT(inst_info.p_mode).tested = False Or VBIN_RESULT(inst_info.p_mode).GRADE > 0 Then
                If TheExec.enableWord("Vddbin_PTE_Debug") = True Then
                    If Flag_BinX_Info_Parsed = True And DYNAMIC_VBIN_IDS_ZONE(inst_info.p_mode).step_Mapping(VBIN_RESULT(inst_info.p_mode).passBinCut, AllBinCut(inst_info.p_mode).Last_Bin1_EQN) <> -1 Then        ''BinX merge, set last bin1 EQN step
                        step_inherit = DYNAMIC_VBIN_IDS_ZONE(inst_info.p_mode).step_Mapping(1, AllBinCut(inst_info.p_mode).Last_Bin1_EQN)
                    ElseIf Flag_BinX_Info_Parsed = False And VBIN_RESULT(inst_info.p_mode).passBinCut = 1 And DYNAMIC_VBIN_IDS_ZONE(inst_info.p_mode).step_Mapping(VBIN_RESULT(inst_info.p_mode).passBinCut, 1) <> -1 Then '''//"Vddbin_PTE_Debug" should use step of Bin1 EQN1 in Dynamic_IDS_Zone.
                        step_inherit = DYNAMIC_VBIN_IDS_ZONE(inst_info.p_mode).step_Mapping(VBIN_RESULT(inst_info.p_mode).passBinCut, 1)
                    Else
                        Call Print_Error_Message(Error_Info, "LIB_VDD_BINNING", "find_start_voltage", "site:" & site & "," & VddBinName(inst_info.p_mode) & ", it should use Bin1 DUT for Vddbin_PTE_Debug. Error!!!")
                    End If
                ElseIf inst_info.StartBin1E1_Flag = True Then   '''20220125, add for set start with bin1 eq1
                    ''20230901: Added saving bin1 EQN, for pre-flow in bincut
                    If Flag_BinX_Info_Parsed = True Then        ''BinX merge, set last bin1 EQN step
                        step_inherit = DYNAMIC_VBIN_IDS_ZONE(inst_info.p_mode).step_Mapping(1, AllBinCut(inst_info.p_mode).Last_Bin1_EQN)
                    Else                                        ''Regular
                        If CurrentPassBinCutNum(site) = 1 And DYNAMIC_VBIN_IDS_ZONE(inst_info.p_mode).step_Mapping(VBIN_RESULT(inst_info.p_mode).passBinCut, 1) <> -1 Then
                            step_inherit = DYNAMIC_VBIN_IDS_ZONE(inst_info.p_mode).step_Mapping(VBIN_RESULT(inst_info.p_mode).passBinCut, 1)
                        Else
                            step_inherit = VBIN_RESULT(inst_info.p_mode).step_in_IDS_Zone
                        End If
                    End If
                ElseIf VBIN_RESULT(inst_info.p_mode).tested = False And bincutNum(site) = 1 Then
                    '''+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
                    ''' 1. If the current performance mode has not been tested ==> we will use last EQ Number (step 0 in ids zone) to be the "step_inherit".
                    '''+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
                    '''//Since p_mode is interpolated but not tested, the revised Interpolated method changes IDS_Start_step Dynamic_IDS_ZONE, it needs to revise step_inherit.
                    If DYNAMIC_VBIN_IDS_ZONE(inst_info.p_mode).interpolated = True Then
                        step_inherit = DYNAMIC_VBIN_IDS_ZONE(inst_info.p_mode).step_Interpolated_Start
                    Else
                        step_inherit = 0 '''but EcidVddExecuted(P_mode)(Site) might be true, just keep the same level.
                    End If
                    
                    VBIN_RESULT(inst_info.p_mode).step_in_IDS_Zone = 0
                Else '''If VBIN_RESULT(inst_info.p_mode).tested =True or bincutNum(site) <> 1
                    '''+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
                    ''' 2. If the current performance mode had been tested  ==> we will use the result(p_mode).step_in_ids_zone to inherit the previous test items to be the "step_inherit".
                    ''' 3. If the PREVIOUS Performance Mode had been tested ==> we will compare the efuse value (product value) to inherit the step in ids zone to be the "step_inherit".
                    ''' 4. If the current pass bincut number is Bin2        ==> we will adjust the current performance mode to the bin2 last EQ number in the ids zone.
                    '''+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
                    step_inherit = VBIN_RESULT(inst_info.p_mode).step_in_IDS_Zone
                End If
                
                If VBIN_RESULT(inst_info.p_mode).tested = False Then
                    VBIN_RESULT(inst_info.p_mode).GRADE = DYNAMIC_VBIN_IDS_ZONE(inst_info.p_mode).Voltage(step_inherit)     '''use the ids zone step
                    EQ_Num = DYNAMIC_VBIN_IDS_ZONE(inst_info.p_mode).EQ_Num(step_inherit)
                    '''//PRODUCT = CP LVCC + CPGB.
                    VBIN_RESULT(inst_info.p_mode).GRADEVDD = DYNAMIC_VBIN_IDS_ZONE(inst_info.p_mode).Product_Voltage(step_inherit)
                End If
                
                '''//The current performance-mode's efuse voltage should be greater than lower performance mode's efuse voltage.
                exit_while_flag = False
                
                '''//Check if Allow_Equal and previous performance mode of p_mode are tested.
                If AllBinCut(inst_info.p_mode).PREVIOUS_Performance_Mode = AllBinCut(inst_info.p_mode).Allow_Equal And AllBinCut(inst_info.p_mode).Allow_Equal <> 0 Then
                    '''//If Voltage of current performance mode (Grade and GradeVDD) are lower than previous performance mode.
                    '''//Note: If the vbt of checking GRADE is masked, please set globalVariable "Public Const Flag_Only_Check_PV_for_VoltageHeritage As Boolean = True".
                    While (CDec(VBIN_RESULT(inst_info.p_mode).GRADEVDD) < CDec(VBIN_RESULT(AllBinCut(inst_info.p_mode).PREVIOUS_Performance_Mode).GRADEVDD) And exit_while_flag = False) _
                    'Or (CDec(VBIN_RESULT(inst_info.p_mode).GRADE) < CDec(VBIN_RESULT(AllBinCut(inst_info.p_mode).PREVIOUS_Performance_Mode).GRADE) And exit_while_flag = False)
                        step_inherit = step_inherit + 1 'it will be increased
                        
                        If step_inherit > DYNAMIC_VBIN_IDS_ZONE(inst_info.p_mode).Max_Step - 1 Then
                            Call Print_Error_Message(Error_Info, "LIB_VDD_BINNING", "find_start_voltage", "site:" & site & "," & VddBinName(inst_info.p_mode) & ",find_start_voltage failed. Error!!!")
                            
                            '''//If no step avaiable, exit the while...
                            exit_while_flag = True
                            '''20210611: Modified to bin out the failed site in find_start_voltage.
                            '''Warning!!!please check SortNumber and binNumber of Flag_Vddbinning_Fail_Stop in Bin_Table before using this.
                            '''20240303: Use modular function to bin out site
                            Call BinoutSite(site, VddbinFailStopSortNumber, VddbinFailStopBinNumber, strGlb_Flag_Vddbinning_Fail_Stop)
                            
                        Else
                            '''//If Voltage of current performance mode (Grade and GradeVDD) are lower than previous performance mode.
                            VBIN_RESULT(inst_info.p_mode).GRADE = DYNAMIC_VBIN_IDS_ZONE(inst_info.p_mode).Voltage(step_inherit)
                            EQ_Num = DYNAMIC_VBIN_IDS_ZONE(inst_info.p_mode).EQ_Num(step_inherit)
                            '''//PRODUCT = CP LVCC + CPGB.
                            VBIN_RESULT(inst_info.p_mode).GRADEVDD = DYNAMIC_VBIN_IDS_ZONE(inst_info.p_mode).Product_Voltage(step_inherit)
                        End If
                    Wend
                Else '''//If p_mode has no Allow_Equal.
                    '''//If Voltage of current performance mode is lower than previous performance mode, step-looping until gradevdd is greater than gradevdd of previous mode.
                    '''//Note: If the vbt of checking GRADE is masked, please set globalVariable "Public Const Flag_Only_Check_PV_for_VoltageHeritage As Boolean = True".
                    While ((CDec(VBIN_RESULT(inst_info.p_mode).GRADEVDD) <= CDec(VBIN_RESULT(AllBinCut(inst_info.p_mode).PREVIOUS_Performance_Mode).GRADEVDD)) And exit_while_flag = False) _
                    'Or (CDec(VBIN_RESULT(inst_info.p_mode).GRADE) <= CDec(VBIN_RESULT(AllBinCut(inst_info.p_mode).PREVIOUS_Performance_Mode).GRADE) And exit_while_flag = False)
                        step_inherit = step_inherit + 1
                        
                        TheExec.Datalog.WriteComment "site:" & site & ", BinCut voltage for pmode:" & VddBinName(inst_info.p_mode) & " is equal or smaller than the previous pmode:" _
                                                        & VddBinName(AllBinCut(inst_info.p_mode).PREVIOUS_Performance_Mode) & ". So that adjust the BinCut step for find_start_voltage."
                        
                        If step_inherit > DYNAMIC_VBIN_IDS_ZONE(inst_info.p_mode).Max_Step - 1 Then
                            Call Print_Error_Message(Error_Info, "LIB_VDD_BINNING", "find_start_voltage", "site:" & site & ", " & VddBinName(inst_info.p_mode) & ", it has the incorrect step. Error!!!")
                            
                            '''//If no step avaiable, exit the while...
                            exit_while_flag = True
                            '''20210611: Modified to bin out the failed site in find_start_voltage.
                            '''Warning!!!please check SortNumber and binNumber of Flag_Vddbinning_Fail_Stop in Bin_Table before using this.
                            '''20240303: Use modular function to bin out site
                            Call BinoutSite(site, VddbinFailStopSortNumber, VddbinFailStopBinNumber, strGlb_Flag_Vddbinning_Fail_Stop)

                        Else
                            VBIN_RESULT(inst_info.p_mode).GRADE = DYNAMIC_VBIN_IDS_ZONE(inst_info.p_mode).Voltage(step_inherit) '''until it is greater than gradeVDD of previous mode.
                            EQ_Num = DYNAMIC_VBIN_IDS_ZONE(inst_info.p_mode).EQ_Num(step_inherit)
                            '''//PRODUCT = CP LVCC + CPGB.
                            VBIN_RESULT(inst_info.p_mode).GRADEVDD = DYNAMIC_VBIN_IDS_ZONE(inst_info.p_mode).Product_Voltage(step_inherit)
                        End If
                    Wend
                End If

                VBIN_RESULT(inst_info.p_mode).step_in_IDS_Zone = step_inherit
                EQ_Num = DYNAMIC_VBIN_IDS_ZONE(inst_info.p_mode).EQ_Num(VBIN_RESULT(inst_info.p_mode).step_in_IDS_Zone)
                VBIN_RESULT(inst_info.p_mode).step_in_BinCut = EQ_Num - 1
            End If
        Else '''If VBIN_RESULT(AllBinCut(inst_info.p_mode).PREVIOUS_Performance_Mode).tested = False
            If TheExec.enableWord("Vddbin_PTE_Debug") = True Then
                If Flag_BinX_Info_Parsed = True And DYNAMIC_VBIN_IDS_ZONE(inst_info.p_mode).step_Mapping(VBIN_RESULT(inst_info.p_mode).passBinCut, AllBinCut(inst_info.p_mode).Last_Bin1_EQN) <> -1 Then        ''BinX merge, set last bin1 EQN step
                    step_inherit = DYNAMIC_VBIN_IDS_ZONE(inst_info.p_mode).step_Mapping(1, AllBinCut(inst_info.p_mode).Last_Bin1_EQN)
                ElseIf Flag_BinX_Info_Parsed = False And VBIN_RESULT(inst_info.p_mode).passBinCut = 1 And DYNAMIC_VBIN_IDS_ZONE(inst_info.p_mode).step_Mapping(VBIN_RESULT(inst_info.p_mode).passBinCut, 1) <> -1 Then '''//"Vddbin_PTE_Debug" should use step of Bin1 EQN1 in Dynamic_IDS_Zone.
                    step_inherit = DYNAMIC_VBIN_IDS_ZONE(inst_info.p_mode).step_Mapping(VBIN_RESULT(inst_info.p_mode).passBinCut, 1)
                Else
                    Call Print_Error_Message(Error_Info, "LIB_VDD_BINNING", "find_start_voltage", "site:" & site & "," & VddBinName(inst_info.p_mode) & ", it should use Bin1 DUT for Vddbin_PTE_Debug. Error!!!")
                End If
            ElseIf inst_info.StartBin1E1_Flag = True Then   '''20220125, add for set start with bin1 eq1
                ''20230901: Added saving bin1 EQN, for pre-flow in bincut
                If Flag_BinX_Info_Parsed = True Then         ''BinX merge, set last bin1 EQN step
                    If VBIN_RESULT(inst_info.p_mode).tested = False Then
                        step_inherit = DYNAMIC_VBIN_IDS_ZONE(inst_info.p_mode).step_Mapping(1, AllBinCut(inst_info.p_mode).Last_Bin1_EQN)
                    Else
                        step_inherit = VBIN_RESULT(inst_info.p_mode).step_in_IDS_Zone
                    End If
                Else
                    If CurrentPassBinCutNum(site) = 1 And DYNAMIC_VBIN_IDS_ZONE(inst_info.p_mode).step_Mapping(VBIN_RESULT(inst_info.p_mode).passBinCut, 1) <> -1 Then
                        step_inherit = DYNAMIC_VBIN_IDS_ZONE(inst_info.p_mode).step_Mapping(VBIN_RESULT(inst_info.p_mode).passBinCut, 1)
                    Else
                        step_inherit = VBIN_RESULT(inst_info.p_mode).step_in_IDS_Zone
                    End If
                End If
            Else
                '''=======================================================================================================================================================
                ''' If the PREVIOUS Performance Mode has not been tested,
                ''' we do not need to consider the inherit step from PREVIOUS Performance Mode and just use the step in ids zone to be the "step_inherit".
                '''=======================================================================================================================================================
                step_inherit = VBIN_RESULT(inst_info.p_mode).step_in_IDS_Zone
            End If
        End If
'''//End of determining step_inherit//'''
    Next site

    For Each site In TheExec.sites
        '''//Update step_inherit to DYNAMIC_VBIN_IDS_ZONE(inst_info.p_mode).step_Lowest.
        DYNAMIC_VBIN_IDS_ZONE(inst_info.p_mode).step_inherit(site) = step_inherit(site)
        
        '''=============================================================================================================================================================
        ''' B. consider the IDS start EQ for current performance mode to define the step in ids zone, step, Grade and GradeVDD.
        '''=============================================================================================================================================================
        '''//Find start point according to IDS
        If VBIN_RESULT(inst_info.p_mode).tested = False _
        Or (VBIN_RESULT(inst_info.p_mode).tested = True And LCase(TheExec.DataManager.instancename) Like "*spi*") _
        Or (VBIN_RESULT(inst_info.p_mode).tested = True And LCase(TheExec.DataManager.instancename) Like "*rtos*") Then 'use IDS for only the first instance of every performance mode
            '''+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
            ''' 1. If the current performance mode had not been tested.
            '''   ==> we need to consider the start EQ number in the IDS Distribution Table and base on start EQ number to find out the step in ids zone to be the "Step_start".
            ''' 2. If the current performance mode had been tested but the test type is "SPI".
            '''   ==> we need to consider the start EQ number in the IDS Distribution Table and base on start EQ number to find out the step in ids zone to be the "Step_start".
            '''+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
            If TheExec.enableWord("Vddbin_PTE_Debug") = True Then
                If Flag_BinX_Info_Parsed = True And DYNAMIC_VBIN_IDS_ZONE(inst_info.p_mode).step_Mapping(VBIN_RESULT(inst_info.p_mode).passBinCut, AllBinCut(inst_info.p_mode).Last_Bin1_EQN) <> -1 Then  ''BinX merge, set last bin1 EQN step
                    step_inherit = DYNAMIC_VBIN_IDS_ZONE(inst_info.p_mode).step_Mapping(1, AllBinCut(inst_info.p_mode).Last_Bin1_EQN)
                    step_Start = step_inherit
                ElseIf Flag_BinX_Info_Parsed = False And VBIN_RESULT(inst_info.p_mode).passBinCut = 1 And DYNAMIC_VBIN_IDS_ZONE(inst_info.p_mode).step_Mapping(VBIN_RESULT(inst_info.p_mode).passBinCut, 1) <> -1 Then  '''//"Vddbin_PTE_Debug" should use step of Bin1 EQN1 in Dynamic_IDS_Zone.
                    step_inherit = DYNAMIC_VBIN_IDS_ZONE(inst_info.p_mode).step_Mapping(VBIN_RESULT(inst_info.p_mode).passBinCut, 1)
                    step_Start = step_inherit
                Else
                    Call Print_Error_Message(Error_Info, "LIB_VDD_BINNING", "find_start_voltage", "site:" & site & "," & VddBinName(inst_info.p_mode) & ", it should use Bin1 DUT for Vddbin_PTE_Debug. Error!!!")
                End If
            ElseIf inst_info.StartBin1E1_Flag = True Then   '''20220125, add for set start with bin1 eq1
                ''20230901: Added saving bin1 EQN, for pre-flow in bincut
                If Flag_BinX_Info_Parsed = True Then        ''BinX merge, set last bin1 EQN step
                    step_inherit = DYNAMIC_VBIN_IDS_ZONE(inst_info.p_mode).step_Mapping(1, AllBinCut(inst_info.p_mode).Last_Bin1_EQN)
                    step_Start = step_inherit
                Else
                    If CurrentPassBinCutNum(site) = 1 And DYNAMIC_VBIN_IDS_ZONE(inst_info.p_mode).step_Mapping(VBIN_RESULT(inst_info.p_mode).passBinCut, 1) <> -1 Then
                        step_inherit = DYNAMIC_VBIN_IDS_ZONE(inst_info.p_mode).step_Mapping(VBIN_RESULT(inst_info.p_mode).passBinCut, 1)
                        step_Start = step_inherit
                    End If
                End If
            Else
                '''from step number to get the LVCC search voltage (refer to "IDS_Distribution table")
                step_Start = DYNAMIC_VBIN_IDS_ZONE(inst_info.p_mode).IDS_START_STEP(inst_info.Test_Type) '''Base on ids zone number and column of "Start Bin" to get bin number
            End If '''then from bin number to know which step no. is corresponding
        Else '''test instance is not "spi" or "rtos".
            '''+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
            '''3. If the current performance mode had been tested but the test type is not "SPI".
            '''   ==> we don't need to consider the IDS Distribution Table. And set the "Step_start" to 0. (Just let the "step_inherit" to be the final step in ids zone.)
            '''+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
            step_Start(site) = 0
        End If
        
        If bincutNum > 1 Then  '''IDS distribution has no start point prediction for non-Bin1 DUT.
            step_Start = VBIN_RESULT(inst_info.p_mode).step_in_IDS_Zone
        End If
        
        '''//Decide start point by considering previous performance result, previous lower performance and IDS.
        '''=============================================================================================================================================================
        ''' A. algorithm = IDS condition:
        '''   1. the current performance mode and previous performance mode did not be tested and the ids_start_step is not 0.
        '''
        ''' B. algorithm = Linear condition:
        '''   1. the current performance mode had been tested.
        '''   2. the ids_start_step is 0.
        '''   3. the previous performance mode had been tested and the step had been adjusted to grater or equal to the ids_start_step.
        ''' C. compare the "Step_start" and "step_inherit" to define the step in ids zone, and calculate the start voltage.
        '''    If the "Step_start is greater than "step_inherit", the Start Voltage = voltage(Step_start) and the Stop Voltage = Voltage(step_inherit).
        '''=============================================================================================================================================================
        If VBIN_RESULT(inst_info.p_mode).tested = False And VBIN_RESULT(AllBinCut(inst_info.p_mode).PREVIOUS_Performance_Mode).tested = False Then
            If step_Start(site) <> 0 And DYNAMIC_VBIN_IDS_ZONE(inst_info.p_mode).passBinCut(step_Start) = 1 And TheExec.enableWord("Vddbin_PTE_Debug") = False And inst_info.StartBin1E1_Flag = False Then
                inst_info.step_Start(site) = step_Start
                '''20210419: Modified to set step_stop = step_inherit because voltage heritance between p_mode and previous mode.
                inst_info.step_Stop(site) = step_inherit '''lowest available step
                inst_info.gradeAlg(site) = GradeSearchAlgorithm.ids '''IDS search
            Else '''Start from EQN-based voltage with the lowest step, ex: Eqn7
                inst_info.step_Start(site) = step_Start
                inst_info.step_Stop(site) = DYNAMIC_VBIN_IDS_ZONE(inst_info.p_mode).Max_Step - 1 '''highest available step
                inst_info.gradeAlg(site) = GradeSearchAlgorithm.linear '''Linear search
            End If
        Else '''If VBIN_RESULT(inst_info.p_mode).tested = True Or VBIN_RESULT(AllBinCut(inst_info.p_mode).PREVIOUS_Performance_Mode).tested = True
            If step_Start(site) <= step_inherit Then
                step_Start(site) = step_inherit
                inst_info.step_Start(site) = step_Start
                inst_info.step_Stop(site) = DYNAMIC_VBIN_IDS_ZONE(inst_info.p_mode).Max_Step - 1
                inst_info.gradeAlg(site) = GradeSearchAlgorithm.linear ' Linear search
            Else '''for SPI test
                inst_info.step_Start(site) = step_Start
                inst_info.step_Stop(site) = step_inherit
                inst_info.gradeAlg(site) = GradeSearchAlgorithm.ids ' IDS search
            End If
        End If
        
        '''//Adjust step according to Pass step_inherit Category.
        inst_info.Step_Current(site) = step_Start '''If step_start is defined, the IndexLevelPerSite can directly use step_start.
        
        '''//IDS calculation uses the scale and the unit in "mA".
        TheExec.Datalog.WriteComment VddBinName(inst_info.p_mode) & "," & site & "," & _
                            "Alg=" & inst_info.gradeAlg & "," & _
                            DYNAMIC_VBIN_IDS_ZONE(inst_info.p_mode).Voltage(inst_info.step_Start) & "mV," & _
                            DYNAMIC_VBIN_IDS_ZONE(inst_info.p_mode).Voltage(inst_info.step_Stop) & "mV," & _
                            Format(inst_info.ids_current(site), ".0") & "mA," & _
                            DYNAMIC_VBIN_IDS_ZONE(inst_info.p_mode).Voltage(step_Start) & "mV," & _
                            VBIN_RESULT(inst_info.p_mode).GRADE & "mV," & _
                            VBIN_RESULT(AllBinCut(inst_info.p_mode).PREVIOUS_Performance_Mode).GRADE & "mV," & _
                            "bin" & bincutNum(site) & "," & _
                            "ids_zone " & DYNAMIC_VBIN_IDS_ZONE(inst_info.p_mode).IDS_ZONE_NUMBER
    Next site
Exit Function
errHandler:
    Call Print_Error_Message(Error_Info, "LIB_VDD_BINNING", "find_start_voltage") 'Add ErrHandler 2023/08/01
    If AbortTest Then Exit Function Else Resume Next
End Function

'20210903: Modified to merge properties of "Public Type Instance_Step_Control" into Public Type Instance_Info "Public Type Instance_Info".
'20210901: Modified to rename "stepcountMax As Long" as "maxStep As New SiteLong" for Public Type Instance_Info.
'20210805: Modified to remove the redundant vbt function initialize_step_control since it initialized step_control.All_Site_Mask = 0 in the vbt function decide_binSearch_and_start_voltage.
'20210805: Modified to check if inst_info.is_BinSearch=True for the vbt function decide_binSearch_and_start_voltage.
'20210803: Modified to update inst_info.ids_current = IDS_for_BinCut(VddBinStr2Enum(powerDomain)).Real in the vbt function initialize_inst_info and remove the redundant vbt function set_IDS_current.
'20210706: Modified to replace is_BinCutJob_for_StepSearch with AllBinCut(p_mode).is_for_BinSearch...
'20210126: Modified to revise the vbt code for DevChar.
'20201210: Modified to use the flag "is_BinCutJob_for_StepSearch" for "check_bincutJob_for_StepSearch" to check if the test program is binSearch or functional test.
'20201210: Modified to use the arguments "inst_info As Instance_Info" and "step_control As Instance_Step_Control" for ResetPmodePowerforBincut, set_IDS_current, and find_start_voltage.
'20201208: Modified to use "initialize_step_control".
'20201207: Created to decide the flag is binSearch and find start_voltage.
'20201203: Modified to revise the vbt code for the undefined testJobs.
Public Function decide_binSearch_and_start_voltage(inst_info As Instance_Info, FuncTestOnly As Boolean)
On Error GoTo errHandler
'''//==================================================================================================================================================================================//'''
'''//Note:
'''//The flag "is_BinCutJob_for_StepSearch" is True if any testCondition from the table "Non_Binning_Rail" has the keyword "*Evaluate*Bin*".
'''//inst_info.is_BinSearch =True if testCondition for PowerDomain of the binning p_mode contains the keyword "*Evaluate*Bin*".
'''inst_info.is_BinSearch = True    : find the start voltage, stop voltage, algorithm(Linear or IDS distribution).
'''inst_info.is_BinSearch = False   : only do functional test.
'''//AllBinCut(inst_info.p_mode).is_for_BinSearch = True is defined if testCondition from BinCut flow table(sheet "Non_Binning_Rail") with the keyword "*Evaluate*Bin*".
'''//==================================================================================================================================================================================//'''
    Dim site As Variant 'Carter, 20240304
    '''//inst_info.is_BinSearch=True is determined in the vbt function initialize_inst_info if testCondition for powerDomain of the binning p_mode contains the keyword "*evaluate*bin*".
    If inst_info.is_BinSearch = True Then
        If inst_info.is_DevChar_Running = True Then '''for DevChar.
            inst_info.maxStep = 0
        Else
            '''//Initialize flags of Grade_Found and AnySiteGradeFound.
            '''//If EnableWord "Vddbin_DoAll_DebugCollection" is enabled, initial performance mode result for Char. BinCut search voltage(Grade) and efuse product voltage(GradeVdd) as 0.
            Call ResetPmodePowerforBincut(inst_info)
        
            '''//Get the IDS value of PowerDomain for the binning p_mode.
            '''//IDS calculation uses the scale and the unit in "mA".
            '''20210803: Modified to update inst_info.ids_current = IDS_for_BinCut(VddBinStr2Enum(powerDomain)).Real in the vbt function initialize_inst_info and remove the redundant vbt function set_IDS_current.
            If inst_info.powerDomain <> "" Then
            
                If Flag_IDS_Mapping_enable = True Then 'WWW20220630 for ids mapping
                    inst_info.ids_current = IDS_for_BinCut(VddBinStr2Enum(AllBinCut(inst_info.p_mode).IDS_MAPPING)).Real
                Else
                    inst_info.ids_current = IDS_for_BinCut(VddBinStr2Enum(inst_info.powerDomain)).Real '''unit: mA
                End If
                
                
                '''//Check if IDS values of the binning powerDomain > 0.
                For Each site In TheExec.sites
                    If inst_info.ids_current(site) <= 0 Then
                        Call Print_Error_Message(Error_Info, "LIB_VDD_BINNING", "decide_binSearch_and_start_voltage", "site:" & site & ",Instance:" & inst_info.inst_name & ", it doesn't get the correct IDS value for the binning performance mode:" & inst_info.performance_mode & ". Please check the argument about performance mode for the instance and IDS values from DC and Efuse. Error!!!")
                        'TheExec.ErrorLogMessage "site:" & site & ",Instance:" & inst_info.inst_name & ", it can't get the correct powerDomain for the binning performance mode:" & inst_info.performance_mode & ". Please check the argument about performance mode for the instance and IDS values from DC and Efuse. Error!!!"    ''20230807: Removed by sanity check
                    End If
                Next site
            Else
                Call Print_Error_Message(Error_Info, "LIB_VDD_BINNING", "decide_binSearch_and_start_voltage", "Instance:" & inst_info.inst_name & ", it can't get the correct powerDomain for the binning performance mode:" & inst_info.performance_mode & ". Please check the argument about performance mode for the instance. Error!!!")
                'TheExec.ErrorLogMessage "Instance:" & inst_info.inst_name & ", it can't get the correct powerDomain for the binning performance mode:" & inst_info.performance_mode & ". Please check the argument about performance mode for the instance. Error!!!"    ''20230807: Removed by sanity check
            End If
            
            '''//Find the start voltage, stop voltage, algorithm.
            '''===========================================================================================================================================
            '''Start serach the LVCC based on IDS Zone, the VDD_BIN_ALL(P_mode).MODE_STEP means the maximum EQ count for all BinCut Tables.
            '''Before the searching start, we had based on the IDS current to find out the IDS Zone number, search algorithm and start step in the IDS Zone.
            '''We use the stop voltage to be the stop EQN and when the Grade_Found_Mask = all site, we exit the loop to judge the PF for all sites.
            '''===========================================================================================================================================
            find_start_voltage inst_info
            
            '''//Decide max step for stepcount-loop.
            inst_info.maxStep = AllBinCut(inst_info.p_mode).Mode_Step
        End If
    ElseIf inst_info.is_BinSearch = False Or FuncTestOnly = True Then '''Only do functional test.
        '''//Decide max step for stepcount-loop.
        inst_info.maxStep = 0
    Else '''For the undefined BinCut settings...
        inst_info.maxStep = -1
    End If
Exit Function
errHandler:
    Call Print_Error_Message(Error_Info, "LIB_VDD_BINNING", "decide_binSearch_and_start_voltage") 'Add ErrHandler 2023/08/01
    'TheExec.ErrorLogMessage "Error encountered in VBT Function of decide_binSearch_and_start_voltage"
    If AbortTest Then Exit Function Else Resume Next
End Function

'20191127: Modified for the revised InitVddBinTable.
Public Function IsExcludedVddBin(p_mode As Integer) As Boolean
On Error GoTo errHandler
    If BinCut(p_mode, CurrentPassBinCutNum).ExcludedPmode = True Then
        IsExcludedVddBin = True '''The Performance mode doesn't exist in the CurrentPassBinCutNum.
    Else
        IsExcludedVddBin = False '''The Performance mode exists in the CurrentPassBinCutNum.
    End If
Exit Function
errHandler:
    Call Print_Error_Message(Error_Info, "LIB_VDD_BINNING", "IsExcludedVddBin") 'Add ErrHandler 2023/08/01
    If AbortTest Then Exit Function Else Resume Next
End Function

'20191127: Modified for the revised InitVddBinTable.
Public Function SkipTestBin2Site(p_mode As Integer, Active_site_count As Long)
On Error GoTo errHandler 'Add ErrHandler 2023/08/01
    Dim site As Variant
    Dim EnableSites As New SiteBoolean
'''//==================================================================================================================================================================================//'''
'''//Note:
'''PassBinCutNum is global variable and will be decided if this site already run into Bin2.
'''If this site has been in Bin2, and Bin2 doesn't have this performance mode in BinCut voltage table, this site will be disabled.
'''//==================================================================================================================================================================================//'''
    RestoredSites = TheExec.sites.Selected
    EnableSites = TheExec.sites.Selected
    Active_site_count = 0
    '''**********************************************************************************************************************************************************'''
    '''If the CurrentPassBinCutNum is Bin2 but the current performance mode does not exist in the Bin2 Table, set th site disable and skip in this Test Instance.
    '''**********************************************************************************************************************************************************'''
    For Each site In TheExec.sites
        If IsExcludedVddBin(p_mode) Then '''Do not test the performance mode which do not exist in the Bin2 table if it is Bin2 device.
            EnableSites(site) = False
        Else
            Active_site_count = Active_site_count + 1
        End If
    Next site
    TheExec.sites.Selected = EnableSites
Exit Function
errHandler:
    Call Print_Error_Message(Error_Info, "LIB_VDD_BINNING", "SkipTestBin2Site") 'Add ErrHandler 2023/08/01
    If AbortTest Then Exit Function Else Resume Next
End Function

'20191127: Modified for the revised InitVddBinTable.
Public Function RestoreSkipTestBin2Site(p_mode As Integer)
On Error GoTo errHandler
    'if IsExcludedVddBin(P_mode) Then
        TheExec.sites.Selected = RestoredSites
    'End If
Exit Function
errHandler:
    Call Print_Error_Message(Error_Info, "LIB_VDD_BINNING", "RestoreSkipTestBin2Site") 'Add ErrHandler 2023/08/01
    If AbortTest Then Exit Function Else Resume Next
End Function

'20191127: Modified for the revised InitVddBinTable.
'20191126: Modified to use the dictionary to store powerPin and pmode.
'20190516: Modified for temporary use due the incosistent pin name in the BinCut flow and EFUSE_BitDef_Table.
Public Function VddBinStr2Enum(performance_mode As String) As Integer
On Error GoTo errHandler 'Add ErrHandler 2023/08/01
    Dim strTemp As String
    strTemp = UCase(performance_mode)

    If VddbinPmodeDict.Exists(strTemp) Then
        VddBinStr2Enum = CInt(VddbinPmodeDict.item(strTemp))
    Else
        VddBinStr2Enum = cntVddbinPmode + 1
         Call Print_Error_Message(Error_Info, "LIB_VDD_BINNING", "VddBinStr2Enum", performance_mode & " doesn't have the matched definition of Enum p_mode in VddBinStr2Enum. Error!!!")
        'TheExec.Datalog.WriteComment performance_mode & " doesn't have the matched definition of Enum p_mode in VddBinStr2Enum. Error!!!"
        'TheExec.ErrorLogMessage performance_mode & " doesn't have the matched definition of Enum p_mode in VddBinStr2Enum. Error!!!"
    End If
Exit Function
errHandler:
    Call Print_Error_Message(Error_Info, "LIB_VDD_BINNING", "VddBinStr2Enum") 'Add ErrHandler 2023/08/01
    If AbortTest Then Exit Function Else Resume Next
End Function

'20210719: Modified to check AllBinCut(p_mode).Mode_Step with TotalStepPerMode, as requested by ZYLINI and ZQLIN.
'20210414: Modified to add "is_for_BinSearch as Boolean" for AllBinCut(p_mode).
'20210312: Modified to check BinCutList in Vdd_Binning_Def tables.
'20210106: Modified to support the new format with "+" of BinCutList in Vdd_Binning_Def tables, requested by AutoGen team.
'20201021: Modified to use "dict_IsCorePower" to store and check CorePower/OtherRail.
'20200703: Modiifed to use "check_Sheet_Range".
'20200528: Modified to check header of the table.
'20200423: Modified to remove the unused argument "col_lvcc as Long".
'20200421: Modified to remove "Init AllBinCut(p_mode).allow_equal".
'20200421: Modified to check the column of "CPIDSMax".
'20200415: Modified to check "col_soft_bin".
'20191127: Modified for the revised InitVddBinTable.
'20191126: Modified to use the dictionary to store powerDomain and pmode.
'20191125: Modified to check if the items of the header exist in the table.
'20190426: Modified to use the function "Find_Sheet".
'20190321: Modified to add the utility to check if the sheet "Vdd_Binning_Def_appA_1" and bincutlist exist.
Public Function initVddBinTable()
On Error GoTo errHandler 'Add ErrHandler 2023/08/01
    Dim wb As Workbook
    Dim ws_def As Worksheet
    Dim sheetName As String
    Dim col_binned As Integer
    Dim col_domain As Integer
    Dim col_eqn As Integer '20230526
    Dim col_mode As Integer
    Dim col_cpids As Integer
    Dim col_sort As Integer
    Dim str_PassBinCut As String
    Dim bincutNum As Variant
    Dim p_mode As Integer
    Dim strAry_PassBinCut() As String
    Dim passBinCut As Variant
    Dim i As Long
    Dim Row As Long, Col As Long
    Dim MaxRow As Long
    Dim maxcol As Long
    Dim Row_of_BasicInfo As Integer
    Dim Row_of_AdditionalInfo As Integer
    Dim Row_of_Title As Integer
    Dim powerDomain As String
    Dim pmodeName As String
    Dim str_pmode_FullName As String
    Dim PinTemp As String
    Dim pmodeTemp As String
    Dim pmodeAllTemp As String
    Dim split_array0() As String
    Dim split_array1() As String
    Dim idxArray As Long
    Dim enableRowParsing As Boolean
    Dim isSheetFound As Boolean
'''//==================================================================================================================================================================================//'''
'''//Note:
'''Sheets "Vdd_Binning_Def" (for CorePower) and "Other_Rail" (for OtherRail) are merged into sheets "Vdd_Binning_Def" (for CorePower and OtherRail).
'''//==================================================================================================================================================================================//'''
    '''*****************************************************************'''
    '''//Check if the sheet exists
    sheetName = "Vdd_Binning_Def_appA_1_" & TheExec.CurrentJob
    Set wb = Application.ActiveWorkbook
    Call check_Sheet_Range(sheetName, wb, ws_def, MaxRow, maxcol, isSheetFound, False)
    If isSheetFound = False Then        ''' if "Vdd_Binning_Def_appA_1_JOBNAME" does not exist
        sheetName = "Vdd_Binning_Def_appA_1"
        Call check_Sheet_Range(sheetName, wb, ws_def, MaxRow, maxcol, isSheetFound)
    End If
    '''*****************************************************************'''
    If isSheetFound = True Then
        '''//init
        '''Since all col_XXX and row_XXX related variables with default values=0, no need to initialize them as 0.
        Version_Vdd_Binning_Def = vbNullString
        str_PassBinCut = vbNullString
        BV_StepVoltage = 0
        VddbinningBaseVoltage = 0
        Total_Bincut_Num = 0  'initialize the variable
        i = 0
        pmodeName = vbNullString
        str_pmode_FullName = vbNullString
        PinTemp = vbNullString
        pmodeTemp = vbNullString
        pmodeAllTemp = vbNullString
        cntVddbinPin = 0
        cntVddbinPmode = 0
        idxArray = -1
        enableRowParsing = False
        
        For Row = 1 To MaxRow
            For Col = 1 To maxcol
                If LCase(ws_def.Cells(Row, Col).value) Like LCase("Rev*") Then '''//Revision of BinCut tables
                    Version_Vdd_Binning_Def = LCase(ws_def.Cells(Row, Col + 1).value)
                    Row_of_BasicInfo = Row
                End If
                
                If Row_of_BasicInfo > 0 Then
                    If LCase(ws_def.Cells(Row_of_BasicInfo, Col).value) Like LCase("Bin*Cut*List*") Then '''//Number of BinCut tables
                        str_PassBinCut = LCase(ws_def.Cells(Row_of_BasicInfo, Col + 1).value)
                        
                        '''//Check if the sheet Vdd_Binning_Def_appA_1 contains the correct BinCutList.
                        '''If that, parse BinCutList to get BinCut number.
                        If str_PassBinCut <> "" Then
                            '''20210106: Modified to support the new format with "+" of BinCutList in Vdd_Binning_Def tables, requested by AutoGen team.
                            If str_PassBinCut Like "*,*" Then
                                strAry_PassBinCut = Split(str_PassBinCut, ",")   '//BinCut number
                            Else
                                strAry_PassBinCut = Split(str_PassBinCut, "+")   '//BinCut number
                            End If
                            
                            '''//Parse the string in the cell to decide PassBinCut
                            ReDim PassBinCut_ary(UBound(strAry_PassBinCut)) '''//redefine size of array PassBinCut_ary().
                            Total_Bincut_Num = strAry_PassBinCut(UBound(strAry_PassBinCut)) '20230526
                            For Each bincutNum In strAry_PassBinCut '''put how many bincut tables to the array
                                idxArray = idxArray + 1
                                
                                '''//Check PassBin number from BinCutList in Vdd_Binning_Def tables.
                                If bincutNum = CStr(idxArray + 1) Then
                                    PassBinCut_ary(idxArray) = CLng(bincutNum)
                                Else
                                    Call Print_Error_Message(Error_Info, "LIB_VDD_BINNING", "initVddBinTable", "BinCutList:" & str_PassBinCut & " of sheet:" & sheetName & " doesn't have any correct sequence of BinCut passBin numbers. Error!!!")
                                    'TheExec.ErrorLogMessage "BinCutList:" & str_PassBinCut & " of sheet:" & sheetName & " doesn't have any correct sequence of BinCut passBin numbers. Error!!!"    ''20230807: Removed by sanity check
                                End If
                            Next bincutNum
                        Else
                            Call Print_Error_Message(Error_Info, "LIB_VDD_BINNING", "initVddBinTable", sheetName & " doesn't contain any correct BinCutList. Error!!!")
                            'TheExec.Datalog.WriteComment sheetName & " doesn't contain any correct BinCutList. Error!!!"
                            'TheExec.ErrorLogMessage sheetName & " doesn't contain any correct BinCutList. Error!!!"    ''20230807: Removed by sanity check
                        End If
                    ElseIf LCase(ws_def.Cells(Row_of_BasicInfo, Col).value) Like LCase("col_soft_bin*") Then '''//Stat column of sort bin
                        col_sort = CLng(ws_def.Cells(Row_of_BasicInfo, Col + 1).value)
                    End If
                End If
            Next Col
            MaxPassBinCut = UBound(PassBinCut_ary) + 1 '20230626 redim MaxPassBinCut
                
            If Row_of_BasicInfo > 0 Then
                If idxArray > -1 And col_sort > 0 Then
                    '''Do nothing...
                Else
                    Row_of_BasicInfo = 0
                    If idxArray = -1 Then
                        Call Print_Error_Message(Error_Info, "LIB_VDD_BINNING", "initVddBinTable", "Column Bin Cut List doesn't exist in header of " & sheetName & ". Error!!!")
                        'TheExec.ErrorLogMessage "Column Bin Cut List doesn't exist in header of " & sheetName & ". Error!!!"    ''20230807: Removed by sanity check
                    End If
                    
                    If col_sort = 0 Then
                        Call Print_Error_Message(Error_Info, "LIB_VDD_BINNING", "initVddBinTable", "Column col_soft_bin doesn't exist in header of " & sheetName & ". Error!!!")
                        'TheExec.ErrorLogMessage "Column col_soft_bin doesn't exist in header of " & sheetName & ". Error!!!"    ''20230807: Removed by sanity check
                    End If
                End If
                
                Exit For
            End If
        Next Row

        If Row_of_BasicInfo > 0 Then
            For Row = Row_of_BasicInfo + 1 To MaxRow
                For Col = 1 To maxcol
                    If LCase(ws_def.Cells(Row, Col).value) Like LCase("Base*Voltage*") Then '''//Base Voltage
                        VddbinningBaseVoltage = CDbl(ws_def.Cells(Row, Col + 1).value)
                        Row_of_AdditionalInfo = Row
                    End If

                    If Row_of_AdditionalInfo > 0 Then
                        If LCase(ws_def.Cells(Row_of_AdditionalInfo, Col).value) Like LCase("Step*Size*") Then '''//Step Size
                            BV_StepVoltage = CDbl(ws_def.Cells(Row_of_AdditionalInfo, Col + 1).value)
                            
                            '''//Check if StepVoltage from Vdd_Binning_Def matches the definition of "gC_StepVoltage" in globalVariable.
                            If BV_StepVoltage = gC_StepVoltage Then
                                TheExec.Datalog.WriteComment "Step Size Voltage and gC_StepVoltage are " & BV_StepVoltage
                            Else
                                BV_StepVoltage = 0
                                Call Print_Error_Message(Error_Info, "LIB_VDD_BINNING", "initVddBinTable", "The Step Size Voltage in BinCut voltage table = " & BV_StepVoltage & ", The Gc_Stepvoltage = " & gC_StepVoltage & ". Error!!!")
                                'TheExec.ErrorLogMessage "The Step Size Voltage in BinCut voltage table = " & BV_StepVoltage & ", The Gc_Stepvoltage = " & gC_StepVoltage & ". Error!!!"    ''20230807: Removed by sanity check
                            End If
                        End If
                    End If
                Next Col
                    
                '''//Row of the header
                If Row_of_AdditionalInfo > 0 Then
                    If BV_StepVoltage > 0 Then
                        '''Do nothing...
                    Else
                        Row_of_AdditionalInfo = 0
                        Call Print_Error_Message(Error_Info, "LIB_VDD_BINNING", "initVddBinTable", "Column Base Voltage doesn't exist in header of " & sheetName & ". Error!!!")
                        'TheExec.ErrorLogMessage "Column Base Voltage doesn't exist in header of " & sheetName & ". Error!!!"    ''20230807: Removed by sanity check
                    End If
                    
                    Exit For
                End If
            Next Row
        Else
            Call Print_Error_Message(Error_Info, "LIB_VDD_BINNING", "initVddBinTable", "Column Rev doesn't exist in header of " & sheetName & ". Error!!!")
            'TheExec.ErrorLogMessage "Column Rev doesn't exist in header of " & sheetName & ". Error!!!"    ''20230807: Removed by sanity check
        End If

        If Row_of_AdditionalInfo > 0 Then
            For Row = Row_of_AdditionalInfo + 1 To MaxRow
                For Col = 1 To maxcol
                    If LCase(ws_def.Cells(Row, Col).value) = LCase("Binned") Then
                        col_binned = Col
                        Row_of_Title = Row
                    End If

                    If Row_of_Title > 0 Then
                        '''//Check if the items of the header exist in the table.
                        If LCase(ws_def.Cells(Row_of_Title, Col).value) = LCase("Domain") Then '''//Domain
                            col_domain = Col
                        ElseIf LCase(ws_def.Cells(Row_of_Title, Col).value) = LCase("Mode") Then '''//Mode
                            col_mode = Col
                        ElseIf LCase(ws_def.Cells(Row_of_Title, Col).value) = LCase("EQN") Then '20230526
                            col_eqn = Col
                        '''//Check the column of "CPIDSMax".
                        ElseIf LCase(ws_def.Cells(Row_of_Title, Col).value) = "cpidsmax" Then '''//CP IDS limit
                            col_cpids = Col
                        End If

                        '''//Check col_soft_bin
                        If col_sort > 0 Then
                            If LCase(ws_def.Cells(Row_of_Title, col_sort - 1).value) = LCase("comment") Then '''//Comment
                                '''Do nothing
                            Else
                                Call Print_Error_Message(Error_Info, "LIB_VDD_BINNING", "initVddBinTable", "col_soft_bin " & col_sort & " doesn't match the start column of sort bin in " & sheetName & ". Error!!!")
                                'TheExec.ErrorLogMessage "col_soft_bin " & col_sort & " doesn't match the start column of sort bin in " & sheetName & ". Error!!!"    ''20230807: Removed by sanity check
                            End If
                        Else
                            Call Print_Error_Message(Error_Info, "LIB_VDD_BINNING", "initVddBinTable", sheetName & " doesn't contain the correct column position of col_soft_bin. Error!!!")
                            'TheExec.Datalog.WriteComment sheetName & " doesn't contain the correct column position of col_soft_bin. Error!!!"
                            'TheExec.ErrorLogMessage sheetName & " doesn't contain the correct column position of col_soft_bin. Error!!!"    ''20230807: Removed by sanity check
                        End If
                    End If
                Next Col

                '''//If items are found, exit the for-loop.
                If Row_of_Title > 0 Then
                    If col_domain > 0 And col_mode > 0 And col_sort > 0 And LCase(ws_def.Cells(Row_of_Title, col_sort - 1).value) = LCase("comment") Then
                        enableRowParsing = True
                    Else
                        enableRowParsing = False
                        If col_binned = 0 Then
                            Call Print_Error_Message(Error_Info, "LIB_VDD_BINNING", "initVddBinTable", "Column col_soft_bin doesn't exist in header of " & sheetName & ". Error!!!")
                            'TheExec.ErrorLogMessage "Column col_soft_bin doesn't exist in header of " & sheetName & ". Error!!!"    ''20230807: Removed by sanity check
                        End If
                        
                        If col_domain = 0 Then
                            Call Print_Error_Message(Error_Info, "LIB_VDD_BINNING", "initVddBinTable", "Column Domain doesn't exist in header of " & sheetName & ". Error!!!")
                            'TheExec.ErrorLogMessage "Column Domain doesn't exist in header of " & sheetName & ". Error!!!"    ''20230807: Removed by sanity check
                        End If
                        
                        If col_mode = 0 Then
                            Call Print_Error_Message(Error_Info, "LIB_VDD_BINNING", "initVddBinTable", "Column Mode doesn't exist in header of " & sheetName & ". Error!!!")
                            'TheExec.ErrorLogMessage "Column Mode doesn't exist in header of " & sheetName & ". Error!!!"    ''20230807: Removed by sanity check
                        End If
                        
                        If LCase(ws_def.Cells(Row_of_Title, col_sort - 1).value) <> LCase("comment") Then
                            Call Print_Error_Message(Error_Info, "LIB_VDD_BINNING", "initVddBinTable", "Column Softbin doesn't exist in header of " & sheetName & ". Error!!!")
                            'TheExec.ErrorLogMessage "Column Softbin doesn't exist in header of " & sheetName & ". Error!!!"    ''20230807: Removed by sanity check
                        End If
                        
                        If col_cpids = 0 Then
                            Call Print_Error_Message(Error_Info, "LIB_VDD_BINNING", "initVddBinTable", "Column CPIDSMax doesn't exist in header of " & sheetName & ". Error!!!")
                            'TheExec.ErrorLogMessage "Column CPIDSMax doesn't exist in header of " & sheetName & ". Error!!!"    ''20230807: Removed by sanity check
                        End If
                    End If
                    
                    Exit For
                Else
                    Call Print_Error_Message(Error_Info, "LIB_VDD_BINNING", "initVddBinTable", "Column Binned doesn't exist in header of " & sheetName & ". Error!!!")
                    'TheExec.ErrorLogMessage "Column Binned doesn't exist in header of " & sheetName & ". Error!!!"    ''20230807: Removed by sanity check
                End If
            Next Row
        Else
            Call Print_Error_Message(Error_Info, "LIB_VDD_BINNING", "initVddBinTable", "Column Base Voltage doesn't exist in header of " & sheetName & ". Error!!!")
            'TheExec.ErrorLogMessage "Column Base Voltage doesn't exist in header of " & sheetName & ". Error!!!"    ''20230807: Removed by sanity check
        End If

        If enableRowParsing = True Then
            '''//Parse the table to enumerate powerDomain(from column Domain) and pmode(from column Mode) for BinCut.
            For Row = Row_of_Title + 1 To MaxRow
                If ws_def.Cells(Row, col_domain).value <> "" And ws_def.Cells(Row, col_mode).value <> "" Then
                    '''//column binned="true" is CorePower, and column binned="false" or "ate" is OtherRail.
                    If LCase(Trim(ws_def.Cells(Row, col_binned).value)) = "true" _
                    Or LCase(Trim(ws_def.Cells(Row, col_binned).value)) = "ate" _
                    Or LCase(Trim(ws_def.Cells(Row, col_binned).value)) = "false" Then
                        '''******************************************************'''
                        '''//Create the dictionary to store powerDomain and pmode
                        '''******************************************************'''
                        '''//powerDomain
                        If UCase(ws_def.Cells(Row, col_domain).value) Like UCase("VDD_*") Then
                            powerDomain = UCase(ws_def.Cells(Row, col_domain).value)
                        Else
                            powerDomain = UCase("VDD_" & ws_def.Cells(Row, col_domain).value)
                        End If
                        
                        '''//Pmode name
                        pmodeName = UCase(UCase(ws_def.Cells(Row, col_mode).value))
                        
                        '''//Full Pmode name
                        str_pmode_FullName = powerDomain & "_" & UCase(ws_def.Cells(Row, col_mode).value)
                        
                        '''//Add powerDomain and Pmode to the temporary group
                        If PinTemp <> "" Then
                            If LCase("*," & PinTemp & ",*") Like LCase("*," & powerDomain & ",*") Then
                                '''Do nothing
                            Else
                                PinTemp = PinTemp & "," & powerDomain
                            End If
                        Else
                            PinTemp = powerDomain
                        End If
                        
                        '''//Store DomainType (CorePower or OtherRail of PowerDomain into the dictionary "dict_IsCorePower".
                        If dict_IsCorePower.Exists(UCase(powerDomain)) = True Then
                            '''Do nothing...
                        Else
                            If LCase(Trim(ws_def.Cells(Row, col_binned).value)) = "true" Then
                                dict_IsCorePower.Add UCase(powerDomain), True
                            ElseIf LCase(Trim(ws_def.Cells(Row, col_binned).value)) = "ate" Or LCase(Trim(ws_def.Cells(Row, col_binned).value)) = "false" Then
                                dict_IsCorePower.Add UCase(powerDomain), False
                            End If
                        End If

                        '''//Only pmode of CorePower
                        If LCase(Trim(ws_def.Cells(Row, col_binned).value)) = "true" Then
                            If pmodeTemp <> "" Then
                                If LCase("*," & pmodeTemp & ",*") Like LCase("*," & pmodeName & ",*") Then
                                    '''Do nothing
                                Else
                                    pmodeTemp = pmodeTemp & "," & pmodeName
                                End If
                            Else
                                pmodeTemp = pmodeName
                            End If
                            
                            If pmodeAllTemp <> "" Then
                                If LCase("*," & pmodeAllTemp & ",*") Like LCase("*," & str_pmode_FullName & ",*") Then
                                    '''Do nothing
                                Else
                                    pmodeAllTemp = pmodeAllTemp & "," & str_pmode_FullName
                                End If
                            Else
                                pmodeAllTemp = str_pmode_FullName
                            End If
                        End If
                    End If
                    If MaxEqnNum < CLng(Replace(UCase(ws_def.Cells(Row, col_eqn).value), "E", "")) Then '20230526
                        MaxEqnNum = CLng(Replace(UCase(ws_def.Cells(Row, col_eqn).value), "E", ""))
                    End If
                End If
            Next Row
        Else
            Call Print_Error_Message(Error_Info, "LIB_VDD_BINNING", "initVddBinTable", sheetName & " doesn't contain the correct columns of Binned, Domain, and Mode. Error!!!")
            'TheExec.Datalog.WriteComment sheetName & " doesn't contain the correct columns of Binned, Domain, and Mode. Error!!!"
            'TheExec.ErrorLogMessage sheetName & " doesn't contain the correct columns of Binned, Domain, and Mode. Error!!!"    ''20230807: Removed by sanity check
        End If
        
        '''//Create the dictionary for pmode (replace str2enum)
        '''We also fill the array "VddBinName" (replace enum2str).
        If PinTemp <> "" And pmodeTemp <> "" And pmodeAllTemp <> "" Then
            split_array0 = Split(PinTemp, ",")
            For i = 0 To UBound(split_array0)
                If Not (VddbinPinDict.Exists(split_array0(i))) Then
                    cntVddbinPin = cntVddbinPin + 1
                    VddbinPinDict.Add split_array0(i), cntVddbinPin
                End If
            Next i
            
            pmodeTemp = PinTemp & "," & pmodeTemp
            pmodeAllTemp = PinTemp & "," & pmodeAllTemp
            split_array0 = Split(pmodeTemp, ",")
            split_array1 = Split(pmodeAllTemp, ",")
            
            If UBound(split_array0) = UBound(split_array1) Then
                For i = 0 To UBound(split_array1)
                    If Not (VddbinPmodeDict.Exists(split_array1(i))) Then
                        cntVddbinPmode = cntVddbinPmode + 1
                        VddbinPmodeDict.Add split_array1(i), cntVddbinPmode
                        
                        ReDim Preserve VddBinName(cntVddbinPmode)
                        VddBinName(cntVddbinPmode) = split_array1(i)
                        
                        If Not (VddbinPmodeDict.Exists(split_array0(i))) Then
                            VddbinPmodeDict.Add split_array0(i), cntVddbinPmode
                        End If
                    End If
                Next i
            Else
                Call Print_Error_Message(Error_Info, "LIB_VDD_BINNING", "initVddBinTable", sheetName & " doesn't contain the correct format of Domain and Mode. Error!!!")
                'TheExec.Datalog.WriteComment sheetName & " doesn't contain the correct format of Domain and Mode. Error!!!"
                'TheExec.ErrorLogMessage sheetName & " doesn't contain the correct format of Domain and Mode. Error!!!"    ''20230807: Removed by sanity check
            End If
        End If

        '''MaxBincutPowerdomainCount
        
        initDynamicArray
        
        '''initilize the MODE_STEP for AllBinCut
        For p_mode = 0 To MaxPerformanceModeCount - 1
            AllBinCut(p_mode).Mode_Step = 0
            AllBinCut(p_mode).is_for_BinSearch = False
        Next p_mode
        
        For Each passBinCut In PassBinCut_ary
            '''Parsing Vdd_Binning_Def sheets for CorePower
            initVddBinTableOneMod CLng(passBinCut), col_cpids, col_sort
            
            '''Parsing Vdd_Binning_Def sheets for OtherRail
            initVddotherrailOneMod CLng(passBinCut)
            
            For p_mode = 0 To MaxPerformanceModeCount - 1
                '''20210719: Modified to check BinCut(p_mode, passBinCut).Mode_Step with TotalStepPerMode, as requested by ZYLINI and ZQLIN.
                If BinCut(p_mode, passBinCut).Mode_Step > MaxEqnNum Then
                    Call Print_Error_Message(Error_Info, "LIB_VDD_BINNING", "", "bin" & CLng(passBinCut) & "," & VddBinName(p_mode) & ", it has steps(EQNs)=" & AllBinCut(p_mode).Mode_Step & ", but it is greater than BinCut globalVariable TotalStepPerMode=" & Max_IDS_Step & ", please check check tables Vdd_Binning_Def and update TotalStepPerMode. Error!!!")
                    'TheExec.ErrorLogMessage "bin" & CLng(passBinCut) & "," & VddBinName(p_mode) & ", it has steps(EQNs)=" & AllBinCut(p_mode).Mode_Step & ", but it is greater than BinCut globalVariable TotalStepPerMode=" & TotalStepPerMode & ", please check check tables Vdd_Binning_Def and update TotalStepPerMode. Error!!!"    ''20230807: Removed by sanity check
                ElseIf BinCut(p_mode, passBinCut).Mode_Step >= 0 Then '''if the bincut exists, add to AllBinCut
                    AllBinCut(p_mode).Mode_Step = AllBinCut(p_mode).Mode_Step + BinCut(p_mode, passBinCut).Mode_Step + 1
                End If
            Next p_mode
            
            '20230504: Added for parsing sheet defining BinX EQNs and IDS value
            If Flag_BinX_Info_Parsed = True And passBinCut > 1 Then
                TheExec.Datalog.WriteComment "Header 'EQN_Bin' & 'BinX_IDSMax' exists, can only support bin1 sheet!!! Please disable non-bin1 sheet 'Bin Cut List =1,2'->'Bin Cut List =1', Error!!!"
                TheExec.AddOutput "Header 'EQN_Bin' & 'BinX_IDSMax' exists, can only support bin1 sheet!!! Please disable non-bin1 sheet 'Bin Cut List =1,2'->'Bin Cut List =1', Error!!!"
            End If
        Next passBinCut
        
        '''Corrected for max step of per mode
        For p_mode = 0 To MaxPerformanceModeCount - 1
            AllBinCut(p_mode).Mode_Step = AllBinCut(p_mode).Mode_Step - 1
            
            '''//Check if max step of p_mode is greater than TotalStepPerMode.
            '''20210719: Modified to check AllBinCut(p_mode).Mode_Step with TotalStepPerMode, as requested by ZYLINI and ZQLIN.
            If AllBinCut(p_mode).Mode_Step > MaxEqnNum Then
                Call Print_Error_Message(Error_Info, "LIB_VDD_BINNING", "initVddBinTable", VddBinName(p_mode) & ", all BinCut voltage tables have steps(EQNs)=" & AllBinCut(p_mode).Mode_Step & ", it is greater than BinCut globalVariable Max_IDS_Step=" & Max_IDS_Step & ", please check check tables Vdd_Binning_Def and update Max_IDS_Step. Error!!!")
                'TheExec.Datalog.WriteComment VddBinName(p_mode) & ", all BinCut voltage tables have steps(EQNs)=" & AllBinCut(p_mode).Mode_Step & ", it is greater than BinCut globalVariable Max_IDS_Step=" & Max_IDS_Step & ", please check check tables Vdd_Binning_Def and update Max_IDS_Step. Error!!!"
                'TheExec.ErrorLogMessage VddBinName(p_mode) & ", all BinCut voltage tables have steps(EQNs)=" & AllBinCut(p_mode).Mode_Step & ", it is greater than BinCut globalVariable Max_IDS_Step=" & Max_IDS_Step & ", please check check tables Vdd_Binning_Def and update Max_IDS_Step. Error!!!"    ''20230807: Removed by sanity check
            End If
        Next p_mode
        
        '''//Check if p_mode is ExcludedPmode.
        For Each passBinCut In PassBinCut_ary
            For p_mode = 0 To MaxPerformanceModeCount - 1
                If BinCut(p_mode, passBinCut).ExcludedPmode = True And ExcludedPmode(p_mode) = False Then
                    Call Print_Error_Message(Error_Info, "LIB_VDD_BINNING", "initVddBinTable", "Test performance mode " & VddBinName(p_mode) & " doesn't exist in BinCut " & passBinCut - 1 & ". Error!!!")
                    'TheExec.ErrorLogMessage "Test performance mode " & VddBinName(p_mode) & " doesn't exist in BinCut " & passBinCut - 1 & ". Error!!!"    ''20230807: Removed by sanity check
                End If
            Next p_mode
        Next passBinCut
    End If '''If isSheetFound = True
Exit Function
errHandler:
    Call Print_Error_Message(Error_Info, "LIB_VDD_BINNING", "initVddBinTable") 'Add ErrHandler 2023/08/01
    If AbortTest Then Exit Function Else Resume Next
End Function

'20191227: Modified to remove checking AllBinCut(pmode).Used=true.
'20191127: Modified for the revised InitVddBinTable.
Public Function InitVddBinInherit(Power_Seq() As String)
On Error GoTo errHandler 'Add ErrHandler 2023/08/01
    Dim i As Integer
    For i = 0 To UBound(Power_Seq)
        If i = 0 Then
            AllBinCut(VddBinStr2Enum(Power_Seq(i))).PREVIOUS_Performance_Mode = cntVddbinPmode + 1
        Else
            AllBinCut(VddBinStr2Enum(Power_Seq(i))).PREVIOUS_Performance_Mode = VddBinStr2Enum(Power_Seq(i - 1))
        End If
    Next i
Exit Function
errHandler:
    Call Print_Error_Message(Error_Info, "LIB_VDD_BINNING", "InitVddBinInherit") 'Add ErrHandler 2023/08/01
    If AbortTest Then Exit Function Else Resume Next
End Function

'20210530: Modified to replace typo "Multisftp_Binout" with "MultiFstp_NoBinout".
'20210514: Modified to overwrite failStop if failflag "MultiFstp_NoBinout" is enabled for MultiFSTP.
'20191127: Modified for the revised InitVddBinTable.
'20190422: Modified to check if alarmFail(site) is triggered or not.
Public Function judge_PF_func(p_mode As Integer, Test_Type As testType, patt_result As SiteBoolean, Optional isPostBV As Boolean = False, Optional InstName As String = vbNullString)
On Error GoTo errHandler 'Add ErrHandler 2023/08/01
    Dim site As Variant
    Dim inst_name As String
'''//==================================================================================================================================================================================//'''
'''//Note:
'''20210514: Modified to overwrite failStop if failflag "MultiFstp_NoBinout" is enabled for MultiFSTP.
'''We modified the vbt function to mask the failed site by theExec.sites.Selected and trig the failFlag if theExec.sites.item(site).FlagState("MultiFstp_NoBinout") = logicTrue for MultiFSTP instance.
'''//==================================================================================================================================================================================//'''
    If TheExec.flow.enableWord("Vddbin_Skip_BinCheck_Enable") = True Then
        Exit Function
    Else
        If glb_TesterType = "UltraFLEXplus" Then
            Dim alarmOccurred As New SiteBoolean
            alarmOccurred = False
            alarmOccurred = TheHdw.Alarms.GetAlarmingSites(True)        ''GetAlarmingSites(clearAlarm = True), only plus or above 10_30_90 can use
            For Each site In TheExec.sites
                If alarmOccurred(site) = True Then
                    alarmFail(site) = True      ''Update alarm info to alarmFail array for UFP
                End If
            Next site
        End If
            
    
        For Each site In TheExec.sites
        #If IGXL_VER_1030 = True Then
            If patt_result(site) = False Or alarmOccurred(site) = True Then
        #Else
            If patt_result(site) = False Or alarmFail(site) = True Then
        #End If
                    '''//If the flag "MultiFstp_NoBinout" = True, skip SortNumber and fail-stop for MultiFSTP instances.
        '''ToDo: Please check if the failFlag ""MultiFstp_NoBinout"" exists in the flow table!!!
            If TheExec.sites.item(site).FlagState("MultiFstp_NoBinout") = logicTrue Then 'MultiFstp without Binout
                inst_name = TheExec.DataManager.instancename
                TheExec.Datalog.WriteComment "Site:" & site & "," & inst_name & ", test failed, but MultiFSTP bypassed BinOut!"
            ''20220922: Coll FT2 uses shadow voltage, no binout
            ElseIf bincutJobName = "ft_hot" And Flag_Vddbin_FT2_Shadow_NoBinout = True Then
                                inst_name = TheExec.DataManager.instancename
                TheExec.Datalog.WriteComment "Site:" & site & "," & inst_name & ", test failed, but FT2 bypassed BinOut!"
            Else
                '''//Check if alarmFail(site) is triggered or not.
                #If IGXL_VER_1030 = True Then
                    If alarmOccurred(site) = True Then
                #Else
                    If alarmFail(site) = True Then
                #End If
                TheExec.Datalog.WriteComment "Site:" & site & ", alarmFail!!!"
                End If
                        '''//Bin out the failed DUT with SoftBin and HardBin defined in Vdd_Binning_Def tables.
                        '''Warning!!!please check SortNumber and binNumber of Flag_Vddbinning_Fail_Stop in Bin_Table before using this.
                If isPostBV = False Then        '' 20220215: skip postBV update sort number. Use BinTable instead.
                    '''20230821: Modified to support multi harvest flag, other pin group
                    Call Set_MultiHarvestFlag_ForBV(NonHarvest, site, , VddBinName(p_mode))
                    
                    Call BinoutSite(site, BinCut(p_mode, VBIN_RESULT(p_mode).passBinCut).SBIN_LVCC_FAIL(0, Test_Type), _
                                        BinCut(p_mode, VBIN_RESULT(p_mode).passBinCut).HBIN_LVCC_FAIL(0, Test_Type), strGlb_Flag_Vddbinning_Fail_Stop)


                Else
                    If Not UCase(InstName) Like "*_ELB_PBC_NO_*" Then
                        '''20230821: Modified to support multi harvest flag, other pin group
                        Call Set_MultiHarvestFlag_ForBV(NonHarvest, site, , VddBinName(p_mode))
                        
                        TheExec.sites.item(site).FlagState(strGlb_Flag_Vddbinning_Fail_Stop) = logicTrue
                    End If
                End If
            End If
        End If
            TheExec.sites.item(site).IncrementTestNumber
        Next site
    End If
Exit Function
errHandler:
    Call Print_Error_Message(Error_Info, "LIB_VDD_BINNING", "judge_PF_func") 'Add ErrHandler 2023/08/01
    If AbortTest Then Exit Function Else Resume Next
End Function

'20210903: Modified to merge properties of "Public Type Instance_Step_Control" into Public Type Instance_Info "Public Type Instance_Info".
'20210831: Modified to update HarvestBinningFlag for Harvest in BinCut.
'20210830: As per discussion with TSMC ZYLINI, we decided to use step inherited from Judge_stored_IDS and updated by CurrentPassBinCutNum for HarvestBinningFlag.
'20210830: Modified to revised the vbt code for Harvest in BinCut, as requested by C651 Toby.
'20210820: Modified to remove the redundant GB_delta from the vbt function judge_PF.
'20210803: Modified to use GB_delta for search in non-cp1.
'20210729: Modified to replace "powerDomain = AllBinCut(inst_info.p_mode).powerPin" with inst_info.powerDomain.
'20210722: Modified to use VBIN_IDS_ZONE(p_mode).Product_Voltage(Max_IDS_Zone, Max_IDS_Step) and DYNAMIC_VBIN_IDS_ZONE(p_mode).Product_Voltage(Max_IDS_Step) for GradeVDD.
'20210629: Modified to print the message about skipping Sort Number / Bin Number / fail-stop.
'20210531: Modified to adjust the format of "EQN" items in Judge_PF.
'20210530: Modified to update theExec.sites.Selected for MultiFSTP before exiting Judge_PF.
'20210530: Modified to replace typo "Multisftp_Binout" with "MultiFstp_NoBinout".
'20210529: Modified to check if inst_info.Pattern_Pmode and inst_info.By_Mode exist for MultiFSTP.
'20210529: Modified to unifiy the naming rule of failFlags for MultiFSTP with prefix "F_Multifstp_".
'20210528: Modified to assemble the FailFlag of CP1 MultiFSTP for the failed site.
'20210528: Modified to replace testLimit with theExec.Datalog.WriteParametricResult because testLimit latched FailFlag in flow table incorrectly.
'20210528: Modified to update TheExec.sites.Selected.
'20210526: Modified to remove Monotonicity_Offset check from find_start_voltage because C651 Si revised the check rules.
'20210525: Modified to update siteMask for MultiFSTP in CP1.
'20210514: Modified to check if Montonicity_Offset is triggered.
'20210401: Modified to separate "step_1stPass_in_IDS_Zone" into "Bin_1stPass" and "EQN_1stPass".
'20210325: Modified to print info about Bin and EQN for "COF_StepInheritance" and "Vddbin_DoAll_DebugCollection" if grade_found=False.
'20210325: Modified to use Flag_Vddbin_DoAll_DebugCollection for TheExec.EnableWord("Vddbin_DoAll_DebugCollection").
'20210325: Modified to merge branches of "COF_StepInheritance" and "Vddbin_DoAll_DebugCollection".
'20210324: Modified to use step_mapping(passBin, EQN1) as step_IDS_Zone for COF_StepInheritance and overwrite step_1stPass_in_IDS_Zone.
'20210322: Modified to decide Flag_Vddbin_COF_StepInheritance by checking TheExec.Flow.EnableWord("Vddbin_COF_StepInheritance").
'20210317: Modified to use VBIN_RESULT(inst_info.p_mode).step_in_IDS_Zone to decide PassBinCut.
'20210315: Modified to overwrite VBIN_RESULT(p_mode) for the new COF method requested by C651 Si Li if TheExec.Flow.EnableWord("Vddbin_COF_StepInheritance") = True.
'20201210: Modified to use the arguments "inst_info As Instance_Info" and "step_control As Instance_Step_Control" for judge_PF.
'20201102: Modified to use "Public Type Instance_Info".
'20201020: Modified to add the variable "COFInstance" and "PerEqnLog" for COFInstance.
'20201016: Modified to use "print_info_for_COFInstance".
'20201015: Modified to print the summary for "COFInstance".
'20201015: Modified to add the argument "is_COFInstance_enabled".
'20200811: Modified to align the naming rule of failFlag in BinTable with Ellis and JC-Chop.
'20200615: Modified to get dynamic_offset type from the argument "offsetTestTypeIdx As Integer" for judge_PF.
'20200429: Modified to print info while "Vddbin_DoAll_DebugCollection".
'20200212: Modified to print DSSC_Dec when DSSC_Dec=-1.
'20200102: Modified to print the message for C651 PE and checkscript.
'20191219: Modified to add the EnableWord "Vddbin_DoAll_DebugCollection" for Bincut_DoAll_debug.
'20191127: Modified for the revised InitVddBinTable.
'20190722: Modified to printout the scale and the unit for BinCut voltages and IDS values.
'20190716: Modified to unify the unit for IDS. ids_current with unit mA.
'20190417: Modified to rename the output string "DSSCDEC" with "SELSRAM_DSSC".
'20190226: Modified the calculation for dynamic offset.
'20180821: Modified for BinCut testjob mapping.
Public Function judge_PF(inst_info As Instance_Info, passBinCut As SiteLong)
On Error GoTo errHandler 'Add ErrHandler 2023/08/01
    Dim site As Variant
    Dim Flag As Variant
    Dim ids_step As Long
    Dim lvcc_step As New SiteLong
    Dim strChannel As String
    Dim voltage_Temp As Double
    Dim str_MultiFSTP_FailFlag As String
    Dim PassBinNum As Long
    Dim dbl_BV_lo_limit As Double
    Dim dbl_BV_hi_limit As Double
    Dim str_testJob_Keyword As String
    '''for failFlag
    Dim ModeArr() As String
    Dim mode As String
    Dim Flag_EQN_UpToX As New SiteBoolean: Flag_EQN_UpToX = False '20240201 record each site is BinX EQN
    Dim Flag_Print_EQN_UpToX As Boolean: Flag_Print_EQN_UpToX = False '20240201 if any site is BinX need print by instance
'''//==================================================================================================================================================================================//'''
'''//Note:
'''1. Modified to replace theExec.Flow.TestLimit with theExec.Datalog.WriteParametricResult because theExec.Flow.TestLimit latched FailFlag in flow table incorrectly, 20210528.
'''We modified the vbt function to mask the failed site by theExec.sites.Selected and trig the failFlag if theExec.Flow.EnableWord("Multifstp_Datacollection") = True for MultiFSTP instance.
'''Warning!!! ToDo: Contact TER Expert and factory to solve the issue that FailFlag was triggered incorrectly due to theExec.Flow.TestLimit.
'''2. C651 Toby updated the rules of step voltage calculation. It should not use GB_delta, 20210728.
'''//==================================================================================================================================================================================//'''
'''//==================================================================================================================================================================================//'''
'''//Note:
'''20200811: Modified to align the naming rule of failFlag in BinTable with Ellis and JC-Chop.
'''The unpredictable outside force changed the naming rule of failFlag in BinTable, so that we had to change the vbt code...
'''Warning!!! Please check failFlag in BinTable for "Record the test type and performance mode for Bin2 or Bin3 binning".
'''//==================================================================================================================================================================================//'''
    '''****************************************************************************************************************'''
    ''' Judge IDS Fail first
    ''' If the grade is not found and the EQ number is not 1 then set Binning Fail Bin
    '''****************************************************************************************************************'''
    '''IDS Pass/Fail Check
    For Each site In TheExec.sites
        If inst_info.grade_found = True Then
            ids_step = VBIN_RESULT(inst_info.p_mode).step_in_BinCut
        Else
            ids_step = DYNAMIC_VBIN_IDS_ZONE(inst_info.p_mode).EQ_Num(DYNAMIC_VBIN_IDS_ZONE(inst_info.p_mode).Max_Step - 1) - 1 '''STEP = EQ -1. If the grade is not found, use the last step to identify IDS or LVCC fail.
        End If
        
        '''//IDS calculation uses the scale and the unit in "mA", but TheExec.Datalog.WriteParametricResult should convert IDS value into "A" with settings "unit:=unitAmp" and "scaleMilli".
        If inst_info.grade_found = True Then
            TheExec.Datalog.WriteParametricResult site, TheExec.sites.item(site).TestNumber, logTestPass, parmTestLim, inst_info.powerDomain, strChannel, _
                                                    0, inst_info.ids_current(site) / 1000, BinCut(inst_info.p_mode, passBinCut).IDS_CP_LIMIT(ids_step) / 1000, _
                                                    unitAmp, 0, unitAmp, 0, , , "IDS", scaleMilli
            
        '''****************************************************************************************************************'''
        ''' <LVCC Fail>
        ''' If the C and M of the last step in the IDS Zone are the same with the EQ1 in the BinCut. We fail belong LVCC Fail
        '''****************************************************************************************************************'''
        ElseIf inst_info.grade_found = False And DYNAMIC_VBIN_IDS_ZONE(inst_info.p_mode).c(DYNAMIC_VBIN_IDS_ZONE(inst_info.p_mode).Max_Step - 1) = BinCut(inst_info.p_mode, DYNAMIC_VBIN_IDS_ZONE(inst_info.p_mode).passBinCut(DYNAMIC_VBIN_IDS_ZONE(inst_info.p_mode).Max_Step - 1)).c(0) _
        And DYNAMIC_VBIN_IDS_ZONE(inst_info.p_mode).m(DYNAMIC_VBIN_IDS_ZONE(inst_info.p_mode).Max_Step - 1) = BinCut(inst_info.p_mode, DYNAMIC_VBIN_IDS_ZONE(inst_info.p_mode).passBinCut(DYNAMIC_VBIN_IDS_ZONE(inst_info.p_mode).Max_Step - 1)).m(0) Then
            TheExec.Datalog.WriteParametricResult site, TheExec.sites.item(site).TestNumber, logTestPass, parmTestLim, inst_info.powerDomain, strChannel, _
                                                    0, inst_info.ids_current(site) / 1000, BinCut(inst_info.p_mode, passBinCut).IDS_CP_LIMIT(0) / 1000, _
                                                    unitAmp, 0, unitAmp, 0, , , "IDS", scaleMilli

        '''****************************************************************************************************************'''
        ''' <IDS fail>
        ''' If the C and M of the last step in the IDS Zone are different from the EQ1 in the BinCut.
        ''' It means some EQ numbers can not be tested in this IDS zone becuase the IDS current is over CPIDSMAX spec. It belongs to Binning Fail.
        '''****************************************************************************************************************'''
        Else
            '''//use ids limit of one step less (less current limit)
            TheExec.Datalog.WriteParametricResult site, TheExec.sites.item(site).TestNumber, logTestFail, parmTestLim, inst_info.powerDomain, strChannel, _
                                                    0, inst_info.ids_current(site) / 1000, BinCut(inst_info.p_mode, passBinCut).IDS_CP_LIMIT(ids_step - 1) / 1000, _
                                                    unitAmp, 0, unitAmp, 0, , , "IDS", scaleMilli
        End If
        TheExec.sites.item(site).IncrementTestNumber
    Next site

    '''****************************************************************************************************************'''
    ''' If IDS PASS then judge the LVCC.
    ''' If the grade is not found and the IDS is not fail then set LVCC Fail Bin.
    '''****************************************************************************************************************'''
    '''LVCC Pass/Fail Check
    For Each site In TheExec.sites
        If inst_info.grade_found = True Then
            lvcc_step = VBIN_RESULT(inst_info.p_mode).step_in_BinCut
            PassBinNum = VBIN_RESULT(inst_info.p_mode).passBinCut
        Else
            lvcc_step = DYNAMIC_VBIN_IDS_ZONE(inst_info.p_mode).EQ_Num(DYNAMIC_VBIN_IDS_ZONE(inst_info.p_mode).Max_Step - 1) - 1
            PassBinNum = 1
        End If
        
        '''//Select the keyword about the current testJob for TName of testLimit.
        Select Case LCase(bincutJobName)
            Case "cp1": str_testJob_Keyword = "CP1"
            Case "cp2": str_testJob_Keyword = "CP2"
            Case "ft_room": str_testJob_Keyword = "FT1"
            Case "ft_hot": str_testJob_Keyword = "FT2"
            Case "qa": str_testJob_Keyword = "QA"
            Case Else: str_testJob_Keyword = "CP1"
            Call Print_Error_Message(Error_Info, "LIB_VDD_BINNING", "judge_PF", "site:" & site & ", " & bincutJobName & " is the incorrect BinCut testJob for judge_PF. Error!!!")
                'TheExec.ErrorLogMessage "site:" & site & ", " & bincutJobName & " is the incorrect BinCut testJob for judge_PF. Error!!!"
        End Select
        
        '''//Calculate BV_lo_limit and BV_hi_limit to check BinCut voltage(Grade).
        dbl_BV_lo_limit = BinCut(inst_info.p_mode, PassBinNum).CP_Vmin(lvcc_step)
        dbl_BV_hi_limit = BinCut(inst_info.p_mode, PassBinNum).CP_Vmax(lvcc_step)
        
        '''****************************************************************************************************************'''
        ''' BinCut voltage (Grade)
        '''****************************************************************************************************************'''
        '''//BinCut voltage calculation uses the scale and the unit in "mV", but TheExec.Flow.TestLimit should convert voltage value into "V" with settings "unit:=unitVolt" and "scaleMilli".
        If inst_info.grade_found = True Then
            TheExec.Datalog.WriteParametricResult site, TheExec.sites.item(site).TestNumber, logTestPass, parmTestLim, inst_info.powerDomain, strChannel, _
                                    dbl_BV_lo_limit / 1000, VBIN_RESULT(inst_info.p_mode).GRADE / 1000, dbl_BV_hi_limit / 1000, unitVolt, 0, unitVolt, 0, , , str_testJob_Keyword, scaleMilli, "%.4f"
        Else
        
            TheExec.Datalog.WriteParametricResult site, TheExec.sites.item(site).TestNumber, logTestFail, parmTestLim, inst_info.powerDomain, strChannel, _
                                    dbl_BV_lo_limit / 1000, 0, dbl_BV_hi_limit / 1000, unitVolt, 0, unitVolt, 0, , , str_testJob_Keyword, scaleMilli, "%.4f"
        End If
        
        '''//Align testNumber.
        TheExec.sites.item(site).IncrementTestNumber
    Next site
    
    '''****************************************************************************************************************'''
    ''' Dynamic Offset
    '''****************************************************************************************************************'''
    For Each site In TheExec.sites
        '''//Dynamic offset is not related to product voltage or efuse, so it doesn't need take the least multiple of stepVoltage.
        '''BinCut voltage calculation uses the scale and the unit in "mV", but TheExec.Flow.TestLimit should convert voltage value into "V" with settings "unit:=unitVolt" and "scaleMilli".
        voltage_Temp = BinCut(inst_info.p_mode, VBIN_RESULT(inst_info.p_mode).passBinCut).DYNAMIC_OFFSET(inst_info.jobIdx, inst_info.offsetTestTypeIdx)
        
        TheExec.Datalog.WriteParametricResult site, TheExec.sites.item(site).TestNumber, logTestPass, parmTestLim, inst_info.powerDomain, strChannel, _
                                                   -0.1, voltage_Temp / 1000, 0.1, unitVolt, 0, unitVolt, 0, , , "OFFSET", scaleMilli, "%.4f"
                                                   
        TheExec.sites.item(site).IncrementTestNumber
    Next site
    
    '''****************************************************************************************************************'''
    ''' EQN Result
    '''****************************************************************************************************************'''
    For Each site In TheExec.sites
        '''********************************************************************************************************************************************************'''
        '''20210528: Modified to replace theExec.Flow.TestLimit with theExec.Datalog.WriteParametricResult because theExec.Flow.TestLimit latched FailFlag in flow table incorrectly.
        '''Warning!!! ToDo: Contact TER Expert and factory to solve the issue that FailFlag was triggered incorrectly due to theExec.Flow.TestLimit.
        '''********************************************************************************************************************************************************'''
        If inst_info.grade_found = True Then
            TheExec.Datalog.WriteParametricResult site, TheExec.sites.item(site).TestNumber, logTestPass, parmTestLim, inst_info.powerDomain, strChannel, _
                                                    1, VBIN_RESULT(inst_info.p_mode).step_in_BinCut + 1, BinCut(inst_info.p_mode, VBIN_RESULT(inst_info.p_mode).passBinCut).Mode_Step + 1, _
                                                    unitNone, 0, unitNone, 0, , , "EQN", scaleNoScaling, "%.0f"
        Else
            TheExec.Datalog.WriteParametricResult site, TheExec.sites.item(site).TestNumber, logTestFail, parmTestLim, inst_info.powerDomain, strChannel, _
                                                    1, 0, BinCut(inst_info.p_mode, VBIN_RESULT(inst_info.p_mode).passBinCut).Mode_Step + 1, _
                                                    unitNone, 0, unitNone, 0, , , "EQN", scaleNoScaling, "%.0f"
        End If
        
        TheExec.sites.item(site).IncrementTestNumber
    Next site
    
    '''****************************************************************************************************************'''
    ''' BinCut PassBinNumber Result
    '''****************************************************************************************************************'''
    For Each site In TheExec.sites
        If inst_info.grade_found = True Then
            '''//BinCut PASSBIN values.
            '''********************************************************************************************************************************************************'''
            '''20210528: Modified to replace theExec.Flow.TestLimit with theExec.Datalog.WriteParametricResult because theExec.Flow.TestLimit latched FailFlag in flow table incorrectly.
            '''Warning!!! ToDo: Contact TER Expert and factory to solve the issue that FailFlag was triggered incorrectly due to theExec.Flow.TestLimit.
            '''********************************************************************************************************************************************************'''
            TheExec.Datalog.WriteParametricResult site, TheExec.sites.item(site).TestNumber, logTestPass, parmTestLim, inst_info.powerDomain, strChannel, _
                                                    1, VBIN_RESULT(inst_info.p_mode).passBinCut, PassBinCut_ary(UBound(PassBinCut_ary)), _
                                                    unitNone, 0, unitNone, 0, , , "PASSBIN", scaleNoScaling, "%.0f"
            
            '''****************************************************************************************************************'''
            ''' Record the test type and performance mode for Bin2 or Bin3 binning
            '''****************************************************************************************************************'''
            '''20200811: Modified to align the naming rule of failFlag in BinTable with Ellis and JC-Chop.
            '''The unpredictable outside force changed the naming rule of failFlag in BinTable, so that we had to change the vbt code...
            '''Warning!!! Please check failFlag in BinTable.
            If VBIN_RESULT(inst_info.p_mode).passBinCut = 2 And Binx_fail_flag(site) = False Then
                ModeArr = Split(UCase(VddBinName(inst_info.p_mode)), "_")
                mode = ModeArr(UBound(ModeArr))
                
                If inst_info.Test_Type = testType.TD Then
                    Binx_fail_power(site) = mode & "_TD_BV"
                    Binx_fail_flag(site) = True
                ElseIf inst_info.Test_Type = testType.Mbist Then
                    Binx_fail_power(site) = mode & "_Mbist_BV"
                    Binx_fail_flag(site) = True
                ElseIf inst_info.Test_Type = testType.SPI Then
                    Binx_fail_power(site) = mode & "_SPI_BV"
                    Binx_fail_flag(site) = True
                ElseIf inst_info.Test_Type = testType.RTOS Then
                    Binx_fail_power(site) = mode & "_RTOS_BV"
                    Binx_fail_flag(site) = True
                End If
            ElseIf VBIN_RESULT(inst_info.p_mode).passBinCut = 3 And Biny_fail_flag(site) = False Then
                ModeArr = Split(UCase(VddBinName(inst_info.p_mode)), "_")
                mode = ModeArr(UBound(ModeArr))
                
                If inst_info.Test_Type = testType.TD Then
                    Biny_fail_power(site) = mode & "_TD_BV"
                    Biny_fail_flag(site) = True
                ElseIf inst_info.Test_Type = testType.Mbist Then
                    Biny_fail_power(site) = mode & "_Mbist_BV"
                    Biny_fail_flag(site) = True
                ElseIf inst_info.Test_Type = testType.SPI Then
                    Biny_fail_power(site) = mode & "_SPI_BV"
                    Biny_fail_flag(site) = True
                ElseIf inst_info.Test_Type = testType.RTOS Then
                    Biny_fail_power(site) = mode & "_RTOS_BV"
                    Biny_fail_flag(site) = True
                End If
            End If
            
            If Flag_BinX_Info_Parsed = True Then
                '20240201 check if EQN is BinX
                If Pmode2BinX_EQN_Dict.Exists(inst_info.p_mode) And Check_EQN_UpToBinX(inst_info.p_mode, site) = True Then
                    Flag_Fuse_BinX(site) = True
                    Flag_EQN_UpToX(site) = True
                    Flag_Print_EQN_UpToX = True
                End If
            End If
        Else
            '''********************************************************************************************************************************************************'''
            '''20210528: Modified to replace theExec.Flow.TestLimit with theExec.Datalog.WriteParametricResult because theExec.Flow.TestLimit latched FailFlag in flow table incorrectly.
            '''Warning!!! ToDo: Contact TER Expert and factory to solve the issue that FailFlag was triggered incorrectly due to theExec.Flow.TestLimit.
            '''********************************************************************************************************************************************************'''
            TheExec.Datalog.WriteParametricResult site, TheExec.sites.item(site).TestNumber, logTestFail, parmTestLim, inst_info.powerDomain, strChannel, _
                                                    1, 0, PassBinCut_ary(UBound(PassBinCut_ary)), _
                                                    unitNone, 0, unitNone, 0, , , "PASSBIN", scaleNoScaling, "%.0f"
        End If
        TheExec.sites.item(site).IncrementTestNumber
    Next site
    
    '20240201 print each site is Bin1 or BinX if any site is BinX
    If Flag_Print_EQN_UpToX = True Then
        For Each site In TheExec.sites
            If Flag_EQN_UpToX(site) Then
                TheExec.Datalog.WriteParametricResult site, TheExec.sites.item(site).TestNumber, logTestPass, parmTestLim, "", "", _
                                                0, 1, 1, unitNone, 0, unitNone, 0, , , VddBinName(inst_info.p_mode) & " EQN BinX", scaleNoScaling, "%.0f"
            Else
                TheExec.Datalog.WriteParametricResult site, TheExec.sites.item(site).TestNumber, logTestPass, parmTestLim, "", "", _
                                                0, 0, 1, unitNone, 0, unitNone, 0, , , VddBinName(inst_info.p_mode) & " EQN Bin1", scaleNoScaling, "%.0f"
            End If
            TheExec.sites.item(site).IncrementTestNumber
        Next site
    End If
    '''****************************************************************************************************************'''
    ''' SELSRM DSSC DigiSrc (converted into decimal).
    '''****************************************************************************************************************'''
    For Each site In TheExec.sites
        If VBIN_RESULT(inst_info.p_mode).DSSC_Dec <> -1 Then
            '''********************************************************************************************************************************************************'''
            '''20210528: Modified to replace theExec.Flow.TestLimit with theExec.Datalog.WriteParametricResult because theExec.Flow.TestLimit latched FailFlag in flow table incorrectly.
            '''Warning!!! ToDo: Contact TER Expert and factory to solve the issue that FailFlag was triggered incorrectly due to theExec.Flow.TestLimit.
            '''********************************************************************************************************************************************************'''
            TheExec.Datalog.WriteParametricResult site, TheExec.sites.item(site).TestNumber, logTestPass, parmTestLim, inst_info.powerDomain, strChannel, _
                                                    VBIN_RESULT(inst_info.p_mode).DSSC_Dec, VBIN_RESULT(inst_info.p_mode).DSSC_Dec, VBIN_RESULT(inst_info.p_mode).DSSC_Dec, _
                                                    unitNone, 0, unitNone, 0, , , "SELSRAM_DSSC", scaleNoScaling, "%.0f"
        Else
            VBIN_RESULT(inst_info.p_mode).DSSC_Dec = -1
        End If
    Next site
    
    '''****************************************************************************************************************'''
    '''//Print the summary about COFInstance with PTR format into STDF. Requested by C651 Si Li.
    '''****************************************************************************************************************'''
    If inst_info.enable_PerEqnLog = True Then
        Call print_info_for_COFInstance(inst_info)
    End If
    
    
    For Each site In TheExec.sites
        If inst_info.HarvestBinningFlag <> "" And inst_info.Harvest_With_BIN1E1 = True And Check_EQN_UpToBinX(inst_info.p_mode, site) _
        And inst_info.grade_found(site) = True And TheExec.sites.item(site).FlagState(strGlb_Flag_HarvestBinningFlag_SingleFSTP) = logicFalse _
        And TheExec.sites.item(site).FlagState(strGlb_Flag_HarvestBinningFlag_AllCorePass) = logicTrue Then
            inst_info.grade_found(site) = False '20240403 BIN1 fail and BINX pass need to harvest to BIN1
            VBIN_RESULT(inst_info.p_mode).step_BINX_MFSTP(site) = VBIN_RESULT(inst_info.p_mode).step_1stPass_in_IDS_Zone '20240403 record BINX step if harvest on bin1 but pass on BINX, and reset step to BIN1 EQN1
        End If
    Next site
    
    '''****************************************************************************************************************'''
    ''' Decide SortNumber/BinNumber, and bin out the failed DUT.
    '''//SortNumber and BinNumber are generated from "BinNumberConfig" sheet by Tautogen into "Vdd_Binning_Def".
    '''****************************************************************************************************************'''
    '''ToDo: Maybe we can merge the vbt code of this part to the vbt function judge_PF_func...
    autoShmoo.start: Set autoShmoo = Nothing ''' start auto shmoo
    For Each site In TheExec.sites
        If inst_info.grade_found = False Then
            If Flag_Vddbin_COF_StepInheritance = True Then
                '''*******************************************************************************************************************************************************************'''
                '''20210315: Modified to overwrite VBIN_RESULT(p_mode) for the new COF method requested by C651 Si Li if TheExec.Flow.EnableWord("Vddbin_COF_StepInheritance") = True.
                '''//If p_mode has found PassBin and EQN in previous instances, it will overwrite VBIN_RESULT with 1st_Pass_Step of p_mode for COF_StepInstance.
                '''//If p_mode didn't find any PassBin and EQN in previous instances, it will use step from DYNAMIC_VBIN_IDS_ZONE(inst_info.p_mode).step_Mapping(CurrentPassBinCutNum, EQN1)
                '''*******************************************************************************************************************************************************************'''
                VBIN_RESULT(inst_info.p_mode).FLAGFAIL = False
                
                If (VBIN_RESULT(inst_info.p_mode).step_1stPass_in_IDS_Zone > -1) = True Then '''If p_mode was tested and has found Grade in previous instances...
                    VBIN_RESULT(inst_info.p_mode).step_in_IDS_Zone = VBIN_RESULT(inst_info.p_mode).step_1stPass_in_IDS_Zone
                    VBIN_RESULT(inst_info.p_mode).passBinCut = DYNAMIC_VBIN_IDS_ZONE(inst_info.p_mode).passBinCut(VBIN_RESULT(inst_info.p_mode).step_in_IDS_Zone)
                Else '''If It didn't find any 1st_Pass_Step in previous instances...
                    '''*******************************************************************************************************************************************************************'''
                    '''//step_control.step_Start(site) stores start step inherited from previous and current p_mode.
                    '''It can get passBin from step_control.step_Start(site), then it can set step_mapping(passBin, EQN1) as step_IDS_Zone.
                    '''*******************************************************************************************************************************************************************'''
                    VBIN_RESULT(inst_info.p_mode).passBinCut = DYNAMIC_VBIN_IDS_ZONE(inst_info.p_mode).passBinCut(inst_info.step_Start(site))
                    VBIN_RESULT(inst_info.p_mode).step_in_IDS_Zone = DYNAMIC_VBIN_IDS_ZONE(inst_info.p_mode).step_Mapping(VBIN_RESULT(inst_info.p_mode).passBinCut, 1)
                    '20230605: Modified for new COF_StepInheritence rule, do not set E1 when there's no grade. Requested by C651 Oscar and TSMC ZYLINI
                    'VBIN_RESULT(inst_info.p_mode).step_1stPass_in_IDS_Zone = VBIN_RESULT(inst_info.p_mode).step_in_IDS_Zone
                    TheExec.Datalog.WriteComment "site:" & site & ", Pmode has no step, bincut set E1 temporary until we find the real step!!!"
                End If
                
                '''//Update PassBin, BinCut voltage(Grade), and Efuse product voltage(GradeVDD).
                VBIN_RESULT(inst_info.p_mode).step_in_BinCut = DYNAMIC_VBIN_IDS_ZONE(inst_info.p_mode).EQ_Num(VBIN_RESULT(inst_info.p_mode).step_in_IDS_Zone) - 1
                VBIN_RESULT(inst_info.p_mode).GRADE = DYNAMIC_VBIN_IDS_ZONE(inst_info.p_mode).Voltage(VBIN_RESULT(inst_info.p_mode).step_in_IDS_Zone)
                VBIN_RESULT(inst_info.p_mode).GRADEVDD = DYNAMIC_VBIN_IDS_ZONE(inst_info.p_mode).Product_Voltage(VBIN_RESULT(inst_info.p_mode).step_in_IDS_Zone)
                CurrentPassBinCutNum(site) = VBIN_RESULT(inst_info.p_mode).passBinCut
                
                '''//Print info about bin, Eqn, Grade, GradeVDD.
                TheExec.Datalog.WriteComment "site:" & site & "," & VddBinName(inst_info.p_mode) & ",Vddbin_COF_StepInheritance overwrites test result of the failed DUT" & _
                                                ",bin=" & VBIN_RESULT(inst_info.p_mode).passBinCut & _
                                                ",EQN=" & VBIN_RESULT(inst_info.p_mode).step_in_BinCut + 1 & _
                                                ",Grade=" & VBIN_RESULT(inst_info.p_mode).GRADE & ",GradeVDD=" & VBIN_RESULT(inst_info.p_mode).GRADEVDD
                
            ElseIf Flag_Vddbin_DoAll_DebugCollection = True Then
                TheExec.Datalog.WriteComment "site:" & CStr(site) & "," & VddBinName(inst_info.p_mode) & ",is forced to Bin1 EQN1 CPVmax because Vddbin_DoAll_DebugCollection is enabled, and this test item couldn't find any grade-search result."
                VBIN_RESULT(inst_info.p_mode).step_in_BinCut = 0
                VBIN_RESULT(inst_info.p_mode).passBinCut = 1
                
                '''//VddBin_DoAll_DebugCollection should forces CP1 BinCut voltage to Bin1 EQN1 CPVmax.
                VBIN_RESULT(inst_info.p_mode).GRADE = BinCut(inst_info.p_mode, VBIN_RESULT(inst_info.p_mode).passBinCut).CP_Vmax(VBIN_RESULT(inst_info.p_mode).step_in_BinCut)
                '''//Efuse product voltage(GradeVDD) = BinningVoltage(Grade) + binning_GuardBand.
                VBIN_RESULT(inst_info.p_mode).GRADEVDD = BinCut(inst_info.p_mode, VBIN_RESULT(inst_info.p_mode).passBinCut).CP_Vmax(VBIN_RESULT(inst_info.p_mode).step_in_BinCut) + BinCut(inst_info.p_mode, VBIN_RESULT(inst_info.p_mode).passBinCut).CP_GB(VBIN_RESULT(inst_info.p_mode).step_in_BinCut)
                
                '''//Print info about bin, Eqn, Grade, GradeVDD.
                TheExec.Datalog.WriteComment "site:" & site & "," & VddBinName(inst_info.p_mode) & ",Vddbin_DoAll_DebugCollection overwrites test result of the failed DUT" & _
                                                ",bin=" & VBIN_RESULT(inst_info.p_mode).passBinCut & _
                                                ",EQN=" & VBIN_RESULT(inst_info.p_mode).step_in_BinCut + 1 & _
                                                ",Grade=" & VBIN_RESULT(inst_info.p_mode).GRADE & ",GradeVDD=" & VBIN_RESULT(inst_info.p_mode).GRADEVDD
            
            '''//If the flag "MultiFstp_NoBinout" = True, skip SortNumber and fail-stop for MultiFSTP instances.
            '''ToDo: Please check if the failFlag ""MultiFstp_NoBinout"" exists in the flow table!!!
            ElseIf TheExec.sites.item(site).FlagState("MultiFstp_NoBinout") = logicTrue Then 'MultiFstp without Binout
                '''//Mask the failed site for MultiFSTP instances.
                gb_siteMask_current(site) = False
                
                '''//Print info about overwriting fail-stop(BinOut) for MultiFSTP.
                TheExec.Datalog.WriteComment "site:" & site & "," & VddBinName(inst_info.p_mode) & ", test failed, but overwrite siteMask not to directly bin out the failed site for MultiFSTP."
                    
                '''//It should assemble and update the failFalg of MultiFSTP instances for the failed site.
                '''*******************************************************************************************************************************************************************'''
                '''//Unify the naming rule of failFlags for MultiFSTP with prefix "F_Multifstp_", ex: F_Multifstp_MGX001_X4_BV, F_Multifstp_MGX003_X6_BV, F_Multifstp_MGX008_X10_BV.
                '''Warning: Remember to check if all related FailFlags for MultiFSTP exist in BinCut test flow tables!!!
                '''*******************************************************************************************************************************************************************'''
                '''//Check if keywords of Pattern_Pmode and By_Mode for MultiFSTP (Harvest Core DSSC) are available in the instance name.
                If inst_info.Pattern_Pmode <> "" And inst_info.By_Mode <> "" Then
                    '''//Update test result to the FailFlag assembled with prefix "F_Multifstp_" for MultiFSTP instances.
                    str_MultiFSTP_FailFlag = "F_Multifstp_" & inst_info.Pattern_Pmode & "_" & inst_info.By_Mode & "_BV"
                    TheExec.sites.item(site).FlagState(str_MultiFSTP_FailFlag) = logicTrue
                Else
                    '''//Print the message about skipping Sort Number / Bin Number / fail-stop.
                    TheExec.Datalog.WriteComment "site:" & site & "," & inst_info.inst_name & "," & VddBinName(inst_info.p_mode) & ", test failed, but flag MultiFstp_NoBinout is enabled to overwrite Judge_PF."
                    'TheExec.ErrorLogMessage "site:" & site & "," & inst_info.inst_name & "," & VddBinName(inst_info.p_mode) & ", test failed, but flag MultiFstp_NoBinout is enabled to overwrite Judge_PF."
                End If
                
            '''//If HarvestBinningFlag is not empty, skip updating sortNumber, and skip fail-stop.
            '''20210830: Modified to revised the vbt code for Harvest in BinCut, as requested by C651 Toby.
            ElseIf inst_info.HarvestBinningFlag <> "" Then '''HarvestBinning
                VBIN_RESULT(inst_info.p_mode).FLAGFAIL = False
                
                '''20230821: Modified to support multi harvest flag, other pin group
                Call Set_MultiHarvestFlag_ForBV(GFXHarvest, site, inst_info.HarvestBinningFlag)
                
                '''//Update the flagstate of strGlb_Flag_HarvestBinningFlag_AllCorePass because one Core fails.
                If strGlb_Flag_HarvestBinningFlag_AllCorePass <> "" Then
                    If TheExec.sites.item(site).FlagState(strGlb_Flag_HarvestBinningFlag_AllCorePass) = logicTrue Then
                        TheExec.sites.item(site).FlagState(strGlb_Flag_HarvestBinningFlag_AllCorePass) = logicFalse
                        TheExec.sites.item(site).FlagState(strGlb_Flag_HarvestBinningFlag_SingleFSTP) = logicTrue ''' this is to run single FSTP once MFSTP failed
                        TheExec.Datalog.WriteComment "site:" & site & "," & inst_info.inst_name & "," & VddBinName(inst_info.p_mode) & ", test failed, so that update Harvest AllCorePass failFlag:" & strGlb_Flag_HarvestBinningFlag_AllCorePass & "=False."
                        
                        If (VBIN_RESULT(inst_info.p_mode).step_1stPass_in_IDS_Zone > -1) = True Then '''If p_mode was tested and has found Grade in previous instances...
                            If inst_info.Harvest_With_BIN1E1 = True Then
                                VBIN_RESULT(inst_info.p_mode).step_in_IDS_Zone = inst_info.step_Start(site) '20240403 BINX pass but need to harv on BIN1 so reset step to BIN1 EQN1
                                Flag_Fuse_BinX(site) = False
                            Else
                                VBIN_RESULT(inst_info.p_mode).step_in_IDS_Zone = VBIN_RESULT(inst_info.p_mode).step_1stPass_in_IDS_Zone
                            End If
                        Else '''If It didn't find any 1st_Pass_Step in previous instances...
                            '''*******************************************************************************************************************************************************************'''
                            '''//Note:
                            '''20210830: As per discussion with TSMC ZYLINI, we decided to use step inherited from Judge_stored_IDS and updated by CurrentPassBinCutNum for HarvestBinningFlag.
                            '''*******************************************************************************************************************************************************************'''
                            VBIN_RESULT(inst_info.p_mode).step_in_IDS_Zone = inst_info.step_Start(site)
                            '''//Overwrite step_1stPass_in_IDS_zone.
                            VBIN_RESULT(inst_info.p_mode).step_1stPass_in_IDS_Zone = VBIN_RESULT(inst_info.p_mode).step_in_IDS_Zone
                        End If
                        
                        '''//Update PassBin, BinCut voltage(Grade), and Efuse product voltage(GradeVDD).
                        VBIN_RESULT(inst_info.p_mode).passBinCut = DYNAMIC_VBIN_IDS_ZONE(inst_info.p_mode).passBinCut(VBIN_RESULT(inst_info.p_mode).step_in_IDS_Zone)
                        VBIN_RESULT(inst_info.p_mode).step_in_BinCut = DYNAMIC_VBIN_IDS_ZONE(inst_info.p_mode).EQ_Num(VBIN_RESULT(inst_info.p_mode).step_in_IDS_Zone) - 1
                        VBIN_RESULT(inst_info.p_mode).GRADE = DYNAMIC_VBIN_IDS_ZONE(inst_info.p_mode).Voltage(VBIN_RESULT(inst_info.p_mode).step_in_IDS_Zone)
                        VBIN_RESULT(inst_info.p_mode).GRADEVDD = DYNAMIC_VBIN_IDS_ZONE(inst_info.p_mode).Product_Voltage(VBIN_RESULT(inst_info.p_mode).step_in_IDS_Zone)
                        CurrentPassBinCutNum(site) = VBIN_RESULT(inst_info.p_mode).passBinCut
                        
                        '''//Print info about HarvestBinningFlag, bin, Eqn, Grade, GradeVDD.
                        TheExec.Datalog.WriteComment "site:" & site & "," & VddBinName(inst_info.p_mode) & ",HarvestBinningFlag:" & inst_info.HarvestBinningFlag & ", it overwrites test result of the failed DUT" & _
                                                        ",bin=" & VBIN_RESULT(inst_info.p_mode).passBinCut & _
                                                        ",EQN=" & VBIN_RESULT(inst_info.p_mode).step_in_BinCut + 1 & _
                                                        ",Grade=" & VBIN_RESULT(inst_info.p_mode).GRADE & ",GradeVDD=" & VBIN_RESULT(inst_info.p_mode).GRADEVDD
                    Else '''20240403 If more than two Cores fails, but MFSTP BINX pass, set BINX result to VBIN_RESULT and disable SingleFSTP instance behind
                        If inst_info.Harvest_With_BIN1E1 = True And VBIN_RESULT(inst_info.p_mode).step_BINX_MFSTP(site) <> -1 Then
                            VBIN_RESULT(inst_info.p_mode).step_in_IDS_Zone = VBIN_RESULT(inst_info.p_mode).step_BINX_MFSTP(site)
                            VBIN_RESULT(inst_info.p_mode).passBinCut = DYNAMIC_VBIN_IDS_ZONE(inst_info.p_mode).passBinCut(VBIN_RESULT(inst_info.p_mode).step_in_IDS_Zone)
                            VBIN_RESULT(inst_info.p_mode).step_in_BinCut = DYNAMIC_VBIN_IDS_ZONE(inst_info.p_mode).EQ_Num(VBIN_RESULT(inst_info.p_mode).step_in_IDS_Zone) - 1
                            VBIN_RESULT(inst_info.p_mode).GRADE = DYNAMIC_VBIN_IDS_ZONE(inst_info.p_mode).Voltage(VBIN_RESULT(inst_info.p_mode).step_in_IDS_Zone)
                            VBIN_RESULT(inst_info.p_mode).GRADEVDD = DYNAMIC_VBIN_IDS_ZONE(inst_info.p_mode).Product_Voltage(VBIN_RESULT(inst_info.p_mode).step_in_IDS_Zone)
                            CurrentPassBinCutNum(site) = VBIN_RESULT(inst_info.p_mode).passBinCut
                            Flag_Fuse_BinX(site) = True
                            TheExec.sites.item(site).FlagState(strGlb_Flag_HarvestBinningFlag_SingleFSTP) = logicFalse
                            
                            TheExec.Datalog.WriteComment "site:" & site & "," & inst_info.inst_name & "," & VddBinName(inst_info.p_mode) & ", Harvest Core failFlag:" & inst_info.HarvestBinningFlag & _
                                                         ", more than one Harvest Core failed, but MFSTP pass on BINX, EQN =" & VBIN_RESULT(inst_info.p_mode).step_in_BinCut + 1 & _
                                                         ", so that force ,bin=" & VBIN_RESULT(inst_info.p_mode).passBinCut & _
                                                         ", EQN=" & VBIN_RESULT(inst_info.p_mode).step_in_BinCut + 1 & _
                                                         ", Grade=" & VBIN_RESULT(inst_info.p_mode).GRADE & ",GradeVDD=" & VBIN_RESULT(inst_info.p_mode).GRADEVDD
                        Else '''If more than two Cores fails, bin out the failed DUT...
                            TheExec.Datalog.WriteComment "site:" & site & "," & inst_info.inst_name & "," & VddBinName(inst_info.p_mode) & ", Harvest Core failFlag:" & inst_info.HarvestBinningFlag & ", more than one Harvest Core failed, so that bin out the failed DUT."
                        
                            '''Warning!!!please check SortNumber and binNumber of Flag_Vddbinning_Fail_Stop in Bin_Table before using this.
                        Call BinoutSite(site, BinCut(inst_info.p_mode, PassBinCut_ary(UBound(PassBinCut_ary))).SBIN_LVCC_FAIL(lvcc_step, inst_info.Test_Type), _
                                            BinCut(inst_info.p_mode, PassBinCut_ary(UBound(PassBinCut_ary))).HBIN_LVCC_FAIL(lvcc_step, inst_info.Test_Type), strGlb_Flag_Vddbinning_Fail_Stop)
                        End If
                    End If
                Else
                    TheExec.Datalog.WriteComment "site:" & site & "," & inst_info.inst_name & "," & VddBinName(inst_info.p_mode) & ", test failed, but Judge_PF can't update the failFlag of AllCorePass because failFlag:" & strGlb_Flag_HarvestBinningFlag_AllCorePass & " doesn't exist. Error!!!"
                    TheExec.ErrorLogMessage "site:" & site & "," & inst_info.inst_name & "," & VddBinName(inst_info.p_mode) & ", test failed, but Judge_PF can't update the failFlag of AllCorePass because failFlag:" & strGlb_Flag_HarvestBinningFlag_AllCorePass & " doesn't exist. Error!!!"
                End If

                If inst_info.Harvest_With_BIN1E1 = True And Flag_Fuse_BinX(site) = True Then
                    TheExec.Datalog.WriteComment "site:" & site & "," & inst_info.inst_name & "," & VddBinName(inst_info.p_mode) & ", More than two Core failed. But MFSTP BinX Pass. Force to BinX and Clear Single Core Flag"
                    For Each Flag In Array("F_GFX_Core0", "F_GFX_TD_CORE0", "F_GFX_Core1", "F_GFX_TD_CORE1", "F_GFX_Core2", "F_GFX_TD_CORE2", "F_GFX_Core3", "F_GFX_TD_CORE3", "F_GFX_Core4", "F_GFX_TD_CORE4", "F_GFX_Core5", "F_GFX_TD_CORE5")
                        TheExec.sites.item(site).FlagState(Flag) = logicFalse
                    Next Flag
                Else
                    '''20210831: Modified to update HarvestBinningFlag for Harvest in BinCut.
                    TheExec.sites.item(site).FlagState(inst_info.HarvestBinningFlag) = logicTrue
                    TheExec.Datalog.WriteComment "site:" & site & "," & inst_info.inst_name & "," & VddBinName(inst_info.p_mode) & ", test failed, it has to update Harvest Core failFlag:" & inst_info.HarvestBinningFlag & "=True."
                End If
            Else
                If inst_info.enable_HarvPinGrp = False Then
                    '''20230821: Modified to support multi harvest flag, other pin group
                    Call Set_MultiHarvestFlag_ForBV(NonHarvest, site, , inst_info.performance_mode)
                End If
                
                '''Warning!!!please check SortNumber and binNumber of Flag_Vddbinning_Fail_Stop in Bin_Table before using this.
                Call BinoutSite(site, BinCut(inst_info.p_mode, PassBinCut_ary(UBound(PassBinCut_ary))).SBIN_LVCC_FAIL(lvcc_step, inst_info.Test_Type), _
                                    BinCut(inst_info.p_mode, PassBinCut_ary(UBound(PassBinCut_ary))).HBIN_LVCC_FAIL(lvcc_step, inst_info.Test_Type), strGlb_Flag_Vddbinning_Fail_Stop)

            End If
        End If
    Next site
    
    '''//Update theExec.sites.Selected for MultiFSTP before exiting Judge_PF.
    '''Warning!!! It can update theExec.sites.Selected outside site-loop only.
    '''ToDo: Please check if EnableWord("Multifstp_Datacollection") exists in the flow table!!!
    '''20210530: Modified to update theExec.sites.Selected for MultiFSTP before exiting Judge_PF.
    If TheExec.flow.enableWord("Multifstp_Datacollection") = True Then
        TheExec.sites.Selected = gb_siteMask_current
    End If
Exit Function
errHandler:
    Call Print_Error_Message(Error_Info, "LIB_VDD_BINNING", "judge_PF") 'Add ErrHandler 2023/08/01
    If AbortTest Then Exit Function Else Resume Next
End Function

'20210903: Modified to merge properties of "Public Type Instance_Step_Control" into Public Type Instance_Info "Public Type Instance_Info".
'20210126: Modified to revise the vbt code for DevChar.
'20201211: Created to align testNumber, then do judge_PF for binSearch and judge_PF_func for functional test.
'20201127: Modified to remove the redundant argument "IfStoreData As SiteBoolean".
'20201125: As suggestion from Chihome, modified to clear capture Memory (CMEM) after PostTestIPF.
Public Function update_sort_result(inst_info As Instance_Info, pattPass As SiteBoolean, Org_Test_Number As Long, Optional failpins As String, Optional CollectOnEachStep As Boolean, Optional Test_Number_Per_Step As Long = 0)
On Error GoTo errHandler 'Add ErrHandler 2023/08/01
    Dim site As Variant
    Dim Site_Align As Long
    '''//Only BinSearch can use CMEM.
    If inst_info.enable_CMEM_Collection = True And CollectOnEachStep = False Then
        Call PostTestIPF(inst_info.performance_mode, failpins, inst_info.PrintSize, inst_info.BC_CMEM_StoreData, inst_info.str_patternPins, inst_info.PrintSize)
        TheHdw.Digital.CMEM.SetCaptureConfig 0, CmemCaptNone '''CmemCaptNone: Capture no cycles.
        TheHdw.Digital.CMEM.CentralFields = tlCMEMNone  '20220512, add for turn off cmem capture
    End If
    
    If inst_info.enable_HarvPinGrp = True Then
        TheHdw.Digital.CMEM.SetCaptureConfig 0, CmemCaptNone '''CmemCaptNone: Capture no cycles.
        TheHdw.Digital.CMEM.CentralFields = tlCMEMNone  '20220512, add for turn off cmem capture
    End If
    
    If inst_info.is_DevChar_Running = False Then
        '''//Update sort number and bin out the failed DUT.
        If inst_info.is_BinSearch = True Then '''BinCut search.
            '''//Update PassBinCut for DUT "grade_found=false".
            Call Update_PassBinCut_for_GradeNotFound(inst_info)
            
            '''************************************************************************************************************************************************'''
            ''' (2) For TestNumber align, calculate the EQs in BinCut Tables and get a guard band to avoid the TestNumber is different from another touch down.
            '''************************************************************************************************************************************************'''
            If Test_Number_Per_Step = 0 Then
                Site_Align = Org_Test_Number + (inst_info.count_PrePatt_decomposed + inst_info.count_FuncPat_decomposed) * Max_V_Step_per_IDS_Zone + 10
            Else
                Site_Align = Org_Test_Number + Test_Number_Per_Step * Max_V_Step_per_IDS_Zone + 10
            End If
            ''20230807: Modified to read test number after called instance returns back to bincut
            For Each site In TheExec.sites
                Site_Align = WorksheetFunction.Max(Site_Align, TheExec.sites(site).TestNumber)
            Next site
            
            For Each site In TheExec.sites
                TheExec.sites(site).TestNumber = Site_Align
            Next site
            
            '''************************************************************************************************************************************************'''
            ''' Base on the search result, print EQN and BinCut CP voltage to datalog and bin out the failed DUT.
            '''************************************************************************************************************************************************'''
            judge_PF inst_info, CurrentPassBinCutNum
            
            '''//For the performance mode that does not exist in the bincut table
            RestoreSkipTestBin2Site inst_info.p_mode
        Else '''BinCut check.
            If UCase(inst_info.inst_name) Like "*ELB_PBC*AMPLP5P*" Or UCase(inst_info.inst_name) Like "*ILB_PBC*AMPLP5P*" Then
                judge_PF_func inst_info.p_mode, inst_info.Test_Type, pattPass, True, inst_info.inst_name
            Else
                judge_PF_func inst_info.p_mode, inst_info.Test_Type, pattPass
            End If
        End If
    End If
Exit Function
errHandler:
    Call Print_Error_Message(Error_Info, "LIB_VDD_BINNING", "update_sort_result") 'Add ErrHandler 2023/08/01
    If AbortTest Then Exit Function Else Resume Next
End Function

'20210901: Modified to rename "IndexLevelPerSite As New SiteLong" as "Step_Current As New SiteLong" for Public Type Instance_Info.
'20201204: Modified to initialize "inst_info.IndexLevelPerSite = -1" in the vbt function initialize_inst_info.
'20201117: Modified to use "tlResultModeDomain" for pattern burst=Yes and decomposePatt=No. Requested by Leon Weng.
'20201111: Modified to replace the vbt function name "print_bincut_power" with "print_bincut_voltage".
'20201029: Modified to use inst_info.is_BV_Safe_Voltage_printed and inst_info.is_BV_Payload_Voltage_printed.
'20201027: Modified to use "Public Type Instance_Info".
'20201022: Modified to fix the vbt code for PatternBurst result issues. Requested by TSMC PCLINZG.
'20200827: Modified to remove the redundant site-loop.
'20200319: Modified to switch off save_core_power_vddbinning and restore_core_power_vddbinning if Flag_Enable_Rail_Switch = True.
'20200203: Modified to use the function "print_bincut_power".
'20200113: Modified for pattern bursted without decomposing pattern.
'20191127: Modified for the revised InitVddBinTable.
'20190627: Modified to use the global variable "pinGroup_BinCut" for BinCut powerPins.
'20190617: Modified to use siteDouble "CorePowerStored()" to save/restore voltages for BinCut powerPins.
Public Function run_prepatt(PrePatt As Pattern, inst_name As String, p_mode As Integer, PrePattPass As SiteBoolean, result_mode As tlResultMode, Optional Special_Voltage_setup As Boolean)
On Error GoTo errHandler 'Add ErrHandler 2023/08/01
    Dim site As Variant
    Dim CorePowerStored() As New SiteDouble
    Dim i As Integer
    Dim siteResult As New SiteBoolean
    Dim inst_info As Instance_Info
    If PrePatt.value <> "" Then
        '''init
        inst_info.is_BV_Safe_Voltage_printed = False
        inst_info.is_BV_Payload_Voltage_printed = False
        inst_info.inst_name = inst_name
        inst_info.p_mode = p_mode
        inst_info.Special_Voltage_setup = Special_Voltage_setup
        inst_info.PrePatt = PrePatt
        inst_info.Step_Current = -1
    
        '''//siteDouble "CorePowerStored()" is used to save/restore voltages for BinCut powerDomains.
        ReDim CorePowerStored(UBound(pinGroup_BinCut))
        
        For i = 0 To UBound(pinGroup_BinCut)
            CorePowerStored(i) = 0
        Next i
        
        '''//Save payload voltages of CorePower and OtherRail powerPins before init pattern.
        If Flag_noRestoreVoltageForPrepatt = False Then
            save_core_power_vddbinning CorePowerStored
        End If
                
        '''//Set to nominal voltage (NV).
'''ToDo: If initial voltages and safe voltage(init voltage) use the same DC category, we will skip "set_core_power_vddbinning_VT" after initial voltages...
        set_core_power_vddbinning_VT VddBinName(p_mode), "NV"
        TheHdw.Wait 0.0001
        
        '''//Print safe voltages(init voltages) for PrePatt(init patt).
        print_bincut_voltage inst_info, , Flag_Remove_Printing_BV_voltages
                
        '''//Set "result_mode = tlResultModeModule" (return a unique pass/fail result for each module and time domain) if pattern bursted without decomposing pattern.
        '''20201117: Modified to use "tlResultModeDomain" for pattern burst=Yes and decomposePatt=No. Requested by Leon Weng.
        result_mode = tlResultModeDomain
        
        '''//Run the pattern
        Call TheHdw.patterns(PrePatt).test(pfAlways, 0, result_mode)
        DebugPrintFunc PrePatt.value
        
        '''//Check pattern Pass/Fail.
        '''//Warning!!! currently "TheHdw.Digital.Patgen.PatternBurstPassedPerSite" doesn't support "result_mode=tlResultModeModule" with PatternBurst=Yes and DecomposePatt=No.
        '''20201022: Modified to fix the vbt code for PatternBurst result issues. Requested by TSMC PCLINZG.
        PrePattPass = TheHdw.Digital.Patgen.PatternBurstPassedPerSite
        
        '''//Restore the BinCut voltages for payload patterns after PrePatt.
        If Flag_noRestoreVoltageForPrepatt = False Then
            restore_core_power_vddbinning CorePowerStored
        End If
    End If
Exit Function
errHandler:
    Call Print_Error_Message(Error_Info, "LIB_VDD_BINNING", "run_prepatt") 'Add ErrHandler 2023/08/01
    If AbortTest Then Exit Function Else Resume Next
End Function

'20210803: Modified to update inst_info.ids_current = IDS_for_BinCut(VddBinStr2Enum(powerDomain)).Real in the vbt function initialize_inst_info and remove the redundant vbt function set_IDS_current.
'20200923: Modified to remove "clear_after_patt".
'20200921: Discussed "RTOS_bootup_relay" / "KeepAliveFlag" / "spi_ttr_flag" with SWLINZA and PCLINZG. We decided to remove these SPI/RTOS branches because RTOS didn't use pattern test since Cebu/Sicily/Tonga/JC-Chop/Ellis/Bora.
'20190706: Modified for the new datatype of power_seq.
Public Function sort_power_seqence(Power_List As String, Power_Seq() As String) As Long
On Error GoTo errHandler 'Add ErrHandler 2023/08/01
    Dim strAry_Performance_Mode() As String
    Dim idxA As Integer
    Dim idxB As Integer
    Dim idxC As Integer
    Dim replace_flag As Boolean
    Dim strTemp As String
    Dim strReplaced As String
    '''init
    idxA = 0
    idxB = 0
    idxC = 0
    
    strAry_Performance_Mode = Split(Power_List, ",")
    ReDim Power_Seq(UBound(strAry_Performance_Mode))
    
    For idxA = 0 To UBound(strAry_Performance_Mode)
        replace_flag = False
        If idxA = 0 Then
            Power_Seq(idxB) = strAry_Performance_Mode(idxA)
            idxB = idxB + 1
        Else
            For idxC = 0 To UBound(Power_Seq)
                If Power_Seq(idxC) <> "" Then
                    If BinCut(VddBinStr2Enum(strAry_Performance_Mode(idxA)), 1).MAX_ID < BinCut(VddBinStr2Enum(Power_Seq(idxC)), 1).MAX_ID Then
                        If replace_flag = False Then
                            replace_flag = True
                            strTemp = Power_Seq(idxC)
                            Power_Seq(idxC) = strAry_Performance_Mode(idxA)
                            idxC = idxC + 1
                        End If
                    End If
                    If replace_flag = True Then
                        strReplaced = Power_Seq(idxC)
                        Power_Seq(idxC) = strTemp
                        strTemp = strReplaced
                    End If
                Else
                   If replace_flag = False Then
                        Power_Seq(idxC) = strAry_Performance_Mode(idxA)
                   Else
                        Power_Seq(idxC) = strTemp
                   End If
                   Exit For
                End If
            Next idxC
        End If
    Next idxA
Exit Function
errHandler:
    Call Print_Error_Message(Error_Info, "LIB_VDD_BINNING", "sort_power_seqence") 'Add ErrHandler 2023/08/01
    If AbortTest Then Exit Function Else Resume Next
End Function

'20210812: C651 Toby updated the rules that Product voltage=BinCut voltage+binning_GuardBand.
'20210803: Modified to calculate step voltages with BV_StepVoltage.
'20210728: C651 Toby updated the rules of step voltage calculation. It should not use GB_delta.
'20210722: Modified to update VBIN_IDS_ZONE(p_mode).Product_Voltage.
'20210223: Modified to replace "Dim step As Long" with "Dim idx_step As Long".
'20191127: Modified for the revised InitVddBinTable.
'20190716: Modified to unify the unit for IDS. ids_current with unit mA.
'20190507: Modified to add "Cdec" to avoid double format accuracy issues.
Public Function Generate_IDS_ZONE_Voltage_Per_Site(ids_current As SiteDouble, p_mode As Integer)
On Error GoTo errHandler 'Add ErrHandler 2023/08/01
    Dim site As Variant
    Dim Test_Type As testType
    Dim idx_step As Long
    Dim Remainder As Double
    Dim voltage_Temp As Double
    Dim Zone_Num As Integer
    Dim dbl_CPVmax As Double
    Dim dbl_CPVmin As Double
    Dim PassBinNum As Long
    Dim step_in_BinCut As Long
'''//==================================================================================================================================================================================//'''
'''//Note:
'''1. Calculate the voltage of each step and each zone by IDS values for p_mode.
'''2. C651 Toby updated the rules of step voltage calculation. It should not use GB_delta, 20210728.
'''3. C651 Toby updated the rules that Product voltage=BinCut voltage+binning_GuardBand, 20210812.
'''//==================================================================================================================================================================================//'''
    '''//The default Testtype is TD
    Test_Type = testType.TD
    
    If VBIN_IDS_ZONE(p_mode).Used = True Then
        For Each site In TheExec.sites
            For Zone_Num = 0 To VBIN_IDS_ZONE(p_mode).IDS_RANGE_COUNT(Test_Type)        'loop IDS Range for all IDS Zone
                For idx_step = 0 To VBIN_IDS_ZONE(p_mode).Max_Step(Zone_Num) - 1        'loop the step in the IDS Zone
                    '''************************************************************************************************************************************************'''
                    '''//Formula: CP voltage = C-M*log10(IDS).
                    '''CP voltage with unit: mV. ids_current with unit mA.
                    '''************************************************************************************************************************************************'''
                    ''20230201: Avoid IDS < 0 scenario
                    If ids_current < 0 Then ids_current = 99999
                    
                    voltage_Temp = VBIN_IDS_ZONE(p_mode).c(Zone_Num, idx_step) - VBIN_IDS_ZONE(p_mode).m(Zone_Num, idx_step) * (log(ids_current) / log(10))
                    
                    '''//For LVCC, floor the value by step_voltage defined in the header of sheet "Vdd_Binning_Def".
                    '''//Floor step voltages of each step in Dynamic_IDS_Zone by BV_StepVoltage.
                    Remainder = Floor(voltage_Temp / BV_StepVoltage)
                    voltage_Temp = Remainder * BV_StepVoltage
                    
                    '''//Update PassBinNum and step_in_BinCut for each step in IDS_Zone.
                    PassBinNum = VBIN_IDS_ZONE(p_mode).passBinCut(Zone_Num, idx_step)
                    step_in_BinCut = VBIN_IDS_ZONE(p_mode).EQ_Num(Zone_Num, idx_step) - 1
                    
                    '''//Check if voltage of each step in IDS_Zone is between CPVmax and CPVmin.
                    dbl_CPVmax = BinCut(p_mode, PassBinNum).CP_Vmax(step_in_BinCut)
                    dbl_CPVmin = BinCut(p_mode, PassBinNum).CP_Vmin(step_in_BinCut)
                    
                    If CDec(voltage_Temp) > CDec(dbl_CPVmax) Then
                        voltage_Temp = dbl_CPVmax
                    ElseIf CDec(voltage_Temp) < CDec(dbl_CPVmin) Then
                        voltage_Temp = dbl_CPVmin
                    End If
                    
                    '''//Calculate GradeVDD for each step.
                    '''Efuse product voltage(GradeVDD) = BinCut voltage(Grade) + binning_GuardBand.
                    '''20210812: C651 Toby updated the rules that Product voltage=BinCut voltage+binning_GuardBand.
                    VBIN_IDS_ZONE(p_mode).Voltage(Zone_Num, idx_step) = voltage_Temp
                    VBIN_IDS_ZONE(p_mode).Product_Voltage(Zone_Num, idx_step) = voltage_Temp + BinCut(p_mode, PassBinNum).CP_GB(step_in_BinCut)
                Next idx_step
            Next Zone_Num
        Next site
    Else
        Call Print_Error_Message(Error_Info, "LIB_VDD_BINNING", "Generate_IDS_ZONE_Voltage_Per_Site", VddBinName(p_mode) & ", it doesn't have any correct IDS ZONE for Generate_IDS_ZONE_Voltage_Per_Site. Error!!!")
        'TheExec.Datalog.WriteComment VddBinName(p_mode) & ", it doesn't have any correct IDS ZONE for Generate_IDS_ZONE_Voltage_Per_Site. Error!!!"
    End If
Exit Function
errHandler:
    Call Print_Error_Message(Error_Info, "LIB_VDD_BINNING", "Generate_IDS_ZONE_Voltage_Per_Site") 'Add ErrHandler 2023/08/01
    If AbortTest Then Exit Function Else Resume Next
End Function

'20181101: Modified for the format of .CSV file.
'20181030: Modified to integrate all IDS distribution related functions.
'20181026: Modified for IDS, by MSLi.
Public Function Print_IDS_ZONE_Table_to_sheet()
On Error GoTo errHandler 'Add ErrHandler 2023/08/01
    Dim site As Variant
    Dim wb As Workbook
    Dim Test_Type As testType
    Dim p_mode As Integer
    Dim i As Long, k As Long, L As Long
    Dim ids_range_step() As Long 'Dim ids_range_step(MaxPerformanceModeCount) As Long
    Dim IDS_current_Max() As Double 'Dim IDS_current_Max(MaxPerformanceModeCount) As Double
    Dim p_col As Integer, p_row As Integer
    Dim max_print_step As Integer
    Dim SheetCnt As Long
    Dim str_CurPath As String
    Dim SheetExist As Boolean
    Dim str_Output As String
    Dim str_header As String
    Dim IsHeaderPrinted As Boolean
    '''init
    str_CurPath = "D:\IDS_ZONE_TABLE.csv"
    Open str_CurPath For Output As #1
    Set wb = Application.ActiveWorkbook
    SheetCnt = ActiveWorkbook.Sheets.Count
    SheetExist = False
    Test_Type = testType.TD
    p_col = 1
    p_row = 1
    str_Output = vbNullString
    str_header = vbNullString
    IsHeaderPrinted = False
    
    '''use the max step count to print the table
    max_print_step = Max_V_Step_per_IDS_Zone
    
    For i = 0 To MaxPerformanceModeCount - 1
        ids_range_step(i) = 0
        IDS_current_Max(i) = 0
    Next i
    
    For p_mode = 0 To MaxPerformanceModeCount - 1
        If VBIN_IDS_ZONE(p_mode).Used = True Then
            IsHeaderPrinted = False
        
            '''//Performance mode//
            Print #1, VddBinName(p_mode)
            
            For k = 0 To VBIN_IDS_ZONE(p_mode).IDS_RANGE_COUNT(Test_Type)
                '''//IDS range & Start bin
                For L = 0 To testType.Total_Count - 1
                    If L = 0 Then
                        str_header = "IDS Range"
                        str_Output = VBIN_IDS_ZONE(p_mode).Ids_range(k, L)
                    Else
                        str_header = str_header & "," & "IDS Range"
                        str_Output = str_Output & "," & VBIN_IDS_ZONE(p_mode).Ids_range(k, L)
                    End If
                    
                    str_header = str_header & "," & "Start Bin"
                    str_Output = str_Output & "," & VBIN_IDS_ZONE(p_mode).IDS_Start_EQ_Num(k, L)
                Next L
                
                '''//C
                For i = 0 To VBIN_IDS_ZONE(p_mode).Max_Step(0) - 1
                    str_header = str_header & "," & "C"
                    If (VBIN_IDS_ZONE(p_mode).Max_Step(k) - 1) <= (VBIN_IDS_ZONE(p_mode).Max_Step(0) - 1) Then
                        str_Output = str_Output & "," & VBIN_IDS_ZONE(p_mode).c(k, i)
                    Else
                        str_Output = str_Output & "," & " "
                    End If
                Next i
                
                '''//M
                For i = 0 To VBIN_IDS_ZONE(p_mode).Max_Step(0) - 1
                    str_header = str_header & "," & "M"
                     If (VBIN_IDS_ZONE(p_mode).Max_Step(k) - 1) <= (VBIN_IDS_ZONE(p_mode).Max_Step(0) - 1) Then
                        str_Output = str_Output & "," & VBIN_IDS_ZONE(p_mode).m(k, i)
                    Else
                        str_Output = str_Output & "," & " "
                    End If
                Next i
                
                '''//EQN
                For i = 0 To VBIN_IDS_ZONE(p_mode).Max_Step(0) - 1
                    str_header = str_header & "," & "EQN"
                    If (VBIN_IDS_ZONE(p_mode).Max_Step(k) - 1) <= (VBIN_IDS_ZONE(p_mode).Max_Step(0) - 1) Then
                        str_Output = str_Output & "," & VBIN_IDS_ZONE(p_mode).EQ_Num(k, i)
                    Else
                        str_Output = str_Output & "," & " "
                    End If
                Next i
                
                '''//PASSBINCUT
                For i = 0 To VBIN_IDS_ZONE(p_mode).Max_Step(0) - 1
                    str_header = str_header & "," & "PASSBINCUT"
                     If (VBIN_IDS_ZONE(p_mode).Max_Step(k) - 1) <= (VBIN_IDS_ZONE(p_mode).Max_Step(0) - 1) Then
                        str_Output = str_Output & "," & VBIN_IDS_ZONE(p_mode).passBinCut(k, i)
                    Else
                        str_Output = str_Output & "," & " "
                    End If
                Next i
                
                '''//print out the header for each performance mode
                If IsHeaderPrinted = False Then
                    Print #1, str_header
                    str_header = vbNullString
                    IsHeaderPrinted = True
                End If
                
                Print #1, str_Output
            Next k
            
            Print #1, ""
        End If
    Next p_mode
    
    '''//Close the csv file//
    Close #1
    
    '''//Import csv to the sheet"IDS_ZONE_TABLE"//
    'Open str_CurPath For Input As #1
Exit Function
errHandler:
    Call Print_Error_Message(Error_Info, "LIB_VDD_BINNING", "Print_IDS_ZONE_Table_to_sheet") 'Add ErrHandler 2023/08/01
    If AbortTest Then Exit Function Else Resume Next
End Function

'20210223: Modified to replace "Dim step As Long" with "Dim idx_step As Long".
'20191204: Modified for the revised initVddBinTable.
'20181101: Modified for the format of .CSV file.
'20181031: Modified for the sheet "IDS_ZONE_voltage".
Public Function Print_IDS_ZONE_voltage_to_sheet()
On Error GoTo errHandler 'Add ErrHandler 2023/08/01
    Dim site As Variant
    Dim Test_Type As testType
    Dim idx_step As Long
    Dim p_mode As Integer
    Dim i As Long
    Dim ids_range_step() As Long 'Dim ids_range_step(MaxPerformanceModeCount) As Long
    Dim IDS_current_Max() As Double 'Dim IDS_current_Max(MaxPerformanceModeCount) As Double
    Dim p_col As Integer, p_row As Integer
    Dim Zone_Number As Long
    Dim SheetCnt As Long
    Dim str_CurPath As String
    Dim SheetExist As Boolean
    Dim str_Output As String
    Dim str_header As String
    Dim IsHeaderPrinted As Boolean
    '''init
    str_CurPath = "D:\IDS_ZONE_Voltage.csv"
    Open str_CurPath For Output As #1
    p_col = 1
    p_row = 1
    Test_Type = testType.TD
    SheetCnt = ActiveWorkbook.Sheets.Count
    SheetExist = False
    str_Output = vbNullString
    str_header = vbNullString
    IsHeaderPrinted = False
        
    For i = 0 To MaxPerformanceModeCount - 1
        ids_range_step(i) = 0
        IDS_current_Max(i) = 0
    Next i
    
    For p_mode = 0 To MaxPerformanceModeCount - 1
        If VBIN_IDS_ZONE(p_mode).Used = True Then
            IsHeaderPrinted = False
            
            '''//Performance mode//
            Print #1, VddBinName(p_mode)
            
            For Each site In TheExec.sites
                str_header = "Site"
                str_Output = site
                
                str_header = str_header & "," & "IDS_ZONE_NUMBER"
                str_Output = str_Output & "," & VBIN_IDS_ZONE(p_mode).IDS_ZONE_NUMBER
                
                For Zone_Number = 0 To VBIN_IDS_ZONE(p_mode).IDS_RANGE_COUNT(Test_Type)
                    For idx_step = 0 To VBIN_IDS_ZONE(p_mode).Max_Step(Zone_Number) - 1
                        str_header = str_header & "," & "V" & "_" & "step" & idx_step
                        str_Output = str_Output & "," & VBIN_IDS_ZONE(p_mode).Voltage(Zone_Number, idx_step)
                    Next idx_step
                Next Zone_Number
                
                If IsHeaderPrinted = False Then
                    Print #1, str_header
                    str_header = vbNullString
                    IsHeaderPrinted = True
                End If
                
                Print #1, str_Output
            Next site
            Print #1, ""
        End If
    Next p_mode
    
    '''//Close the csv file//
    Close #1
Exit Function
errHandler:
    Call Print_Error_Message(Error_Info, "LIB_VDD_BINNING", "Print_IDS_ZONE_voltage_to_sheet") 'Add ErrHandler 2023/08/01
    If AbortTest Then Exit Function Else Resume Next
End Function

'20210722: Modified to use VBIN_IDS_ZONE(p_mode).Product_Voltage(Max_IDS_Zone, Max_IDS_Step) and DYNAMIC_VBIN_IDS_ZONE(p_mode).Product_Voltage(Max_IDS_Step) for GradeVDD.
'20210405: Modified to remove "PassBinCutList_per_Zone(Max_IDS_Zone) As Long"
'20210303: Modified to replace "Dim steps As Long" with "Dim idx_step As Long".
'20191127: Modified for the revised InitVddBinTable.
Public Function init_IDS_ZONE()
On Error GoTo errHandler 'Add ErrHandler 2023/08/01
    Dim ids_zone_num As Long
    Dim Test_Type As Long
    Dim idx_step As Long
    Dim p_mode As Integer
    For p_mode = 0 To MaxPerformanceModeCount - 1
        VBIN_IDS_ZONE(p_mode).Used = False
        
        For idx_step = 0 To Max_IDS_Step
            For ids_zone_num = 0 To Max_IDS_Zone - 1
                VBIN_IDS_ZONE(p_mode).Max_Step(ids_zone_num) = 0
                VBIN_IDS_ZONE(p_mode).c(ids_zone_num, idx_step) = 0
                VBIN_IDS_ZONE(p_mode).m(ids_zone_num, idx_step) = 0
                VBIN_IDS_ZONE(p_mode).passBinCut(ids_zone_num, idx_step) = 0
                VBIN_IDS_ZONE(p_mode).EQ_Num(ids_zone_num, idx_step) = 0
                VBIN_IDS_ZONE(p_mode).Voltage(ids_zone_num, idx_step) = 0
                VBIN_IDS_ZONE(p_mode).Product_Voltage(Max_IDS_Zone, Max_IDS_Step) = 0
                
                For Test_Type = 0 To testType.Total_Count - 1
                    VBIN_IDS_ZONE(p_mode).Ids_range(ids_zone_num, Test_Type) = 0
                    VBIN_IDS_ZONE(p_mode).IDS_Start_EQ_Num(ids_zone_num, Test_Type) = 0
                    VBIN_IDS_ZONE(p_mode).IDS_START_STEP(ids_zone_num, Test_Type) = 0
                    VBIN_IDS_ZONE(p_mode).IDS_RANGE_COUNT(Test_Type) = 0
                Next Test_Type
            Next ids_zone_num
        Next idx_step
    Next p_mode
Exit Function
errHandler:
    Call Print_Error_Message(Error_Info, "LIB_VDD_BINNING", "init_IDS_ZONE") 'Add ErrHandler 2023/08/01
    If AbortTest Then Exit Function Else Resume Next
End Function

'20210405: Modified to remove "PassBinCutList_per_Zone(Max_IDS_Zone) As Long"
'20210223: Modified to replace "Dim step As Long" with "Dim idx_step As Long".
'20200703: Modiifed to use "check_Sheet_Range".
'20200506: Modified to revise the vbt code for parsing the sheet.
'20200505: Modified to check if "IDS_Distribution" exists in the workbook.
'20191127: Modified for the revised InitVddBinTable.
'20160614: Modified by TSMC Jack.
Public Function initIDSTable()
On Error GoTo errHandler 'Add ErrHandler 2023/08/01
    Dim ws_def As Worksheet
    Dim wb As Workbook
    Dim sheetName As String
    Dim MaxRow As Long
    Dim maxcol As Long
    Dim Row As Long, Col As Long
    Dim p_mode As Integer
    Dim idx_step As Long
    Dim Test_Type As Long
    Dim col_testTypeSelected As Long
    Dim zoneTemp As Integer
    Dim Row_of_Title As Integer
    Dim Row_of_step0 As Integer
    Dim col_TestType() As Integer 'Dim col_TestType(MaxTestType) As Integer20230526
    Dim strTemp As String
    Dim isSheetFound As Boolean
    '''*****************************************************************'''
    '''//Check if the sheet exists
    sheetName = "IDS_Distribution_" & TheExec.CurrentJob
    Set wb = Application.ActiveWorkbook
    Call check_Sheet_Range(sheetName, wb, ws_def, MaxRow, maxcol, isSheetFound, False)
    If isSheetFound = False Then
        sheetName = "IDS_Distribution"
        Call check_Sheet_Range(sheetName, wb, ws_def, MaxRow, maxcol, isSheetFound)
    End If
    ReDim col_TestType(testType.Total_Count) As Integer
    
    '''*****************************************************************'''
    If isSheetFound = True Then
        '''//init
        Version_IDS_Distribution = vbNullString
        
        '''//Initialize the array
        '''//Please check "Enum TestType" and "MaxTestType" in GlobalVariable.
        For Test_Type = 0 To testType.Total_Count - 1
            For p_mode = 0 To MaxPerformanceModeCount - 1
                For zoneTemp = 0 To Max_IDS_Zone - 1
                    IDS_Distribution_Table(p_mode).range(zoneTemp, Test_Type) = 0
                    IDS_Distribution_Table(p_mode).Start_Bin(zoneTemp, Test_Type) = 0
                    IDS_Distribution_Table(p_mode).RANGE_COUNT = 0
                    IDS_Distribution_Table(p_mode).Used = False
                Next zoneTemp
            Next p_mode
            
            '''init the array to store the column number of each TestType.
            col_TestType(Test_Type) = 0
        Next Test_Type
    
        '''//Find the start point of the header.
        For Row = 1 To MaxRow
            For Col = 1 To maxcol
                If LCase(ws_def.Cells(Row, Col).value) Like "*rev*" And LCase(ws_def.Cells(Row + 1, Col).value) Like "td" Then
                    Version_IDS_Distribution = ws_def.Cells(1, 2).value
                    Row_of_Title = Row + 1
                End If
                
                If Row_of_Title > 0 And Row = Row_of_Title Then
                    If LCase(ws_def.Cells(Row, Col).value) <> "" Then
                        col_TestType(decide_test_type_for_string(ws_def.Cells(Row, Col).value)) = Col
                    End If
                End If
            Next Col
        Next Row
        
        If Row_of_Title > 0 And col_TestType(testType.TD) > 0 Then
            For Row = Row_of_Title + 1 To MaxRow
                If LCase(ws_def.Cells(Row, col_TestType(testType.TD)).value) = "ids range" And LCase(ws_def.Cells(Row, col_TestType(testType.TD) + 1).value) = "start bin" Then '''//Find the row with "IDS Range","Start Bin"
                    p_mode = 0
                    strTemp = ws_def.Cells(Row - 1, col_TestType(testType.TD)).value
                    
                    If VddbinPmodeDict.Exists(strTemp) Then
                        p_mode = VddBinStr2Enum(strTemp)
                        Row_of_step0 = Row + 1
                    Else
                        p_mode = 0
                        Call Print_Error_Message(Error_Info, "LIB_VDD_BINNING", "initIDSTable", sheetName & " doesn't have any correct Performance_mode in row" & Row & ". Error!!!")
                        'TheExec.Datalog.WriteComment sheetName & " doesn't have any correct Performance_mode in row" & Row & ". Error!!!"
                        'TheExec.ErrorLogMessage SheetName & " doesn't have any correct Performance_mode in row" & Row & ". Error!!!"
                    End If
                        
                    If p_mode > 0 Then
                        For Test_Type = 0 To testType.Total_Count - 1
                            col_testTypeSelected = col_TestType(Test_Type)
                            
                            If col_testTypeSelected > 0 Then
                                If ws_def.Cells(Row_of_step0 - 2, col_testTypeSelected).value = strTemp Then
                                    idx_step = 0
                                    Row = Row_of_step0
                                    
                                    If IsNumeric(ws_def.Cells(Row, col_testTypeSelected).value) And (IsEmpty(ws_def.Cells(Row, col_testTypeSelected).value) = False) Then
                                        While (LCase(ws_def.Cells(Row, col_testTypeSelected).value) <> "end" And (ws_def.Cells(Row, col_testTypeSelected).value) <> "")
                                            IDS_Distribution_Table(p_mode).range(idx_step, Test_Type) = CDbl(ws_def.Cells(Row, col_testTypeSelected).value)
                                            IDS_Distribution_Table(p_mode).Start_Bin(idx_step, Test_Type) = CDbl(ws_def.Cells(Row, col_testTypeSelected + 1).value)
                                            IDS_Distribution_Table(p_mode).RANGE_COUNT = idx_step
                                            IDS_Distribution_Table(p_mode).Used = True
                                            idx_step = idx_step + 1
                                            Row = Row + 1 '''Row Offset
                                        Wend
                                    End If
                                    
                                    Row = Row_of_step0
                                Else
                                    Call Print_Error_Message(Error_Info, "LIB_VDD_BINNING", "initIDSTable", sheetName & " doesn't have the correct Performance_mode in row" & (Row_of_step0 - 2) & ", col" & col_testTypeSelected & " consistent with other TestType columns. Error!!!")
                                    'TheExec.Datalog.WriteComment sheetName & " doesn't have the correct Performance_mode in row" & (Row_of_step0 - 2) & ", col" & col_testTypeSelected & " consistent with other TestType columns. Error!!!"
                                    'TheExec.ErrorLogMessage SheetName & " doesn't have the correct Performance_mode in row" & Row & ", col" & col_TestType(TestType.TD) & " consistent with other TestType columns. Error!!!"
                                End If
                            End If
                        Next Test_Type
                    End If
                End If
            Next Row
        Else
            Call Print_Error_Message(Error_Info, "LIB_VDD_BINNING", "initIDSTable", sheetName & " doesn't have correct format of the header. Error!!!")
            'TheExec.Datalog.WriteComment sheetName & " doesn't have correct format of the header. Error!!!"
            'TheExec.ErrorLogMessage sheetName & " doesn't have correct format of the header. Error!!!"    ''20230807: Removed by sanity check
        End If
    End If '''If isSheetFound = True
Exit Function
errHandler:
    Call Print_Error_Message(Error_Info, "LIB_VDD_BINNING", "initIDSTable") 'Add ErrHandler 2023/08/01
    If AbortTest Then Exit Function Else Resume Next
End Function

'20210722: Modified to use VBIN_IDS_ZONE(p_mode).Product_Voltage(Max_IDS_Zone, Max_IDS_Step) and DYNAMIC_VBIN_IDS_ZONE(p_mode).Product_Voltage(Max_IDS_Step) for GradeVDD.
'20210303: Modified to replace "Dim steps As Long" with "Dim idx_step As Long".
'20191127: Modified for the revised InitVddBinTable.
Public Function init_IDS_ZONE_Voltage()
On Error GoTo errHandler 'Add ErrHandler 2023/08/01
    Dim ids_zone_num As Long
    Dim idx_step As Long
    Dim p_mode As Integer
    For p_mode = 0 To MaxPerformanceModeCount - 1
        For idx_step = 0 To Max_IDS_Step
            For ids_zone_num = 0 To Max_IDS_Zone - 1
                VBIN_IDS_ZONE(p_mode).Voltage(ids_zone_num, idx_step) = 0 '''siteDouble
                VBIN_IDS_ZONE(p_mode).Product_Voltage(ids_zone_num, idx_step) = 0 '''siteDouble
            Next ids_zone_num
        Next idx_step
    Next p_mode
Exit Function
errHandler:
    Call Print_Error_Message(Error_Info, "LIB_VDD_BINNING", "init_IDS_ZONE_Voltage") 'Add ErrHandler 2023/08/01
    If AbortTest Then Exit Function Else Resume Next
End Function

'20210906: Modified to remove the unused variables for the vbt function Generate_IDS_ZONE_RANGE.
'20200512: Modified to merge Generate_IDS_Zone_with_IDS_Distribution_Table and Generate_IDS_Zone_NO_IDS_Distribution_Table into "Generate_IDS_Zone_with_IDS_Distribution_Table".
'20160614: Modified by TSMC Jack.
Public Function Generate_IDS_ZONE_RANGE()
On Error GoTo errHandler
    '''//Generate IDS_Zone for each p_mode with/without IDS_Distribution_Table.
    Call Generate_IDS_Zone_with_IDS_Distribution_Table
Exit Function
errHandler:
    Call Print_Error_Message(Error_Info, "LIB_VDD_BINNING", "Generate_IDS_ZONE_RANGE") 'Add ErrHandler 2023/08/01
    If AbortTest Then Exit Function Else Resume Next
End Function

'20210303: Modified to replace "Dim step As Long" with "Dim idx_step As Long".
'20200512: Created to merge "Generate_IDS_Zone_with_IDS_Distribution_Table" and "Generate_IDS_Zone_NO_IDS_Distribution_Table" into "Generate_IDS_Zone_with_IDS_Distribution_Table".
'20190422: Modified to define the bin number for DUT with IDS on the IDS_limit.
Public Function Generate_IDS_Zone_with_IDS_Distribution_Table()
On Error GoTo errHandler 'Add ErrHandler 2023/08/01
    Dim Test_Type As Integer
    Dim i As Integer
    Dim j As Integer
    Dim RngNum As Integer
    Dim Bincut_step As Integer
    Dim idx_step As Integer
    Dim Ids_zone_cnt As Integer
    Dim DblTemp As Double
    Dim Zone_Num As Integer
    Dim Duplicate_Flag As Boolean
    Dim bincutNum As Variant
    Dim p_mode As Integer
    Dim Zone As Integer
    Dim Max_mode_step As Integer
    Dim Ids_range As Double
    Dim bincut_max_step As Integer
'''//==================================================================================================================================================================================//'''
'''//Note:
'''Determine how many ids zone would be after considering bincut and ids_distribution.
'''//==================================================================================================================================================================================//'''
    For p_mode = 0 To MaxPerformanceModeCount - 1
        '''***********************************************************************'''
        '''[Step0] Initialize the IDS Zone value and Start-Search bin.
        '''***********************************************************************'''
        For Test_Type = 0 To testType.Total_Count - 1
            For Zone_Num = 0 To Max_IDS_Zone
                VBIN_IDS_ZONE(p_mode).Ids_range(Zone_Num, Test_Type) = 0
                VBIN_IDS_ZONE(p_mode).IDS_Start_EQ_Num(Zone_Num, Test_Type) = 0
                VBIN_IDS_ZONE_Temp(p_mode).Ids_range(Zone_Num, Test_Type) = -999
                VBIN_IDS_ZONE_Temp(p_mode).IDS_Start_EQ_Num(Zone_Num, Test_Type) = -999
            Next Zone_Num
        Next Test_Type
        
        '''//If IDS distribution table is parsed, and p_mode is defined in the table.
        If IDS_Distribution_Table(p_mode).Used = True Then
            For Each bincutNum In PassBinCut_ary
                If bincutNum = 1 Then
                    '''***********************************************************************'''
                    '''[Step1] For offline simulation mode, we make man-made data for simulation.
                    '''***********************************************************************'''
                    If TheExec.TesterMode = testModeOffline Then
                        '''If last parameter =-1 means bincut ids limit less than ids distribution range
                        '''If last parameter =0 means bincut ids limit is same as ids distribution range
                    End If
                            
                    '''***********************************************************************'''
                    '''[Step2] Copy ids_distribution table into a ids zone array.
                    '''***********************************************************************'''
                    For Test_Type = 0 To testType.Total_Count - 1 '''For all test type, TD, MBIST, SPI, RTOS, TMPS and LDCBFD
                        Ids_zone_cnt = 0
                        
                        For RngNum = 0 To IDS_Distribution_Table(p_mode).RANGE_COUNT
                            VBIN_IDS_ZONE_Temp(p_mode).Ids_range(RngNum, Test_Type) = IDS_Distribution_Table(p_mode).range(RngNum, Test_Type)
                            Ids_zone_cnt = Ids_zone_cnt + 1
                        Next RngNum
                        
                        VBIN_IDS_ZONE_Temp(p_mode).IDS_RANGE_COUNT(Test_Type) = Ids_zone_cnt
                    Next Test_Type
                End If '''If BincutNum = 1 Then
                                        
                '''***********************************************************************'''
                '''[Step3] Copy bincut table into ids zone array.
                '''***********************************************************************'''
                For Test_Type = 0 To testType.Total_Count - 1 '''For all test type "TD, MBIST, SPI, RTOS, TMPS and LDCBFD"
                    Ids_zone_cnt = VBIN_IDS_ZONE_Temp(p_mode).IDS_RANGE_COUNT(Test_Type)
                    
                    For idx_step = 0 To BinCut(p_mode, bincutNum).Mode_Step
                        VBIN_IDS_ZONE_Temp(p_mode).Ids_range(Ids_zone_cnt, Test_Type) = BinCut(p_mode, bincutNum).IDS_CP_LIMIT(idx_step)
                        Ids_zone_cnt = Ids_zone_cnt + 1
                    Next idx_step
                    
                    VBIN_IDS_ZONE_Temp(p_mode).IDS_RANGE_COUNT(Test_Type) = Ids_zone_cnt
                Next Test_Type
            Next bincutNum
                                    
            '''***********************************************************************'''
            '''[Step4] Screen out the duplicate ids limit from vbin_ids_zone().
            '''***********************************************************************'''
            For Test_Type = 0 To testType.Total_Count - 1
                Zone_Num = 0
                Ids_zone_cnt = VBIN_IDS_ZONE_Temp(p_mode).IDS_RANGE_COUNT(Test_Type)
                For i = 0 To Ids_zone_cnt - 1
                    Duplicate_Flag = False
                    
                    For j = i + 1 To Ids_zone_cnt
                        If VBIN_IDS_ZONE_Temp(p_mode).Ids_range(i, Test_Type) = VBIN_IDS_ZONE_Temp(p_mode).Ids_range(j, Test_Type) Then
                            Duplicate_Flag = True
                            Exit For
                        End If
                    Next j
                    
                    If Duplicate_Flag = False Then
                        VBIN_IDS_ZONE(p_mode).Ids_range(Zone_Num, Test_Type) = VBIN_IDS_ZONE_Temp(p_mode).Ids_range(i, Test_Type)
                        Zone_Num = Zone_Num + 1
                    End If
                Next i
                
                VBIN_IDS_ZONE(p_mode).IDS_RANGE_COUNT(Test_Type) = Zone_Num - 1 'This is for fit in with original bincut vbt structure.
                VBIN_IDS_ZONE(p_mode).Used = True                               'This auxiliary setting is to ensure no error in check_ids test instance.
            Next Test_Type

            '''***********************************************************************'''
            '''[Step5] Sorting the merged array in ascending way because IDS Zone is from low to high.
            '''***********************************************************************'''
            For Test_Type = 0 To testType.Total_Count - 1
                Ids_zone_cnt = VBIN_IDS_ZONE(p_mode).IDS_RANGE_COUNT(Test_Type)
                For i = 0 To Ids_zone_cnt - 1
                    For j = i + 1 To Ids_zone_cnt
                        If (VBIN_IDS_ZONE(p_mode).Ids_range(i, Test_Type) > VBIN_IDS_ZONE(p_mode).Ids_range(j, Test_Type)) Then
                            DblTemp = VBIN_IDS_ZONE(p_mode).Ids_range(j, Test_Type)
                            VBIN_IDS_ZONE(p_mode).Ids_range(j, Test_Type) = VBIN_IDS_ZONE(p_mode).Ids_range(i, Test_Type)
                            VBIN_IDS_ZONE(p_mode).Ids_range(i, Test_Type) = DblTemp
                        End If
                    Next j
                Next i
            Next Test_Type
            
            '''***********************************************************************'''
            '''[Step6] Determine the start-search level for merged bincut ids limits.
            '''***********************************************************************'''
            For Each bincutNum In PassBinCut_ary
                bincut_max_step = BinCut(p_mode, bincutNum).Mode_Step
                
                For Test_Type = 0 To testType.Total_Count - 1
                    For Zone_Num = 0 To VBIN_IDS_ZONE(p_mode).IDS_RANGE_COUNT(Test_Type)
                        Ids_range = VBIN_IDS_ZONE(p_mode).Ids_range(Zone_Num, Test_Type)
                        
                        '''//If any ids zone execeed the maximum ids of bincut, its start bin is set to 0.
                        If Ids_range >= BinCut(p_mode, bincutNum).IDS_CP_LIMIT(bincut_max_step) Then
                            VBIN_IDS_ZONE(p_mode).IDS_Start_EQ_Num(Zone_Num, Test_Type) = 0
                        Else
                            For j = 0 To IDS_Distribution_Table(p_mode).RANGE_COUNT - 1
                                If Ids_range >= IDS_Distribution_Table(p_mode).range(j, Test_Type) And Ids_range < IDS_Distribution_Table(p_mode).range(j + 1, Test_Type) Then
                                    VBIN_IDS_ZONE(p_mode).IDS_Start_EQ_Num(Zone_Num, Test_Type) = IDS_Distribution_Table(p_mode).Start_Bin(j, Test_Type)
                                    Exit For
                                End If
                            Next j
                        End If '''If Ids_range >= VDD_BIN(P_mode, BincutNum).IDS_CP_LIMIT(bincut_max_step) Then
                    Next Zone_Num
                Next Test_Type
            Next bincutNum
            
        '''//If IDS distribution table is not parsed, or p_mode is not defined in the table.
        Else
            '''//The default Testtype is TD
            Test_Type = testType.TD
            Zone_Num = 0
        
            For Each bincutNum In PassBinCut_ary
                If BinCut(p_mode, bincutNum).ExcludedPmode = False Then
                    VBIN_IDS_ZONE(p_mode).Used = True '''With this statement, then no argue(error) happens in instance of Check_IDS
                        
                    '''***********************************************************************'''
                    '''[Step2] Copy bincut table into a ids zone array.
                    '''***********************************************************************'''
                    If Zone_Num = 0 Then
                        VBIN_IDS_ZONE_Temp(p_mode).Ids_range(Zone_Num, Test_Type) = 0 '''Ids zone value starts from 0 based on current structure
                        Zone_Num = Zone_Num + 1
                    End If
                                
                    For idx_step = 1 To BinCut(p_mode, bincutNum).Mode_Step + 1
                        VBIN_IDS_ZONE_Temp(p_mode).Ids_range(Zone_Num, Test_Type) = BinCut(p_mode, bincutNum).IDS_CP_LIMIT(idx_step - 1)
                        Zone_Num = Zone_Num + 1
                    Next idx_step
                    VBIN_IDS_ZONE_Temp(p_mode).IDS_RANGE_COUNT(Test_Type) = Zone_Num
                End If '''If VDD_BIN(P_mode, BincutNum).ExcludedPmode = False Then
            Next bincutNum

            '''***********************************************************************'''
            '''[Step3] Screen out the duplicate ids limit from vbin_ids_zone().
            '''***********************************************************************'''
            If BinCut(p_mode, bincutNum).ExcludedPmode = False Then
                Zone_Num = 0
                Ids_zone_cnt = VBIN_IDS_ZONE_Temp(p_mode).IDS_RANGE_COUNT(Test_Type)
                
                For i = 0 To Ids_zone_cnt - 1
                    Duplicate_Flag = False
                    
                    For j = i + 1 To Ids_zone_cnt
                        If VBIN_IDS_ZONE_Temp(p_mode).Ids_range(i, Test_Type) = VBIN_IDS_ZONE_Temp(p_mode).Ids_range(j, Test_Type) Then
                            Duplicate_Flag = True
                            Exit For
                        End If
                    Next j
                    
                    If Duplicate_Flag = False Then
                        VBIN_IDS_ZONE(p_mode).Ids_range(Zone_Num, Test_Type) = VBIN_IDS_ZONE_Temp(p_mode).Ids_range(i, Test_Type)
                        Zone_Num = Zone_Num + 1
                    End If
                Next i
                
                VBIN_IDS_ZONE(p_mode).IDS_RANGE_COUNT(Test_Type) = Zone_Num - 1 '''Minus 1 is for fit in with the original vdd-binning code's counting
                
                '''***********************************************************************'''
                '''[Step5] Sorting the merged array in ascending way.
                '''***********************************************************************'''
                '''This is because IDS Zone is from low to high.
                Ids_zone_cnt = VBIN_IDS_ZONE(p_mode).IDS_RANGE_COUNT(Test_Type)
                For i = 0 To Ids_zone_cnt - 1
                    For j = i + 1 To Ids_zone_cnt
                        If (VBIN_IDS_ZONE(p_mode).Ids_range(i, Test_Type) > VBIN_IDS_ZONE(p_mode).Ids_range(j, Test_Type)) Then
                            DblTemp = VBIN_IDS_ZONE(p_mode).Ids_range(j, Test_Type)
                            VBIN_IDS_ZONE(p_mode).Ids_range(j, Test_Type) = VBIN_IDS_ZONE(p_mode).Ids_range(i, Test_Type)
                            VBIN_IDS_ZONE(p_mode).Ids_range(i, Test_Type) = DblTemp
                        End If
                    Next j
                Next i
                
                '''***********************************************************************'''
                '''[Step6] Copt TD to the remaining test type (MBIST, SPI, ...)
                '''***********************************************************************'''
                For Test_Type = 1 To testType.Total_Count - 1
                    VBIN_IDS_ZONE(p_mode).IDS_RANGE_COUNT(Test_Type) = VBIN_IDS_ZONE(p_mode).IDS_RANGE_COUNT(testType.TD)
                Next Test_Type
                
                For Test_Type = 1 To testType.Total_Count - 1
                    For RngNum = 0 To VBIN_IDS_ZONE(p_mode).IDS_RANGE_COUNT(Test_Type)
                        VBIN_IDS_ZONE(p_mode).Ids_range(RngNum, Test_Type) = VBIN_IDS_ZONE(p_mode).Ids_range(RngNum, testType.TD)
                    Next RngNum
                Next Test_Type
                                                
                '''***********************************************************************'''
                '''[Step7] Assign the start-search level for each ids zone.
                '''***********************************************************************'''
                '''Because start-search level is totally referred to bincut 1, we don't need to have bincut loop here.
                Test_Type = testType.TD '''//The default Testtype is TD
                Max_mode_step = BinCut(p_mode, 1).Mode_Step
                
                If BinCut(p_mode, 1).ExcludedPmode = False Then
                    For RngNum = 0 To VBIN_IDS_ZONE(p_mode).IDS_RANGE_COUNT(Test_Type)
                        If VBIN_IDS_ZONE(p_mode).Ids_range(RngNum, Test_Type) < BinCut(p_mode, 1).IDS_CP_LIMIT(Max_mode_step) Then
                            VBIN_IDS_ZONE(p_mode).IDS_Start_EQ_Num(RngNum, Test_Type) = Max_mode_step + 1
                        Else
                            VBIN_IDS_ZONE(p_mode).IDS_Start_EQ_Num(RngNum, Test_Type) = 0
                        End If
                    Next RngNum
                End If
                
                '''***********************************************************************'''
                '''[Step8] Copy TD to the remaining test type (MBIST, SPI, ...)
                '''***********************************************************************'''
                For Test_Type = 1 To testType.Total_Count - 1
                    For RngNum = 0 To VBIN_IDS_ZONE(p_mode).IDS_RANGE_COUNT(Test_Type)
                        VBIN_IDS_ZONE(p_mode).IDS_Start_EQ_Num(RngNum, Test_Type) = VBIN_IDS_ZONE(p_mode).IDS_Start_EQ_Num(RngNum, testType.TD)
                    Next RngNum
                Next Test_Type
            End If '''If VDD_BIN(P_mode, BincutNum).ExcludedPmode = False Then
        End If '''If IDS_Distribution_Table(P_mode).Used = True Then
    Next p_mode
Exit Function
errHandler:
    Call Print_Error_Message(Error_Info, "LIB_VDD_BINNING", "Generate_IDS_Zone_with_IDS_Distribution_Table") 'Add ErrHandler 2023/08/01
    If AbortTest Then Exit Function Else Resume Next
End Function

'20210223: Modified to replace "Dim step As Long" with "Dim idx_step As Long".
'20191204: Modified for the revised InitVddBinTable.
'20190314: Modified to check the following condition to avoid any C vlaue in Bin2 < Bin1 maxC in each ids zone.
'20181031: Modified to add TotalPmodeRangeCnt for Print_IDS_ZONE_Table_to_sheet.
'20160614: Modified by TSMC Jack.
Public Function Generate_IDS_ZONE_CONTENT()
On Error GoTo errHandler 'Add ErrHandler 2023/08/01
    Dim Test_Type As testType
    Dim idx_step As Long
    Dim Srch_Step As Long
    Dim i As Long
    Dim ids_range_step() As Long 'Dim ids_range_step(MaxPerformanceModeCount) As Long
    Dim IDS_current_Max() As Double 'Dim IDS_current_Max(MaxPerformanceModeCount) As Double
    Dim IdsRng As Long
    Dim Zone_Num As Long
    Dim Mode_Step As Long
    Dim p_mode As Integer
    Dim isPmodeUsed As Boolean
    Dim bincutNum As Variant
    Dim StepCnt As Long
    ReDim ids_range_step(MaxPerformanceModeCount) As Long
    ReDim IDS_current_Max(MaxPerformanceModeCount) As Double
    '''init
    Max_V_Step_per_IDS_Zone = 0
    
    '''//The default Testtype is TD
    Test_Type = testType.TD
            
    For i = 0 To MaxPerformanceModeCount - 1
        ids_range_step(i) = 0
        IDS_current_Max(i) = 0
    Next i
    
    For p_mode = 0 To MaxPerformanceModeCount - 1
        '''***********************************************************************'''
        '''[Step1] Skip those unused mode from mode select loop.
        '''***********************************************************************'''
        isPmodeUsed = False
        If Flag_IDS_Distribution_enable = True Then '''If IDS_Distribution_Table_table is available
            If IDS_Distribution_Table(p_mode).Used = True Or BinCut(p_mode, bincutNum).ExcludedPmode = False Then
                isPmodeUsed = True
            End If
        Else '''If IDS_Distribution_Table_table is not available
            If BinCut(p_mode, bincutNum).ExcludedPmode = False Then
                isPmodeUsed = True
            End If
        End If
        
        '''***********************************************************************'''
        '''[Step2]
        ''' 2.1: Copy C and M of Bincut to VBIN_IDS_Zone().
        ''' 2.2: Assign Level and bincutnum search sequence to VBIN_IDS_Zone().
        '''***********************************************************************'''
        If isPmodeUsed = True Then
            For Zone_Num = 0 To VBIN_IDS_ZONE(p_mode).IDS_RANGE_COUNT(Test_Type) - 1
               Srch_Step = VBIN_IDS_ZONE(p_mode).Max_Step(Zone_Num)         'This is a count-up step across multiple bincuts
               
               For Each bincutNum In PassBinCut_ary
                   Mode_Step = BinCut(p_mode, bincutNum).Mode_Step          'This is mode step number of different bincut table
                   StepCnt = 0                                              'Count up step of per ids zone
                   For idx_step = 0 To Mode_Step
                       If VBIN_IDS_ZONE(p_mode).Ids_range(Zone_Num, Test_Type) < BinCut(p_mode, bincutNum).IDS_CP_LIMIT(idx_step) Then
                           If bincutNum = 1 Then
                               VBIN_IDS_ZONE(p_mode).c(Zone_Num, Srch_Step) = BinCut(p_mode, bincutNum).c(Mode_Step - StepCnt)
                               VBIN_IDS_ZONE(p_mode).m(Zone_Num, Srch_Step) = BinCut(p_mode, bincutNum).m(Mode_Step - StepCnt)
                               VBIN_IDS_ZONE(p_mode).EQ_Num(Zone_Num, Srch_Step) = (Mode_Step + 1) - StepCnt
                               VBIN_IDS_ZONE(p_mode).passBinCut(Zone_Num, Srch_Step) = bincutNum
                               Srch_Step = Srch_Step + 1
                               StepCnt = StepCnt + 1
                           Else
                               '''20190314: Modified. The following condition is to avoid any C value in bin2 < bin1 maxC in each ids zone.
                               If BinCut(p_mode, bincutNum).CP_Vmax(Mode_Step - idx_step) >= BinCut(p_mode, bincutNum - 1).CP_Vmax(Mode_Step - idx_step) Then
                                   VBIN_IDS_ZONE(p_mode).c(Zone_Num, Srch_Step) = BinCut(p_mode, bincutNum).c(Mode_Step - StepCnt)
                                   VBIN_IDS_ZONE(p_mode).m(Zone_Num, Srch_Step) = BinCut(p_mode, bincutNum).m(Mode_Step - StepCnt)
                                   VBIN_IDS_ZONE(p_mode).EQ_Num(Zone_Num, Srch_Step) = (Mode_Step + 1) - StepCnt
                                   VBIN_IDS_ZONE(p_mode).passBinCut(Zone_Num, Srch_Step) = bincutNum
                                   Srch_Step = Srch_Step + 1
                              End If
                              
                              StepCnt = StepCnt + 1
                           End If 'If BincutNum = 1 Then
                       End If 'If VBIN_IDS_ZONE(P_mode).Ids_range(Zone_Num, Test_Type)
                   Next idx_step 'For Step = 0 To Mode_Step
               Next bincutNum
               
               VBIN_IDS_ZONE(p_mode).Max_Step(Zone_Num) = Srch_Step
               
               If VBIN_IDS_ZONE(p_mode).Max_Step(Zone_Num) > Max_V_Step_per_IDS_Zone Then
                   Max_V_Step_per_IDS_Zone = VBIN_IDS_ZONE(p_mode).Max_Step(Zone_Num)
               End If
            Next Zone_Num
        End If 'If IsPmodeUsed = True Then
    Next p_mode
    
    '''***********************************************************************'''
    '''[Step3]
    ''' 3.1: Determine VBIN_IDS_ZONE().IDS_Start_Bin.
    ''' 3.2: While all VBIN_IDS_ZONE().IDS_Start_Bin() are 0 if we only refer to bincut alone.
    '''***********************************************************************'''
    If Flag_IDS_Distribution_enable = True Then
        For p_mode = 0 To MaxPerformanceModeCount - 1
            If BinCut(p_mode, 1).ExcludedPmode = False Then
                Mode_Step = BinCut(p_mode, 1).Mode_Step
                
                For Test_Type = 0 To testType.Total_Count - 1
                    '''//Because all test type has same ids range, we select td as representative.
                    For Zone_Num = 0 To VBIN_IDS_ZONE(p_mode).IDS_RANGE_COUNT(Test_Type)
                        For IdsRng = 0 To IDS_Distribution_Table(p_mode).RANGE_COUNT
                            '''//Zone_Num always start from 0 and it is same as IDS_Distribution_Table table at its idsrng=0.
                            If Zone_Num = 0 Then
                                VBIN_IDS_ZONE(p_mode).IDS_Start_EQ_Num(Zone_Num, Test_Type) = IDS_Distribution_Table(p_mode).Start_Bin(0, Test_Type)
                                Exit For 'This zone has been given start_bin
                            ElseIf VBIN_IDS_ZONE(p_mode).Ids_range(Zone_Num, Test_Type) < IDS_Distribution_Table(p_mode).range(IdsRng, Test_Type) Then
                                VBIN_IDS_ZONE(p_mode).IDS_Start_EQ_Num(Zone_Num, Test_Type) = IDS_Distribution_Table(p_mode).Start_Bin(IdsRng - 1, Test_Type)
                                If VBIN_IDS_ZONE(p_mode).Ids_range(Zone_Num, Test_Type) >= BinCut(p_mode, 1).IDS_CP_LIMIT(Mode_Step) Then
                                    VBIN_IDS_ZONE(p_mode).IDS_Start_EQ_Num(Zone_Num, Test_Type) = 0
                                End If
                                Exit For 'This zone has been given start_bin
                            ElseIf Zone_Num > 0 And VBIN_IDS_ZONE(p_mode).Ids_range(Zone_Num, Test_Type) = IDS_Distribution_Table(p_mode).range(IdsRng, Test_Type) Then
                                VBIN_IDS_ZONE(p_mode).IDS_Start_EQ_Num(Zone_Num, Test_Type) = IDS_Distribution_Table(p_mode).Start_Bin(IdsRng + 1, Test_Type)
                                If VBIN_IDS_ZONE(p_mode).Ids_range(Zone_Num, Test_Type) < BinCut(p_mode, 1).IDS_CP_LIMIT(Mode_Step) Then
                                    VBIN_IDS_ZONE(p_mode).IDS_Start_EQ_Num(Zone_Num, Test_Type) = IDS_Distribution_Table(p_mode).Start_Bin(IdsRng, Test_Type)
                                Else
                                    VBIN_IDS_ZONE(p_mode).IDS_Start_EQ_Num(Zone_Num, Test_Type) = 0
                                End If
                                Exit For 'This zone has been given start_bin
                            End If
                        Next IdsRng
                    Next Zone_Num
                Next Test_Type
            End If '''If VDD_BIN(P_mode, 1).ExcludedPmode = False Then
        Next p_mode
    End If '''If Flag_IDS_Distribution_Table_enable = True Then

    '''***********************************************************************'''
    '''[Step4] Determine the start search step from VBIN_IDS_ZONE().IDS_Start_EQ_Num and VBIN_IDS_ZONE().EQ_Num.
    '''***********************************************************************'''
    If Flag_IDS_Distribution_enable = True Then
        For p_mode = 0 To MaxPerformanceModeCount - 1
            If BinCut(p_mode, 1).ExcludedPmode = False Then
                For Zone_Num = 0 To VBIN_IDS_ZONE(p_mode).IDS_RANGE_COUNT(0)
                    For idx_step = 0 To VBIN_IDS_ZONE(p_mode).Max_Step(Zone_Num) - 1
                        For Test_Type = 0 To testType.Total_Count - 1
                            If VBIN_IDS_ZONE(p_mode).IDS_Start_EQ_Num(Zone_Num, Test_Type) = 0 Then
                                VBIN_IDS_ZONE(p_mode).IDS_START_STEP(Zone_Num, Test_Type) = 0
                            Else
                                If VBIN_IDS_ZONE(p_mode).EQ_Num(Zone_Num, idx_step) = VBIN_IDS_ZONE(p_mode).IDS_Start_EQ_Num(Zone_Num, Test_Type) And VBIN_IDS_ZONE(p_mode).passBinCut(Zone_Num, idx_step) = 1 Then
                                    VBIN_IDS_ZONE(p_mode).IDS_START_STEP(Zone_Num, Test_Type) = idx_step
                                End If
                            End If
                        Next Test_Type
                    Next idx_step
                Next Zone_Num
            End If
        Next p_mode
    End If
    
    If Flag_Print_Out_tables_enable = True Then
        Print_IDS_ZONE_Table_to_sheet
    End If
Exit Function
errHandler:
    Call Print_Error_Message(Error_Info, "LIB_VDD_BINNING", "Generate_IDS_ZONE_CONTENT") 'Add ErrHandler 2023/08/01
    If AbortTest Then Exit Function Else Resume Next
End Function

Public Function GetInstrument_BV(PinList As String, site As Variant) As String
On Error GoTo errHandler 'Add ErrHandler 2023/08/01
    Dim strChannel As String
    Dim strAry_PinName() As String
    Dim NumberPins As Long
    Dim strAry_slot() As String
    Dim slot As Long
    Call TheExec.DataManager.DecomposePinList(PinList, strAry_PinName(), NumberPins)
    Call TheExec.DataManager.GetChannelStringFromPinAndSite(strAry_PinName(0), site, strChannel)
        
    If strChannel = "" Then
        MsgBox ("Please check pin type of " & PinList & " in channel map")
    Else
        strAry_slot = Split(strChannel, ".")
        slot = CLng(strAry_slot(0))
        GetInstrument_BV = TheHdw.config.Slots(slot).type
    End If
Exit Function
errHandler:
    Call Print_Error_Message(Error_Info, "LIB_VDD_BINNING", "GetInstrument_BV") 'Add ErrHandler 2023/08/01
    If AbortTest Then Exit Function Else Resume Next
End Function

'20210810: Modified to add the property "step_Lowest As New SiteLong" to Public Type DYNAMIC_VBIN_IDS_ZONE.
'20210722: Modified to use VBIN_IDS_ZONE(p_mode).Product_Voltage(Max_IDS_Zone, Max_IDS_Step) and DYNAMIC_VBIN_IDS_ZONE(p_mode).Product_Voltage(Max_IDS_Step) for GradeVDD.
'20210407: Modified to add "interpolated as new SiteBoolean" and "step_Interpolated_Start as new SiteLong" for "Public Type DYNAMIC_VBIN_IDS_ZONE".
'20210223: Modified to replace "Dim step As Long" with "Dim idx_step As Long".
'20210223: Modified to map DYNAMIC_VBIN_IDS_ZONE(p_mode).step_Mapping(BinNum, EQN) to step in DYNAMIC_IDS_Zone.
'20201113: Modified to use "Last_Bin1_Step" to store last Bin1 Step (ex: EQN1) in Dynamic_IDS_ZONE for Interpolation.
'20191127: Modified for the revised InitVddBinTable.
Public Function Generate_DYNAMIC_IDS_ZONE_Voltage_Per_Site(p_mode As Integer)
On Error GoTo errHandler 'Add ErrHandler 2023/08/01
    Dim site As Variant
    Dim Test_Type As testType
    Dim idx_step As Long
    Dim Zone_Num As Integer
'''//==================================================================================================================================================================================//'''
'''//Note:
'''Calculate the voltage of each step and each zone from IDS_Zone for p_mode.
'''//==================================================================================================================================================================================//'''
    '''init
    Test_Type = testType.TD

    If VBIN_IDS_ZONE(p_mode).Used = True Then
        For Each site In TheExec.sites
            DYNAMIC_VBIN_IDS_ZONE(p_mode).Used = True
            DYNAMIC_VBIN_IDS_ZONE(p_mode).IDS_ZONE_NUMBER = VBIN_IDS_ZONE(p_mode).IDS_ZONE_NUMBER
            Zone_Num = DYNAMIC_VBIN_IDS_ZONE(p_mode).IDS_ZONE_NUMBER
            
            '''//Initialize the lowest step in Dynamic_IDS_Zone as 0.
            DYNAMIC_VBIN_IDS_ZONE(p_mode).step_inherit(site) = 0
            
            '''//Use "Last_Bin1_Step" to store last Bin1 Step (ex: EQN1) in Dynamic_IDS_ZONE. Initialize it as -1.
            '''For interpolation.
            DYNAMIC_VBIN_IDS_ZONE(p_mode).interpolated = False
            DYNAMIC_VBIN_IDS_ZONE(p_mode).step_Interpolated_Start = -1
            
            For idx_step = 0 To VBIN_IDS_ZONE(p_mode).Max_Step(Zone_Num) - 1    'loop the step in the IDS Zone
                DYNAMIC_VBIN_IDS_ZONE(p_mode).c(idx_step) = VBIN_IDS_ZONE(p_mode).c(Zone_Num, idx_step)
                DYNAMIC_VBIN_IDS_ZONE(p_mode).m(idx_step) = VBIN_IDS_ZONE(p_mode).m(Zone_Num, idx_step)
                DYNAMIC_VBIN_IDS_ZONE(p_mode).passBinCut(idx_step) = VBIN_IDS_ZONE(p_mode).passBinCut(Zone_Num, idx_step)
                DYNAMIC_VBIN_IDS_ZONE(p_mode).EQ_Num(idx_step) = VBIN_IDS_ZONE(p_mode).EQ_Num(Zone_Num, idx_step)
                DYNAMIC_VBIN_IDS_ZONE(p_mode).Max_Step = VBIN_IDS_ZONE(p_mode).Max_Step(Zone_Num)
                DYNAMIC_VBIN_IDS_ZONE(p_mode).Voltage(idx_step) = VBIN_IDS_ZONE(p_mode).Voltage(Zone_Num, idx_step)
                DYNAMIC_VBIN_IDS_ZONE(p_mode).Product_Voltage(idx_step) = VBIN_IDS_ZONE(p_mode).Product_Voltage(Zone_Num, idx_step)
                DYNAMIC_VBIN_IDS_ZONE(p_mode).step_Mapping(DYNAMIC_VBIN_IDS_ZONE(p_mode).passBinCut(idx_step), DYNAMIC_VBIN_IDS_ZONE(p_mode).EQ_Num(idx_step)) = idx_step
               
                For Test_Type = 0 To testType.Total_Count - 1
                    DYNAMIC_VBIN_IDS_ZONE(p_mode).IDS_Start_EQ_Num(Test_Type) = VBIN_IDS_ZONE(p_mode).IDS_Start_EQ_Num(Zone_Num, Test_Type)
                    DYNAMIC_VBIN_IDS_ZONE(p_mode).IDS_START_STEP(Test_Type) = VBIN_IDS_ZONE(p_mode).IDS_START_STEP(Zone_Num, Test_Type)
                Next Test_Type
            Next idx_step
        Next site
    Else
        Call Print_Error_Message(Error_Info, "LIB_VDD_BINNING", "Generate_DYNAMIC_IDS_ZONE_Voltage_Per_Site", VddBinName(p_mode) & ", it doesn't have any correct IDS_Zone for Generate_DYNAMIC_IDS_ZONE_Voltage_Per_Site. Error!!!")
        'TheExec.Datalog.WriteComment VddBinName(p_mode) & ", it doesn't have any correct IDS_Zone for Generate_DYNAMIC_IDS_ZONE_Voltage_Per_Site. Error!!!"
    End If
Exit Function
errHandler:
    Call Print_Error_Message(Error_Info, "LIB_VDD_BINNING", "Generate_DYNAMIC_IDS_ZONE_Voltage_Per_Site") 'Add ErrHandler 2023/08/01
    If AbortTest Then Exit Function Else Resume Next
End Function

'20210812: Modified to rename the property "step_lowest As New SiteLong" as "step_inherit As New SiteLong".
'20210810: Modified to add the property "step_Lowest As New SiteLong" to Public Type DYNAMIC_VBIN_IDS_ZONE.
'20210728: Modified to remove the redundant vbt function Clear_Dynamic_IDS_ZONE_by_Site.
'20210722: Modified to use VBIN_IDS_ZONE(p_mode).Product_Voltage(Max_IDS_Zone, Max_IDS_Step) and DYNAMIC_VBIN_IDS_ZONE(p_mode).Product_Voltage(Max_IDS_Step) for GradeVDD.
'20210407: Modified to add "interpolated as new SiteBoolean" and "step_Interpolated_Start as new SiteLong" for "Public Type DYNAMIC_VBIN_IDS_ZONE".
'20210223: Modified to replace "Dim step As Long" with "Dim idx_step As Long".
'20210223: Modified to map DYNAMIC_VBIN_IDS_ZONE(p_mode).step_Mapping(BinNum, EQN) to step in DYNAMIC_IDS_Zone.
'20191127: Modified for the revised InitVddBinTable.
Public Function init_Dynamic_IDS_ZONE()
On Error GoTo errHandler 'Add ErrHandler 2023/08/01
    Dim Test_Type As Long
    Dim idx_step As Long
    Dim p_mode As Integer
    Dim idx_binNum As Long
    Dim idx_EQN As Long
    For p_mode = 0 To MaxPerformanceModeCount - 1
        DYNAMIC_VBIN_IDS_ZONE(p_mode).Used = False
        DYNAMIC_VBIN_IDS_ZONE(p_mode).IDS_ZONE_NUMBER = 99
        DYNAMIC_VBIN_IDS_ZONE(p_mode).step_inherit = 0
        '''For interpolation.
        DYNAMIC_VBIN_IDS_ZONE(p_mode).interpolated = False
        DYNAMIC_VBIN_IDS_ZONE(p_mode).step_Interpolated_Start = -1
        
        For idx_step = 0 To Max_IDS_Step
            DYNAMIC_VBIN_IDS_ZONE(p_mode).c(idx_step) = 0
            DYNAMIC_VBIN_IDS_ZONE(p_mode).m(idx_step) = 0
            DYNAMIC_VBIN_IDS_ZONE(p_mode).passBinCut(idx_step) = 0
            DYNAMIC_VBIN_IDS_ZONE(p_mode).EQ_Num(idx_step) = 0
            DYNAMIC_VBIN_IDS_ZONE(p_mode).Max_Step = 0
            DYNAMIC_VBIN_IDS_ZONE(p_mode).Voltage(idx_step) = 0
            DYNAMIC_VBIN_IDS_ZONE(p_mode).Product_Voltage(idx_step) = 0
        Next idx_step
        
        '''//Initialize array of mapping (binNum, EQN) to step in Dynamic_IDS_Zone.
        For idx_binNum = 0 To MaxPassBinCut
            For idx_EQN = 0 To Max_IDS_Step + 1
                DYNAMIC_VBIN_IDS_ZONE(p_mode).step_Mapping(idx_binNum, idx_EQN) = -1
            Next idx_EQN
        Next idx_binNum
        
        For Test_Type = 0 To testType.Total_Count - 1
             DYNAMIC_VBIN_IDS_ZONE(p_mode).IDS_Start_EQ_Num(Test_Type) = 0
             DYNAMIC_VBIN_IDS_ZONE(p_mode).IDS_START_STEP(Test_Type) = 0
        Next Test_Type
    Next p_mode
Exit Function
errHandler:
    Call Print_Error_Message(Error_Info, "LIB_VDD_BINNING", "init_Dynamic_IDS_ZONE") 'Add ErrHandler 2023/08/01
    If AbortTest Then Exit Function Else Resume Next
End Function

'20210831: Modified to remove the vbt code related to CPVmax and CPVmin.
'20210813: Modified to use Set_VBinResult_by_Step for updating PassBin, Pass step, and voltage to VBIN_Result.
'20210813: Modified to print interpolation info while SkipTest=True for BinX and BinY.
'20210812: Modified to rename the property "step_lowest As New SiteLong" as "step_inherit As New SiteLong".
'20210726: Modified to add the argument "bincutNum As Long" to the vbt function Adjust_Multi_PassBinCut_Per_Site.
'20210722: Modified to use VBIN_IDS_ZONE(p_mode).Product_Voltage(Max_IDS_Zone, Max_IDS_Step) and DYNAMIC_VBIN_IDS_ZONE(p_mode).Product_Voltage(Max_IDS_Step) for GradeVDD.
'20210618: Modified to update SortNumber and binNumber if F_Vddbinning_Interpolation_fail is triggered in the vbt function ReGenerate_DYNAMIC_IDS_ZONE_Voltage_Per_Site.
'20210412: Modified to remove the vbt code not to check if the voltage_Calc is in the range between CPVmin and CPVmax, requested by PCLINZG.
'20210409: Modified to separate voltage calculation for interpolation and check steps.
'20210408: Modified to update step for BinX/Y if AllBinCut(p_mode).INTP_SKIPTEST = True.
'20210408: Modified to overwrite step_inherit and VBIN_RESULT(inst_info.p_mode).step_in_IDS_Zone if DYNAMIC_VBIN_IDS_ZONE(inst_info.p_mode).interpolated = True.
'20210407: Modified to revise the vbt code for the new Interpolation method proposed by C651 Toby.
'20210407: Modified to add "interpolated as new SiteBoolean" and "step_Interpolated_Start as new SiteLong" for "Public Type DYNAMIC_VBIN_IDS_ZONE".
'20210322: Modified to initialize Last_Bin1_Step = -1 for Interpolation.
'20210305: Modified to check if enableCalcInterpolation = True.
'20210223: Modified to replace "Dim step As Long" with "Dim idx_step As Long".
'20210223: Modified to map DYNAMIC_VBIN_IDS_ZONE(p_mode).step_Mapping(BinNum, EQN) to step in DYNAMIC_IDS_Zone.
'20210120: Modified to use VBIN_RESULT(p_mode).step_1stPass_in_IDS_Zone to store the first pass step in Dynamic IDS Zone and find the correspondent PassBinCut number.
'20201229: Modified to check if start_p_mode and end_p_mode were tested or not, requested by Verity.
'20201218: Modified to update "FIRSTPASSBINCUT(p_mode) = VBIN_RESULT(p_mode).passBinCut", requested by Leon Weng.
'20201113: C651 Toby requested the new method to check BinCut voltages montonicity, so that we need to keep all available steps in Dynamic_IDS_Zone.
'20201113: Modified to use "Last_Bin1_Step" to store last Bin1 Step (ex: EQN1) in Dynamic_IDS_ZONE for Interpolation.
'20200502: Modified to replace "VBIN_IDS_ZONE" with "DYNAMIC_VBIN_IDS_ZONE".
'20200501: Modified to replace "BinCut(p_mode, PassBinNum(site)).INTP_MFACTOR(0)" with "AllBinCut(p_mode).INTP_SKIPTEST".
'20200425: Modified to change the output format of the interpolated voltage string.
'20200423: Modified to replace "BinCut(p_mode, bincutNum(site)).tested = True" with "VBIN_RESULT(p_mode).tested=True".
'20200414: Modified the branches of Intrpolation for Bin1 and non-Bin1.
'20200410: Modified to control "Exit Function" by the siteFlag "enableCalcInterpolation".
'20200406: Modified to revise the format of interpolation output strings.
'20200401: Modified to use IDS to calculate BinCut payload voltage for BinX and BinY.
'20200330: Modified to use "Int_Offset","Int_SkipTest" for interpolation.
'20200319: Modified to find start p_mode, end p_mode, and interpolation factor by site.
'20191127: Modified for the revised InitVddBinTable.
'20190507: Modified to add "Cdec" to avoid double format accuracy issues.
'20190226: Modified to skip interpolation for start_p_mode and end_p_mode if INT_MF=0.
'20181107: Modified by Oscar. Based on Customer's new Vdd_Binning_Def tables (Add columns of Interpolation factor and Info).
'20181029: Modified for Cebu interpolation, by Oscar.
'20171222: SWLINZA modified CPVmax to use E1.
Public Function ReGenerate_DYNAMIC_IDS_ZONE_Voltage_Per_Site(p_mode As Integer, PassBinNum As SiteLong)
On Error GoTo errHandler 'Add ErrHandler 2023/08/01
    Dim site As Variant
    Dim Test_Type As testType
    Dim start_p_mode As Integer '''start p_mode of interpolation
    Dim end_p_mode As Integer   '''end p_mode of interpolation
    Dim enableCalcInterpolation As New SiteBoolean
    '''
    Dim Remainder As Double
    Dim voltage_INTP_L As Double
    Dim voltage_INTP_H As Double
    Dim voltage_Calc As Double
    Dim dbl_Interpolation_MF As Double
    Dim dbl_Interpolation_Offset As Double
    '''
    Dim gotCorretPmode As Boolean
    Dim step_Calc As Long
    Dim idx_step As Long
    Dim EQ_Num As Long
'''//==================================================================================================================================================================================//'''
'''//Note:
'''1. C651 Si added the column "INT_SKIPTEST" of Interpolation in table "Vdd_Binning_Def_appA_1" only, so that we only check this from Bin1 table.
'''2. Only Bin1 DUT uses Interpolation, and Bin1 has the Interpolation factor.
'''3. Warning!!!please check SortNumber and binNumber of Flag_Vddbinning_Interpolation_fail in Bin_Table before using this.
'''//==================================================================================================================================================================================//'''
    For Each site In TheExec.sites
        '''init the flag
        enableCalcInterpolation = False
        DYNAMIC_VBIN_IDS_ZONE(p_mode).interpolated(site) = False
        step_Calc = -1
        gotCorretPmode = False
        
        '''//Check if p_mode is tested.
        If VBIN_RESULT(p_mode).tested = True Then
            gotCorretPmode = False
            '''ToDo: Maybe we can remove the error message...
            Call Print_Error_Message(Error_Info, "LIB_VDD_BINNING", "ReGenerate_DYNAMIC_IDS_ZONE_Voltage_Per_Site", "site:" & site & "," & VddBinName(p_mode) & " is tested. It should not be interpolated in Interpolatetion instance again. Error!!!")
            'TheExec.ErrorLogMessage "site:" & site & "," & VddBinName(p_mode) & " is tested. It should not be interpolated in Interpolatetion instance again. Error!!!"    ''20230807: Removed by sanity check
        Else
            gotCorretPmode = True
        End If
        
        If gotCorretPmode = True Then
            '''//Align PassBinCutNum.
            '''20210408: Modified to update step for BinX/Y if AllBinCut(p_mode).INTP_SKIPTEST = True.
            If VBIN_RESULT(p_mode).passBinCut <> PassBinNum(site) Then
                '''20210726: Modified to add the argument "bincutNum As Long" to the vbt function Adjust_Multi_PassBinCut_Per_Site.
                Adjust_Multi_PassBinCut_Per_Site p_mode, site, PassBinNum(site)
            End If
            
            '''======================================================================================
            '''Only Bin1 DUT uses Interpolation, and Bin1 has the Interpolation factor.
            '''//If the Interpolation factor of p_mode is 0, it will skip interpolation calculation.
            '''======================================================================================
            If PassBinNum(site) = 1 And BinCut(p_mode, PassBinNum(site)).INTP_MFACTOR(0) <> 0 Then '''Bin1
                '''//All Equation in the interpolation item has the start P_mode, end P_mode and factor info. so step can be 0 ~ Bin1 maxstep.
                enableCalcInterpolation = True
                start_p_mode = BinCut(p_mode, PassBinNum(site)).INTP_MODE_L(0) 'only step0
                end_p_mode = BinCut(p_mode, PassBinNum(site)).INTP_MODE_H(0) 'only step0
                dbl_Interpolation_MF = BinCut(p_mode, PassBinNum(site)).INTP_MFACTOR(0)
                dbl_Interpolation_Offset = BinCut(p_mode, PassBinNum(site)).INTP_OFFSET(0)
                
                '''//start_p_mode and end_p_mode should be tested prior to p_mode interpolated.
                If enableCalcInterpolation = True And (VBIN_RESULT(start_p_mode).tested = False Or VBIN_RESULT(end_p_mode).tested = False) Then
                    enableCalcInterpolation = False
                    Call Print_Error_Message(Error_Info, "LIB_VDD_BINNING", "ReGenerate_DYNAMIC_IDS_ZONE_Voltage_Per_Site", "site:" & site & "," & VddBinName(start_p_mode) & " or " & VddBinName(end_p_mode) & " wasn't tested or failed. It should not be interpolated. Error!!!")
                    'TheExec.ErrorLogMessage "site:" & site & "," & VddBinName(start_p_mode) & " or " & VddBinName(end_p_mode) & " wasn't tested or failed. It should not be interpolated. Error!!!"    ''20230807: Removed by sanity check
                    '''Warning!!!please check SortNumber and binNumber of Flag_Vddbinning_Interpolation_fail in Bin_Table before using this.
                    '''20210618: Modified to update SortNumber and binNumber if F_Vddbinning_Interpolation_fail is triggered in the vbt function ReGenerate_DYNAMIC_IDS_ZONE_Voltage_Per_Site.
                    TheExec.sites.item(site).FlagState(strGlb_Flag_Vddbinning_Interpolation_fail) = logicTrue
                End If
            Else '''non-Bin1
                enableCalcInterpolation = False
                step_Calc = VBIN_RESULT(p_mode).step_in_IDS_Zone
            End If
        End If
        
        '''//Only Bin1 DUT can calulate voltages for Interpolation.
        If enableCalcInterpolation = True Then
            '''************************************************************************************************************'''
            '''//Check if start P_mode and end P_mode of Interpolation exist (and are used for BinCut).
            '''(Errors occur when one of the interpolated performance mode and its start/end performance mode are not used)
            '''************************************************************************************************************'''
            If (DYNAMIC_VBIN_IDS_ZONE(p_mode).Used = True And DYNAMIC_VBIN_IDS_ZONE(start_p_mode).Used = True And DYNAMIC_VBIN_IDS_ZONE(end_p_mode).Used = True) Then
                '''======================================================================================
                '''[Step1] Vi = Vlow + (Vhigh - Vlow) * MF
                '''======================================================================================
                '''//Calcute voltage by Interpolation from start_p_mode and end_p_mode.
                voltage_INTP_L = VBIN_RESULT(start_p_mode).GRADE(site)
                voltage_INTP_H = VBIN_RESULT(end_p_mode).GRADE(site)
                voltage_Calc = voltage_INTP_L + (voltage_INTP_H - voltage_INTP_L) * dbl_Interpolation_MF
                
                '''======================================================================================
                '''[Step2] Vx = Vi + Int_Offset
                '''======================================================================================
                If dbl_Interpolation_Offset <> 0 Then
                    voltage_Calc = voltage_Calc + dbl_Interpolation_Offset
                End If
                
                '''//Ceiling voltage_Calc by BV_StepVoltage defined in BinCut voltage table(sheet "Vdd_Binning_Def").
                '''20210412: Modified to remove the vbt code not to check if the voltage_Calc is in the range between CPVmin and CPVmax, requested by PCLINZG.
                '''20210831: Modified to remove the vbt code related to CPVmax and CPVmin.
                Remainder = Ceiling(voltage_Calc / BV_StepVoltage)
                voltage_Calc = Remainder * BV_StepVoltage
                
                '''======================================================================================
                '''[Step3] Round Vx up to next bincut equation: Vx_rounded = RoundUpToEQN(Vx).
                ''' 1. Find the new_En_step in IDS Zone(most close to and higher than Vx EBB).
                ''' 2. Find the last_bin1_step in IDS Zone(last step of Bin1 in IDS ZOne).
                '''======================================================================================
                '''//Find the nearest step for interpolated voltage (step_interpolated).
                For idx_step = 0 To DYNAMIC_VBIN_IDS_ZONE(p_mode).Max_Step - 1
                    If CDec(DYNAMIC_VBIN_IDS_ZONE(p_mode).Voltage(idx_step)) >= CDec(voltage_Calc) And step_Calc = -1 Then
                        step_Calc = idx_step
                        Exit For
                    End If
                Next idx_step
            Else
                Call Print_Error_Message(Error_Info, "LIB_VDD_BINNING", "ReGenerate_DYNAMIC_IDS_ZONE_Voltage_Per_Site", "site:" & site & "," & VddBinName(p_mode) & ", Dynamic_IDS ZONE doesn't exist. it can't do Interpolation. Error!!!")
                'TheExec.ErrorLogMessage "site:" & site & "," & VddBinName(p_mode) & ", Dynamic_IDS ZONE doesn't exist. it can't do Interpolation. Error!!!"    ''20230807: Removed by sanity check
            End If
        End If
        
        '''//Check if step_Calc is valid in Dynamic_IDS_Zone.
        If step_Calc <> -1 Then
            DYNAMIC_VBIN_IDS_ZONE(p_mode).interpolated(site) = True
            
            If TheExec.sites.item(site).FlagState("BIN4_CAND") = logicTrue Then
                step_Calc = DYNAMIC_VBIN_IDS_ZONE(p_mode).step_Mapping(PassBinNum(site), 1)
                TheExec.Datalog.WriteComment "Site:" & site & ", BIN4_CAND Flage is True, Set selected Eqn to Bin1 EQ1"
            Else
                '''==================================================================================================================================================================
                '''//For SkipTest = True  : If no step in Bin1/X/Y available, do not use step Bin1 EQN1, just bin out the failed DUT with failFlag "F_Vddbinning_Interpolation_fail".
                '''//For SkipTest = False : If step_Calc is greater than step(PassBinNum(site),EQN1), set step(PassBinNum(site),EQN1) as 1st step of Interpolation.
                '''==================================================================================================================================================================
                If AllBinCut(p_mode).INTP_SKIPTEST = True Then
                    '''//Maybe step_Calc is in the higher PassBin, it should update VBIN_RESULT(p_mode).step_in_IDS_Zone and CurrentPassBinCutNum(site).
                    If CurrentPassBinCutNum(site) <> DYNAMIC_VBIN_IDS_ZONE(p_mode).passBinCut(step_Calc) Then
                        CurrentPassBinCutNum(site) = DYNAMIC_VBIN_IDS_ZONE(p_mode).passBinCut(step_Calc)
                        '''20210726: Modified to add the argument "bincutNum As Long" to the vbt function Adjust_Multi_PassBinCut_Per_Site.
                        Adjust_Multi_PassBinCut_Per_Site p_mode, site, CurrentPassBinCutNum(site)
                    End If
                Else '''If AllBinCut(p_mode).INTP_SKIPTEST = False
                    '''//If step_Calc is greater than step(PassBinNum(site),EQN1), set step(PassBinNum(site),EQN1) as 1st step of Interpolation.
                    If step_Calc > DYNAMIC_VBIN_IDS_ZONE(p_mode).step_Mapping(PassBinNum(site), 1) Then
                        step_Calc = DYNAMIC_VBIN_IDS_ZONE(p_mode).step_Mapping(PassBinNum(site), 1)
                    End If
                End If
            End If
        Else '''//If step_Calc = -1, it means no matched step for Interpolation, it only overwrite step with step_mapping(Bin1,EQN1) for Bin1 DUT; otherwise, bin out the failed DUT.
            If PassBinNum(site) = 1 Then
                If AllBinCut(p_mode).INTP_SKIPTEST = True Then
                    DYNAMIC_VBIN_IDS_ZONE(p_mode).interpolated(site) = False
                    TheExec.sites.item(site).FlagState(strGlb_Flag_Vddbinning_Interpolation_fail) = logicTrue
                    TheExec.Datalog.WriteComment "site:" & site & ", " & "CurrentPassBinCutNum:" & CurrentPassBinCutNum(site) & ", Binning Mode:" & VddBinName(p_mode) & ", no step available in Bin1/X/Y for " & VddBinName(start_p_mode) & ". Bin out the failed DUT."
                    TheExec.sites.item(site).FlagState(strGlb_Flag_Vddbinning_Fail_Stop) = logicTrue
                Else '''If AllBinCut(p_mode).INTP_SKIPTEST = False
                    DYNAMIC_VBIN_IDS_ZONE(p_mode).interpolated(site) = True
                    step_Calc = DYNAMIC_VBIN_IDS_ZONE(p_mode).step_Mapping(PassBinNum(site), 1)
                End If
            Else
                DYNAMIC_VBIN_IDS_ZONE(p_mode).interpolated(site) = False
                TheExec.sites.item(site).FlagState(strGlb_Flag_Vddbinning_Interpolation_fail) = logicTrue
                TheExec.Datalog.WriteComment "site:" & site & ", " & "CurrentPassBinCutNum:" & CurrentPassBinCutNum(site) & ", Binning Mode:" & VddBinName(p_mode) & ", no step available in Bin1/X/Y for " & VddBinName(start_p_mode) & ". Bin out the failed DUT."
                TheExec.sites.item(site).FlagState(strGlb_Flag_Vddbinning_Fail_Stop) = logicTrue
            End If
        End If
        
        '''//Update step_Calc to VBIN_RESULT(p_mode).step_in_IDS_Zone.
        If DYNAMIC_VBIN_IDS_ZONE(p_mode).interpolated(site) = True Then '''If Interpolation is completed, update VBIN_RESULT(p_mode) and print out the related info.
            ''20240304 BIN1 BINX merge if Vddbin_PTE_Debug, BINX force to BIN1 E1
            If TheExec.enableWord("Vddbin_PTE_Debug") = True And Flag_BinX_Info_Parsed = True Then
                VBIN_RESULT(p_mode).step_in_BinCut = DYNAMIC_VBIN_IDS_ZONE(p_mode).EQ_Num(step_Calc) - 1
                If Check_EQN_UpToBinX(p_mode, site) Then
                    step_Calc = DYNAMIC_VBIN_IDS_ZONE(p_mode).step_Mapping(1, AllBinCut(p_mode).Last_Bin1_EQN)
                End If
            End If
            DYNAMIC_VBIN_IDS_ZONE(p_mode).step_Interpolated_Start = step_Calc
            VBIN_RESULT(p_mode).step_in_IDS_Zone = step_Calc
            '''20210812: Modified to rename the property "step_lowest As New SiteLong" as "step_inherit As New SiteLong".
            DYNAMIC_VBIN_IDS_ZONE(p_mode).step_inherit = 0
            
            '''==================================================================================================================================================================
            '''//Default: test_type = testType.TD
            '''//DYNAMIC_VBIN_IDS_ZONE(p_mode).IDS_START_STEP(test_type) will be used as step_inherit for the vbt function "find_start_voltage".
            '''Remember to check the branches in the vbt function "find_start_voltage"!!!
            '''==================================================================================================================================================================
            '''20210408: Modified to overwrite step_inherit and VBIN_RESULT(inst_info.p_mode).step_in_IDS_Zone if DYNAMIC_VBIN_IDS_ZONE(inst_info.p_mode).interpolated = True.
            For Test_Type = 0 To testType.Total_Count - 1
                DYNAMIC_VBIN_IDS_ZONE(p_mode).IDS_START_STEP(Test_Type) = step_Calc
            Next Test_Type
            
            '''//Print info about result of the interpolation.
            '''ex: Site:1,CurrentPassBinCutNum:1,Binning Mode:VDD_GPU_MG003,The lowest Performance Mode:MG002,The highest Performance Mode:MG007,The MFx:0.19,SkipTest:True,The interpolated Voltage:565.625,The selected Eqn:5,The selected Voltage:565.625
            If CurrentPassBinCutNum(site) = 1 Then
                TheExec.Datalog.WriteComment "Site:" & site & ", " & _
                                            "CurrentPassBinCutNum:" & CurrentPassBinCutNum(site) & ", " & _
                                            "Binning Mode:" & VddBinName(p_mode) & ", " & _
                                            "The lowest Performance Mode:" & VddBinName(CInt(start_p_mode)) & ", " & _
                                            "The highest Performance Mode:" & VddBinName(CInt(end_p_mode)) & ", " & _
                                            "The MFx:" & dbl_Interpolation_MF & ", " & _
                                            "SkipTest:" & AllBinCut(p_mode).INTP_SKIPTEST & ", " & _
                                            "The interpolated Voltage:" & voltage_Calc & " mV" & ", " & _
                                            "The selected Eqn:" & DYNAMIC_VBIN_IDS_ZONE(p_mode).EQ_Num(step_Calc) & ", " & _
                                            "The selected Voltage:" & DYNAMIC_VBIN_IDS_ZONE(p_mode).Voltage(step_Calc) & " mV"
            ElseIf AllBinCut(p_mode).INTP_SKIPTEST = True Then
                '''20210813: Modified to print interpolation info while SkipTest=True for BinX and BinY.
                TheExec.Datalog.WriteComment "site:" & site & ", " & _
                                            "CurrentPassBinCutNum:" & CurrentPassBinCutNum(site) & ", " & _
                                            "Binning Mode:" & VddBinName(p_mode) & ", " & _
                                            "The lowest Performance Mode:" & VddBinName(CInt(start_p_mode)) & ", " & _
                                            "The highest Performance Mode:" & VddBinName(CInt(end_p_mode)) & ", " & _
                                            "The MFx:" & dbl_Interpolation_MF & ", " & _
                                            "SkipTest:" & AllBinCut(p_mode).INTP_SKIPTEST & ", " & _
                                            "The interpolated Voltage:" & DYNAMIC_VBIN_IDS_ZONE(p_mode).Voltage(step_Calc) & " mV" & ", " & _
                                            "The selected Eqn:" & DYNAMIC_VBIN_IDS_ZONE(p_mode).EQ_Num(step_Calc) & ", " & _
                                            "The selected Voltage:" & DYNAMIC_VBIN_IDS_ZONE(p_mode).Voltage(step_Calc) & " mV"
            End If
            
            '''************************************************************************************************************************************************************'''
            '''//If Interpolation SkipTest="Yes" (AllBinCut(p_mode).INTP_SKIPTEST = True) ==> Just put step0 of "DYNAMIC_VBIN_IDS_ZONE(p_mode)" into "VBIN_RESULT(p_mode)".
            '''************************************************************************************************************************************************************'''
            If AllBinCut(p_mode).INTP_SKIPTEST = True Then
                '''//Update PassBin, Pass step, flag"VBIN_Result(p_mode).tested", and voltage to VBIN_Result by the step in Dynamic_IDS_Zone.
                '''20210813: Modified to use Set_VBinResult_by_Step for updating PassBin, Pass step, and voltage to VBIN_Result.
                Call Set_VBinResult_by_Step(site, p_mode, step_Calc)
                
                VBIN_RESULT(p_mode).FLAGFAIL = False
            End If
        End If '''DYNAMIC_VBIN_IDS_ZONE(p_mode).interpolated = True
    Next site
Exit Function
errHandler:
    Call Print_Error_Message(Error_Info, "LIB_VDD_BINNING", "ReGenerate_DYNAMIC_IDS_ZONE_Voltage_Per_Site") 'Add ErrHandler 2023/08/01
    If AbortTest Then Exit Function Else Resume Next
End Function

'20200211: Modified to replace "cntFlowTestCond" with "cntAdditionalMode".
'20200211: Modified to replace the function name "FlowTestCondStr2Enum" with "AdditionalModeStr2Enum".
'20191202: Modified for the revised initVddBinCondition.
Public Function AdditionalModeStr2Enum(additional_mode As String) As Integer
On Error GoTo errHandler
    additional_mode = UCase(additional_mode)
    
    If AdditionalModeDict.Exists(additional_mode) Then
        AdditionalModeStr2Enum = AdditionalModeDict.item(additional_mode)
    Else
        AdditionalModeStr2Enum = cntAdditionalMode + 1
        Call Print_Error_Message(Error_Info, "LIB_VDD_BINNING", "AdditionalModeStr2Enum", "Enum:" & AdditionalModeStr2Enum & ", TestCond = " & additional_mode & " doesn't exist in enum FlowSheetCondition. Error!!!")
        'TheExec.ErrorLogMessage "Enum:" & AdditionalModeStr2Enum & ", TestCond = " & additional_mode & " doesn't exist in enum FlowSheetCondition. Error!!!"    ''20230807: Removed by sanity check
    End If
Exit Function
errHandler:
    Call Print_Error_Message(Error_Info, "LIB_VDD_BINNING", "AdditionalModeStr2Enum") 'Add ErrHandler 2023/08/01
    If AbortTest Then Exit Function Else Resume Next
End Function

'20190502: Modified for OFFSET_FUNC, ex: "MG005_GFXTD_BPL_BV".
Public Function decide_offset_testType(strInput As String) As Integer
On Error GoTo errHandler 'Add ErrHandler 2023/08/01
    Dim strTemp As String
    strTemp = LCase(strInput)
    
    If strTemp Like "*_td*" Then
        decide_offset_testType = testType.TD
    ElseIf strTemp Like "*_bist*" Then
        decide_offset_testType = testType.Mbist
    ElseIf strTemp Like "*_func*" Then '''added for dynamic offset_Func, 20190502
        decide_offset_testType = testType.Func
    Else
        decide_offset_testType = testType.ldcbfd '''It's a pseudo testType. If we read "Offset_CP1_FUNC" from sheet, set it as test type "LDCBFD".
    End If
Exit Function
errHandler:
    Call Print_Error_Message(Error_Info, "LIB_VDD_BINNING", "decide_offset_testType") 'Add ErrHandler 2023/08/01
    If AbortTest Then Exit Function Else Resume Next
End Function

'20200817: Modified to added "strTemp=Lcase(inst_name)".
'20191107: Modified to add FUNC.
'20190502: Modified for OFFSET_FUNC, ex: "MG005_GFXTD_BPL_BV".
Public Function decide_offset_testType_byInstName(inst_name As String) As Integer
On Error GoTo errHandler 'Add ErrHandler 2023/08/01
    Dim strTemp As String
'''//==================================================================================================================================================================================//'''
'''//Note:
'''For those instances with special offsets different from TD or Mbist, it should use keywords in the instance names same as special keywords added in column of FUNC, ex: BPL.
'''//==================================================================================================================================================================================//'''
    strTemp = LCase(inst_name)
    
    If strTemp Like "*gfxtd*" Or strTemp Like "*cputd*" Or strTemp Like "*soctd*" Or strTemp Like "*gputd*" Then
        decide_offset_testType_byInstName = testType.TD
    ElseIf strTemp Like "*cpu*bist*" Or strTemp Like "*gfx*bist*" Or strTemp Like "*soc*bist*" Or strTemp Like "*gpu*bist*" Then
        decide_offset_testType_byInstName = testType.Mbist
    Else
        'decide_offset_testType_byInstName = TestType.ldcbfd 'fake test type, if no match td/mbist/spi..., pretend as test type "LDCBFD"
        decide_offset_testType_byInstName = testType.Func
    End If
Exit Function
errHandler:
    Call Print_Error_Message(Error_Info, "LIB_VDD_BINNING", "decide_offset_testType_byInstName") 'Add ErrHandler 2023/08/01
    If AbortTest Then Exit Function Else Resume Next
End Function

'20210531: Modified to update theExec.sites.Selected for MultiFSTP before running PrePatt in run_prepatt_decompose_VT.
'20210129: Modified to revise the vbt code for DevChar.
'20201124: Modified to remove "Dim CorePowerStored_Init As New SiteDouble".
'20201118: Modified to use "TheHdw.Digital.Patgen.PatternBurstPassedPerSite" to get siteResult of pattern pass/fail.
'20201029: Modified to remove the argument "result_mode As tlResultMode" and use inst_info.result_mode.
'20201029: Modified to use inst_info.previousDcvsOutput and inst_info.currentDcvsOutput.
'20201029: Modified to remove the argument "Optional idxBlock_Selsrm_singlePatt As Integer".
'20201029: Modified to use inst_info.is_BV_Safe_Voltage_printed and inst_info.is_BV_Payload_Voltage_printed.
'20201027: Modified to add the argument "IndexLevelPerSite As SiteLong".
'20201027: Modified to use "Public Type Instance_Info".
'20201026: Modified to revise the vbt code for TD pattern burst proposed by C651 Toby.
'20201012: Modified to use "check_patt_Pass_Fail" to check pattern Pass/Fail.
'20200924: Modified to move "select_DCVS_output_for_powerDomain" from GradeSearch_VT to "run_prepatt_decompose_VT".
'20200923: Modified to move the position of "Check_Pattern_NoBurst_NoDecompose".
'20200923: Modified to remove "run_prepatt" and keep "run_prepatt_decompose_VT".
'20200921: Modified to check if "Test_Type = TestType.Mbist".
'20200520: Modified to use Check_Pattern_NoBurst_NoDecompose to show the errorLogMessage if "burst=no" and "Decompose_Pattern=false".
'20200319: Modified to switch off save_core_power_vddbinning and restore_core_power_vddbinning if Flag_Enable_Rail_Switch = True.
'20200203: Modified to use the function "print_bincut_power".
'20200115: Modified to skip applying safe voltages to non-selsram powerpin for project with selsrm_mapping_table.
'20200113: Modified for pattern bursted without decomposing pattern.
'20200106: Modified to add "TheHdw.Alarms.Check".
'20191202: Modified for the revised initVddBinCondition.
'20191127: Modified for the revised InitVddBinTable.
'20191125: Modified PrePattPass to avoid pseudo pass.
'20191125: Modified to remove IGSIM block.
'20190627: Modified to use the global variable "pinGroup_BinCut" for BinCut powerPins.
'20190617: Modified to use siteDouble "CorePowerStored()" to save/restore voltages for BinCut powerPins.
'20190606: Modified to add the argument "DcSpecsCategoryForInitPat as string" for Init patterns with the new test setting DC Specs.
Public Function run_prepatt_decompose_VT(inst_info As Instance_Info, PrePatt As String, ary_PrePatt_decomposed() As String, count_PrePatt_decomposed As Long, PrePattPass As SiteBoolean, Optional DcSpecsCategoryForInitPat As String)
On Error GoTo errHandler 'Add ErrHandler 2023/08/01
    Dim site As Variant
    Dim CorePowerStored() As New SiteDouble
    Dim indexPatt As Long
    Dim i As Integer
    Dim offsetTestTypeIdx As Integer
    Dim sitePatPass As New SiteBoolean
'''//==================================================================================================================================================================================//'''
'''//Note:
'''20210531: Modified to update theExec.sites.Selected for MultiFSTP before running PrePatt in run_prepatt_decompose_VT.
'''It seemed that theExec.sites.Selected masked the failed site (not siteShutDown). But the site still ran pattern.test without updating test results.
'''Discussed this with Chihome. He saw this is ancient projects, and he suggested us to check if test results were correct.
'''We checked test results, and it seemed no error with PassBin and EQN.
'''//==================================================================================================================================================================================//'''
    If PrePatt <> "" Then
        '''//Update theExec.sites.Selected for MultiFSTP before running PrePatt.
        '''20210531: Modified to update theExec.sites.Selected for MultiFSTP before running PrePatt in run_prepatt_decompose_VT.
        '''ToDo: Please check if EnableWord("Multifstp_Datacollection") exists in the flow table!!!
        If TheExec.flow.enableWord("Multifstp_Datacollection") = True Then
            TheExec.sites.Selected = gb_siteMask_current
        End If
    
        '''//For BIST instance, it saves current DCVS Vmain values into the array "CorePowerStored", and then set safe voltage to DCVS Vmain by DC Category.
        If inst_info.Test_Type = testType.Mbist Then '''ex: "*cpu*bist*", "*gfx*bist*", "*gpu*bist*", "*soc*bist*".
            '''init
            '''//siteDouble "CorePowerStored()" is used to save/restore voltages for BinCut powerDomains.
            ReDim CorePowerStored(UBound(pinGroup_BinCut))
            
            For i = 0 To UBound(pinGroup_BinCut)
                CorePowerStored(i) = 0
            Next i
        
            '''//Save payload voltages of CorePower and OtherRail powerPins before init pattern.
            If Flag_noRestoreVoltageForPrepatt = False Then
                save_core_power_vddbinning CorePowerStored
            End If
            
            '''//Get BinCut INIT voltages (safe voltages for SELSRM DSSC), usually set to nominal voltage.
            '''//If initial voltages and safe voltage(init voltage) use the same DC category, it can skip "set_core_power_vddbinning_VT" after initial voltages...
            '''Note: For PTE/TTR, it can use the flag "Flag_Skip_ReApplyInitVolageToDCVS" to skip "set_core_power_vddbinning_VT".
            'T-Col TTR approve by Si -- 230413
            If Flag_noRestoreVoltageForPrepatt = False Then
                set_core_power_vddbinning_VT "NV", DcSpecsCategoryForInitPat
            End If
            TheHdw.Wait 0.0001
        End If
        
        '''//Print safe voltages(init voltages) for PrePatt(init patt).
        print_bincut_voltage inst_info, , Flag_Remove_Printing_BV_voltages, Flag_PrintDcvsShadowVoltage, BincutVoltageType.SafeVoltage
        
'**********************************************
'@@PrePatt pattern-loop Start
'**********************************************
        For indexPatt = 0 To count_PrePatt_decomposed - 1
            '''//Sync up DCVS output and print BinCut payload voltage for projects with Rail Switch for TD instance.
            Call prepare_DCVS_Output_for_RailSwitch(inst_info, ary_PrePatt_decomposed(indexPatt), inst_info.idxBlock_Selsrm_PrePatt)
            
            '''//Run pattern.
            Call TheHdw.patterns(inst_info.ary_PrePatt_decomposed(indexPatt)).test(pfAlways, 0, inst_info.result_mode)
            
            '''//Get siteResult of pattern pass/fail.
            '''//Warning!!! currently "TheHdw.Digital.Patgen.PatternBurstPassedPerSite" doesn't support "result_mode=tlResultModeModule" with PatternBurst=Yes and DecomposePatt=No.
            sitePatPass = TheHdw.Digital.Patgen.PatternBurstPassedPerSite
            
            If inst_info.enable_HarvPinGrp = True Then
                'if bin1 initial pat fail turn com fail result to true
                For Each site In TheExec.sites
                    If sitePatPass(site) = False And CurrentPassBinCutNum(site) = 1 Then
                        If inst_info.HarvPinGrpCommnFailResult(site) = False Then
                            inst_info.HarvPinGrpCommnFailResult(site) = True
                        End If
                    End If
                Next
            End If
            
            '''//Check alarmFail for pattern.
            Call check_alarmFail_for_pattern(sitePatPass)

            '''for DevChar.
            If inst_info.is_DevChar_Running = False Then
                '''//Update pattern pass/fail to the flag.
                Call update_Pattern_result_to_PattPass(sitePatPass, PrePattPass)
            End If
        Next indexPatt
        
        DebugPrintFunc PrePatt
'**********************************************
'@@PrePatt pattern-loop End
'**********************************************
        '''//Check if running Pattern with "burst=no" and "Decompose_Pattern=false".
        Call Check_Pattern_NoBurst_NoDecompose(inst_info.PrePatt, inst_info.count_PrePatt_decomposed, inst_info.enable_DecomposePatt)
        
        '''//Restore the BinCut voltages for payload patterns after init pattern.
        If inst_info.Test_Type = testType.Mbist Then '''ex: "*cpu*bist*", "*gfx*bist*", "*gpu*bist*", "*soc*bist*".
            If Flag_noRestoreVoltageForPrepatt = False Then
                restore_core_power_vddbinning CorePowerStored
            End If
        End If
    End If
Exit Function
errHandler:
    Call Print_Error_Message(Error_Info, "LIB_VDD_BINNING", "run_prepatt_decompose_VT") 'Add ErrHandler 2023/08/01
    If AbortTest Then Exit Function Else Resume Next
End Function

'20210823: Modified to revise the vbt code for C651 new rules of testJobs naming, as requested by C651 Toby and TSMC ZYLINI.
'20210819: Modified to add the special case "wlft2" mapping to "ft_room" for T-Cre A0, as requested by C651 Toby and TSMC ZYLINI.
'20200827: Modified to replace "If..Else" with "Select Case".
'20200731: Modified to merge MappingBincutJobName and Mapping_TPJobName_to_BincutJobName into Mapping_TestJobName_to_BincutJobName.
'20180704: Created for BinCut testjob mapping.
Public Function Mapping_TestJobName_to_BincutJobName(strTestJob As String) As String
On Error GoTo errHandler
'''//==================================================================================================================================================================================//'''
'''//Note:
'''//Please discuss TestJobs of test program with C651 project DRI and decide how to mapping testjobs (especially for FT3 and XXX_QA) to 5 BinCut testJobs (CP1, CP2, FT_ROOM, FT_HOT, QA)!!!
'''For example, BinCut doesn't have FT3, but DC_TEST_IDS still uses BinCut IDS limit as IDS HiLimit.
'''//==================================================================================================================================================================================//'''
    '''//This subroutine needs to be kept maintaining depending on individual project// ''''
    Select Case LCase(strTestJob)
        '''//cp1
        Case "cp1": Mapping_TestJobName_to_BincutJobName = "cp1"
        '''//cp2
        Case "cp2": Mapping_TestJobName_to_BincutJobName = "cp2"
        '''//ft_room
        Case "ft_room": Mapping_TestJobName_to_BincutJobName = "ft_room"
        Case "wlft1": Mapping_TestJobName_to_BincutJobName = "ft_room"
        '''20210819: Modified to add the special case "wlft2" mapping to "ft_room" for T-Cre A0, as requested by C651 Toby and TSMC ZYLINI.
        Case "wlft2": Mapping_TestJobName_to_BincutJobName = "ft_room" '''the special case "wlft2" mapping to "ft_room" for T-Cre A0, as requested by C651 Toby and TSMC ZYLINI.
        Case "ft1": Mapping_TestJobName_to_BincutJobName = "ft_room"
        Case "ft2_25c": Mapping_TestJobName_to_BincutJobName = "ft_room"
        Case "ft3": Mapping_TestJobName_to_BincutJobName = "ft_room"
        Case "ft4": Mapping_TestJobName_to_BincutJobName = "ft_room"
        Case "ft5": Mapping_TestJobName_to_BincutJobName = "ft_room"
        '''20210823: Modified to revise the vbt code for C651 new rules of testJobs naming, as requested by C651 Toby and TSMC ZYLINI.
        Case "rma_room": Mapping_TestJobName_to_BincutJobName = "ft_room"
        '''//ft_hot
        Case "ft_hot": Mapping_TestJobName_to_BincutJobName = "ft_hot"
        Case "ft2": Mapping_TestJobName_to_BincutJobName = "ft_hot"
        Case "ft2_85c": Mapping_TestJobName_to_BincutJobName = "ft_hot"
        '''20210823: Modified to revise the vbt code for C651 new rules of testJobs naming, as requested by C651 Toby and TSMC ZYLINI.
        Case "rma_hot": Mapping_TestJobName_to_BincutJobName = "ft_hot"
        '''//qa
        Case "qa": Mapping_TestJobName_to_BincutJobName = "qa"
        '''20200827: Modified to mask "wlft1_qa" and "ft2_85c_qa" because these testJob are not the common testJobs in recent projects. These need to be discussed with C651 project DRIs!!!
        'Case "wlft1_qa": Mapping_TestJobName_to_BincutJobName = "qa"
        'Case "ft2_85c_qa": Mapping_TestJobName_to_BincutJobName = "qa"
'''ToDo: Discuss rules of testJobs mapping for "FT1_FQA", "FT2_FQA", "T0TX_ROOM", "T0TX_HOT" with C651 and TSMC...
        '''//others
        Case Else:
                Mapping_TestJobName_to_BincutJobName = vbNullString
                Call Print_Error_Message(Error_Info, "LIB_VDD_BINNING", "Mapping_TestJobName_to_BincutJobName", "job:" & strTestJob & ", it doesn't have any matched definition for Mapping_TestJobName_to_BincutJobName. Error!!!")
                'TheExec.ErrorLogMessage "job:" & strTestJob & ", it doesn't have any matched definition for Mapping_TestJobName_to_BincutJobName. Error!!!"    ''20230807: Removed by sanity check
    End Select
Exit Function
errHandler:
    Call Print_Error_Message(Error_Info, "LIB_VDD_BINNING", "Mapping_TestJobName_to_BincutJobName") 'Add ErrHandler 2023/08/01
    If AbortTest Then Exit Function Else Resume Next
End Function

'20210727: C651 Toby provided the BinCut flow with testCondition "M*### E1 voltage" for non BinCut search in CP1, so that judge_stored_IDS(check_IDS) should be compatible with all BinCut testJobs.
'20210707: Modified to check if ids_name (Efuse category) exists in dict_EfuseCategory2BinCutTestJob.
'20210617: Discussed this with TSMC T-Cre team and C651 Si. C651 Si said that Check_IDS and judge_IDS are dedicated to Efuse processed IDS for BinCut search.
'20201228: Patty asked us to adapt Efuse object vbt code.
'20201210: Modified to use the flag "is_BinCutJob_for_StepSearch" for "check_bincutJob_for_StepSearch" to check if the test program is binSearch or functional test.
'20201005: Modified to set Lo_Limit=0 for non-CP1.
'20200827: Modified to replace "If..Else" with "Select Case".
'20200106: Modified to remove the ErrorLogMessage.
'20190813: Modified to use different IDS lo_limit by BinCut testjobs.
'20190722: Modified to printout the scale and the unit for BinCut voltages and IDS values.
'20190716: Modified to unify the unit for IDS.
'20190612: Modified to use the new datatype of IDS.
'20180917: Due to data with double format accuracy issue, we follow the suggestion from Microsoft official document to use "Cdec".
'20180209: SWLINZA modified this because we should read efuse to identify BinX/BinY.
'20170810: SWLINZA modified to get Resolution from EFUSE_BitDef_Table and set it as IDS low limit for CorePower.
Public Function judge_IDS(ids_current As SiteDouble, performance_mode As String, site As Variant)
On Error GoTo errHandler 'Add ErrHandler 2023/08/01
    Dim Hi_Limit As Double
    Dim Lo_Limit As Double
    Dim p_mode As Integer
    Dim powerDomain As String
    Dim i As Long
    Dim Mode_Step As Integer
    Dim str_IDS_PowerDomain As String
    '''for Efuse obj vbt.
    Dim field As eFuseBdfField
'''//==================================================================================================================================================================================//'''
'''//Note:
'''1. Check_IDS and judge_IDS are dedicated to Efuse processed IDS for BinCut search, as requested by C651 Si, 20210617.
'''2. C651 Toby provided the BinCut flow with testCondition "M*### E1 voltage" for non BinCut search in CP1, so that judge_stored_IDS(check_IDS) should be compatible with all BinCut testJobs, 20210727.
'''//==================================================================================================================================================================================//'''
    '''//Get p_mode from performance mode
    p_mode = VddBinStr2Enum(performance_mode)
    powerDomain = AllBinCut(p_mode).powerPin
    
    '''//Get IDS name by powerDomain for each site.
    str_IDS_PowerDomain = IDS_for_BinCut(VddBinStr2Enum(powerDomain)).ids_name(site)  'Modify to site variant
        
    '''************************************************************************************************************************************************************'''
    '''//Use Resolution of IDS from EFUSE_BitDef_Table as IDS low limit with the scale and the unit in "mA".
    '''//The flag "is_BinCutJob_for_StepSearch" is True if any testCondition from the table "Non_Binning_Rail" has the keyword "*Evaluate*Bin*".
    '''20210617: Discussed this with TSMC T-Cre team and C651 Si. C651 Si said that Check_IDS and judge_IDS are dedicated to Efuse processed IDS for BinCut search.
    '''************************************************************************************************************************************************************'''
    '''//Check if ids_name of PowerDomain is the correct Efuse category in Efuse_BitDef_Table.
    If dict_EfuseCategory2BinCutTestJob.Exists(UCase(str_IDS_PowerDomain)) = True Then
        '''For project with Efuse obj vbt code.
        '''20201228: Patty asked us to adapt Efuse object vbt code.
        ''====20201230 add for efuse new code====
        Set field = BdfDataBase.Bank_Cfg.Fields(str_IDS_PowerDomain)
        Lo_Limit = 1# * field.Resolution

        '''************************************************************************************************************************************************'''
        ''' Loop the BinCut table and use the different CP current limit to print out the datalog and only judge the IDS by AllBinCut(P_mode).IDS_CP_LIMIT
        '''************************************************************************************************************************************************'''
        For i = 0 To UBound(PassBinCut_ary)
            '''//IDS calculation uses the scale and the unit in "mA", but TheExec.Flow.TestLimit should convert IDS value into "A" with settings "unit:=unitAmp" and "scaleMilli".
            If i = UBound(PassBinCut_ary) Then
                Hi_Limit = AllBinCut(p_mode).IDS_CP_LIMIT '''unit: mA
                '20240131 Bin1 BinX merge judge Bin1 IDS or BinX IDS
                If Flag_BinX_Info_Parsed = True Then
                    If Pmode2BinX_IDS_Dict.Exists(p_mode) Then
                        If ids_current(site) < Pmode2BinX_IDS_Dict(p_mode) Then 'Bin1 IDS judge
                            TheExec.flow.TestLimit ids_current(site) / 1000, Lo_Limit / 1000, Pmode2BinX_IDS_Dict(p_mode) / 1000, , tlSignLess, scaleMilli, unit:=unitAmp, _
                                        PinName:=powerDomain, Tname:=VddBinName(p_mode) & " BinCut1 IDS", ForceUnit:=unitAmp
                        Else 'BinX IDS judge
                            TheExec.flow.TestLimit ids_current(site) / 1000, Lo_Limit / 1000, Pmode2BinX_IDS_Dict(p_mode) / 1000, , tlSignLess, scaleMilli, unit:=unitAmp, _
                                            PinName:=powerDomain, Tname:=VddBinName(p_mode) & " BinCut1 IDS", ForceResults:=tlForcePass, ForceUnit:=unitAmp

                            TheExec.flow.TestLimit ids_current(site) / 1000, Lo_Limit / 1000, Hi_Limit / 1000, , tlSignLess, scaleMilli, unit:=unitAmp, _
                                        PinName:=powerDomain, Tname:=VddBinName(p_mode) & " BinCut2 IDS", ForceUnit:=unitAmp
                            Flag_Fuse_BinX(site) = True
                        End If
                    Else
                        Call Print_Error_Message(Error_Info, "LIB_VDD_BINNING", "judge_IDS", powerDomain & performance_mode & " doesn't have CP1 IDS limit, just use BinX CP1 IDS limit, please check definition sheet.")
                        TheExec.flow.TestLimit ids_current(site) / 1000, Lo_Limit / 1000, Hi_Limit / 1000, , tlSignLess, scaleMilli, unit:=unitAmp, _
                                        PinName:=powerDomain, Tname:=VddBinName(p_mode) & " BinCut2 IDS", ForceUnit:=unitAmp
                    End If
                Else
                    TheExec.flow.TestLimit ids_current(site) / 1000, Lo_Limit / 1000, Hi_Limit / 1000, , tlSignLess, scaleMilli, unit:=unitAmp, _
                                        PinName:=powerDomain, Tname:=VddBinName(p_mode) & " BinCut" & PassBinCut_ary(i) & " IDS", ForceUnit:=unitAmp
                End If
            Else
                Mode_Step = BinCut(p_mode, PassBinCut_ary(i)).Mode_Step
                
                '''//IDS calculation uses the scale and the unit in "mA".
                Hi_Limit = BinCut(p_mode, PassBinCut_ary(i)).IDS_CP_LIMIT(Mode_Step) '''unit: mA
                
                '''//TheExec.Flow.TestLimit should convert IDS value into "A" with settings "unit:=unitAmp" and "scaleMilli".
                '''20180917: Due to data with double format accuracy issue, we follow the suggestion from Microsoft official document to use "Cdec".
                If CDec(ids_current) < CDec(Hi_Limit) Then
                    TheExec.flow.TestLimit ids_current(site) / 1000, Lo_Limit / 1000, Hi_Limit / 1000, , tlSignLess, scaleMilli, unit:=unitAmp, _
                                            PinName:=powerDomain, Tname:=VddBinName(p_mode) & " BinCut" & PassBinCut_ary(i) & " IDS", ForceUnit:=unitAmp
                Else
                    TheExec.flow.TestLimit ids_current(site) / 1000, Lo_Limit / 1000, Hi_Limit / 1000, , tlSignLess, scaleMilli, unit:=unitAmp, _
                                            PinName:=powerDomain, Tname:=VddBinName(p_mode) & " BinCut" & PassBinCut_ary(i) & " IDS", ForceResults:=tlForcePass, ForceUnit:=unitAmp
                End If
            End If
        Next i
    Else
        Call Print_Error_Message(Error_Info, "LIB_VDD_BINNING", "judge_IDS", performance_mode & ",Efuse category:" & str_IDS_PowerDomain & ",it can't use Efuse category to get IDS values for judge_IDS. Error!!!")
        'TheExec.Datalog.WriteComment performance_mode & ",Efuse category:" & str_IDS_PowerDomain & ",it can't use Efuse category to get IDS values for judge_IDS. Error!!!"
        'TheExec.ErrorLogMessage performance_mode & ",Efuse category:" & str_IDS_PowerDomain & ",it can't use Efuse category to get IDS values for judge_IDS. Error!!!"    ''20230807: Removed by sanity check
    End If '''If dict_EfuseCategory2BinCutTestJob.Exists(UCase(str_IDS_PowerDomain)) = True
Exit Function
errHandler:
    Call Print_Error_Message(Error_Info, "LIB_VDD_BINNING", "judge_IDS") 'Add ErrHandler 2023/08/01
    If AbortTest Then Exit Function Else Resume Next
End Function

'20210924: Modified to remove the redundant vbt code about "_85" for IDS values.
'20210907: Modified to check if siteNumber from the argument site is correct.
'20210727: C651 Toby provided the BinCut flow with testCondition "M*### E1 voltage" for non BinCut search in CP1, so that judge_stored_IDS(check_IDS) should be compatible with all BinCut testJobs.
'20210707: Modified to check if ids_name (Efuse category) exists in dict_EfuseCategory2BinCutTestJob.
'20210707: Modified to add site-loop to trig the failFlag strGlb_Flag_Vddbinning_IDS_fail for each site.
'20210629: Modified to adjust the sequence of the vbt code in get_I_VDD_values.
'20210507: Modified to replace testLimit with failStop for get_I_VDD_values.
'20201228: Patty asked us to adapt Efuse object vbt code.
'20201210: Modified to use the flag "is_BinCutJob_for_StepSearch" for "check_bincutJob_for_StepSearch" to check if the test program is binSearch or functional test.
'20200815: Modified to check powerPin exists.
'20200812: Modified to check powerDomain and powerPin.
'20200717: Modified to use the globalVariable "Flag_Vddbinning_IDS_fail".
'20200712: Modified to bin out the DUT due to IDS failed, suggested and requested by PCLIN.
'20200430: Modified to print the string about the incorrect IDS.
'20200130: Modified to get 1st powerPin from powerDomain.
'20200114: Modified to check if powerDomain exists in domain2pinDict or pin2domainDict.
'20200106: Modified to remove the ErrorLogMessage.
'20190716: Modified to unify the unit for IDS.
'20190630: Modified to use the real IDS values for non-CP1 tests.
'20190624: Modified to unify the unit of IDS with "A".
'20190615: Modified to align "Efuse Read Write Decimal" with I_VDD_xxx values
'20190523: Modified the argument for the new IDS datatype.
Public Function get_I_VDD_values(site As Variant, powerDomain As String, I_VDD_val As SiteDouble)
On Error GoTo errHandler 'Add ErrHandler 2023/08/01
    Dim str_IDS_PowerDomain As String
    Dim powerPin As String
    '''for Efuse obj vbt.
    Dim field As eFuseBdfField
'''//==================================================================================================================================================================================//'''
'''//Note:
'''1. Check_IDS and judge_IDS are dedicated to Efuse processed IDS for BinCut search, as requested by C651 Si, 20210617.
'''2. Warning!!!please check SortNumber and binNumber of Flag_Vddbinning_IDS_fail in Bin_Table before using this.
'''3. As per discussion with TSMC SWLINZA, for powerPin group, it should use 1st powerPin to check IDS limit of powerPin group, 20210707.
'''ex: powerGroup: VDD_FIXED_GRP, and its 1st powerPin: VDD_FIXED, so that compare IDS value of VDD_FIXED with IDS_limit of VDD_FIXED_GRP. It must have Efuse category in Efuse_BitDef_Table to store IDS for VDD_FIXED.
'''4. C651 Toby provided the BinCut flow with testCondition "M*### E1 voltage" for non BinCut search in CP1, so that judge_stored_IDS(check_IDS) should be compatible with all BinCut testJobs, 20210727.
'''//==================================================================================================================================================================================//'''
    '''//Get the selected site.
    If CLng(site) < 0 Then
        Call Print_Error_Message(Error_Info, "LIB_VDD_BINNING", "get_I_VDD_values", "Please check the site number of get_I_VDD_values. Error!!!")
        'TheExec.ErrorLogMessage "Please check the site number of get_I_VDD_values. Error!!!"
    End If
    
    '''//Warning!!! Please contact project Efuser owner to see if CFG Read/Write with scale mA or not.
    '''//IDS calculation uses the scale and the unit in "mA", but TheExec.Flow.TestLimit should convert IDS value into "A" with settings "unit:=unitAmp" and "scaleMilli".
    '''//The flag "is_BinCutJob_for_StepSearch" is True if any testCondition from the table "Non_Binning_Rail" has the keyword "*Evaluate*Bin*".
    '''ToDo: Please discuss this with project Efuse owner to see if rules about "Programming Stage" in Efuse_BitDef_Table are changed.
    '''20210727: C651 Toby provided the BinCut flow with testCondition "M*### E1 voltage" for non BinCut search in CP1, so that judge_stored_IDS(check_IDS) should be compatible with all BinCut testJobs.
    '''***************************************************************************************************************************************************'''
    '''[CP1] get IDS values from Efuse CFG data structure by IDS name.
    '''***************************************************************************************************************************************************'''
    '''//Step1: Get IDS name of powerDomain for each site.
    str_IDS_PowerDomain = IDS_for_BinCut(VddBinStr2Enum(powerDomain)).ids_name(site)

    '''//Step2: Get IDS value from Efuse CFG by IDS name of powerDomain.
    '''//If IDS_name of PowerDomain is correct, it will get IDS real values from Efuse Read.Decimal or Write.Decimal.
    If dict_EfuseCategory2BinCutTestJob.Exists(UCase(str_IDS_PowerDomain)) = True Then
        '''==========================================================================================================
        '''Note: Read IDS from efuse Read Decimal.
        '''If the IDS from efuse Read Decimal is 0, it can get IDS from efuse Write Decimal. Make sure the IDS is not 0.
        '''==========================================================================================================
        '''For project with Efuse obj vbt code.
        '''20201228: Patty asked us to adapt Efuse object vbt code.
        Set field = BdfDataBase.Bank_Cfg.Fields(str_IDS_PowerDomain)    ''202004xx for ap

        '''20201228: Patty asked us to adapt Efuse object vbt code.
        If field.DsscDecValue * field.Resolution <> 0 Then
            I_VDD_val(site) = field.DsscDecValue * field.Resolution '''unit: mA
        ElseIf field.TrimAteDecValue * field.Resolution <> 0 Then
            I_VDD_val(site) = field.TrimAteDecValue * field.Resolution '''unit: mA
        Else
            I_VDD_val(site) = 0
            Call Print_Error_Message(Error_Info, "LIB_VDD_BINNING", "get_I_VDD_values", "site:" & site & ",powerDomain:" & powerDomain & ",Efuse category:" & str_IDS_PowerDomain & ", IDS value from Efuse CFG is 0. Please check DC_TEST_IDS or EFuse. Error!!!")
            '''//Use the globalVariable "Flag_Vddbinning_IDS_fail" to bin out the DUT due to IDS failed, as suggested and requested by PCLIN.
            TheExec.sites.item(site).FlagState(strGlb_Flag_Vddbinning_IDS_fail) = logicTrue
        End If
    Else
        I_VDD_val(site) = 0
        Call Print_Error_Message(Error_Info, "LIB_VDD_BINNING", "get_I_VDD_values", "site:" & site & ",powerDomain:" & powerDomain & ",Efuse category:" & str_IDS_PowerDomain & ", it can't use Efuse category for get_I_VDD_values, Please check Efuse_BitDef_Table. Error!!!")
        'TheExec.ErrorLogMessage "site:" & site & ",powerDomain:" & powerDomain & ",Efuse category:" & str_IDS_PowerDomain & ", it can't use Efuse category for get_I_VDD_values, Please check Efuse_BitDef_Table. Error!!!"
    End If
Exit Function
errHandler:
    Call Print_Error_Message(Error_Info, "LIB_VDD_BINNING", "get_I_VDD_values") 'Add ErrHandler 2023/08/01
    If AbortTest Then Exit Function Else Resume Next
End Function

'20200106: Modified to remove the ErrorLogMessage.
'20180807: Modified for BinCut testjob mapping.
Public Function getBinCutJobDefinition(strInput As String) As Integer
On Error GoTo errHandler 'Add ErrHandler 2023/08/01
    Dim strTemp As String
'''//==================================================================================================================================================================================//'''
'''//Note:
'''We use strInput (the keyword about testJob) to decide the jobidx, Enum BinCutJobDefinition in global variable.
'''define CP1 = 0, CP2 = 1, FT1 = 2, FT2 = 3, QA = 4.
'''//==================================================================================================================================================================================//'''
    strTemp = LCase(Trim(strInput))

    If strTemp Like "*cp1*" Or strTemp Like "*binsearch*" Then
        getBinCutJobDefinition = BinCutJobDefinition.CP1
    ElseIf strTemp Like "*cp2*" Then
        getBinCutJobDefinition = BinCutJobDefinition.CP2
    ElseIf strTemp Like "*ft_room*" Or strTemp Like "*ft1*" Or strTemp Like "*ft2_25c" Then
        getBinCutJobDefinition = BinCutJobDefinition.FT1
    ElseIf strTemp Like "*ft_hot*" Or strTemp Like "*ft2_85c*" Or strTemp Like "*ft2*" Then
        getBinCutJobDefinition = BinCutJobDefinition.FT2
    ElseIf strTemp Like "*qa*" Then
        getBinCutJobDefinition = BinCutJobDefinition.QA
    Else
        getBinCutJobDefinition = BinCutJobDefinition.COND_ERROR
        Call Print_Error_Message(Error_Info, "LIB_VDD_BINNING", "getBinCutJobDefinition", "getBinCutJobDefinition = " & strInput & " doesn't exist in enum BinCutJobDefinition. Error!!!")
        'TheExec.ErrorLogMessage "getBinCutJobDefinition = " & strInput & " doesn't exist in enum BinCutJobDefinition. Error!!!"
    End If
Exit Function
errHandler:
    Call Print_Error_Message(Error_Info, "LIB_VDD_BINNING", "getBinCutJobDefinition") 'Add ErrHandler 2023/08/01
    If AbortTest Then Exit Function Else Resume Next
End Function

'20210722: Modified to use VBIN_IDS_ZONE(p_mode).Product_Voltage(Max_IDS_Zone, Max_IDS_Step) and DYNAMIC_VBIN_IDS_ZONE(p_mode).Product_Voltage(Max_IDS_Step) for GradeVDD.
'20210701: Modified to remove "Exit Function" from the vbt function check_voltageInheritance_for_powerDomain, as requested by TER Verity.
'20210621: Modified to use the argument "powerDomain as string" for the input powerDomain of the vbt function check_voltageInheritance_for_powerDomain.
'20210621: Modified to remove the redundant argument "ids_PowerDomain As SiteDouble" from the vbt function check_voltageInheritance_for_powerDomain.
'20210621: Modified to rename the vbt function check_pmode_for_adjust_VddBinning as check_voltageInheritance_for_powerDomain.
'20210621: Modified to merge the vbt code from the vbt function find_next_bin_eq_interpolation.
'20210303: Modified to remove the redundant "ids_current As SiteDouble" from arguments of the vbt function "find_next_bin_eq_interpolation".
'20201113: Modified to rename the argument "I_VDD_core_power" as "ids_PowerDomain".
'20200825: Modified to remove the redundant branches.
'20191127: Modified for the revised InitVddBinTable.
'20190227: For complete BinCut search steps in DYNAMIC_VBIN_IDS_ZONE, we replace "find_next_bin_eq" with "find_next_bin_eq_interpolation" according to BinCut monthly meeting Dec-2018.
'20181026: Added for interpolation by Oscar.
'20180816: Created this function to simplify adjust_VddBinning.
Public Function check_voltageInheritance_for_powerDomain(powerDomain As String)
On Error GoTo errHandler 'Add ErrHandler 2023/08/01
    Dim site As Variant
    Dim strAry_Pmode_Seq() As String
    Dim p_mode As Integer
    Dim gradevdd_last As New SiteDouble
    Dim grade_last As New SiteDouble
    Dim i As Long
    '''
    Dim idx_step As Long
    Dim next_bin_flag As Boolean
    Dim exit_while_flag As Boolean
'''//==================================================================================================================================================================================//'''
'''//Note:
'''1. Only testjob with keyword "*Evaluate*Bin*" in the testCondition of BinCut flow can check voltage inheritance in Adjust_VddBinning.
'''2. 20210610 C651 Si defined the rule of voltage inheritance check: (GradeVDD(P_mode)-GradeVDD(previous perfromance_mode))> Monotonicity_Offset if without Allow_Equal.
'''//==================================================================================================================================================================================//'''
    '''//Check if powerDomain is the BinCut CorePower defined in the header of BinCut flow(sheet "Non_Binning_Rail").
    If dict_IsCorePowerInBinCutFlowSheet.Exists(UCase(powerDomain)) = True Then
        If dict_IsCorePowerInBinCutFlowSheet.item(UCase(powerDomain)) = True Then
            '20211022: Added to support SRAM Voltage calculate with equation
            If dict_IsCorePower(powerDomain) = True And UCase(powerDomain) Like "*VDD*_SRAM*" Then
                check_voltageInheritance_for_SrampowerDomain (powerDomain)
                Exit Function
            End If
            strAry_Pmode_Seq = BinCut_Power_Seq(VddBinStr2Enum(powerDomain)).Power_Seq
            
            If UBound(strAry_Pmode_Seq) >= 1 Then
                '''BinCut powerDomain has more than one p_mode, and it's ready to check voltage inheritance for the Power_Seq of BinCut powerDomain.
            Else
                '''If UBound(strAry_Pmode_Seq)=0, skip voltage inheritance check for BinCut powerDomain.
                Exit Function
            End If
        Else
            Call Print_Error_Message(Error_Info, "LIB_VDD_BINNING", "check_voltageInheritance_for_powerDomain", "PowerDomain:" & powerDomain & ",it isn't BinCut CorePower for check_voltageInheritance_for_powerDomain. Error!!!")
            'TheExec.ErrorLogMessage "PowerDomain:" & powerDomain & ",it isn't BinCut CorePower for check_voltageInheritance_for_powerDomain. Error!!!"
            Exit Function
        End If
    Else
        Call Print_Error_Message(Error_Info, "LIB_VDD_BINNING", "check_voltageInheritance_for_powerDomain", "PowerDomain:" & powerDomain & ",it isn't BinCut CorePower for check_voltageInheritance_for_powerDomain. Error!!!")
        'TheExec.ErrorLogMessage "PowerDomain:" & powerDomain & ",it isn't BinCut CorePower for check_voltageInheritance_for_powerDomain. Error!!!"
        Exit Function
    End If
    
    '''//If UBound(strAry_Pmode_Seq) >= 1, adjust the vdd binning value to make voltage of each p_mode is always greater than its previous performance_mode for BinCut powerDomain.
    For i = 1 To UBound(strAry_Pmode_Seq)
        p_mode = VddBinStr2Enum(strAry_Pmode_Seq(i))
        
        '''//Get BinCut voltage and Efuse product voltage of the previous performance_mode.
        gradevdd_last = VBIN_RESULT(AllBinCut(VddBinStr2Enum(strAry_Pmode_Seq(i))).PREVIOUS_Performance_Mode).GRADEVDD
        grade_last = VBIN_RESULT(AllBinCut(VddBinStr2Enum(strAry_Pmode_Seq(i))).PREVIOUS_Performance_Mode).GRADE
        
        If BinCut(p_mode, VBIN_RESULT(p_mode).passBinCut(site)).ExcludedPmode = True Then
            '''Set Grade and GradeVdd to 0 if the Performance Mode is not enabled in the Test Result (Bin1 or BinX).
            VBIN_RESULT(p_mode).GRADE = 0
            VBIN_RESULT(p_mode).step_in_BinCut = -1
            VBIN_RESULT(p_mode).GRADEVDD = 0
        Else
            '''//Check the voltage heritance between p_mode and the previous performance_mode.
            For Each site In TheExec.sites
                next_bin_flag = False
            
                If CDec(VBIN_RESULT(p_mode).GRADE) > 0 Then
                    idx_step = VBIN_RESULT(p_mode).step_in_IDS_Zone
                    exit_while_flag = False
                    
                    '''//Check if p_mode has Allow_Equal with the previous performance_mode.
                    If AllBinCut(p_mode).PREVIOUS_Performance_Mode = AllBinCut(p_mode).Allow_Equal And AllBinCut(p_mode).Allow_Equal <> 0 Then '''for AllowEqual
                        '''//Note: If the vbt of checking GRADE is masked, please set globalVariable "Public Const Flag_Only_Check_PV_for_VoltageHeritage As Boolean = True".
                        '''//Print the status of VBIN_RESULT(p_mode).is_Monotonicity_Offset_triggered(site) with PTR format in Adjust_Binning for datalogs.
                        '''20210526: C651 Si revised the check rules to ensure that: (GradeVDD(P_mode)-GradeVDD(previous perfromance_mode))> Monotonicity_Offset.
                        If Flag_Get_column_Monotonicity_Offset = True Then
                            If (CDec(VBIN_RESULT(p_mode).GRADEVDD - gradevdd_last(site)) < CDec(BinCut(p_mode, DYNAMIC_VBIN_IDS_ZONE(p_mode).passBinCut(idx_step)).Monotonicity_Offset(DYNAMIC_VBIN_IDS_ZONE(p_mode).EQ_Num(idx_step) - 1))) Then
                                VBIN_RESULT(p_mode).is_Monotonicity_Offset_triggered(site) = True
                                TheExec.Datalog.WriteComment "site:" & site & "," & VddBinName(p_mode) & "=" & VBIN_RESULT(p_mode).GRADEVDD
                            End If
                        End If
                        
                        '''//Check if GradeVDD(p_mode) => GradeVDD(previous_performance_mode) + Monotonicity_Offset(p_mode).
                        Do While (CDec(VBIN_RESULT(p_mode).GRADEVDD) < CDec(gradevdd_last + BinCut(p_mode, DYNAMIC_VBIN_IDS_ZONE(p_mode).passBinCut(idx_step)).Monotonicity_Offset(DYNAMIC_VBIN_IDS_ZONE(p_mode).EQ_Num(idx_step) - 1)) And exit_while_flag = False) _
                        'Or (CDec(VBIN_RESULT(p_mode).GRADE) < CDec(grade_last) And exit_while_flag = False)
                            idx_step = idx_step + 1
                            
                            If DYNAMIC_VBIN_IDS_ZONE(p_mode).passBinCut(idx_step) <> CurrentPassBinCutNum Then
                                Call Print_Error_Message(Error_Info, "LIB_VDD_BINNING", "", "site:" & site & "," & VddBinName(p_mode) & ",bin=" & CurrentPassBinCutNum(site) & ",but it can't find any step to adjust the product voltage for voltage inheritance check. Error!!!")
                                exit_while_flag = True
                                '''Warning!!!please check SortNumber and binNumber of Flag_Vddbinning_Fail_Stop in Bin_Table before using this.
                                '''20240303: Use modular function to bin out site
                                Call BinoutSite(site, VddbinFailStopSortNumber, VddbinFailStopBinNumber, strGlb_Flag_Vddbinning_Fail_Stop)
                                Exit Do     ''20221020: Added to avoid DoAll no shutdown scenario
                            Else
                                VBIN_RESULT(p_mode).GRADE = DYNAMIC_VBIN_IDS_ZONE(p_mode).Voltage(idx_step)
                                VBIN_RESULT(p_mode).GRADEVDD = DYNAMIC_VBIN_IDS_ZONE(p_mode).Product_Voltage(idx_step)
                                next_bin_flag = True
                            End If
                        Loop
                    Else
                        '''//Note: If the vbt of checking GRADE is masked, please set globalVariable "Public Const Flag_Only_Check_PV_for_VoltageHeritage As Boolean = True".
                        '''20210526: C651 Si revised the check rules to ensure that: (GradeVDD(P_mode)-GradeVDD(previous perfromance_mode))> Monotonicity_Offset.
                        '''20210610: Modified for the rule: (GradeVDD(P_mode)-GradeVDD(previous perfromance_mode))> Monotonicity_Offset if without Allow_Equal.
                        If Flag_Get_column_Monotonicity_Offset = True Then
                            If (CDec(VBIN_RESULT(p_mode).GRADEVDD - gradevdd_last(site)) <= CDec(BinCut(p_mode, DYNAMIC_VBIN_IDS_ZONE(p_mode).passBinCut(idx_step)).Monotonicity_Offset(DYNAMIC_VBIN_IDS_ZONE(p_mode).EQ_Num(idx_step) - 1))) Then
                                VBIN_RESULT(p_mode).is_Monotonicity_Offset_triggered(site) = True
                                TheExec.Datalog.WriteComment "site:" & site & "," & VddBinName(p_mode) & "=" & VBIN_RESULT(p_mode).GRADEVDD
                            End If
                        End If
                        
                        '''//Check if GradeVDD(p_mode) > GradeVDD(previous_performance_mode) + Monotonicity_Offset(p_mode).
                        Do While (CDec(VBIN_RESULT(p_mode).GRADEVDD) <= CDec(gradevdd_last + BinCut(p_mode, DYNAMIC_VBIN_IDS_ZONE(p_mode).passBinCut(idx_step)).Monotonicity_Offset(DYNAMIC_VBIN_IDS_ZONE(p_mode).EQ_Num(idx_step) - 1)) And exit_while_flag = False) _
                        'Or (CDec(VBIN_RESULT(p_mode).GRADE) <= CDec(grade_last) And exit_while_flag = False)
                            idx_step = idx_step + 1
                            
                            If DYNAMIC_VBIN_IDS_ZONE(p_mode).passBinCut(idx_step) <> CurrentPassBinCutNum Then
                                Call Print_Error_Message(Error_Info, "LIB_VDD_BINNING", "check_voltageInheritance_for_powerDomain", "site:" & site & "," & VddBinName(p_mode) & ",bin=" & CurrentPassBinCutNum(site) & ",but it can't find any step to adjust the product voltage for voltage inheritance check. Error!!!")
                                exit_while_flag = True
                                '''Warning!!!please check SortNumber and binNumber of Flag_Vddbinning_Fail_Stop in Bin_Table before using this.
                                 '''20240303: Use modular function to bin out site
                                Call BinoutSite(site, VddbinFailStopSortNumber, VddbinFailStopBinNumber, strGlb_Flag_Vddbinning_Fail_Stop)
                                Exit Do     ''20221020: Added to avoid DoAll no shutdown scenario
                            Else
                                VBIN_RESULT(p_mode).GRADE = DYNAMIC_VBIN_IDS_ZONE(p_mode).Voltage(idx_step)
                                VBIN_RESULT(p_mode).GRADEVDD = DYNAMIC_VBIN_IDS_ZONE(p_mode).Product_Voltage(idx_step)
                                next_bin_flag = True
                            End If
                        Loop
                    End If
                    
                    If next_bin_flag = True Then
                        VBIN_RESULT(p_mode).step_in_BinCut = DYNAMIC_VBIN_IDS_ZONE(p_mode).EQ_Num(idx_step) - 1
                        VBIN_RESULT(p_mode).step_in_IDS_Zone = idx_step
                    End If
                End If '''If CDec(VBIN_RESULT(p_mode).GRADE) > 0
            Next site
        End If '''If BinCut(p_mode, VBIN_RESULT(p_mode).passBinCut).ExcludedPmode = True
    Next i
Exit Function
errHandler:
    Call Print_Error_Message(Error_Info, "LIB_VDD_BINNING", "check_voltageInheritance_for_powerDomain") 'Add ErrHandler 2023/08/01
    If AbortTest Then Exit Function Else Resume Next
End Function

'20211022: Added to support SRAM Voltage calculate with equation
Public Function check_voltageInheritance_for_SrampowerDomain(powerDomain As String)
On Error GoTo errHandler 'Add ErrHandler 2023/08/01
    Dim site As Variant
    Dim strAry_Pmode_Seq() As String
    Dim p_mode As Integer
    Dim gradevdd_last As New SiteDouble
    Dim grade_last As New SiteDouble
    Dim i As Long, j As Long
    '''
    Dim idx_step As Long
    Dim next_bin_flag As Boolean
    Dim exit_while_flag As Boolean

    For j = 0 To BinCut_Sram_Power_Seq.Count - 1
    '''//==================================================================================================================================================================================//'''
    '''//Note:
    '''1. Only testjob with keyword "*Evaluate*Bin*" in the testCondition of BinCut flow can check voltage inheritance in Adjust_VddBinning.
    '''2. 20210610 C651 Si defined the rule of voltage inheritance check: (GradeVDD(P_mode)-GradeVDD(previous perfromance_mode))> Monotonicity_Offset if without Allow_Equal.
    '''//==================================================================================================================================================================================//'''
        '''//Check if powerDomain is the BinCut CorePower defined in the header of BinCut flow(sheet "Non_Binning_Rail").
        If dict_IsCorePowerInBinCutFlowSheet.Exists(UCase(powerDomain)) = True Then
            If dict_IsCorePowerInBinCutFlowSheet.item(UCase(powerDomain)) = True Then
                '20211022: Added to support SRAM Voltage calculate with equation
                If dict_IsCorePower(powerDomain) = True And UCase(powerDomain) Like "*VDD*_SRAM*" Then
                    strAry_Pmode_Seq = BinCut_Sram_Power_Seq.Items(j)
                End If
    
                If UBound(strAry_Pmode_Seq) >= 1 Then
                    '''BinCut powerDomain has more than one p_mode, and it's ready to check voltage inheritance for the Power_Seq of BinCut powerDomain.
                Else
                    '''If UBound(strAry_Pmode_Seq)=0, skip voltage inheritance check for BinCut powerDomain.
                    Exit Function
                End If
            Else
                Call Print_Error_Message(Error_Info, "LIB_VDD_BINNING", "check_voltageInheritance_for_SrampowerDomain", "PowerDomain:" & powerDomain & ",it isn't BinCut CorePower for check_voltageInheritance_for_powerDomain. Error!!!")
                'TheExec.ErrorLogMessage "PowerDomain:" & powerDomain & ",it isn't BinCut CorePower for check_voltageInheritance_for_powerDomain. Error!!!"
                Exit Function
            End If
        Else
            Call Print_Error_Message(Error_Info, "LIB_VDD_BINNING", "check_voltageInheritance_for_SrampowerDomain", "PowerDomain:" & powerDomain & ",it isn't BinCut CorePower for check_voltageInheritance_for_powerDomain. Error!!!")
            'TheExec.ErrorLogMessage "PowerDomain:" & powerDomain & ",it isn't BinCut CorePower for check_voltageInheritance_for_powerDomain. Error!!!"
            Exit Function
        End If
        
        '''//If UBound(strAry_Pmode_Seq) >= 1, adjust the vdd binning value to make voltage of each p_mode is always greater than its previous performance_mode for BinCut powerDomain.
        For i = 1 To UBound(strAry_Pmode_Seq)
            p_mode = VddBinStr2Enum(strAry_Pmode_Seq(i))
            
            '''//Get BinCut voltage and Efuse product voltage of the previous performance_mode.
            gradevdd_last = VBIN_RESULT(VddBinStr2Enum(strAry_Pmode_Seq(i - 1))).GRADEVDD
            grade_last = VBIN_RESULT(VddBinStr2Enum(strAry_Pmode_Seq(i - 1))).GRADE
            
            If BinCut(p_mode, VBIN_RESULT(p_mode).passBinCut(site)).ExcludedPmode = True Then
                '''Set Grade and GradeVdd to 0 if the Performance Mode is not enabled in the Test Result (Bin1 or BinX).
                VBIN_RESULT(p_mode).GRADE = 0
                VBIN_RESULT(p_mode).step_in_BinCut = -1
                VBIN_RESULT(p_mode).GRADEVDD = 0
            Else
                '''//Check the voltage heritance between p_mode and the previous performance_mode.
                For Each site In TheExec.sites
                    next_bin_flag = False
                
                    If CDec(VBIN_RESULT(p_mode).GRADE) > 0 Then
                        idx_step = VBIN_RESULT(p_mode).step_in_IDS_Zone
                        exit_while_flag = False
                        
                        '''//Check if p_mode has Allow_Equal with the previous performance_mode.
                        If AllBinCut(p_mode).Allow_Equal <> 0 Then  '''for AllowEqual
                            '''//Note: If the vbt of checking GRADE is masked, please set globalVariable "Public Const Flag_Only_Check_PV_for_VoltageHeritage As Boolean = True".
                            '''//Print the status of VBIN_RESULT(p_mode).is_Monotonicity_Offset_triggered(site) with PTR format in Adjust_Binning for datalogs.
                            '''20210526: C651 Si revised the check rules to ensure that: (GradeVDD(P_mode)-GradeVDD(previous perfromance_mode))> Monotonicity_Offset.
                            If Flag_Get_column_Monotonicity_Offset = True Then
                                If (CDec(VBIN_RESULT(p_mode).GRADEVDD - gradevdd_last(site)) < CDec(BinCut(p_mode, DYNAMIC_VBIN_IDS_ZONE(p_mode).passBinCut(idx_step)).Monotonicity_Offset(DYNAMIC_VBIN_IDS_ZONE(p_mode).EQ_Num(idx_step) - 1))) Then
                                    VBIN_RESULT(p_mode).is_Monotonicity_Offset_triggered(site) = True
                                    TheExec.Datalog.WriteComment "site:" & site & "," & VddBinName(p_mode) & "=" & VBIN_RESULT(p_mode).GRADEVDD
                                End If
                            End If
                            
                            '''//Check if GradeVDD(p_mode) => GradeVDD(previous_performance_mode) + Monotonicity_Offset(p_mode).
                            Do While (CDec(VBIN_RESULT(p_mode).GRADEVDD) < CDec(gradevdd_last + BinCut(p_mode, DYNAMIC_VBIN_IDS_ZONE(p_mode).passBinCut(idx_step)).Monotonicity_Offset(DYNAMIC_VBIN_IDS_ZONE(p_mode).EQ_Num(idx_step) - 1)) And exit_while_flag = False) _
                            'Or (CDec(VBIN_RESULT(p_mode).GRADE) < CDec(grade_last) And exit_while_flag = False)
                                idx_step = idx_step + 1
                                
                                If DYNAMIC_VBIN_IDS_ZONE(p_mode).passBinCut(idx_step) <> CurrentPassBinCutNum Then
                                    Call Print_Error_Message(Error_Info, "LIB_VDD_BINNING", "check_voltageInheritance_for_SrampowerDomain", "site:" & site & "," & VddBinName(p_mode) & ",bin=" & CurrentPassBinCutNum(site) & ",but it can't find any step to adjust the product voltage for voltage inheritance check. Error!!!")
                                    exit_while_flag = True
                                    '''Warning!!!please check SortNumber and binNumber of Flag_Vddbinning_Fail_Stop in Bin_Table before using this.
                                    '''20240303: Use modular function to bin out site
                                    Call BinoutSite(site, VddbinFailStopSortNumber, VddbinFailStopBinNumber, strGlb_Flag_Vddbinning_Fail_Stop)
                                    
                                    Exit Do
                                Else
                                    VBIN_RESULT(p_mode).GRADE = DYNAMIC_VBIN_IDS_ZONE(p_mode).Voltage(idx_step)
                                    VBIN_RESULT(p_mode).GRADEVDD = DYNAMIC_VBIN_IDS_ZONE(p_mode).Product_Voltage(idx_step)
                                    next_bin_flag = True
                                End If
                            Loop
                        Else
                            '''//Note: If the vbt of checking GRADE is masked, please set globalVariable "Public Const Flag_Only_Check_PV_for_VoltageHeritage As Boolean = True".
                            '''20210526: C651 Si revised the check rules to ensure that: (GradeVDD(P_mode)-GradeVDD(previous perfromance_mode))> Monotonicity_Offset.
                            '''20210610: Modified for the rule: (GradeVDD(P_mode)-GradeVDD(previous perfromance_mode))> Monotonicity_Offset if without Allow_Equal.
                            If Flag_Get_column_Monotonicity_Offset = True Then
                                If (CDec(VBIN_RESULT(p_mode).GRADEVDD - gradevdd_last(site)) <= CDec(BinCut(p_mode, DYNAMIC_VBIN_IDS_ZONE(p_mode).passBinCut(idx_step)).Monotonicity_Offset(DYNAMIC_VBIN_IDS_ZONE(p_mode).EQ_Num(idx_step) - 1))) Then
                                    VBIN_RESULT(p_mode).is_Monotonicity_Offset_triggered(site) = True
                                    TheExec.Datalog.WriteComment "site:" & site & "," & VddBinName(p_mode) & "=" & VBIN_RESULT(p_mode).GRADEVDD
                                End If
                            End If
                            
                            '''//Check if GradeVDD(p_mode) > GradeVDD(previous_performance_mode) + Monotonicity_Offset(p_mode).
                            Do While (CDec(VBIN_RESULT(p_mode).GRADEVDD) <= CDec(gradevdd_last + BinCut(p_mode, DYNAMIC_VBIN_IDS_ZONE(p_mode).passBinCut(idx_step)).Monotonicity_Offset(DYNAMIC_VBIN_IDS_ZONE(p_mode).EQ_Num(idx_step) - 1)) And exit_while_flag = False) _
                            'Or (CDec(VBIN_RESULT(p_mode).GRADE) <= CDec(grade_last) And exit_while_flag = False)
                                idx_step = idx_step + 1
                                
                                If DYNAMIC_VBIN_IDS_ZONE(p_mode).passBinCut(idx_step) <> CurrentPassBinCutNum Then
                                    Call Print_Error_Message(Error_Info, "LIB_VDD_BINNING", "", "site:" & site & "," & VddBinName(p_mode) & ",bin=" & CurrentPassBinCutNum(site) & ",but it can't find any step to adjust the product voltage for voltage inheritance check. Error!!!")
                                    exit_while_flag = True
                                    '''Warning!!!please check SortNumber and binNumber of Flag_Vddbinning_Fail_Stop in Bin_Table before using this.
                                    '''20240303: Use modular function to bin out site
                                    Call BinoutSite(site, VddbinFailStopSortNumber, VddbinFailStopBinNumber, strGlb_Flag_Vddbinning_Fail_Stop)
                                    Exit Do
                                Else
                                    VBIN_RESULT(p_mode).GRADE = DYNAMIC_VBIN_IDS_ZONE(p_mode).Voltage(idx_step)
                                    VBIN_RESULT(p_mode).GRADEVDD = DYNAMIC_VBIN_IDS_ZONE(p_mode).Product_Voltage(idx_step)
                                    next_bin_flag = True
                                End If
                            Loop
                        End If
                        
                        If next_bin_flag = True Then
                            VBIN_RESULT(p_mode).step_in_BinCut = DYNAMIC_VBIN_IDS_ZONE(p_mode).EQ_Num(idx_step) - 1
                            VBIN_RESULT(p_mode).step_in_IDS_Zone = idx_step
                        End If
                    End If '''If CDec(VBIN_RESULT(p_mode).GRADE) > 0
                Next site
            End If '''If BinCut(p_mode, VBIN_RESULT(p_mode).passBinCut).ExcludedPmode = True
        Next i
    Next j
Exit Function
errHandler:
    Call Print_Error_Message(Error_Info, "LIB_VDD_BINNING", "check_voltageInheritance_for_SrampowerDomain") 'Add ErrHandler 2023/08/01
    If AbortTest Then Exit Function Else Resume Next
End Function

''20211130 Rhodes special case VDD_GPU0~VDD_GPUX
Public Function check_cross_voltageInheritance_for_powerDomain(powerDomain As String)
On Error GoTo errHandler 'Add ErrHandler 2023/08/01
    Dim site As Variant
    Dim i As Long, j As Long
    
    Dim core_power_seq() As String
    Dim p_mode As Integer
    Dim gradevdd_last As New SiteDouble
    Dim grade_last As New SiteDouble
    Dim idx_step As Long
    Dim next_bin_flag As Boolean
    Dim voltage_in_dynamic_ids_zone As Double
    Dim exit_while_flag As Boolean
    
    '''20211130 Rhodes Dict dict_cross_domain_compare
    Dim Maxgradevdd() As Double
    Dim product_V As New SiteDouble
    Dim MG_keys As Variant
    Dim temp_MG_key As String  ''mode
    Dim temp_MG_keys() As String ''key's value
    Dim current_mode As Integer
    

    If dict_IsCorePowerInBinCutFlowSheet.Exists(UCase(powerDomain)) = True Then
        If dict_IsCorePowerInBinCutFlowSheet.item(UCase(powerDomain)) = True Then
            core_power_seq = BinCut_Power_Seq(VddBinStr2Enum(powerDomain)).Power_Seq
            
            If UBound(core_power_seq) >= 1 Then
                '''BinCut powerDomain has more than one p_mode, and it's ready to check voltage inheritance for the Power_Seq of BinCut powerDomain.
            Else
                '''If UBound(core_power_seq)=0, skip voltage inheritance check for BinCut powerDomain.
                Exit Function
            End If
        Else
            Call Print_Error_Message(Error_Info, "LIB_VDD_BINNING", "check_cross_voltageInheritance_for_powerDomain", "PowerDomain:" & powerDomain & ",it isn't BinCut CorePower for check_voltageInheritance_for_powerDomain. Error!!!")
            'TheExec.ErrorLogMessage "PowerDomain:" & powerDomain & ",it isn't BinCut CorePower for check_voltageInheritance_for_powerDomain. Error!!!"
            Exit Function
        End If
    Else
        Call Print_Error_Message(Error_Info, "LIB_VDD_BINNING", "check_cross_voltageInheritance_for_powerDomain", "PowerDomain:" & powerDomain & ",it isn't BinCut CorePower for check_voltageInheritance_for_powerDomain. Error!!!")
        'TheExec.ErrorLogMessage "PowerDomain:" & powerDomain & ",it isn't BinCut CorePower for check_voltageInheritance_for_powerDomain. Error!!!"
        Exit Function
    End If
    
    For Each site In TheExec.sites
        MG_keys = dict_cross_domain_compare.Keys
        ReDim Maxgradevdd(UBound(MG_keys))
        For i = 0 To UBound(MG_keys)
            temp_MG_key = dict_cross_domain_compare(MG_keys(i)) '' MGX00~MGX03
            temp_MG_keys = Split(temp_MG_key, ",")

            For j = 0 To UBound(temp_MG_keys) '"MG001,MG101,MG201,MG301"
                current_mode = VddBinStr2Enum(temp_MG_keys(j))
                product_V = CDec(VBIN_RESULT(current_mode).GRADEVDD)
                Maxgradevdd(i) = Max(Maxgradevdd(i), product_V(site))
                
                If i > 0 Then                ''First cycle don't compare max
                    If CDec(VBIN_RESULT(current_mode).GRADE) > 0 Then
                        idx_step = VBIN_RESULT(current_mode).step_in_IDS_Zone
                        exit_while_flag = False
       
                        '''//Check if GradeVDD(p_mode) => GradeVDD(previous_performance_mode) + Monotonicity_Offset(p_mode).
                        While (CDec(VBIN_RESULT(current_mode).GRADEVDD) <= CDec(Maxgradevdd(i - 1)) And exit_while_flag = False)
                            idx_step = idx_step + 1
                            
                            If DYNAMIC_VBIN_IDS_ZONE(current_mode).passBinCut(idx_step) <> CurrentPassBinCutNum Then
                                '''20210610: Modified to revise the message for the vbt function find_next_bin_eq_interpolation if it can't find any step to define the product voltage.
                                Call Print_Error_Message(Error_Info, "LIB_VDD_BINNING", "check_cross_voltageInheritance_for_powerDomain", "site:" & site & "," & temp_MG_keys(j) & ",bin=" & CurrentPassBinCutNum(site) & ",but it can't find any step to adjust the product voltage for voltage inheritance check. Error!!!")
                                '''20210610: Modified to bin out the failed site in adjust_vddbinning.
                                exit_while_flag = True
                                '''Warning!!!please check SortNumber and binNumber of Flag_Vddbinning_Fail_Stop in Bin_Table before using this.
                                '''20240303: Use modular function to bin out site
                                Call BinoutSite(site, VddbinFailStopSortNumber, VddbinFailStopBinNumber, strGlb_Flag_Vddbinning_Fail_Stop)
                                
                                '''210701 prevent site0 fail then other site skip to check because exit function
'                               Exit Function
                            Else
                                voltage_in_dynamic_ids_zone = DYNAMIC_VBIN_IDS_ZONE(current_mode).Voltage(idx_step)
                                VBIN_RESULT(current_mode).GRADE = voltage_in_dynamic_ids_zone
                                VBIN_RESULT(current_mode).GRADEVDD = voltage_in_dynamic_ids_zone + BinCut(current_mode, VBIN_RESULT(current_mode).passBinCut).CP_GB(DYNAMIC_VBIN_IDS_ZONE(current_mode).EQ_Num(idx_step) - 1)
                                next_bin_flag = True
                            End If
                        Wend
'
                        If next_bin_flag = True Then
                            VBIN_RESULT(current_mode).step_in_BinCut = DYNAMIC_VBIN_IDS_ZONE(current_mode).EQ_Num(idx_step) - 1
                            VBIN_RESULT(current_mode).step_in_IDS_Zone = idx_step
                        End If
                    End If '''If CDec(VBIN_RESULT(p_mode).GRADE) > 0
                End If
            Next j
        Next i
    Next
    
Exit Function
errHandler:
    Call Print_Error_Message(Error_Info, "LIB_VDD_BINNING", "check_cross_voltageInheritance_for_powerDomain") 'Add ErrHandler 2023/08/01
    If AbortTest Then Exit Function Else Resume Next
End Function


'20210909: Modified to merge the branches of the vbt function run_patt_offline_simulation.
'20201210: Modified to remove the redundant branch "If siteResult_Offline(site) = False Then".
'20201118: Modified to remove the redundant argument "offline_flag_patallpass As Boolean".
'20200922: Created to run Pattern offline simulation for GradeSearch_HVCC_VT / GradeSearch_postBinCut_VT / run_patt_only_VT.
'20200730: Modified to add the EnableWord "VDDBinning_Offline_AllPattPass" for Offline simulation with all patterns pass.
'20200106: As per discussion with SWLINZA, he suggested us to add this to check any alarm.
Public Function run_patt_offline_simulation(patt_selected As String, result_mode As tlResultMode, siteResult_Offline As SiteBoolean)
On Error GoTo errHandler 'Add ErrHandler 2023/08/01
    Dim site As Variant
'''//==================================================================================================================================================================================//'''
'''//Note:
'''It can use the EnableWord "VDDBinning_Offline_AllPattPass" for Offline simulation with all patterns pass.
'''//==================================================================================================================================================================================//'''
    If Flag_VDD_Binning_Offline = True Then '''Offline test.
        '''//Generate offline simulation with random pattern Pass/Fail.
        '''Note: It can use the EnableWord "VDDBinning_Offline_AllPattPass" for Offline simulation with all patterns pass.
        If TheExec.enableWord("VDDBinning_Offline_AllPattPass") = True Or TheExec.enableWord("Golden_Default") = True Then
            siteResult_Offline = True
        Else
            For Each site In TheExec.sites
                siteResult_Offline(site) = IIf(Round(WorksheetFunction.Min(1, Rnd * 8), 0) = 1, True, False)
            Next site
        End If
        
        '''//Run the pattern by offline simulation random pattern Pass/Fail.
        Call TheHdw.patterns(patt_selected).test(pfNever, 0, result_mode)
        
        '''20210909: Modified to merge the branches of the vbt function run_patt_offline_simulation.
        For Each site In TheExec.sites
            If siteResult_Offline(site) = False Then
                Call TheExec.Datalog.WriteFunctionalResult(site, TheExec.sites.item(site).TestNumber, logTestFail)
            Else
                Call TheExec.Datalog.WriteFunctionalResult(site, TheExec.sites.item(site).TestNumber, logTestPass)
            End If
        Next
    Else
        Call Print_Error_Message(Error_Info, "LIB_VDD_BINNING", "run_patt_offline_simulation", "Flag_VDD_Binning_Offline is " & CStr(Flag_VDD_Binning_Offline) & ". It is incorrect to use run_patt_offline_simulation. Error!!!")
        'TheExec.ErrorLogMessage "Flag_VDD_Binning_Offline is " & CStr(Flag_VDD_Binning_Offline) & ". It is incorrect to use run_patt_offline_simulation. Error!!!"    ''20230807: Removed by sanity check
    End If
Exit Function
errHandler:
    Call Print_Error_Message(Error_Info, "LIB_VDD_BINNING", "run_patt_offline_simulation") 'Add ErrHandler 2023/08/01
'    TheExec.ErrorLogMessage "Error encountered in VBT Function of run_FuncPat_and_check_PassFail"
    If AbortTest Then Exit Function Else Resume Next
End Function

'20201125: Modified to replace the argument "pattPass as SiteBoolean" with "siteResult As SiteBoolean".
'20201118: Modified to remove the redundant arguments "Optional is_VddBinning_offline As Boolean = False" and "Optional offline_pat_status As SiteBoolean".
'20201118: Created to check alarmFail for pattern.
Public Function check_alarmFail_for_pattern(siteResult As SiteBoolean)
On Error GoTo errHandler 'Add ErrHandler 2023/08/01
    Dim site As Variant

    If glb_TesterType = "UltraFLEXplus" Then
        Dim alarmOccurred As New SiteBoolean
        alarmOccurred = False
        alarmOccurred = TheHdw.Alarms.GetAlarmingSites(True)        ''GetAlarmingSites(clearAlarm = True), only plus or > 10_30_90 can use
        For Each site In TheExec.sites
            If alarmOccurred(site) = True Then
                TheExec.Datalog.WriteComment "site:" & site & ", alarmFail!!!"
                siteResult(site) = False
            End If
        Next site
    Else
    For Each site In TheExec.sites
        '''//Check if alarmFail(site) is triggered or not.
        If alarmFail(site) = True Then
            TheExec.Datalog.WriteComment "Site:" & site & ", alarmFail!!!"
            siteResult(site) = False
        End If
    Next site
    End If
Exit Function
errHandler:
    Call Print_Error_Message(Error_Info, "LIB_VDD_BINNING", "check_alarmFail_for_pattern") 'Add ErrHandler 2023/08/01
    If AbortTest Then Exit Function Else Resume Next
End Function

'20201125: Modified to replace the argument "PrePattPass as SiteBoolean" with "siteResult As SiteBoolean".
'20201015: Modified to rename "check_PrePattPass_for_PattPass" as "update_Pattern_result_to_PattPass".
'20201012: Created to update PrePattPass to PattPass.
Public Function update_Pattern_result_to_PattPass(siteResult As SiteBoolean, pattPass As SiteBoolean)
On Error GoTo errHandler 'Add ErrHandler 2023/08/01
    Dim site As Variant
    For Each site In TheExec.sites
        If pattPass(site) = False Or siteResult(site) = False Then
            pattPass(site) = False
        End If
    Next site
Exit Function
errHandler:
    Call Print_Error_Message(Error_Info, "LIB_VDD_BINNING", "update_Pattern_result_to_PattPass") 'Add ErrHandler 2023/08/01
    If AbortTest Then Exit Function Else Resume Next
End Function

'20210903: Modified to merge properties of "Public Type Instance_Step_Control" into Public Type Instance_Info "Public Type Instance_Info".
'20210809: Modified to remove the redundant property "FoundLevel As New SiteDouble" from Public Type Instance_Step_Control.
'20201210: Modified to use the arguments "inst_info As Instance_Info" and "step_control As Instance_Step_Control" for update_control_flag_for_patt_loop.
'20200923: Created to update the status of "AllSiteFailPatt" and "All_Patt_Pass".
'20200923: Modified to merge the vbt blocks of "AllSiteFailPatt" and "All_Patt_Pass".
'20200923: Modified to remove the unused condition from the branch of "AllSiteFailPatt" and "All_Patt_Pass".
'20200922: Modified to update the status of "AllSiteFailPatt".
Public Function update_control_flag_for_patt_loop(inst_info As Instance_Info, pattPass As SiteBoolean)
On Error GoTo errHandler 'Add ErrHandler 2023/08/01
    Dim site As Variant
'''//==================================================================================================================================================================================//'''
'''//Note:
'''Loop each pattern in the pattern set, and use flag for all site to record if the site had failed.
'''//==================================================================================================================================================================================//'''
'''============================================
''' EQ3 Site0  Site1  Site2
''' loop 5 patterns Start
''' Patt1    F  P  P
''' Patt2    P  F  F
''' Patt3      skip and go to next level
''' Patt4    do not need to test
''' Patt5    do not need to test
'''loop 5 patterns End
'''============================================
    '''//Update the status of "AllSiteFailPatt" and "All_Patt_Pass".
    For Each site In TheExec.sites
        If pattPass(site) = False Then
            '''======================================================================================
            ''' Site 0 pass, Site 1 fail, site 2 fail =>  AllSiteFailPatt = 2^1 or 2^2 = 6.
            ''' Site 0 fail, Site 1 fail, site 2 fail =>  AllSiteFailPatt = 2^0 or 2^1 or 2^2 = 7.
            '''======================================================================================
            inst_info.AllSiteFailPatt = inst_info.AllSiteFailPatt Or 2 ^ site
            
            '=======================================================================================================
            ' If this site had found the grade, but the pattern is failed this step,
            ' we will clear all result and define the site is not found the grade yet, the situation is shmoo hole.
            '  Site0   Site1   Site2
            '  EQ4(F)  EQ4(F)  EQ4(P)
            '  EQ3(F)  EQ3(P)  EQ3(F) => the EX. for site2 EQ3(F)
            '  EQ2(P)  EQ2(P)  EQ2(P)
            '
            '  Grade   Grade   Grade
            '  EQ2     EQ3     EQ2
            '=======================================================================================================
            inst_info.All_Patt_Pass(site) = False
            
            '''//Check if Grade_Found but pattern fails. It means DUT has the shmoo hole.
            '''20210809: Modified to remove the redundant property "FoundLevel As New SiteDouble" from Public Type Instance_Step_Control.
            If inst_info.grade_found(site) = True And inst_info.gradeAlg = GradeSearchAlgorithm.linear Then
                inst_info.grade_found(site) = False '''Shmoo hole: lower VCC passes, but higher VCC fails.
                VBIN_RESULT(inst_info.p_mode).GRADE = 0
                VBIN_RESULT(inst_info.p_mode).GRADEVDD = 0
            End If
        End If
    Next site
Exit Function
errHandler:
    Call Print_Error_Message(Error_Info, "LIB_VDD_BINNING", "update_control_flag_for_patt_loop") 'Add ErrHandler 2023/08/01
    If AbortTest Then Exit Function Else Resume Next
End Function

'20210819: Modified to revise the vbt code for the new format of the header in BinCut voltage tables, requested by C651 Toby.
'20210526: Modified to add "Flag_Get_column_Monotonicity_Offset" for Monotonicity_Offset check because C651 Si revised the check rules.
'20210429: Modified to remove the globalVariable "Flag_Using_Montonicity_Offset".
'20210427: Modified to parse the column of "Monotonicity_Offset".
'20210325: Modified to use the 1-dimension array to store SRAM_Vth.
'20210322: Discussed the vbt code that checked if int_Offset is multiple of StepVoltage, all project BinCut owners decided to remove the vbt code because this was the redundant action.
'20210312: Modified to parse columns of "Softbin" and "HardBin" from column col_sort+1.
'20210305: Modified to set INTP_MODE_L, INTP_MODE_L, and AllowEqual if the cells are empty.
'20210223: Modified to replace "Dim step As Long" with "Dim idx_step As Long".
'20201021: Modified to use "dict_IsCorePower" to store and check CorePower/OtherRail.
'20200824: Modified to check TotalStepPerMode. Revised by Leon Weng.
'20200703: Modiifed to use "check_Sheet_Range".
'20200501: Modified to use "AllBinCut(p_mode).INTP_SKIPTEST".
'20200427: Modified to move "Flag_Interpolation_enable" from "ReGenerate_IDS_ZONE_Voltage_Per_Site_ver2" to "initVddBinning".
'20200421: Modified to check the column of "Allow Equal".
'20200415: Modified to check "col_soft_bin".
'20200331: Modified to check if int_Offset is multiple of 3.125.
'20200330: Modified to parse "Int_Offset","Int_SkipTest" for interpolation.
'20200206: Modified to check if CPVmin, CPVmax, and CPGB are multiple of 3.125.
'20191219: Modified to check powerDomain in the vbt function "initDomain2Pin".
'20191127: Modified for the revised InitVddBinTable.
'20191113: Modified to check pmode, allowEqual, and MaxPV/MinPV.
'20191023: Modified to check if "MaxPV(pmode0/pmode1)" is in the column "Comment" or not.
'20191014: Modified to parse the table with different powerPins when different testjobs.
'20191001: Modified for the new header defined by C651, they separated IDS_limit into "CPIDSMax" and "IDSMax_HOT"
'20190706: Modified to check if the powerPin (Domain) is in the pin_group "FullCorePowerinFlowSheet " or not.
'20190426: Modified to use the function "Find_Sheet".
'20190321: Modified for checking if powerpin exists in pinmap and channelmap.
'20190312: Modified for adding powerpins into power_group "FullCorePowerinFlowSheet".
'20180221: Anderson enhanced this for VddBinDef & OtherRail combine together.
Public Function initVddBinTableOneMod(passBinCut As Long, col_ids As Integer, col_sort As Integer)
On Error GoTo errHandler 'Add ErrHandler 2023/08/01
    Dim wb As Workbook
    Dim ws_def As Worksheet
    Dim sheetName As String
    Dim site As Variant
    Dim strAry_Temp() As String
    Dim Row As Long, Col As Long
    Dim main_p_mode As Integer
    Dim idx_step As Long
    Dim Test_Type As Long
    Dim col_binned As Integer
    Dim col_domain As Integer
    Dim col_mode As Integer
    Dim col_eqn As Integer
    Dim col_id As Integer
    Dim col_eqn_bin As Integer
    Dim col_binx_ids As Integer
    Dim col_binx_ftids As Integer
    Dim col_c As Integer
    Dim col_m As Integer
    Dim col_cpids As Integer
    Dim col_ftids As Integer
    Dim col_cp_vmax As Integer
    Dim col_cp_vmin As Integer
    Dim col_montonicityoffset As Integer '''Monotonicity_Offset
    Dim col_cpgb As Integer
    Dim col_cp2gb As Integer
    Dim col_ft1gb As Integer
    Dim col_ft2gb As Integer
    Dim col_sltgb As Integer
    Dim col_htol_ro_gb As Integer
    Dim col_htol_ro_gb_room As Integer
    Dim col_htol_ro_gb_hot As Integer
    Dim col_ate_ftqa_gb As Integer
    Dim col_slt_ftqa_gb As Integer
    Dim col_cphv As Integer
    Dim col_fthv As Integer
    Dim col_qahv As Integer
    Dim col_intModeL As Integer '''for start p_mode of interpolation
    Dim col_intModeH As Integer '''for end p_mode of interpolation
    Dim col_intMFactor As Integer '''for factor of interpolation
    Dim col_intOffset As Integer
    Dim col_intSkipTest As Integer
    Dim col_allow_equal As Integer
    Dim col_comment As Integer
    Dim col_ids_mapping As Integer              ''' 20211202: For CPU_SRAM IDS Mapping,
    Dim col_ids_mapping_max As Integer          ''' 20211202: Determines which IDS value to use and to calculate C-M*log10(IDS)
    Dim col_sram_vt_spec(1) As Integer '''SRAM_VTH_SPEC(0): for CP1 BV binSearch and postBinCut/OutsideBinCut, SRAM_VTH_SPEC(1): for CP1 HBV and non-CP1 BV/HBV.
    Dim col_dynamic_offset() As Integer 'Dim col_dynamic_offset(MaxJobCountInVbt, MaxTestType) As Integer 20230526
    Dim strTemp As String
    Dim split_content() As String
    Dim p_mode As Integer
    Dim jobIdx As Integer, testTypeIdx As Integer
    Dim powerDomain As String
    Dim MaxRow As Long
    Dim maxcol As Long
    Dim Row_of_Title As Integer
    Dim enableRowParsing As Boolean
    Dim isSheetFound As Boolean
    Dim idx_step_temp As Long
    '''*****************************************************************'''
    '''//Check if the sheet exists
    sheetName = "Vdd_Binning_Def_appA_" & passBinCut & "_" & TheExec.CurrentJob
    Set wb = Application.ActiveWorkbook
    Call check_Sheet_Range(sheetName, wb, ws_def, MaxRow, maxcol, isSheetFound, False)
    ReDim col_dynamic_offset(BinCutJobDefinition.Total_Count, testType.Total_Count) As Integer
    If isSheetFound = False Then                ''' if "Vdd_Binning_Def_appA_2/3_JOBNAME" does not exist
        sheetName = "Vdd_Binning_Def_appA_" & passBinCut
        Call check_Sheet_Range(sheetName, wb, ws_def, MaxRow, maxcol, isSheetFound)
    End If
    '' 20211224: Added to Print Bincut Sheet Name in Print_Bincut_Config
    If LCase(sheetName) Like "*vdd_binning_def_appa_1*" Then
        gS_Vdd_Binning_Def_appA_1 = sheetName
    ElseIf LCase(sheetName) Like "*vdd_binning_def_appa_2*" Then
        gS_Vdd_Binning_Def_appA_2 = sheetName
    ElseIf LCase(sheetName) Like "*vdd_binning_def_appa_3*" Then
        gS_Vdd_Binning_Def_appA_3 = sheetName
    End If
    '''*****************************************************************'''
    MaxPerformanceModeCount = cntVddbinPmode + 1 + cntEqVoltagePmode '20230526
    ReDim ExcludedPmode(MaxPerformanceModeCount) As Boolean 'Public ExcludedPmode(MaxPerformanceModeCount) As Boolean
    ReDim col_dynamic_offset(BinCutJobDefinition.Total_Count, testType.Total_Count) As Integer 'Dim col_dynamic_offset(MaxJobCountInVbt, MaxTestType) As Integer 20230526
    If isSheetFound = True Then
        '''//init
        '''Since all col_XXX and row_XXX related variables with default values=0, no need to initialize them as 0.
        Flag_Adjust_Max_Enable = False
        Flag_Adjust_Min_Enable = False
        Adjust_Power_Max_pmode = vbNullString
        Adjust_Power_Min_pmode = vbNullString
        '20230504: Added for parsing sheet defining BinX EQNs and IDS value
        Pmode2BinX_EQN_Dict.RemoveAll
        Pmode2BinX_IDS_Dict.RemoveAll
        Pmode2BinX_IDSHot_Dict.RemoveAll
        
'        If Total_Bincut_Num < passBinCut Then                   'capture the max BinCut number
'            Total_Bincut_Num = passBinCut
'        End If
        
        For p_mode = 0 To MaxPerformanceModeCount - 1           'initilize the MODE_STEP and ExcludedPmode
            BinCut(p_mode, passBinCut).Mode_Step = -99
            BinCut(p_mode, passBinCut).ExcludedPmode = True     'If you do not assign to True , the default value is False
            ExcludedPmode(p_mode) = True
        Next p_mode
        
        For Row = 1 To MaxRow
            For Col = 1 To maxcol
                '''******************************************************************************************************************'''
                '''//If CorePower and OtherRail are in the same table (only Vdd_Binning_Def), 1st column is "Binned".
                '''//If CorePower and OtherRail are in the different tables (Vdd_Binning_Def and Other_Rail), 1st column is "Domain".
                '''******************************************************************************************************************'''
                '''If 1st column 1 of the header is "Binned", split the line and find out the keyword column.
                If LCase(ws_def.Cells(Row, Col).value) Like "binned" Then
                    col_binned = Col
                    Row_of_Title = Row
                End If
            
                If Row_of_Title > 0 Then
                    If LCase(ws_def.Cells(Row_of_Title, Col).value) = "domain" Then
                        col_domain = Col
                    ElseIf LCase(ws_def.Cells(Row_of_Title, Col).value) = "mode" Then
                        col_mode = Col
                    ElseIf LCase(ws_def.Cells(Row_of_Title, Col).value) = "id" Then
                        col_id = Col
                    ElseIf LCase(ws_def.Cells(Row_of_Title, Col).value) = "eqn" Then
                        col_eqn = Col
                    '20230504: Added for parsing sheet defining BinX EQNs and IDS value
                    ElseIf LCase(ws_def.Cells(Row_of_Title, Col).value) = "eqn_bin" Then
                        col_eqn_bin = Col
                    ElseIf LCase(ws_def.Cells(Row_of_Title, Col).value) = "binx_cpidsmax" Then
                        col_binx_ids = Col
                    ElseIf LCase(ws_def.Cells(Row_of_Title, Col).value) = "binx_idsmax_hot" Then
                        col_binx_ftids = Col
                    ElseIf LCase(ws_def.Cells(Row_of_Title, Col).value) = "c" Then
                        col_c = Col
                    ElseIf LCase(ws_def.Cells(Row_of_Title, Col).value) = "m" Then
                        col_m = Col
                    '''********************************************************'''
                    ElseIf LCase(ws_def.Cells(Row_of_Title, Col).value) = "cpidsmax" Then
                        col_cpids = Col
                    ElseIf LCase(ws_def.Cells(Row_of_Title, Col).value) = "idsmax_hot" Or LCase(ws_def.Cells(Row_of_Title, Col).value) = "ftids" Then
                        col_ftids = Col
                    '''********************************************************'''
                    '''20210819: Modified to revise the vbt code for the new format of the header in BinCut voltage tables, requested by C651 Toby.
                    ElseIf LCase(ws_def.Cells(Row_of_Title, Col).value) = "cpvmax" Or LCase(ws_def.Cells(Row_of_Title, Col).value) = LCase("BinningVmax") Then
                        col_cp_vmax = Col
                    ElseIf LCase(ws_def.Cells(Row_of_Title, Col).value) = "cpvmin" Or LCase(ws_def.Cells(Row_of_Title, Col).value) = LCase("BinningVmin") Then
                        col_cp_vmin = Col
                    ElseIf LCase(ws_def.Cells(Row_of_Title, Col).value) = "cpgb" Or LCase(ws_def.Cells(Row_of_Title, Col).value) = LCase("BinningGB") Then
                        col_cpgb = Col
                    ElseIf LCase(ws_def.Cells(Row_of_Title, Col).value) = "cp2gb" Or LCase(ws_def.Cells(Row_of_Title, Col).value) = LCase("CP_GB_HOT") Then
                        col_cp2gb = Col
                    ElseIf LCase(ws_def.Cells(Row_of_Title, Col).value) = "ft1gb" Or LCase(ws_def.Cells(Row, Col).value) = "ft_gb_room" Then
                        col_ft1gb = Col
                    ElseIf LCase(ws_def.Cells(Row_of_Title, Col).value) = "ft2gb" Or LCase(ws_def.Cells(Row, Col).value) = "ft_gb_hot" Then
                        col_ft2gb = Col
                    ElseIf LCase(ws_def.Cells(Row_of_Title, Col).value) = "sltgb" Then
                        col_sltgb = Col
                    ElseIf LCase(ws_def.Cells(Row_of_Title, Col).value) = "htol_ro_gb" Then
                        col_htol_ro_gb = Col
                    ElseIf LCase(ws_def.Cells(Row_of_Title, Col).value) = "htol_ro_gb_room" Or LCase(ws_def.Cells(Row_of_Title, Col).value) = LCase("HTOL_T0TX_GB_ROOM") Then
                        col_htol_ro_gb_room = Col
                    ElseIf LCase(ws_def.Cells(Row_of_Title, Col).value) = "htol_ro_gb_hot" Or LCase(ws_def.Cells(Row_of_Title, Col).value) = LCase("HTOL_T0TX_GB_HOT") Then
                        col_htol_ro_gb_hot = Col
                    ElseIf LCase(ws_def.Cells(Row_of_Title, Col).value) = "ate_fqagb" Then
                        col_ate_ftqa_gb = Col
                    ElseIf LCase(ws_def.Cells(Row_of_Title, Col).value) = "slt_fqa_gb" Then
                        col_slt_ftqa_gb = Col
                    ElseIf LCase(ws_def.Cells(Row_of_Title, Col).value) = "cphv" Then
                        col_cphv = Col
                    ElseIf LCase(ws_def.Cells(Row_of_Title, Col).value) = "fthv" Then
                        col_fthv = Col
                    ElseIf LCase(ws_def.Cells(Row_of_Title, Col).value) = "qahv" Then
                        col_qahv = Col
                    ElseIf LCase(ws_def.Cells(Row_of_Title, Col).value) Like "offset_*_*" Then
                        jobIdx = getBinCutJobDefinition(LCase(ws_def.Cells(Row_of_Title, Col).value))
                        testTypeIdx = decide_offset_testType(LCase(ws_def.Cells(Row_of_Title, Col).value))
                        col_dynamic_offset(jobIdx, testTypeIdx) = Col
                    '''Allow_Equal
                    ElseIf LCase(ws_def.Cells(Row_of_Title, Col).value) = "allow equal" Then
                        col_allow_equal = Col
                    '''interpolation
                    ElseIf LCase(ws_def.Cells(Row_of_Title, Col).value) = "int_mode_l" Then
                        col_intModeL = Col
                    ElseIf LCase(ws_def.Cells(Row_of_Title, Col).value) = "int_mode_h" Then
                        col_intModeH = Col
                    ElseIf LCase(ws_def.Cells(Row_of_Title, Col).value) = "int_mf" Then
                        col_intMFactor = Col
                    ElseIf LCase(ws_def.Cells(Row_of_Title, Col).value) = "int_offset" Then
                        col_intOffset = Col
                    ElseIf LCase(ws_def.Cells(Row_of_Title, Col).value) = "int_skiptest" Then
                        col_intSkipTest = Col
                    '''Monotonicity
                    ElseIf LCase(ws_def.Cells(Row_of_Title, Col).value) = "monotonicity_offset" Then '''20210427: Modified to parse the column of "Monotonicity_Offset".
                        col_montonicityoffset = Col
                        '''20210526: Modified to add "Flag_Get_column_Monotonicity_Offset" for Monotonicity_Offset check because C651 Si revised the check rules.
                        Flag_Get_column_Monotonicity_Offset = True
                    '''SRAM_Vth
                    ElseIf LCase(ws_def.Cells(Row_of_Title, Col).value) Like "sramthresh*" Then
                        '''************************************************************************************************************************'''
                        '''column "SRAMthresh_CP1" or "SRAMthresh_BinSearch"==> SRAM_VTH_SPEC(0): for BinCut search.
                        '''column "SRAMthresh_Product"                      ==> SRAM_VTH_SPEC(1): for BinCut check.
                        '''************************************************************************************************************************'''
                        '''20210325: Modified to use the 1-dimension array to store SRAM_Vth.
                        If LCase(ws_def.Cells(Row_of_Title, Col).value) Like "*cp1" Or LCase(ws_def.Cells(Row_of_Title, Col).value) Like "*binsearch" Then
                            col_sram_vt_spec(0) = Col
                        ElseIf LCase(ws_def.Cells(Row_of_Title, Col).value) Like "*product" Then
                            col_sram_vt_spec(1) = Col
                        Else
                            enableRowParsing = False
                            Call Print_Error_Message(Error_Info, "LIB_VDD_BINNING", "initVddBinTableOneMod", ws_def.Cells(Row_of_Title, Col).value & " is the undefined column in the sheet:" & sheetName & ". Error!!!")
                            'TheExec.Datalog.WriteComment ws_def.Cells(Row_of_Title, Col).value & " is the undefined column in the sheet:" & sheetName & ". Error!!!"
                            'TheExec.ErrorLogMessage ws_def.Cells(Row_of_Title, Col).value & " is the undefined column in the sheet:" & sheetName & ". Error!!!"    ''20230807: Removed by sanity check
                            Exit For
                        End If
                    ''' 20211202: For CPU_SRAM IDS Mapping, determines which IDS value to use and to calculate C-M*log10(IDS)
                    ElseIf LCase(ws_def.Cells(Row_of_Title, Col).value) Like "*bincutcalc_idsrail*" Then
                        Flag_IDS_Mapping_enable = True          '' when col = "bincutcalc_idsrail", Turn the flag true
                        col_ids_mapping = Col
                    ElseIf LCase(ws_def.Cells(Row_of_Title, Col).value) Like "bincutcalc_idsmax*" Then
                        col_ids_mapping_max = Col
                    ElseIf LCase(ws_def.Cells(Row_of_Title, Col).value) Like "*binx_bincutcalc_idsmax*" Then
                        If Flag_BinX_Info_Parsed = True Then col_ids_mapping_max = Col
                    ''' ****************************************** IDS MAPPING ****************************************** '''
                        
                    ElseIf LCase(ws_def.Cells(Row, Col).value) = "comment" Then
                        col_comment = Col
                        
                        '''//Check if column of "Softbin" is next to column "Comment".
                        If col_comment <> col_sort - 1 Then
                            col_comment = 0
                            Call Print_Error_Message(Error_Info, "LIB_VDD_BINNING", "initVddBinTableOneMod", "col_soft_bin " & col_sort & " doesn't match the start column of sort bin in " & sheetName & ". Error!!!")
                            'TheExec.ErrorLogMessage "col_soft_bin " & col_sort & " doesn't match the start column of sort bin in " & sheetName & ". Error!!!"    ''20230807: Removed by sanity check
                        End If
                    End If
                End If
                
                '''//Check if all columns of the header exist...
                '''Note: col_comment should be checked as the last available column of the table.
                If col_domain > 0 And col_mode > 0 And col_id > 0 And col_eqn > 0 And col_c > 0 And col_m > 0 _
                And col_cpids > 0 And col_ftids > 0 And col_cp_vmax > 0 And col_cp_vmin > 0 _
                And col_cpgb > 0 And col_cp2gb And col_ft1gb > 0 And col_ft2gb > 0 _
                And col_sltgb > 0 And col_ate_ftqa_gb > 0 And col_slt_ftqa_gb > 0 _
                And col_cphv > 0 And col_fthv > 0 And col_qahv > 0 _
                And col_intModeL > 0 And col_intModeH > 0 _
                And col_allow_equal > 0 And col_comment > 0 Then
                    enableRowParsing = True
                End If
            Next Col
            
            '''//If all columns of the header are found, skip the loop and start parsing each row.
            If enableRowParsing = True Then
                Exit For
            End If
            
            If Row = MaxRow And (col_domain = 0 Or col_mode = 0 Or col_cpids = 0 Or col_ftids = 0 Or col_allow_equal = 0 Or col_comment = 0) Then
                enableRowParsing = False
                Call Print_Error_Message(Error_Info, "LIB_VDD_BINNING", "initVddBinTableOneMod", "Columns of header in " & sheetName & " are incorrect. Error!!!")
                'TheExec.ErrorLogMessage "Columns of header in " & sheetName & " are incorrect. Error!!!"    ''20230807: Removed by sanity check
            End If
        Next Row
        
        '20230504: Added for parsing sheet defining BinX EQNs and IDS value
        Dim i As Long
        If col_eqn_bin > 0 And col_binx_ids > 0 And col_binx_ftids > 0 And LCase(ws_def.Cells(Row_of_Title + 1, col_binx_ids).value) <> vbNullString And LCase(ws_def.Cells(Row_of_Title + 1, col_binx_ftids).value) <> vbNullString Then
            Flag_BinX_Info_Parsed = True
            For i = 0 To UBound(AllBinCut)
                AllBinCut(i).Last_Bin1_EQN = 999
            Next i
        Else
            Flag_BinX_Info_Parsed = False
            For i = 0 To UBound(AllBinCut)
                AllBinCut(i).Last_Bin1_EQN = 1
            Next i
        End If
        
        If enableRowParsing = True And Row_of_Title + 1 <= MaxRow Then
            For Row = Row_of_Title + 1 To MaxRow
                '''//If first word in the mode column is M(ex: MC601).
                '''//If column "Binned" is "true", it means that performance_mode of powerDomain is the binning mode of CorePower.
                If ws_def.Cells(Row, col_mode).value Like "M*" And LCase(ws_def.Cells(Row, col_binned).value) = "true" Then
                    '''//performance_mode is enumerated into the p_mode dictionary when "initVddBinTable".
                    main_p_mode = VddBinStr2Enum(ws_def.Cells(Row, col_mode))
                    
                    If UCase(ws_def.Cells(Row, col_domain).value) Like "VDD*" Then
                        powerDomain = UCase(Trim(ws_def.Cells(Row, col_domain)))
                        AllBinCut(main_p_mode).powerPin = powerDomain
                    ElseIf UCase(ws_def.Cells(Row, col_domain).value) <> "" Then
                        powerDomain = "VDD_" & UCase(Trim(ws_def.Cells(Row, col_domain)))
                        AllBinCut(main_p_mode).powerPin = powerDomain
                    Else
                        Call Print_Error_Message(Error_Info, "LIB_VDD_BINNING", "initVddBinTableOneMod", ws_def.Cells(Row, col_domain) & " doesn't have the correct Domain cell in sheet " & sheetName & ". Error!!!")
                        'TheExec.Datalog.WriteComment ws_def.Cells(Row, col_domain) & " doesn't have the correct Domain cell in sheet " & sheetName & ". Error!!!"
                        'TheExec.ErrorLogMessage ws_def.Cells(Row, col_domain) & " doesn't have the correct Domain cell in sheet " & sheetName & ". Error!!!"    ''20230807: Removed by sanity check
                    End If
                    
                    '''//Use "dict_IsCorePower" to check if powerDomain is BinCut CorePower/OtherRail listed in Vdd_Binning_Def_appA_1.
                    If dict_IsCorePower.Exists(UCase(powerDomain)) = True Then
                        If LCase(ws_def.Cells(Row, col_binned).value) = LCase(CStr(dict_IsCorePower.item(UCase(powerDomain)))) Then
                            '''Do nothing...
                        Else
                            Call Print_Error_Message(Error_Info, "LIB_VDD_BINNING", "initVddBinTableOneMod", "column Binned of " & ws_def.Cells(Row, col_domain) & "  in sheet " & sheetName & " isn't consistent with Vdd_Binning_Def sheet_appA_1. Error!!!")
                            'TheExec.ErrorLogMessage "column Binned of " & ws_def.Cells(Row, col_domain) & "  in sheet " & sheetName & " isn't consistent with Vdd_Binning_Def sheet_appA_1. Error!!!"    ''20230807: Removed by sanity check
                        End If
                    Else
                         Call Print_Error_Message(Error_Info, "LIB_VDD_BINNING", "initVddBinTableOneMod", ws_def.Cells(Row, col_domain) & "  in sheet " & sheetName & " doesn't show in other Vdd_Binning_Def sheet. Error!!!")
                        'TheExec.Datalog.WriteComment ws_def.Cells(Row, col_domain) & "  in sheet " & sheetName & " doesn't show in other Vdd_Binning_Def sheet. Error!!!"
                        'TheExec.ErrorLogMessage ws_def.Cells(Row, col_domain) & "  in sheet " & sheetName & " doesn't show in other Vdd_Binning_Def sheet. Error!!!"    ''20230807: Removed by sanity check
                    End If
                    
                    '''//Check if the p_mode is excluded Pmode...
                    '''if the bincut total passbin number is more than 1, and if the performance mode doesn't exist in BinCut 1 but exists in Bin2, it has the error.
                    If passBinCut > 1 And BinCut(main_p_mode, passBinCut - 1).ExcludedPmode = True Then
                        BinCut(main_p_mode, passBinCut).ExcludedPmode = False
                        ExcludedPmode(main_p_mode) = False
                        TheExec.Datalog.WriteComment "Test performance Mode " & VddBinName(main_p_mode) & " do not exist in BinCut " & passBinCut - 1
                        TheExec.ErrorLogMessage "Test Performance Mode " & VddBinName(main_p_mode) & " do not exist in BinCut " & passBinCut - 1
                    Else
                        BinCut(main_p_mode, passBinCut).ExcludedPmode = False
                        ExcludedPmode(main_p_mode) = False
                    End If
                    
                    If ws_def.Cells(Row, col_eqn).value Like "E#*" Then '''read the E1 ~ En
                        strAry_Temp = Split(ws_def.Cells(Row, col_eqn), "E") '''ex: array(0)=E ; array(1)=1
                        idx_step = CLng(strAry_Temp(1)) - 1 '''step: the address for store the EQ number, ex: BinCut(P_mode,passbinnum).EQ_Num(0)=1, step = 0, EQ = 1
                        BinCut(main_p_mode, passBinCut).EQ_Num(idx_step) = CLng(strAry_Temp(1))

                        '''20200824: Modified to check TotalStepPerMode. Revised by Leon Weng.
'                        If BinCut(main_p_mode, passBinCut).EQ_Num(idx_step) > Max_IDS_Step Then
                            'Call Print_Error_Message(Error_Info, "LIB_VDD_BINNING", "initVddBinTableOneMod", sheetName & ", p_mode:" & ws_def.Cells(row, col_mode).value & ", EQ num:" & BinCut(main_p_mode, passBinCut).EQ_Num(idx_step) & " is greater than global variable (TotalStepPerMode:" & CStr(TotalStepPerMode) & "), please check it. Error!!!")
                            'TheExec.Datalog.WriteComment sheetName & ", p_mode:" & ws_def.Cells(Row, col_mode).value & ", EQ num:" & BinCut(main_p_mode, passBinCut).EQ_Num(idx_step) & " is greater than global variable (TotalStepPerMode:" & CStr(TotalStepPerMode) & "), please check it. Error!!!"
                            'TheExec.ErrorLogMessage sheetName & ", p_mode:" & ws_def.Cells(Row, col_mode).value & ", EQ num:" & BinCut(main_p_mode, passBinCut).EQ_Num(idx_step) & " is greater than global variable (TotalStepPerMode:" & CStr(TotalStepPerMode) & "), please check it. Error!!!"    ''20230807: Removed by sanity check
'                        End If
                        
                        BinCut(main_p_mode, passBinCut).MAX_ID = CDbl(ws_def.Cells(Row, col_id).value)
                        BinCut(main_p_mode, passBinCut).c(idx_step) = CDbl(ws_def.Cells(Row, col_c).value)
                        BinCut(main_p_mode, passBinCut).m(idx_step) = CDbl(ws_def.Cells(Row, col_m).value)
                        '''*************************************************************************************'''
                        '''//Check if CPVmin, CPVmax, and CPGB are multiple of Step Size voltage.
                        BinCut(main_p_mode, passBinCut).CP_Vmax(idx_step) = CDbl(ws_def.Cells(Row, col_cp_vmax).value)
                        If BinCut(main_p_mode, passBinCut).CP_Vmax(idx_step) <> (Floor(BinCut(main_p_mode, passBinCut).CP_Vmax(idx_step) / BV_StepVoltage) * BV_StepVoltage) Then
                            Call Print_Error_Message(Error_Info, "LIB_VDD_BINNING", "initVddBinTableOneMod", sheetName & ", p_mode:" & ws_def.Cells(Row, col_mode).value & ", EQN:" & ws_def.Cells(Row, col_eqn).value & ", CPVmax:" & BinCut(main_p_mode, passBinCut).CP_Vmax(idx_step) & " should be multiple of 3.125. Error!!!")
                            'TheExec.Datalog.WriteComment sheetName & ", p_mode:" & ws_def.Cells(Row, col_mode).value & ", EQN:" & ws_def.Cells(Row, col_eqn).value & ", CPVmax:" & BinCut(main_p_mode, passBinCut).CP_Vmax(idx_step) & " should be multiple of 3.125. Error!!!"
                            'TheExec.ErrorLogMessage sheetName & ", p_mode:" & ws_def.Cells(Row, col_mode).value & ", EQN:" & ws_def.Cells(Row, col_eqn).value & ", CPVmax:" & BinCut(main_p_mode, passBinCut).CP_Vmax(idx_step) & " should be multiple of 3.125. Error!!!"    ''20230807: Removed by sanity check
                        End If
                        
                        BinCut(main_p_mode, passBinCut).CP_Vmin(idx_step) = CDbl(ws_def.Cells(Row, col_cp_vmin).value)
                        If BinCut(main_p_mode, passBinCut).CP_Vmin(idx_step) <> (Floor(BinCut(main_p_mode, passBinCut).CP_Vmin(idx_step) / BV_StepVoltage) * BV_StepVoltage) Then
                            Call Print_Error_Message(Error_Info, "LIB_VDD_BINNING", "initVddBinTableOneMod", sheetName & ", p_mode:" & ws_def.Cells(Row, col_mode).value & ", EQN:" & ws_def.Cells(Row, col_eqn).value & ", CPVmin:" & BinCut(main_p_mode, passBinCut).CP_Vmin(idx_step) & " should be multiple of 3.125. Error!!!")
                            'TheExec.Datalog.WriteComment sheetName & ", p_mode:" & ws_def.Cells(Row, col_mode).value & ", EQN:" & ws_def.Cells(Row, col_eqn).value & ", CPVmin:" & BinCut(main_p_mode, passBinCut).CP_Vmin(idx_step) & " should be multiple of 3.125. Error!!!"
                            'TheExec.ErrorLogMessage sheetName & ", p_mode:" & ws_def.Cells(Row, col_mode).value & ", EQN:" & ws_def.Cells(Row, col_eqn).value & ", CPVmin:" & BinCut(main_p_mode, passBinCut).CP_Vmin(idx_step) & " should be multiple of 3.125. Error!!!"    ''20230807: Removed by sanity check
                        End If

                        '''*************************************************************************************'''
                        '''//Parse the column of "Monotonicity_Offset".
                        If col_montonicityoffset <> 0 Then
                            BinCut(main_p_mode, passBinCut).Monotonicity_Offset(idx_step) = CDbl(ws_def.Cells(Row, col_montonicityoffset).value)
                        End If
                        '''*************************************************************************************'''
                        
                        BinCut(main_p_mode, passBinCut).CP_GB(idx_step) = CDbl(ws_def.Cells(Row, col_cpgb).value)

                        If BinCut(main_p_mode, passBinCut).CP_GB(idx_step) <> (Floor(BinCut(main_p_mode, passBinCut).CP_GB(idx_step) / BV_StepVoltage) * BV_StepVoltage) Then
                            Call Print_Error_Message(Error_Info, "LIB_VDD_BINNING", "initVddBinTableOneMod", sheetName & ", p_mode:" & ws_def.Cells(Row, col_mode).value & ", EQN:" & ws_def.Cells(Row, col_eqn).value & ", CPGB:" & BinCut(main_p_mode, passBinCut).CP_GB(idx_step) & " should be multiple of 3.125. Error!!!")
                            'TheExec.Datalog.WriteComment sheetName & ", p_mode:" & ws_def.Cells(Row, col_mode).value & ", EQN:" & ws_def.Cells(Row, col_eqn).value & ", CPGB:" & BinCut(main_p_mode, passBinCut).CP_GB(idx_step) & " should be multiple of 3.125. Error!!!"
                            'TheExec.ErrorLogMessage sheetName & ", p_mode:" & ws_def.Cells(Row, col_mode).value & ", EQN:" & ws_def.Cells(Row, col_eqn).value & ", CPGB:" & BinCut(main_p_mode, passBinCut).CP_GB(idx_step) & " should be multiple of 3.125. Error!!!"    ''20230807: Removed by sanity check
                        End If
                        '''*************************************************************************************'''
                        BinCut(main_p_mode, passBinCut).CP2_GB(idx_step) = CDbl(ws_def.Cells(Row, col_cp2gb).value)
                        BinCut(main_p_mode, passBinCut).FT1_GB(idx_step) = CDbl(ws_def.Cells(Row, col_ft1gb).value)
                        BinCut(main_p_mode, passBinCut).FT2_GB(idx_step) = CDbl(ws_def.Cells(Row, col_ft2gb).value)
                        BinCut(main_p_mode, passBinCut).SLT_GB(idx_step) = CDbl(ws_def.Cells(Row, col_slt_ftqa_gb).value)
                        BinCut(main_p_mode, passBinCut).FTQA_GB(idx_step) = CDbl(ws_def.Cells(Row, col_ate_ftqa_gb).value)
                        BinCut(main_p_mode, passBinCut).SLT_FTQA_GB(idx_step) = CDbl(ws_def.Cells(Row, col_slt_ftqa_gb).value)
                        BinCut(main_p_mode, passBinCut).HVCC_CP(idx_step) = CDbl(ws_def.Cells(Row, col_cphv).value)
                        BinCut(main_p_mode, passBinCut).HVCC_FT(idx_step) = CDbl(ws_def.Cells(Row, col_fthv).value)
                        BinCut(main_p_mode, passBinCut).HVCC_QA(idx_step) = CDbl(ws_def.Cells(Row, col_qahv).value)
                        
                        '20230504: Added for parsing sheet defining BinX EQNs and IDS value
                        If Flag_BinX_Info_Parsed = True Then
                            If UCase(Trim(ws_def.Cells(Row, col_eqn_bin).value)) = "BINX" Then      ''Add this EQN to binX EQN dict
                                If Pmode2BinX_EQN_Dict.Exists(main_p_mode) = False Then
                                    Pmode2BinX_EQN_Dict.Add main_p_mode, CStr(BinCut(main_p_mode, passBinCut).EQ_Num(idx_step))     ''Save dict: pmode(80) -> String "1"
                                Else    ''Concatenate binX EQN "1,2,3"
                                    Dim tmp_str_EQN As String
                                    tmp_str_EQN = Pmode2BinX_EQN_Dict(main_p_mode)
                                    tmp_str_EQN = tmp_str_EQN & "," & CStr(BinCut(main_p_mode, passBinCut).EQ_Num(idx_step))
                                    Pmode2BinX_EQN_Dict.Remove main_p_mode
                                    Pmode2BinX_EQN_Dict.Add main_p_mode, tmp_str_EQN        ''Save dict: pmode(80) -> String "1,2"
                                End If
                            ''20230901: Added saving bin1 EQN, for pre-flow in bincut
                            ElseIf UCase(Trim(ws_def.Cells(Row, col_eqn_bin).value)) = "BIN1" Then
                                AllBinCut(main_p_mode).Last_Bin1_EQN = WorksheetFunction.Min(AllBinCut(main_p_mode).Last_Bin1_EQN, BinCut(main_p_mode, passBinCut).EQ_Num(idx_step))
                            End If
                            
                            If CDbl(Trim(ws_def.Cells(Row, col_cpids).value)) <> 0 Then
                                If Pmode2BinX_IDS_Dict.Exists(main_p_mode) = False Then
                                    Pmode2BinX_IDS_Dict.Add main_p_mode, CDbl(Trim(ws_def.Cells(Row, col_cpids).value))                 ''Save dict: pmode(80) -> 38.8(mA)
                                Else
                                    If Pmode2BinX_IDS_Dict(main_p_mode) <> CDbl(Trim(ws_def.Cells(Row, col_cpids).value)) Then
                                        TheExec.Datalog.WriteComment "'CPIDSMax' is not the same in pmode:" & VddBinName(main_p_mode) & ", ERROR!!"
                                    End If
                                End If
                            End If
                            
                            ''20240118: Added to save IDS hot for binX merge feature
                            If CDbl(Trim(ws_def.Cells(Row, col_ftids).value)) <> 0 Then
                                If Pmode2BinX_IDSHot_Dict.Exists(main_p_mode) = False Then
                                    Pmode2BinX_IDSHot_Dict.Add main_p_mode, CDbl(Trim(ws_def.Cells(Row, col_ftids).value))                 ''Save dict: pmode(80) -> 3888.8(mA)
                                Else
                                    If Pmode2BinX_IDSHot_Dict(main_p_mode) <> CDbl(Trim(ws_def.Cells(Row, col_ftids).value)) Then
                                        TheExec.Datalog.WriteComment "'IDSMax_HOT' is not the same in pmode:" & VddBinName(main_p_mode) & ", ERROR!!"
                                    End If
                                End If
                            End If
                        End If

                        If col_htol_ro_gb > 0 Then
                            BinCut(main_p_mode, passBinCut).HTOL_RO_GB(idx_step) = CDbl(ws_def.Cells(Row, col_htol_ro_gb).value)
                        End If
                        
                        If col_htol_ro_gb_room > 0 Then
                            If TheExec.flow.enableWord("HTOL_TX_ROOM") = True Then
                                BinCut(main_p_mode, passBinCut).FT1_GB(idx_step) = CDbl(ws_def.Cells(Row, col_htol_ro_gb_room).value)
                            End If
                            BinCut(main_p_mode, passBinCut).HTOL_RO_GB_ROOM(idx_step) = CDbl(ws_def.Cells(Row, col_htol_ro_gb_room).value)
                        End If
                        
                        If col_htol_ro_gb_hot > 0 Then
                            If TheExec.flow.enableWord("HTOL_TX_HOT") = True Then
                                BinCut(main_p_mode, passBinCut).FT2_GB(idx_step) = CDbl(ws_def.Cells(Row, col_htol_ro_gb_hot).value)
                            End If
                            BinCut(main_p_mode, passBinCut).HTOL_RO_GB_HOT(idx_step) = CDbl(ws_def.Cells(Row, col_htol_ro_gb_hot).value)
                        End If
                        '''*************************************************************************************'''
                        If (ws_def.Cells(Row, col_allow_equal).value <> "") Then '''20161228, liki
                            BinCut(main_p_mode, passBinCut).Allow_Equal(idx_step) = VddBinStr2Enum(ws_def.Cells(Row, col_allow_equal).value)
                        Else
                            BinCut(main_p_mode, passBinCut).Allow_Equal(idx_step) = 0
                        End If
                        
                        '''//Check the column of "Allow Equal".
                        If passBinCut = 1 And idx_step = 0 Then
                            AllBinCut(main_p_mode).Allow_Equal = BinCut(main_p_mode, passBinCut).Allow_Equal(idx_step)
                        Else
                            If AllBinCut(main_p_mode).Allow_Equal <> BinCut(main_p_mode, passBinCut).Allow_Equal(idx_step) Then
                                Call Print_Error_Message(Error_Info, "LIB_VDD_BINNING", "initVddBinTableOneMod", "The Allow Equal from Eqn" & idx_step + 1 & " of " & VddBinName(main_p_mode) & " in sheet " & sheetName & " doesn't match other step or PassBinCut. Error!!!")
                                'TheExec.ErrorLogMessage "The Allow Equal from Eqn" & idx_step + 1 & " of " & VddBinName(main_p_mode) & " in sheet " & sheetName & " doesn't match other step or PassBinCut. Error!!!"    ''20230807: Removed by sanity check
                            End If
                        End If
                        
                        '' 20211202: For CPU_SRAM Calculate CP Voltage, determines which IDS value to use and to calculate C-M*log10(IDS)
                        If Flag_IDS_Mapping_enable = True Then
                            If col_ids_mapping > 0 And col_ids_mapping_max > 0 Then
                                If ws_def.Cells(Row, col_ids_mapping).value <> "" Then
                                    AllBinCut(main_p_mode).IDS_MAPPING = "VDD_" & ws_def.Cells(Row, col_ids_mapping).value              '' record which IDS to map
                                    AllBinCut(main_p_mode).IDS_MAPPING_LIMIT = CDbl(ws_def.Cells(Row, col_ids_mapping_max).value)           '' record the limit of the IDS (No use for now)
                                Else
                                    AllBinCut(main_p_mode).IDS_MAPPING = AllBinCut(main_p_mode).powerPin
                                End If
                            Else
                                Call Print_Error_Message(Error_Info, "LIB_VDD_BINNING", "initVddBinTableOneMod", "No IDS mapping is determined, but Flag_IDS_Mapping_enable = True. ERROR !!")
                            End If
                        End If
                        '' **************************************************************************************************************
                        
                        If ws_def.Cells(Row, col_comment).value <> "" Then
                            strTemp = LCase(ws_def.Cells(Row, col_comment).value)
                            
                            '''//Check if "Max PV (pmode0/pmode1)" is in the column "Comment" or not. Check check pmode, allowEqual, and MaxPV/MinPV.
                            If LCase(strTemp) Like "max*pv*(*)" Or LCase(strTemp) Like "min*pv*(*)" Then '''//ex: Max PV (MP008/MP009/MP00A/MP105)
                                If ws_def.Cells(Row, col_mode).value <> "" Then
                                    If LCase(strTemp) Like LCase("*" & ws_def.Cells(Row, col_mode).value & "*") _
                                    And LCase(strTemp) Like LCase("*" & ws_def.Cells(Row, col_allow_equal).value & "*") Then
                                        strTemp = UCase(Replace(strTemp, "/", ","))
                                        split_content = Split(strTemp, "(")
                                        split_content = Split(split_content(UBound(split_content)), ")")
                                        
                                        If LCase(strTemp) Like "max*pv*" Then
                                            Flag_Adjust_Max_Enable = Flag_Adjust_Max_Enable Or True
                                            
                                            If Adjust_Power_Max_pmode <> "" Then
                                                If UCase("*+" & Adjust_Power_Max_pmode & "+*") Like UCase("*+" & split_content(0) & "+*") Then
                                                    '''Do nothing...
                                                Else
                                                    Adjust_Power_Max_pmode = Adjust_Power_Max_pmode & "+" & split_content(0)
                                                End If
                                            Else
                                                Adjust_Power_Max_pmode = split_content(0)
                                            End If
                                        ElseIf LCase(strTemp) Like "min*pv*" Then
                                            Flag_Adjust_Min_Enable = Flag_Adjust_Min_Enable Or True
                                            
                                            If Adjust_Power_Min_pmode <> "" Then
                                                If UCase("*+" & Adjust_Power_Min_pmode & "+*") Like UCase("*+" & split_content(0) & "+*") Then
                                                    '''Do nothing...
                                                Else
                                                    Adjust_Power_Min_pmode = Adjust_Power_Min_pmode & "+" & split_content(0)
                                                End If
                                            Else
                                                Adjust_Power_Min_pmode = split_content(0)
                                            End If
                                        End If
                                    Else
                                        Call Print_Error_Message(Error_Info, "LIB_VDD_BINNING", "initVddBinTableOneMod", ws_def.Cells(Row, col_mode).value & " doesn't have the correct AllowEqual cell with MaxPV or MinPV in sheet " & sheetName & ". Error!!!")
                                        'TheExec.Datalog.WriteComment ws_def.Cells(Row, col_mode).value & " doesn't have the correct AllowEqual cell with MaxPV or MinPV in sheet " & sheetName & ". Error!!!"
                                        'TheExec.ErrorLogMessage ws_def.Cells(Row, col_mode).value & " doesn't have the correct AllowEqual cell with MaxPV or MinPV in sheet " & sheetName & ". Error!!!"    ''20230807: Removed by sanity check
                                    End If
                                Else
                                    Call Print_Error_Message(Error_Info, "LIB_VDD_BINNING", "initVddBinTableOneMod", ws_def.Cells(Row, col_mode).value & " doesn't have the correct AllowEqual cell in sheet " & sheetName & ". Error!!!")
                                    'TheExec.Datalog.WriteComment ws_def.Cells(Row, col_mode).value & " doesn't have the correct AllowEqual cell in sheet " & sheetName & ". Error!!!"
                                    'TheExec.ErrorLogMessage ws_def.Cells(Row, col_mode).value & " doesn't have the correct AllowEqual cell in sheet " & sheetName & ". Error!!!"    ''20230807: Removed by sanity check
                                End If
                            End If
                        End If
                        
                        '''***//Interpolation//***'''
                        '''20210305: Modified to set INTP_MODE_L, INTP_MODE_L, and AllowEqual if the cells are empty.
                        '''//Start performance mode
                        If col_intModeL > 0 Then
                            If (ws_def.Cells(Row, col_intModeL).value <> "") Then '//20180312: modified by Anderson.
                                BinCut(main_p_mode, passBinCut).INTP_MODE_L(idx_step) = VddBinStr2Enum(ws_def.Cells(Row, col_intModeL).value)
                            Else
                                BinCut(main_p_mode, passBinCut).INTP_MODE_L(idx_step) = 0
                            End If
                        End If
                        
                        '''//End performance mode
                        If col_intModeH > 0 Then
                            If (ws_def.Cells(Row, col_intModeH).value <> "") Then '//20180312: modified by Anderson.
                                BinCut(main_p_mode, passBinCut).INTP_MODE_H(idx_step) = VddBinStr2Enum(ws_def.Cells(Row, col_intModeH).value)
                            Else
                                BinCut(main_p_mode, passBinCut).INTP_MODE_H(idx_step) = 0
                            End If
                        End If
                        
                        '''//interpolation factor
                        If col_intMFactor > 0 Then
                            If (ws_def.Cells(Row, col_intMFactor).value <> "") Then '//20180312: modified by Anderson.
                                Flag_Interpolation_enable = True
                                BinCut(main_p_mode, passBinCut).INTP_MFACTOR(idx_step) = CDbl(ws_def.Cells(Row, col_intMFactor).value)
                            Else
                                BinCut(main_p_mode, passBinCut).INTP_MFACTOR(idx_step) = 0
                            End If
                        End If
                        
                        '''//offset of interpolation.
                        If col_intOffset > 0 Then
                            If (ws_def.Cells(Row, col_intOffset).value <> "") Then
                                BinCut(main_p_mode, passBinCut).INTP_OFFSET(idx_step) = CDbl(ws_def.Cells(Row, col_intOffset).value)
                                '''20210322: Discussed the vbt code that checked if int_Offset is multiple of StepVoltage, all project BinCut owners decided to remove the vbt code because this was the redundant action.
                            Else
                                BinCut(main_p_mode, passBinCut).INTP_OFFSET(idx_step) = 0
                            End If
                        End If
                        
                        '''//Check "AllBinCut(p_mode).INTP_SKIPTEST".
                        '''//flag to skip interpolation tests of p_mode.
                        If col_intSkipTest > 0 Then
                            If (ws_def.Cells(Row, col_intSkipTest).value <> "") Then
                                If LCase(ws_def.Cells(Row, col_intSkipTest).value) = "yes" Then
                                    If ws_def.Cells(Row, col_intMFactor).value <> "" Then
                                        BinCut(main_p_mode, passBinCut).INTP_SKIPTEST(idx_step) = True
                                        
                                        If passBinCut = 1 Then
                                            If idx_step = 0 Then
                                                AllBinCut(main_p_mode).INTP_SKIPTEST = True
                                            Else
                                                If BinCut(main_p_mode, passBinCut).INTP_SKIPTEST(idx_step) <> AllBinCut(main_p_mode).INTP_SKIPTEST Then
                                                    Call Print_Error_Message(Error_Info, "LIB_VDD_BINNING", "initVddBinTableOneMod", ws_def.Cells(Row, col_mode) & " has interpolation skipTest=yes in row " & Row & ", but all Eqn don't have the same INTP_SKIPTEST setting in sheet " & sheetName & ". Error!!!")
                                                    'TheExec.Datalog.WriteComment ws_def.Cells(Row, col_mode) & " has interpolation skipTest=yes in row " & Row & ", but all Eqn don't have the same INTP_SKIPTEST setting in sheet " & sheetName & ". Error!!!"
                                                End If
                                            End If
                                        End If
                                    Else
                                        BinCut(main_p_mode, passBinCut).INTP_SKIPTEST(idx_step) = False
                                        Call Print_Error_Message(Error_Info, "LIB_VDD_BINNING", "initVddBinTableOneMod", ws_def.Cells(Row, col_mode) & " has interpolation skipTest=yes in row " & Row & ", but interpolation factor doesn't have the correct value in sheet " & sheetName & ". Error!!!")
                                        'TheExec.Datalog.WriteComment ws_def.Cells(Row, col_mode) & " has interpolation skipTest=yes in row " & Row & ", but interpolation factor doesn't have the correct value in sheet " & sheetName & ". Error!!!"
                                    End If
                                Else
                                    BinCut(main_p_mode, passBinCut).INTP_SKIPTEST(idx_step) = False
                                End If
                            Else
                                If passBinCut = 1 And idx_step > 0 Then
                                    If AllBinCut(main_p_mode).INTP_SKIPTEST = True Then
                                        Call Print_Error_Message(Error_Info, "LIB_VDD_BINNING", "initVddBinTableOneMod", ws_def.Cells(Row, col_mode) & " has interpolation skipTest=yes in row " & Row & ", but all Eqn don't have the same INTP_SKIPTEST setting in sheet " & sheetName & ". Error!!!")
                                        'TheExec.Datalog.WriteComment ws_def.Cells(Row, col_mode) & " has interpolation skipTest=yes in row " & Row & ", but all Eqn don't have the same INTP_SKIPTEST setting in sheet " & sheetName & ". Error!!!"
                                        'TheExec.ErrorLogMessage ws_def.Cells(Row, col_mode) & " has interpolation skipTest=yes in row " & Row & ", but all Eqn don't have the same INTP_SKIPTEST setting in sheet " & sheetName & ". Error!!!"    ''20230807: Removed by sanity check
                                    End If
                                End If
                                BinCut(main_p_mode, passBinCut).INTP_SKIPTEST(idx_step) = False
                            End If
                        End If
                        
                        If idx_step = 0 Then
                            AllBinCut(main_p_mode).IDS_CP_LIMIT = 0
                            AllBinCut(main_p_mode).IDS_FT_LIMIT = 0
                            AllBinCut(main_p_mode).IDS_QA_LIMIT = 0
                            AllBinCut(main_p_mode).IDS_FT2_LIMIT = 0
                            AllBinCut(main_p_mode).IDS_FT2_QA_LIMIT = 0
                            '''****************************************************************************************************************'''
                            '''Use ID to define if the inheritance had been changed or not
                            '''****************************************************************************************************************'''
                            If passBinCut = 1 Then                                  '' 20211222: Added to support 7-digit ID Number in Vdd_Def
                                If BinCut(main_p_mode, passBinCut).MAX_ID = 1 Or BinCut(main_p_mode, passBinCut).MAX_ID Mod 100000 = 1001 Then
                                    If Power_List_All <> "" Then
                                        Power_List_All = Power_List_All & "@" & UCase(ws_def.Cells(Row, col_mode))
                                     Else
                                        Power_List_All = UCase(ws_def.Cells(Row, col_mode))
                                    End If
                                Else
                                    If Power_List_All <> "" Then
                                        Power_List_All = Power_List_All & "," & UCase(ws_def.Cells(Row, col_mode))
                                    Else
                                        Power_List_All = UCase(ws_def.Cells(Row, col_mode))
                                    End If
                                End If
                            End If
                        End If
                        
                        ''202307xx: Modified to use BinX IDS for IDS Max
                        If Flag_BinX_Info_Parsed = True Then
                            BinCut(main_p_mode, passBinCut).IDS_CP_LIMIT(idx_step) = CDbl(ws_def.Cells(Row, col_binx_ids).value)
                            BinCut(main_p_mode, passBinCut).IDS_FT_LIMIT(idx_step) = CDbl(ws_def.Cells(Row, col_binx_ftids).value)
                        Else
                            BinCut(main_p_mode, passBinCut).IDS_CP_LIMIT(idx_step) = CDbl(ws_def.Cells(Row, col_cpids).value)
                            BinCut(main_p_mode, passBinCut).IDS_FT_LIMIT(idx_step) = CDbl(ws_def.Cells(Row, col_ftids).value)
                        End If
                        
                        'BinCut(main_p_mode, passBinCut).IDS_FT_LIMIT(idx_step) = CDbl(ws_def.Cells(row, col_ftids).value)
                        
                        'WWW20220630 save IDS to Bincut corepower
                        BinCut(VddBinStr2Enum(AllBinCut(main_p_mode).powerPin), passBinCut).IDS_CP_LIMIT(idx_step) = BinCut(main_p_mode, passBinCut).IDS_CP_LIMIT(idx_step)
                        BinCut(VddBinStr2Enum(AllBinCut(main_p_mode).powerPin), passBinCut).IDS_FT_LIMIT(idx_step) = BinCut(main_p_mode, passBinCut).IDS_FT_LIMIT(idx_step)
     
                        
                        If AllBinCut(main_p_mode).IDS_CP_LIMIT < BinCut(main_p_mode, passBinCut).IDS_CP_LIMIT(idx_step) Then AllBinCut(main_p_mode).IDS_CP_LIMIT = BinCut(main_p_mode, passBinCut).IDS_CP_LIMIT(idx_step)
                        If AllBinCut(main_p_mode).IDS_FT_LIMIT < BinCut(main_p_mode, passBinCut).IDS_FT_LIMIT(idx_step) Then AllBinCut(main_p_mode).IDS_FT_LIMIT = BinCut(main_p_mode, passBinCut).IDS_FT_LIMIT(idx_step)
                        
                        '''20210312: Modified to parse columns of "Softbin" and "HardBin" from column col_sort+1.
                        For Test_Type = 0 To testType.Total_Count - 3 '''only use SPI, Mbist, TD first
                            BinCut(main_p_mode, passBinCut).SBIN_BINNING_FAIL(idx_step, Test_Type) = CLng(ws_def.Cells(Row, (col_comment + 1) + 4 * Test_Type).value)      'large IDS at certain level
                            BinCut(main_p_mode, passBinCut).SBIN_LVCC_FAIL(idx_step, Test_Type) = CLng(ws_def.Cells(Row, (col_comment + 1) + 1 + 4 * Test_Type).value)       'Can find LVCC
                            BinCut(main_p_mode, passBinCut).HBIN_BINNING_FAIL(idx_step, Test_Type) = CLng(ws_def.Cells(Row, (col_comment + 1) + 2 + 4 * Test_Type).value)    'large IDS at certain level
                            BinCut(main_p_mode, passBinCut).HBIN_LVCC_FAIL(idx_step, Test_Type) = CLng(ws_def.Cells(Row, (col_comment + 1) + 3 + 4 * Test_Type).value)       'Can find LVCC

                            '' 20220624: Check Sortbin / Hardbin / Pmode / Result in BinTable and Vdd_Binning_Def
                            If Flag_Enable_Parsing_BinTable_ForBinCut = True And Flag_BinTable_Parsed_ForBinCut = True _
                                And BinCut(main_p_mode, passBinCut).SBIN_LVCC_FAIL(idx_step, Test_Type) <> 0 _
                                And BinCut(main_p_mode, passBinCut).HBIN_LVCC_FAIL(idx_step, Test_Type) <> 0 _
                                And Not UCase(powerDomain) Like "*VDD*_SRAM*" Then
                                '' Check Sortbin -> Hardbin
                                If Sortbin2hardbinDict.Exists(BinCut(main_p_mode, passBinCut).SBIN_LVCC_FAIL(idx_step, Test_Type)) And _
                                    Sortbin2ResultDict.Exists(BinCut(main_p_mode, passBinCut).SBIN_LVCC_FAIL(idx_step, Test_Type)) And _
                                    Sortbin2PmodeDict.Exists(BinCut(main_p_mode, passBinCut).SBIN_LVCC_FAIL(idx_step, Test_Type)) Then '' sortbin exist
                                    '' Sortbin -> "PASS"
                                    If UCase(Sortbin2ResultDict(BinCut(main_p_mode, passBinCut).SBIN_LVCC_FAIL(idx_step, Test_Type))) Like "*PASS*" Then
                                        Call Print_Error_Message(Error_Info, "LIB_VDD_BINNING", "initVddBinTableOneMod", "Error!! " & sheetName & " " & ws_def.Cells(Row, col_mode) & "_" & ws_def.Cells(Row, col_eqn).value & " " & _
                                                                        BinCut(main_p_mode, passBinCut).SBIN_LVCC_FAIL(idx_step, Test_Type) & " has been set to Good Bin in the Bintable !!!")
                                        If isDebugMode = True Then
                                            TheExec.AddOutput "Error!! " & sheetName & " " & ws_def.Cells(Row, col_mode) & "_" & ws_def.Cells(Row, col_eqn).value & " " & _
                                                                BinCut(main_p_mode, passBinCut).SBIN_LVCC_FAIL(idx_step, Test_Type) & " has been set to Good Bin in the Bintable !!!"
                                        End If
                                    '' Sortbin -> Pmode mis-match
                                    ElseIf Not UCase(Sortbin2PmodeDict(BinCut(main_p_mode, passBinCut).SBIN_LVCC_FAIL(idx_step, Test_Type))) Like UCase(Trim(ws_def.Cells(Row, col_mode))) Then
                                        Call Print_Error_Message(Error_Info, "LIB_VDD_BINNING", "initVddBinTableOneMod", "Error!! " & sheetName & " " & ws_def.Cells(Row, col_mode) & "_" & ws_def.Cells(Row, col_eqn).value & " " & _
                                                                        BinCut(main_p_mode, passBinCut).SBIN_LVCC_FAIL(idx_step, Test_Type) & " has a mis-match setting with Bintable! Pmode is different!!")
                                        If isDebugMode = True Then
                                            TheExec.AddOutput "Error!! " & sheetName & " " & ws_def.Cells(Row, col_mode) & "_" & ws_def.Cells(Row, col_eqn).value & " " & _
                                                                        BinCut(main_p_mode, passBinCut).SBIN_LVCC_FAIL(idx_step, Test_Type) & " has a mis-match setting with Bintable! Pmode is different!!"
                                        End If
                                    '' Sortbin -> Hardbin mis-match, overwrite
                                    ElseIf Not Sortbin2hardbinDict(BinCut(main_p_mode, passBinCut).SBIN_LVCC_FAIL(idx_step, Test_Type)) = BinCut(main_p_mode, passBinCut).HBIN_LVCC_FAIL(idx_step, Test_Type) Then
                                        Call Print_Error_Message(Error_Info, "LIB_VDD_BINNING", "initVddBinTableOneMod", "Error!! " & sheetName & " " & ws_def.Cells(Row, col_mode) & "_" & ws_def.Cells(Row, col_eqn).value & " " & _
                                                                BinCut(main_p_mode, passBinCut).SBIN_LVCC_FAIL(idx_step, Test_Type) & " has a mis-match setting with Bintable! Hardbin = " & _
                                                                Sortbin2hardbinDict(BinCut(main_p_mode, passBinCut).SBIN_LVCC_FAIL(idx_step, Test_Type)) & _
                                                                ", not " & BinCut(main_p_mode, passBinCut).HBIN_LVCC_FAIL(idx_step, Test_Type) & "!!")
                                        If isDebugMode = True Then
                                            TheExec.AddOutput "Error!! " & sheetName & " " & ws_def.Cells(Row, col_mode) & "_" & ws_def.Cells(Row, col_eqn).value & " " & _
                                                                BinCut(main_p_mode, passBinCut).SBIN_LVCC_FAIL(idx_step, Test_Type) & " has a mis-match setting with Bintable! Hardbin = " & _
                                                                Sortbin2hardbinDict(BinCut(main_p_mode, passBinCut).SBIN_LVCC_FAIL(idx_step, Test_Type)) & _
                                                                ", not " & BinCut(main_p_mode, passBinCut).HBIN_LVCC_FAIL(idx_step, Test_Type) & "!!"
                                        End If
                                        BinCut(main_p_mode, passBinCut).HBIN_LVCC_FAIL(idx_step, Test_Type) = Sortbin2hardbinDict(BinCut(main_p_mode, passBinCut).SBIN_LVCC_FAIL(idx_step, Test_Type))
                                    End If
                                Else        '' Sortbin doesn't exist
                                    Call Print_Error_Message(Error_Info, "LIB_VDD_BINNING", "initVddBinTableOneMod", "Error!! " & sheetName & " " & ws_def.Cells(Row, col_mode) & "_" & ws_def.Cells(Row, col_eqn).value & " " & _
                                                                    BinCut(main_p_mode, passBinCut).SBIN_LVCC_FAIL(idx_step, Test_Type) & " is missing in Bintable")
                                    If isDebugMode = True Then
                                        TheExec.AddOutput "Error!! " & sheetName & " " & ws_def.Cells(Row, col_mode) & "_" & ws_def.Cells(Row, col_eqn).value & " " & _
                                                                    BinCut(main_p_mode, passBinCut).SBIN_LVCC_FAIL(idx_step, Test_Type) & " is missing in Bintable"
                                    End If
                                End If
                            End If
                        Next Test_Type
                        
                        '''************************************************************************************************************************'''
                        '''column "SRAMthresh_CP1" or "SRAMthresh_BinSearch"==> SRAM_VTH_SPEC(0): for BinCut search.
                        '''column "SRAMthresh_Product"                      ==> SRAM_VTH_SPEC(1): for BinCut check.
                        '''************************************************************************************************************************'''
                        '''20210325: Modified to use the 1-dimension array to store SRAM_Vth.
                        If col_sram_vt_spec(0) > 0 Then
                            BinCut(main_p_mode, passBinCut).SRAM_VTH_SPEC(0) = CDbl(ws_def.Cells(Row, col_sram_vt_spec(0)).value)
                        End If
                        
                        If col_sram_vt_spec(1) > 0 Then
                            BinCut(main_p_mode, passBinCut).SRAM_VTH_SPEC(1) = CDbl(ws_def.Cells(Row, col_sram_vt_spec(1)).value)
                        End If
                        
                        For jobIdx = 0 To BinCutJobDefinition.Total_Count
                            For testTypeIdx = 0 To testType.Total_Count
                                If (col_dynamic_offset(jobIdx, testTypeIdx) <> 0) Then
                                    BinCut(main_p_mode, passBinCut).DYNAMIC_OFFSET(jobIdx, testTypeIdx) = CDbl(ws_def.Cells(Row, col_dynamic_offset(jobIdx, testTypeIdx)).value)
                                End If
                            Next testTypeIdx
                        Next jobIdx
                        
                        BinCut(main_p_mode, passBinCut).Mode_Step = idx_step
                    End If
                End If
            Next Row
        Else
            Call Print_Error_Message(Error_Info, "LIB_VDD_BINNING", "initVddBinTableOneMod", "Columns of the header in the sheet " & sheetName & " might be incorrect. Error!!!")
            'TheExec.ErrorLogMessage "Columns of the header in the sheet " & sheetName & " might be incorrect. Error!!!"    ''20230807: Removed by sanity check
        End If
    End If '''If isSheetFound = True
    
    '''set the last Step to the error value
    For p_mode = 0 To MaxPerformanceModeCount - 1
        BinCut(p_mode, passBinCut).EQ_Num(MaxEqnNum) = 0 'TotalStepPerMode down
        BinCut(p_mode, passBinCut).c(MaxEqnNum) = 0
        BinCut(p_mode, passBinCut).m(MaxEqnNum) = 0
        BinCut(p_mode, passBinCut).CP_Vmax(MaxEqnNum) = 0
        BinCut(p_mode, passBinCut).CP_Vmin(MaxEqnNum) = 0
        BinCut(p_mode, passBinCut).CP_GB(MaxEqnNum) = 0
        BinCut(p_mode, passBinCut).FT1_GB(MaxEqnNum) = 0
        BinCut(p_mode, passBinCut).CP2_GB(MaxEqnNum) = 0
        BinCut(p_mode, passBinCut).FT2_GB(MaxEqnNum) = 0
        BinCut(p_mode, passBinCut).SLT_GB(MaxEqnNum) = 0
        BinCut(p_mode, passBinCut).FTQA_GB(MaxEqnNum) = 0
        BinCut(p_mode, passBinCut).SLT_FTQA_GB(MaxEqnNum) = 0
        BinCut(p_mode, passBinCut).HTOL_RO_GB(MaxEqnNum) = 0
        BinCut(p_mode, passBinCut).HTOL_RO_GB_ROOM(MaxEqnNum) = 0
        BinCut(p_mode, passBinCut).HTOL_RO_GB_HOT(MaxEqnNum) = 0
        BinCut(p_mode, passBinCut).IDS_CP_LIMIT(MaxEqnNum) = 0
        BinCut(p_mode, passBinCut).IDS_FT_LIMIT(MaxEqnNum) = 0
        BinCut(p_mode, passBinCut).IDS_QA_LIMIT(MaxEqnNum) = 0
        BinCut(p_mode, passBinCut).IDS_FT2_LIMIT(MaxEqnNum) = 0
        BinCut(p_mode, passBinCut).IDS_FT2_QA_LIMIT(MaxEqnNum) = 0 ' up TotalStepPerMode
        
        For Test_Type = 0 To testType.Total_Count - 1
            BinCut(p_mode, passBinCut).SBIN_BINNING_FAIL(MaxEqnNum, Test_Type) = 0
            BinCut(p_mode, passBinCut).SBIN_LVCC_FAIL(MaxEqnNum, Test_Type) = 0
            BinCut(p_mode, passBinCut).HBIN_BINNING_FAIL(MaxEqnNum, Test_Type) = 0
            BinCut(p_mode, passBinCut).HBIN_LVCC_FAIL(MaxEqnNum, Test_Type) = 0
            'BinCut(p_mode, passBinCut).SBIN_BINNING_FAIL(TotalStepPerMode, Test_Type) = 0
            'BinCut(p_mode, passBinCut).SBIN_LVCC_FAIL(TotalStepPerMode, Test_Type) = 0
            'BinCut(p_mode, passBinCut).HBIN_BINNING_FAIL(TotalStepPerMode, Test_Type) = 0
            'BinCut(p_mode, passBinCut).HBIN_LVCC_FAIL(TotalStepPerMode, Test_Type) = 0
        Next Test_Type
    Next p_mode
Exit Function
errHandler:
    Call Print_Error_Message(Error_Info, "LIB_VDD_BINNING", "initVddBinTableOneMod") 'Add ErrHandler 2023/08/01
    If AbortTest Then Exit Function Else Resume Next
End Function

'20210819: Modified to revise the vbt code for the new format of the header in BinCut voltage tables, requested by C651 Toby.
'20210720: Modified to revise ids_hi_limit with CPIDSMax because C651 Si asked us to use Efuse processed IDS for search in FT.
'20201021: Modified to use "dict_IsCorePower" to store and check CorePower/OtherRail.
'20200703: Modiifed to use "check_Sheet_Range".
'20200415: Modified to use the same parsing method with "initVddotherrailOneMod".
'20191219: Modified to check powerDomain in the vbt function "initDomain2Pin".
'20191127: Modified for the revised InitVddBinTable.
'20191007: Modified to merge the vbt code for SRAM with/without p_mode.
'20191001: Modified for the new header defined by C651, they separated IDS_limit into "CPIDSMax" and "IDSMax_HOT".
'20191001: Modified to check the column number before storing the data from the column.
'20190706: Modified to check if the powerDomain is in the pin_group "ShtOtherRailinFlowSheet" or not.
'20190606: Modified for CPIDS_Spec_OtherRail and FTIDS_Spec_OtherRail.
'20190603: Modified for IDS limit of SRAM with CorePower p_mode.
'20190527: Modified for C651 new string format of OtherRail, ex: "MCS601 CPVmax", "CPVmax".
'20190426: Modified to use the function "Find_Sheet".
'20190319: Modified to parse dynamic_offset of OtherRail (requested by C651).
'20190311: Modified for CP1_GB of OtherRail.
'20190307: Modified the vbt code for compatible with conventional projects (SRAM_*** with CorePower Pmode), ex: MCS601 for VDD_CPU_SRAM.
Public Function initVddotherrailOneMod(passBinCut As Long)
On Error GoTo errHandler 'Add ErrHandler 2023/08/01
    Dim wb As Workbook
    Dim ws_def As Worksheet
    Dim sheetName As String
    Dim Row As Long, Col As Long
    Dim p_mode As Integer
    Dim other_p_mode As Integer
    Dim col_binned As Integer
    Dim col_domain As Integer
    Dim col_mode As Integer
    Dim col_cp_vmax As Integer
    Dim col_cp_vmin As Integer
    Dim col_cpgb As Integer
    '''20230706: Added to support binX merge into bin1 feature
    Dim col_eqn_bin As Integer
    Dim col_binx_ids As Integer
    Dim col_binx_ftids As Integer
    Dim col_c As Integer
    Dim col_cp2gb As Integer
    Dim col_ft1gb As Integer
    Dim col_ft2gb As Integer
    Dim col_sltgb As Integer
    Dim col_htol_ro_gb As Integer
    Dim col_htol_ro_gb_room As Integer
    Dim col_htol_ro_gb_hot As Integer
    Dim col_ate_ftqa_gb As Integer
    Dim col_slt_ftqa_gb As Integer
    Dim col_cphv As Integer
    Dim col_fthv As Integer
    Dim col_qahv As Integer
    Dim col_cpids As Integer
    Dim col_ftids As Integer
    Dim col_ids_mapping As Integer
    Dim col_ids_mapping_max As Integer
    Dim col_dynamic_offset() As Integer '''Added for dynamic_offset of OtherRail, 20190319  Dim col_dynamic_offset(MaxJobCountInVbt, MaxTestType) As Integer 20230526
    Dim jobIdx As Integer, testTypeIdx As Integer
    Dim powerDomain As String
    Dim i As Long
    Dim Row_of_Title As Integer
    Dim MaxRow As Long
    Dim maxcol As Long
    Dim start_p_mode As Integer
    Dim stop_p_mode As Integer
    Dim enableRowParsing As Boolean
    Dim isSheetFound As Boolean
'''//==================================================================================================================================================================================//'''
'''//Note:
'''1. Check_IDS and judge_IDS are dedicated to Efuse processed IDS for BinCut search, as requested by C651 Si, 20210617
'''//==================================================================================================================================================================================//'''
    '''*****************************************************************'''
    '''//Check if the sheet exists
    sheetName = "Vdd_Binning_Def_appA_" & passBinCut & "_" & TheExec.CurrentJob
    Set wb = Application.ActiveWorkbook
    Call check_Sheet_Range(sheetName, wb, ws_def, MaxRow, maxcol, isSheetFound, False)
    If isSheetFound = False Then                        ''' if "Vdd_Binning_Def_appA_2/3_JOBNAME" does not exist
        sheetName = "Vdd_Binning_Def_appA_" & passBinCut
        Call check_Sheet_Range(sheetName, wb, ws_def, MaxRow, maxcol, isSheetFound)
    End If
    '''*****************************************************************'''
    MaxPerformanceModeCount = cntVddbinPmode + 1 + cntEqVoltagePmode '20230526
    ReDim col_dynamic_offset(BinCutJobDefinition.Total_Count, testType.Total_Count) As Integer 'Dim col_dynamic_offset(MaxJobCountInVbt, MaxTestType) As Integer 20230526
    If isSheetFound = True Then
        '''//init
        '''Since all col_XXX and row_XXX related variables with default values=0, no need to initialize them as 0.
        start_p_mode = -1
        stop_p_mode = -1
        
        '''//Check the header of the table
        '''Get the columns for the diverse coefficient.
        For Row = 1 To MaxRow
            For Col = 1 To maxcol
                '''******************************************************************************************************************'''
                '''//If CorePower and OtherRail are in the same table (only Vdd_Binning_Def), 1st column is "Binned".
                '''//If CorePower and OtherRail are in the different tables (Vdd_Binning_Def and Other_Rail), 1st column is "Domain".
                '''******************************************************************************************************************'''
                '''If 1st column 1 of the header is "Binned", split the line and find out the keyword column.
                If LCase(ws_def.Cells(Row, Col).value) Like "binned" Then
                    col_binned = Col
                    Row_of_Title = Row
                End If
                                    
                If Row_of_Title > 0 Then
                    If LCase(ws_def.Cells(Row_of_Title, Col).value) = "domain" Then
                        col_domain = Col
                    ElseIf LCase(ws_def.Cells(Row_of_Title, Col).value) = "mode" Then
                        col_mode = Col
                    '20230504: Added for parsing sheet defining BinX EQNs and IDS value
                    ElseIf LCase(ws_def.Cells(Row_of_Title, Col).value) = "eqn_bin" Then
                        col_eqn_bin = Col
                    ElseIf LCase(ws_def.Cells(Row_of_Title, Col).value) = "binx_cpidsmax" Then
                        col_binx_ids = Col
                    ElseIf LCase(ws_def.Cells(Row_of_Title, Col).value) = "binx_idsmax_hot" Then
                        col_binx_ftids = Col
                    ElseIf LCase(ws_def.Cells(Row_of_Title, Col).value) = "c" Then
                        col_c = Col
                    '''********************************************************'''
                    ElseIf LCase(ws_def.Cells(Row_of_Title, Col).value) = "cpidsmax" Then
                        col_cpids = Col
                    ElseIf LCase(ws_def.Cells(Row_of_Title, Col).value) = "idsmax_hot" Or LCase(ws_def.Cells(Row_of_Title, Col).value) = "ftids" Then
                        col_ftids = Col
                    '''********************************************************'''
                    '''20210819: Modified to revise the vbt code for the new format of the header in BinCut voltage tables, requested by C651 Toby.
                    ElseIf LCase(ws_def.Cells(Row_of_Title, Col).value) = "cpvmax" Or LCase(ws_def.Cells(Row_of_Title, Col).value) = LCase("BinningVmax") Then
                        col_cp_vmax = Col
                    ElseIf LCase(ws_def.Cells(Row_of_Title, Col).value) = "cpvmin" Or LCase(ws_def.Cells(Row_of_Title, Col).value) = LCase("BinningVmin") Then
                        col_cp_vmin = Col
                    ElseIf LCase(ws_def.Cells(Row_of_Title, Col).value) = "cpgb" Or LCase(ws_def.Cells(Row_of_Title, Col).value) = LCase("BinningGB") Then
                        col_cpgb = Col
                    ElseIf LCase(ws_def.Cells(Row_of_Title, Col).value) = "cp2gb" Or LCase(ws_def.Cells(Row_of_Title, Col).value) = LCase("CP_GB_HOT") Then
                        col_cp2gb = Col
                    ElseIf LCase(ws_def.Cells(Row_of_Title, Col).value) = "ft1gb" Or LCase(ws_def.Cells(Row, Col).value) = "ft_gb_room" Then
                        col_ft1gb = Col
                    ElseIf LCase(ws_def.Cells(Row_of_Title, Col).value) = "ft2gb" Or LCase(ws_def.Cells(Row, Col).value) = "ft_gb_hot" Then
                        col_ft2gb = Col
                    ElseIf LCase(ws_def.Cells(Row_of_Title, Col).value) = "sltgb" Then
                        col_sltgb = Col
                    ElseIf LCase(ws_def.Cells(Row_of_Title, Col).value) = "htol_ro_gb" Then
                        col_htol_ro_gb = Col
                    ElseIf LCase(ws_def.Cells(Row_of_Title, Col).value) = "htol_ro_gb_room" Or LCase(ws_def.Cells(Row_of_Title, Col).value) = LCase("HTOL_T0TX_GB_ROOM") Then
                        col_htol_ro_gb_room = Col
                    ElseIf LCase(ws_def.Cells(Row_of_Title, Col).value) = "htol_ro_gb_hot" Or LCase(ws_def.Cells(Row_of_Title, Col).value) = LCase("HTOL_T0TX_GB_HOT") Then
                        col_htol_ro_gb_hot = Col
                    ElseIf LCase(ws_def.Cells(Row_of_Title, Col).value) = "ate_fqagb" Then
                        col_ate_ftqa_gb = Col
                    ElseIf LCase(ws_def.Cells(Row_of_Title, Col).value) = "slt_fqa_gb" Then
                        col_slt_ftqa_gb = Col
                    ElseIf LCase(ws_def.Cells(Row_of_Title, Col).value) = "cphv" Then
                        col_cphv = Col
                    ElseIf LCase(ws_def.Cells(Row_of_Title, Col).value) = "fthv" Then
                        col_fthv = Col
                    ElseIf LCase(ws_def.Cells(Row_of_Title, Col).value) = "qahv" Then
                        col_qahv = Col
                    ElseIf LCase(ws_def.Cells(Row_of_Title, Col).value) Like "offset_*_*" Then
                        jobIdx = getBinCutJobDefinition(LCase(ws_def.Cells(Row_of_Title, Col).value))
                        testTypeIdx = decide_offset_testType(LCase(ws_def.Cells(Row_of_Title, Col).value))
                        col_dynamic_offset(jobIdx, testTypeIdx) = Col
                    ''' 20211202: For CPU_SRAM IDS Mapping, determines which IDS value to use and to calculate C-M*log10(IDS)
                    ''' 202307: Added for other rail
                    ElseIf LCase(ws_def.Cells(Row_of_Title, Col).value) Like "*bincutcalc_idsrail*" Then
                        Flag_IDS_Mapping_enable = True          '' when col = "bincutcalc_idsrail", Turn the flag true
                        col_ids_mapping = Col
                    ElseIf LCase(ws_def.Cells(Row_of_Title, Col).value) Like "*bincutcalc_idsmax*" Then
                        col_ids_mapping_max = Col
                    End If
                End If
                
                '''//Check if all columns of the header exist...
                '''Note: col_comment should be checked as the last available column of the table.
                If col_domain > 0 And col_mode > 0 And col_c > 0 And col_cpgb > 0 _
                And col_cp2gb > 0 And col_ft1gb > 0 And col_ft2gb > 0 And col_sltgb > 0 _
                And col_ate_ftqa_gb > 0 And col_slt_ftqa_gb > 0 _
                And col_cphv > 0 And col_fthv > 0 And col_qahv > 0 And col_cp_vmax > 0 _
                And col_cp_vmin > 0 And col_cpids > 0 And col_ftids > 0 Then
                    enableRowParsing = True
                End If
            Next Col
            
            '''//If all columns of the header are found, skip the loop and start parsing each row.
            If enableRowParsing = True Then
                Exit For
            End If
            
            If Row = MaxRow And (col_binned = 0 Or col_domain = 0 Or col_mode = 0 Or col_cpids = 0) Then
                enableRowParsing = False
                Call Print_Error_Message(Error_Info, "LIB_VDD_BINNING", "initVddotherrailOneMod", "Columns of header in " & sheetName & " are incorrect. Error!!!")
                'TheExec.ErrorLogMessage "Columns of header in " & sheetName & " are incorrect. Error!!!"    ''20230807: Removed by sanity check
            End If
        Next Row
        
        If enableRowParsing = True And Row_of_Title + 1 <= MaxRow Then
            For Row = Row_of_Title + 1 To MaxRow
                '''//If column "Binned" is "false" or "ate", it means that power Domain is for OtherRail, ex: SRAM_*** and fixed and low...
                If (LCase(ws_def.Cells(Row, col_binned).value) = "false" Or LCase(ws_def.Cells(Row, col_binned).value) = "ate") Then
                    '''=====================================================================================
                    '''[Step1] Get OtherRail from Domain column.
                    '''//performance_mode is enumerated into the p_mode dictionary when "initVddBinTable".
                    '''=====================================================================================
                    If UCase(Trim(ws_def.Cells(Row, col_domain).value)) Like "VDD*" Then '''ex: "VDD_CPU_SRAM"
                        powerDomain = UCase(Trim(ws_def.Cells(Row, col_domain).value))
                        other_p_mode = VddBinStr2Enum(powerDomain) '//p_mode
                        AllBinCut(other_p_mode).powerPin = powerDomain
                        
                    ElseIf UCase(ws_def.Cells(Row, col_domain).value) <> "" Then '''ex: "CPU_SRAM"
                        powerDomain = UCase("VDD_" & Trim(ws_def.Cells(Row, col_domain).value))
                        other_p_mode = VddBinStr2Enum(powerDomain) '//p_mode
                        AllBinCut(other_p_mode).powerPin = powerDomain
                    Else
                        Call Print_Error_Message(Error_Info, "LIB_VDD_BINNING", "initVddotherrailOneMod", ws_def.Cells(Row, col_domain) & " doesn't have the correct Domain cell in sheet " & ws_def & ". Error!!!")
                        'TheExec.Datalog.WriteComment ws_def.Cells(Row, col_domain) & " doesn't have the correct Domain cell in sheet " & ws_def & ". Error!!!"
                        'TheExec.ErrorLogMessage ws_def.Cells(Row, col_domain) & " doesn't have the correct Domain cell in sheet " & ws_def & ". Error!!!"    ''20230807: Removed by sanity check
                    End If
                    
                    '''=====================================================================================
                    '''[Step2] Use "dict_IsCorePower" to check if powerDomain is BinCut CorePower/OtherRail listed in Vdd_Binning_Def_appA_1.
                    '''=====================================================================================
                    If dict_IsCorePower.Exists(UCase(powerDomain)) = True Then
                        If dict_IsCorePower.item(UCase(powerDomain)) = False Then
                            '''Do nothing...
                        Else
                            Call Print_Error_Message(Error_Info, "LIB_VDD_BINNING", "initVddotherrailOneMod", "column Binned of " & ws_def.Cells(Row, col_domain) & "  in sheet " & sheetName & " isn't consistent with Vdd_Binning_Def sheet_appA_1. Error!!!")
                            'TheExec.ErrorLogMessage "column Binned of " & ws_def.Cells(Row, col_domain) & "  in sheet " & sheetName & " isn't consistent with Vdd_Binning_Def sheet_appA_1. Error!!!"    ''20230807: Removed by sanity check
                        End If
                    Else
                        Call Print_Error_Message(Error_Info, "LIB_VDD_BINNING", "initVddotherrailOneMod", ws_def.Cells(Row, col_domain) & "  in sheet " & sheetName & " doesn't show in other Vdd_Binning_Def sheet. Error!!!")
                        'TheExec.Datalog.WriteComment ws_def.Cells(Row, col_domain) & "  in sheet " & sheetName & " doesn't show in other Vdd_Binning_Def sheet. Error!!!"
                        'TheExec.ErrorLogMessage ws_def.Cells(Row, col_domain) & "  in sheet " & sheetName & " doesn't show in other Vdd_Binning_Def sheet. Error!!!"    ''20230807: Removed by sanity check
                    End If
                    
                    '''=====================================================================================
                    '''[Step3] Check if mode column with/without P_mode.
                    '''=====================================================================================
                    If ws_def.Cells(Row, col_domain).value = ws_def.Cells(Row, col_mode).value Then '''//If Mode is same as Domain, ex: "SRAM_CPU"
                        '''//If SRAM without p_mode, col_domain and col_mode are the same.
                        start_p_mode = 0
                        stop_p_mode = MaxPerformanceModeCount - 1
                    ElseIf ws_def.Cells(Row, col_mode).value Like "M*##*" Then '''ex: "MC60A", "MCS60A"
                        '''//If SRAM with p_mode, col_mode has the p_mode related to CorePower, ex: "MC60A".
                        If Len(ws_def.Cells(Row, col_mode).value) = 6 Then
                            start_p_mode = VddBinStr2Enum(mid(ws_def.Cells(Row, col_mode).value, 1, 2) & mid(ws_def.Cells(Row, col_mode).value, 4, 3))
                            stop_p_mode = start_p_mode
                        Else
                            start_p_mode = VddBinStr2Enum(ws_def.Cells(Row, col_mode).value)
                            stop_p_mode = start_p_mode
                        End If
                    Else
                        Call Print_Error_Message(Error_Info, "LIB_VDD_BINNING", "initVddotherrailOneMod", ws_def.Cells(Row, col_domain).value & " doesn't have the correct Domain cell in sheet " & ws_def & ". Error!!!")
                        'TheExec.Datalog.WriteComment ws_def.Cells(Row, col_domain).value & " doesn't have the correct Domain cell in sheet " & ws_def & ". Error!!!"
                        'TheExec.ErrorLogMessage ws_def.Cells(Row, col_domain).value & " doesn't have the correct Domain cell in sheet " & ws_def & ". Error!!!"    ''20230807: Removed by sanity check
                    End If
                    
                    For p_mode = start_p_mode To stop_p_mode
                        BinCut(p_mode, passBinCut).OTHER_PRODUCT_RAIL(other_p_mode) = CDbl(ws_def.Cells(Row, col_c).value) + CDbl(ws_def.Cells(Row, col_cpgb).value)
                        BinCut(p_mode, passBinCut).OTHER_CP1_RAIL(other_p_mode) = CDbl(ws_def.Cells(Row, col_c).value)
                        BinCut(p_mode, passBinCut).OTHER_CP1_GB(other_p_mode) = CDbl(ws_def.Cells(Row, col_cpgb).value) 'added for CP1_GB, 20190311
                        BinCut(p_mode, passBinCut).OTHER_CP2_GB(other_p_mode) = CDbl(ws_def.Cells(Row, col_cp2gb).value)
                        BinCut(p_mode, passBinCut).OTHER_FT1_GB(other_p_mode) = CDbl(ws_def.Cells(Row, col_ft1gb).value)
                        BinCut(p_mode, passBinCut).OTHER_FT2_GB(other_p_mode) = CDbl(ws_def.Cells(Row, col_ft2gb).value)
                        BinCut(p_mode, passBinCut).OTHER_SLT_GB(other_p_mode) = CDbl(ws_def.Cells(Row, col_sltgb).value)
                        '''20230706: Added to support binX merge into bin1 feature
                        '''20230706: When binX merge feature is enabled, col_binx_ids contains MAX ids value
                        If Flag_BinX_Info_Parsed = True Then
                            BinCut(p_mode, passBinCut).OTHER_CPIDS(other_p_mode) = CDbl(ws_def.Cells(Row, col_binx_ids).value)
                            ''Save bin1 limit to dictionary
                            If CDbl(Trim(ws_def.Cells(Row, col_cpids).value)) <> 0 Then
                                If Pmode2BinX_IDS_Dict.Exists(other_p_mode) = False Then
                                    Pmode2BinX_IDS_Dict.Add other_p_mode, CDbl(Trim(ws_def.Cells(Row, col_cpids).value))                 ''Save dict: pmode(80) -> 38.8(mA)
                                Else
                                    If Pmode2BinX_IDS_Dict(other_p_mode) <> CDbl(Trim(ws_def.Cells(Row, col_cpids).value)) Then
                                        TheExec.Datalog.WriteComment "'CPIDSMax' is not the same in pmode:" & VddBinName(other_p_mode) & ", ERROR!!"
                                    End If
                                End If
                            End If
                            
                            ''20240118: Added to save IDS hot for binX merge feature
                            If CDbl(Trim(ws_def.Cells(Row, col_ftids).value)) <> 0 Then
                                If Pmode2BinX_IDSHot_Dict.Exists(other_p_mode) = False Then
                                    Pmode2BinX_IDSHot_Dict.Add other_p_mode, CDbl(Trim(ws_def.Cells(Row, col_ftids).value))                 ''Save dict: pmode(80) -> 38.8(mA)
                                Else
                                    If Pmode2BinX_IDSHot_Dict(other_p_mode) <> CDbl(Trim(ws_def.Cells(Row, col_ftids).value)) Then
                                        TheExec.Datalog.WriteComment "'IDSMax_HOT' is not the same in pmode:" & VddBinName(other_p_mode) & ", ERROR!!"
                                    End If
                                End If
                            End If
                        Else
                            BinCut(p_mode, passBinCut).OTHER_CPIDS(other_p_mode) = CDbl(ws_def.Cells(Row, col_cpids).value)
                        End If
                        BinCut(p_mode, passBinCut).OTHER_FTIDS(other_p_mode) = CDbl(ws_def.Cells(Row, col_ftids).value)
                        BinCut(p_mode, passBinCut).OTHER_ATE_FQA_GB(other_p_mode) = CDbl(ws_def.Cells(Row, col_ate_ftqa_gb).value)
                        
                        If col_htol_ro_gb > 0 Then
                            BinCut(p_mode, passBinCut).OTHER_HTOL_RO_GB(other_p_mode) = CDbl(ws_def.Cells(Row, col_htol_ro_gb).value)
                        End If
                        
                        If col_htol_ro_gb_room > 0 Then
                            If TheExec.flow.enableWord("HTOL_TX_ROOM") = True Then
                                BinCut(p_mode, passBinCut).OTHER_FT1_GB(other_p_mode) = CDbl(ws_def.Cells(Row, col_htol_ro_gb_room).value)
                            End If
                            BinCut(p_mode, passBinCut).OTHER_HTOL_RO_GB_ROOM(other_p_mode) = CDbl(ws_def.Cells(Row, col_htol_ro_gb_room).value)
                        End If
                        
                        If col_htol_ro_gb_hot > 0 Then
                            If TheExec.flow.enableWord("HTOL_TX_HOT") = True Then
                                BinCut(p_mode, passBinCut).OTHER_FT2_GB(other_p_mode) = CDbl(ws_def.Cells(Row, col_htol_ro_gb_hot).value)
                            End If
                            BinCut(p_mode, passBinCut).OTHER_HTOL_RO_GB_HOT(other_p_mode) = CDbl(ws_def.Cells(Row, col_htol_ro_gb_hot).value)
                        End If
                        
                        BinCut(p_mode, passBinCut).OTHER_SLT_FQA_GB(other_p_mode) = CDbl(ws_def.Cells(Row, col_slt_ftqa_gb).value)
                        BinCut(p_mode, passBinCut).HVCC_OTHER_CP_RAIL(other_p_mode) = CDbl(ws_def.Cells(Row, col_cphv).value)
                        BinCut(p_mode, passBinCut).HVCC_OTHER_FT_RAIL(other_p_mode) = CDbl(ws_def.Cells(Row, col_fthv).value)
                        BinCut(p_mode, passBinCut).HVCC_OTHER_QA_RAIL(other_p_mode) = CDbl(ws_def.Cells(Row, col_qahv).value)
                        BinCut(p_mode, passBinCut).OTHER_CP_Vmax(other_p_mode) = CDbl(ws_def.Cells(Row, col_cp_vmax).value)
                        BinCut(p_mode, passBinCut).OTHER_CP_Vmin(other_p_mode) = CDbl(ws_def.Cells(Row, col_cp_vmin).value)
                        
                        '''****************************************************************************************'''
                        '''//Parsing dynamic_offset for OtherRail(request by C651).
                        '''//For storing the values, we take the domain column as OtherRail performance mode index.
                        For jobIdx = 0 To BinCutJobDefinition.Total_Count
                            For testTypeIdx = 0 To testType.Total_Count - 1
                                If (col_dynamic_offset(jobIdx, testTypeIdx) <> 0) Then
                                    BinCut(other_p_mode, passBinCut).DYNAMIC_OFFSET(jobIdx, testTypeIdx) = CDbl(ws_def.Cells(Row, col_dynamic_offset(jobIdx, testTypeIdx)).value)
                                End If
                            Next testTypeIdx
                        Next jobIdx
                        '''****************************************************************************************'''
                    Next p_mode
                    
                    '''=====================================================================================
                    '''[Step4] According to test jobs and PassBinCut, decide the IDS limit for OtherRail.
                    '''=====================================================================================
                    '''20230706: Added to support binX merge into bin1 feature
                    '''20230706: When binX merge feature is enabled, col_binx_ids contains MAX ids value
                    If Flag_BinX_Info_Parsed = True Then
                        CPIDS_Spec(VddBinStr2Enum(powerDomain), passBinCut) = CDbl(ws_def.Cells(Row, col_binx_ids).value)
                        FTIDS_Spec(VddBinStr2Enum(powerDomain), passBinCut) = CDbl(ws_def.Cells(Row, col_binx_ftids).value)
                    Else
                        CPIDS_Spec(VddBinStr2Enum(powerDomain), passBinCut) = CDbl(ws_def.Cells(Row, col_cpids).value)
                        FTIDS_Spec(VddBinStr2Enum(powerDomain), passBinCut) = CDbl(ws_def.Cells(Row, col_ftids).value)
                    End If
                    
                    
                    '' 202307: Added for other rail ids mapping
                    If Flag_IDS_Mapping_enable = True Then
                        If col_ids_mapping > 0 And col_ids_mapping_max > 0 Then
                            If ws_def.Cells(Row, col_ids_mapping).value <> "" Then
                                AllBinCut(other_p_mode).IDS_MAPPING = "VDD_" & ws_def.Cells(Row, col_ids_mapping).value              '' record which IDS to map
                                AllBinCut(other_p_mode).IDS_MAPPING_LIMIT = CDbl(ws_def.Cells(Row, col_ids_mapping_max).value)           '' record the limit of the IDS (No use for now)
                            Else
                                AllBinCut(other_p_mode).IDS_MAPPING = AllBinCut(other_p_mode).powerPin
                            End If
                        Else
                            TheExec.Datalog.WriteComment "No IDS mapping is determined, but Flag_IDS_Mapping_enable = True. ERROR !!"
                        End If
                    End If
                    
                    '''//Choose IDS hi_limit by BinCut testjobs.
                    '''//IDS calculation uses the scale and the unit in "mA".
                    '''20210720: Modified to revise ids_hi_limit with CPIDSMax because C651 Si asked us to use Efuse processed IDS for search in FT.
                    '''<org>
'                    If bincutJobName = "cp1" Or bincutJobName = "ft_room" Or bincutJobName = "qa" Then '''for testjobs with the normal temperature 25C
'                        ids_hi_limit(VddBinStr2Enum(powerDomain), passBinCut) = CDbl(ws_def.Cells(row, col_cpids).Value)
'                    ElseIf bincutJobName = "cp2" Or bincutJobName = "ft_hot" Then '''for testjobs with the high temperature 85C
'                        ids_hi_limit(VddBinStr2Enum(powerDomain), passBinCut) = CDbl(ws_def.Cells(row, col_ftids).Value)
'                    End If

                    '''<new>
                    '''use CPIDSMax as IDS_hi_limit for Efuse processed IDS.
                    '''20230706: Added to support binX merge into bin1 feature
                    '''20230706: When binX merge feature is enabled, col_binx_ids contains MAX ids value
                    If Flag_BinX_Info_Parsed = True Then
                        gb_IDS_hi_limit(VddBinStr2Enum(powerDomain), passBinCut) = CDbl(ws_def.Cells(Row, col_binx_ids).value)
                        If AllBinCut(other_p_mode).IDS_CP_LIMIT = 0 Then
                            AllBinCut(other_p_mode).IDS_CP_LIMIT = CDbl(ws_def.Cells(Row, col_binx_ids).value)
                        ElseIf AllBinCut(VddBinStr2Enum(powerDomain)).IDS_CP_LIMIT < CDbl(ws_def.Cells(Row, col_binx_ids).value) Then
                            AllBinCut(other_p_mode).IDS_CP_LIMIT = CDbl(ws_def.Cells(Row, col_binx_ids).value)
                        End If
                    Else
                        gb_IDS_hi_limit(VddBinStr2Enum(powerDomain), passBinCut) = CDbl(ws_def.Cells(Row, col_cpids).value)
                        If AllBinCut(other_p_mode).IDS_CP_LIMIT = 0 Then
                            AllBinCut(other_p_mode).IDS_CP_LIMIT = CDbl(ws_def.Cells(Row, col_cpids).value)
                        ElseIf AllBinCut(VddBinStr2Enum(powerDomain)).IDS_CP_LIMIT < CDbl(ws_def.Cells(Row, col_cpids).value) Then
                            AllBinCut(other_p_mode).IDS_CP_LIMIT = CDbl(ws_def.Cells(Row, col_cpids).value)
                        End If
                    End If

                    If AllBinCut(other_p_mode).IDS_FT_LIMIT = 0 Then
                        AllBinCut(other_p_mode).IDS_FT_LIMIT = CDbl(ws_def.Cells(Row, col_ftids).value)
                    ElseIf AllBinCut(VddBinStr2Enum(powerDomain)).IDS_FT_LIMIT < CDbl(ws_def.Cells(Row, col_ftids).value) Then
                        AllBinCut(other_p_mode).IDS_FT_LIMIT = CDbl(ws_def.Cells(Row, col_ftids).value)
                    End If
                End If
            Next Row
        Else
            Call Print_Error_Message(Error_Info, "LIB_VDD_BINNING", "initVddotherrailOneMod", "Columns of the header in the sheet " & sheetName & " might be incorrect. Error!!!")
            'TheExec.ErrorLogMessage "Columns of the header in the sheet " & sheetName & " might be incorrect. Error!!!"    ''20230807: Removed by sanity check
        End If
    End If '''If isSheetFound = True
Exit Function
errHandler:
    Call Print_Error_Message(Error_Info, "LIB_VDD_BINNING", "initVddotherrailOneMod") 'Add ErrHandler 2023/08/01
    If AbortTest Then Exit Function Else Resume Next
End Function

'20210819: As per discussion with Jeff, he suggested us to use the keyword "Binning Domain" for each testJob.
'20210819: Modified to move the vbt code about resetting globalVariables of BinCut testCondition from the vbt function initVddBinCondition to the vbt function Reset_BinCut_GlobalVariable_for_initVddBinning.
'20210819: Modified to revise the vbt code for the new format of the header in BinCut flow table, as requested by C651 Toby.
'20210819: Modified to assemble job_keyword by bincutJobName according to the vbt function Mapping_TestJobName_to_BincutJobName.
'20210802: Modified to check if testCondition contains any keyword about PassBin(Bin1/BinX/BinY) greater than the highest bin number.
'20210414: Modified to add "is_for_BinSearch as Boolean" for AllBinCut(p_mode).
'20210201: Modified to check if testCondtion with performance mode for SRAM_Vth, ex: "640mv (MI003)".
'20210131: Modified to check "UCase(Trim(ws_def.Cells(j, Col).value))".
'20210121: Modified to check the format of testConditions, request by TSMC ZQLIN.
'20201222: Modified to use the dictionary "dict_OutsideBinCut_additionalMode" to check if any duplicate additional mode exists in different Outside BinCut flow tables.
'20201222: Modified to revise the vbt function "initVddBinCondition" for multiple "Non_Binning_Rail_Outside_BinCut" sheets.
'20201215: Modified to check if testCondition contain keyword "*Evaluate Bin*" to decide "is_BinCutJob_for_StepSearch" while isParsingOutsideBinCutFlow = False.
'20201210: Modified to use the flag "is_BinCutJob_for_StepSearch" for "check_bincutJob_for_StepSearch" to check if the test program is binSearch or functional test.
'20201112: Modified to use the dictionary "dict_IsCorePowerInBinCutFlowSheet".
'20201023: Modified to store all headers into the dictionary "dict_BinCutFlow_Domain2Column" by "col_jobBlock_start" and "col_jobBlock_end"...
'20201023: Modified to initialize the array of BV and HBV testConditions with the empty string "".
'20201022: Modified to reduce the complity of the loop to find row of the header VDD.
'20201022: Modified to modify the parsing method for "FullCorePowerinFlowSheet" and "FullOtherRailinFlowSheet".
'20201021: As per discussion with TSMC PCLINZG, he suggested us to use the same testCondition for outsideBinCutFlow BV and HBV.
'20201021: Modified to support multiple columns with "IGNORE COLUMN".
'20201021: Modified to use "dict_IsCorePower" to store and check CorePower/OtherRail.
'20201021: Modified to revise the vbt code for parsing "Non_Binning_Rail" and "Non_Binning_Rail_outside_BinCut".
'20201013: Modified to trim string of the testCondition from Non_Binning_Rail.
'20201005: Modified to check "AllBinCut(p_mode).INTP_SKIPTEST = True" to update "AllBinCut(p_mode).used".
'20200827: Modified to check the mapping BinCut testJob.
'20200827: Modified to replace "If..Else" with "Select Case".
'20200711: Modified to check if any testcondition contains "#REF!" or "#NAME?".
'20200703: Modiifed to use "check_Sheet_Range".
'20200611: Modified to check "IGNORE COLUMN".
'20200211: Modified to replace "FlowTestCondName" with "AdditionalModeName".
'20200211: Modified to replace "cntFlowTestCond" with "cntAdditionalMode".
'20191227: Modified to check if allbincut(pmode).used is true.
'20191219: Modified for Domain2Pin and Pin2Domain.
'20191218: Modified to check if trackpower is "N/C" or not.
'20191204: Modified to check if pinGroup_BinCut exists in BV and HBV columns.
'20191129: Modified to check "IGNORE COLUMN".
'20191128: Modified for checking the additional mode (flowTestCondition).
'20191127: Modified for the revised InitVddBinTable.
'20191014: Modified to parse the table with different powerPins when different testjobs.
'20190905: Modified for those projects with bin1/binx in the same block of "Non_Binning_Rail" sheet.
'20190706: Modified to check if the powerPin (Domain) is in the pin_group "ShtCorePowerinFlowSheet" or "ShtOtherRailinFlowSheet".
'20190704: Modified to remove the hard-code "power_list" and "power_seq".
'20190704: As the discussion with SWLINZA, we should check digit1-2 of the p_mode are same as p_modes in the power_list.
'20190521: Modified for the correct column numbers.
'20190426: Modified to use the function "Find_Sheet".
'20190319: Modified for assemblying the BinCut powerpin group "FullBinCutPowerinFlowSheet"
'20190117: Modified for checking if powerpin exists in pinmap and channelmap.
'20181209: Modified for blocking unlisted/untested performance modes.
'20180723: Modified for BinCut testjob mapping.
Public Function initVddBinCondition(sheetName As String)
On Error GoTo errHandler 'Add ErrHandler 2023/08/01
    Dim sheetName_OutsideBinCut As String
    Dim sheetName_EquationVoltages As String
    Dim sheetName_temp As String
    Dim ws_def As Worksheet
    Dim wb As Workbook
    Dim Row As Long, Col As Long
    Dim MaxRow As Long
    Dim maxcol As Long
    Dim split_content() As String
    Dim performance_mode As String
    Dim additional_mode As String
    Dim main_p_mode As Integer
    Dim p_mode As Integer
    Dim addi_mode As Integer '''For the additional mode
    Dim i As Long, j As Long, k As Long, L As Long
    Dim S As Long
    Dim other_voltage_start_point As Long
    Dim HVCC_flag As Long
    Dim passBinCut As Long
    '''for testjob mapping
    Dim Row_of_Title As Long
    Dim row_of_testJob As Long
    Dim job_keyword As String
    Dim row_jobBlock As Long
    '''
    Dim cnt_testJob As Long
    Dim idx_testJob As Long
    Dim col_testJob() As Long
    Dim col_jobBlock_start As Long
    Dim col_jobBlock_end As Long
    '''for Performance Mode
    Dim col_mode As Long
    '''for powerDomain
    Dim got_correct_header As Boolean
    Dim got_CorrectDomain As Boolean
    '''for trackPower
    Dim trackpowerTemp As String
    Dim strAry_trackpower() As String
    '''variables
    Dim COREPOWER As Long
    Dim powerDomain As String
    Dim selected_powerDomain As String
    Dim str_mode_temp As String
    Dim binNumStart As Long
    Dim binNumStop As Long
    Dim isSheetFound As Boolean
    Dim isParsingOutsideBinCutFlow As Boolean
    Dim isIgnoreColumn As Boolean
    Dim testCondition As String
    Dim str_mainColumn_content As String
    Dim strTemp As String
    Dim strSplitted() As String
    Dim bincutNum As Long
'''//==================================================================================================================================================================================//'''
'''//Note:
'''1. Non_Binning_Rail should be parsed prior to sheet Non_Binning_Rail_Post_BinCut.
'''2. As per discussion with TSMC PCLINZG, he suggested us to use the same testCondition for outsideBinCutFlow BV and HBV.
'''3. Please check keyword "Non_Binning_Rail_Outside" of sheetName for the vbt functions "initVddBinCondition" and "parsing_OutsideBinCut_flow_table".
'''//==================================================================================================================================================================================//'''
    '''*****************************************************************'''
    '''//Check if the sheet exists
    'sheetName = "Non_Binning_Rail"
'    dict_BinCutFlow_Domain2Column.RemoveAll '''20220110
'    dict_BinCutFlow_Column2Domain.RemoveAll '''20220110
    FullBinCutPowerinFlowSheet_outside = vbNullString '''20220110
    sheetName_temp = sheetName
    sheetName = sheetName & "_" & TheExec.CurrentJob
    sheetName_OutsideBinCut = "Non_Binning_Rail_Outside" '''for outsiteBinCut.
    Set wb = Application.ActiveWorkbook
    Call check_Sheet_Range(sheetName, wb, ws_def, MaxRow, maxcol, isSheetFound, False)
    If isSheetFound = False Then                        ''' If "Non_Binning_Rail_JOBNAME" does not exist
        sheetName = sheetName_temp
        Call check_Sheet_Range(sheetName, wb, ws_def, MaxRow, maxcol, isSheetFound)
    End If
    If gS_Non_Binning_Rail_sheet = "" Then
        gS_Non_Binning_Rail_sheet = sheetName       '' 20211224: Added to Print Bincut Sheet Name in Print_Bincut_Config
    End If
    '''*****************************************************************'''
    If isSheetFound = True Then
        '''//init
        '''Since all col_XXX and row_XXX related variables with default values=0, no need to initialize them as 0.
        binNumStart = 1
        binNumStop = Total_Bincut_Num
        HVCC_flag = 0
        other_voltage_start_point = 0
        trackpowerTemp = vbNullString
        Row_of_Title = 0
        row_of_testJob = 0
        cnt_testJob = 0
        idx_testJob = -1
        got_correct_header = True
        isIgnoreColumn = False
        got_CorrectDomain = False
        str_mainColumn_content = vbNullString
        str_mode_temp = vbNullString
        performance_mode = vbNullString
        additional_mode = vbNullString
        job_keyword = vbNullString
        col_jobBlock_start = 0
        col_jobBlock_end = 0
        col_mode = 0
        
        '''//Check if sheet "Non_Binning_Rail should be parsed prior to sheet Non_Binning_Rail_Post_BinCut"
        If LCase(sheetName) Like LCase("*" & sheetName_OutsideBinCut & "*") Then
            If dict_BinCutFlow_Domain2Column.Count <> 0 Then
                isParsingOutsideBinCutFlow = True '''Parsing outsideBinCutFlow...
            Else
                isParsingOutsideBinCutFlow = False '''Parsing BinCutFlow (sheet "Non_Binning_Rail")
                TheExec.Datalog.WriteComment "Non_Binning_Rail should be parsed prior to sheet Non_Binning_Rail_Outside_BinCut"
                TheExec.ErrorLogMessage "Non_Binning_Rail should be parsed prior to sheet Non_Binning_Rail_Outside_BinCut"
            End If
        Else
            isParsingOutsideBinCutFlow = False '''//Parsing BinCutFlow (sheet "Non_Binning_Rail")
        End If
        '''20210819: Modified to move the vbt code about resetting globalVariables of BinCut testCondition from the vbt function initVddBinCondition to the vbt function Reset_BinCut_GlobalVariable_for_initVddBinning.
        
        '''//Get keyword for BinCut testJob mapping.
        '''20210819: Modified to assemble job_keyword by bincutJobName according to the vbt function Mapping_TestJobName_to_BincutJobName.
        job_keyword = LCase("*" & bincutJobName & "*") '''ex: "*cp1*", "*cp2*", "*ft_room*", "*ft_hot*", "*qa*".
        If TheExec.enableWord("T0TX_ROOM") = True Then job_keyword = LCase("T0TX_ROOM")  'WWWWW20220631 T0TX
        If TheExec.enableWord("T0TX_HOT") = True And Not TheExec.enableWord("HardIP_RMA") = True Then job_keyword = LCase("T0TX_HOT")   'WWWWW20220631 T0TX
        
    Else
        Exit Function
    End If
    
    '''//Find the keyword of BinCut testjob, and find column of the selected testJob.
    '''20210819: As per discussion with Jeff, he suggested us to use the keyword "Binning Domain" for each testJob.
    If job_keyword <> "" Then
        For Row = 1 To MaxRow
            For Col = 1 To maxcol
                If LCase(ws_def.Cells(Row, Col).value) Like LCase("Binning Domain") Then
                    ReDim Preserve col_testJob(cnt_testJob)
                    col_testJob(cnt_testJob) = Col
                    cnt_testJob = cnt_testJob + 1
                    
                    If row_of_testJob = 0 Then
                        row_of_testJob = Row
                    End If
                End If
            Next Col
            
            If row_of_testJob > 0 Then
                Exit For
            End If
        Next Row
    End If
    
    '''//Check if any matched block for IGXL Job.
    If cnt_testJob > 0 Then
        For i = 0 To cnt_testJob - 1
            If LCase(ws_def.Cells(row_of_testJob, col_testJob(i) + 2).value) Like LCase(job_keyword) Then
                idx_testJob = i
                other_voltage_start_point = col_testJob(i) + 2
                
                '''//Get start/stop columns for block of the selected testJob.
                col_jobBlock_start = col_testJob(idx_testJob) '''column of Domain
                If idx_testJob = UBound(col_testJob) Then
                    col_jobBlock_end = maxcol
                Else
                    col_jobBlock_end = col_testJob(idx_testJob + 1) - 1
                End If
                
                '''//Get column of "Performance Mode"
                For Col = col_jobBlock_start To col_jobBlock_end
                    If LCase(Trim(ws_def.Cells(row_of_testJob, Col))) Like LCase("Performance Mode") Then
                        col_mode = Col
                        
                        '''//Check if column "Performance mode" is defined in the dictionary dict_BinCutFlow_Domain2Column.
                        If dict_BinCutFlow_Domain2Column.Exists(UCase("Performance mode")) = True Then
                            If Col = dict_BinCutFlow_Domain2Column.item(UCase("Performance mode")) Then
                                '''Do nothing...
                            Else
                                col_mode = 0
                                Call Print_Error_Message(Error_Info, "LIB_VDD_BINNING", "initVddBinCondition", "sheet:" & sheetName & ", it doesn't have the correct columns of Domain in the header. Error!!!")
                                'TheExec.ErrorLogMessage "sheet:" & sheetName & ", it doesn't have the correct columns of Domain in the header. Error!!!"    ''20230807: Removed by sanity check
                            End If
                        Else
                            dict_BinCutFlow_Domain2Column.Add UCase("Performance mode"), Col
                        End If
                        
                        Exit For
                    End If
                Next Col
                
                Exit For
            End If
        Next i
    End If
    
    '''//Find the column of BinCut 1st powerDomain.
    If idx_testJob > -1 And col_mode > 0 And other_voltage_start_point > 0 Then
        '''Do nothing...
    Else
        other_voltage_start_point = 0
        Call Print_Error_Message(Error_Info, "LIB_VDD_BINNING", "initVddBinCondition", "sheet:" & sheetName & ", initVddBinCondition doesn't have the correct header for the current testJob:" & bincutJobName & ". Error!!!")
        'TheExec.ErrorLogMessage "sheet:" & sheetName & ", initVddBinCondition doesn't have the correct header for the current testJob:" & bincutJobName & ". Error!!!"    ''20230807: Removed by sanity check
        Exit Function
    End If
            
    '''====================================================-====================================================================
    '''[Step1] Find row of the Header with column of all BinCut powerDomains and pattern keywords.
    '''====================================================-====================================================================
    '''//If column of the selected testJob is found, start to parse the header VDD.
    For Row = row_of_testJob + 1 To MaxRow
        str_mainColumn_content = LCase(ws_def.Cells(Row, other_voltage_start_point).value)
        
        '''//Check if powerDomain is listed in VddbinPinDict (defined by sheet "Vdd_Binning_Def").
        For Col = col_jobBlock_start To col_jobBlock_end
            strTemp = UCase(Trim(ws_def.Cells(Row, Col).value))
        
            If strTemp <> "" Then
                If Col >= other_voltage_start_point And Col < other_voltage_start_point + cntVddbinPin Then '''powerDomain
                    '''//Check if the tracking power exists in the column of the header.
                    If strTemp Like "*,*" Then
                        strAry_trackpower = Split(strTemp, ",")
                        powerDomain = UCase(Trim(strAry_trackpower(0)))
                        trackpowerTemp = UCase(Trim(Replace(strTemp, (UCase(strAry_trackpower(0)) & ","), "")))
                    Else
                        powerDomain = strTemp
                        trackpowerTemp = vbNullString
                    End If
                    
                    '''//Check if powerDomain is CorePower or OtherRail shown in BinCut sheet "Vdd_Binning_Def".
                    '''//dict_IsCorePower is dictionary of BinCut CorePower/OtherRail.
                    If dict_IsCorePower.Exists(UCase(powerDomain)) = True Then
                        '''//Add column of the powerDomain into the dictionary "dict_BinCutFlow_Domain2Column".
                        If dict_BinCutFlow_Domain2Column.Exists(powerDomain) Then
                            If dict_BinCutFlow_Domain2Column.item(powerDomain) = Col Then
                                got_correct_header = got_correct_header And True
                            Else
                                got_correct_header = got_correct_header And False
                                Call Print_Error_Message(Error_Info, "LIB_VDD_BINNING", "", "sheet:" & sheetName & ", it has the duplicate powerdomain:" & powerDomain & " in row" & Row & " of the header VDD. Error!!!")
                                'TheExec.ErrorLogMessage "sheet:" & sheetName & ", it has the duplicate powerdomain:" & powerDomain & " in row" & Row & " of the header VDD. Error!!!"    ''20230807: Removed by sanity check
                            End If
                        Else
                            got_correct_header = got_correct_header And True
                            dict_BinCutFlow_Domain2Column.Add powerDomain, Col
                            dict_BinCutFlow_Column2Domain.Add Col, powerDomain
                        End If
                        
                        '''//Parsing BinCutFlow (sheet "Non_Binning_Rail") to add BinCut powerDomain into "FullBinCutPowerinFlowSheet".
                        If isParsingOutsideBinCutFlow = False Then
                            '''//Check if any "IGNORE COLUMN" exists in column of powerDomain..
                            isIgnoreColumn = False
                            For j = 1 To Row
                                '''20210131: Modified to check "UCase(Trim(ws_def.Cells(j, Col).value))".
                                If UCase(Trim(ws_def.Cells(j, Col).value)) Like UCase("IGNORE*COLUMN") Then
                                    isIgnoreColumn = True
                                    Exit For
                                End If
                            Next j
                            '''''''''''''''''''''''non_binning_rail_outside'''''''''''''''''''''''
                            If LCase(sheetName) Like LCase("*" & "Outside" & "*") Then
                                If isIgnoreColumn = False Then
                                    If FullBinCutPowerinFlowSheet_outside <> "" Then
                                        If LCase("*," & FullBinCutPowerinFlowSheet_outside & ",*") Like LCase("*," & powerDomain & ",*") Then
                                            got_correct_header = got_correct_header And False
                                            Call Print_Error_Message(Error_Info, "LIB_VDD_BINNING", "", "sheet:" & sheetName & ", it has the duplicate powerdomain:" & powerDomain & " in row" & Row & " of the header VDD. Error!!!")
                                            'TheExec.ErrorLogMessage "sheet:" & sheetName & ", it has the duplicate powerdomain:" & powerDomain & " in row" & Row & " of the header VDD. Error!!!"    ''20230807: Removed by sanity check
                                        Else
                                            FullBinCutPowerinFlowSheet_outside = FullBinCutPowerinFlowSheet_outside & "," & powerDomain
                                        End If
                                    Else
                                        FullBinCutPowerinFlowSheet_outside = powerDomain
                                    End If
                                    
                                    '''//TrackPower
                                    If trackpowerTemp <> "" Then
                                        AllBinCut(VddBinStr2Enum(powerDomain)).TRACKINGPOWER = trackpowerTemp
                                    End If
                                End If
                            Else
                                '''''''''''''''''''''''non_binning_rail'''''''''''''''''''''''
                                If isIgnoreColumn = False Then
                                    If FullBinCutPowerinFlowSheet <> "" Then
                                        If LCase("*," & FullBinCutPowerinFlowSheet & ",*") Like LCase("*," & powerDomain & ",*") Then
                                            got_correct_header = got_correct_header And False
                                            Call Print_Error_Message(Error_Info, "LIB_VDD_BINNING", "", "sheet:" & sheetName & ", it has the duplicate powerdomain:" & powerDomain & " in row" & Row & " of the header VDD. Error!!!")
                                            'TheExec.ErrorLogMessage "sheet:" & sheetName & ", it has the duplicate powerdomain:" & powerDomain & " in row" & Row & " of the header VDD. Error!!!"    ''20230807: Removed by sanity check
                                        Else
                                            FullBinCutPowerinFlowSheet = FullBinCutPowerinFlowSheet & "," & powerDomain
                                        End If
                                    Else
                                        FullBinCutPowerinFlowSheet = powerDomain
                                    End If
                                    
                                    '''//TrackPower
                                    If trackpowerTemp <> "" Then
                                        AllBinCut(VddBinStr2Enum(powerDomain)).TRACKINGPOWER = trackpowerTemp
                                    End If
                                End If
                            End If
                        End If
                    Else '''If dict_IsCorePower.Exists(UCase(powerDomain)) = False
                        got_correct_header = got_correct_header And False
                        Call Print_Error_Message(Error_Info, "LIB_VDD_BINNING", "initVddBinCondition", "sheet:" & sheetName & ", the header has the undefined powerDomain:" & strTemp & ", it doesn't show in sheet Vdd_Binning_Def sheet_appA_1. Error!!!")
                        'TheExec.ErrorLogMessage "sheet:" & sheetName & ", the header has the undefined powerDomain:" & strTemp & ", it doesn't show in sheet Vdd_Binning_Def sheet_appA_1. Error!!!"    ''20230807: Removed by sanity check
                    End If '''If dict_IsCorePower.Exists(UCase(powerDomain)) = True
                Else
                    powerDomain = vbNullString
                    trackpowerTemp = vbNullString
                End If
                
                '''//Check columns of non-powerDomain, ex: "All Others".
                If powerDomain = "" Then
                    '''//Add column of the powerDomain into the dictionary "dict_BinCutFlow_Domain2Column".
                    If dict_BinCutFlow_Domain2Column.Exists(strTemp) Then
                        If dict_BinCutFlow_Domain2Column.item(strTemp) = Col Then
                            got_correct_header = got_correct_header And True
                        Else
                            got_correct_header = got_correct_header And False
                            Call Print_Error_Message(Error_Info, "LIB_VDD_BINNING", "initVddBinCondition", "sheet:" & sheetName & ", it has the duplicate powerdomain:" & powerDomain & " in row" & Row & " of the header VDD. Error!!!")
                            'TheExec.ErrorLogMessage "sheet:" & sheetName & ", it has the duplicate powerdomain:" & powerDomain & " in row" & Row & " of the header VDD. Error!!!"    ''20230807: Removed by sanity check
                        End If
                    Else
                        got_correct_header = got_correct_header And True
                        dict_BinCutFlow_Domain2Column.Add strTemp, Col
                        dict_BinCutFlow_Column2Domain.Add Col, strTemp
                    End If
                End If '''If powerDomain = vbNullString
            End If '''If LCase(ws_def.Cells(row, col).Value) <> ""
        Next Col
        If Not FullBinCutPowerinFlowSheet_outside <> "" And FullBinCutPowerinFlowSheet = FullBinCutPowerinFlowSheet_outside Then '''20220110
            'TheExec.ErrorLogMessage "NON_BINNING_RAIL IS NOT MATCH THAN NON_BINNING_RAIL_OUTSIDE Error!!!"    ''20230807: Removed by sanity check
            Call Print_Error_Message(Error_Info, "LIB_VDD_BINNING", "initVddBinCondition", "NON_BINNING_RAIL IS NOT MATCH THAN NON_BINNING_RAIL_OUTSIDE Error!!!")
        End If
        If got_correct_header = True Then
            Row_of_Title = Row
            Exit For
        End If
    Next Row
    
    '''====================================================-====================================================================
    '''[Step2] Parse each row to get testConditions of powerDomain.
    '''====================================================-====================================================================
    If Row_of_Title > 0 Then '''It means that columns of BinCut powerDomains are found.
        While LCase(ws_def.Cells(Row, 1).value) <> "end"
            '''//Check if any testcondition contains "#REF!" or "#NAME?".
            If IsError(ws_def.Cells(Row, other_voltage_start_point).value) Then
                str_mainColumn_content = vbNullString
                Call Print_Error_Message(Error_Info, "LIB_VDD_BINNING", "", "sheet:" & sheetName & ", cell (row:" & Row & ",column:" & other_voltage_start_point & "), content:" & ws_def.Cells(Row, 1).value & ". The cell contains the incorrect content. Error!!!")
                'TheExec.ErrorLogMessage "sheet:" & sheetName & ", cell (row:" & Row & ",column:" & other_voltage_start_point & "), content:" & ws_def.Cells(Row, 1).value & ". The cell contains the incorrect content. Error!!!"    ''20230807: Removed by sanity check
            Else
                str_mainColumn_content = LCase(ws_def.Cells(Row, other_voltage_start_point).value)
                
                '''//HVCC block (for HBV) in "Non_Binning_Rail" sheet
                '''If keyword of "HVCC" exists in the column of other_voltage_start_point or other_voltage_start_point-1, check which bin_number to use...
                If str_mainColumn_content Like "*hvcc*" Or str_mainColumn_content Like "*hbv*" Then '''"Bin1 - HVCC CP1 @ 25'C, mV"
                    HVCC_flag = 1
        
                    If str_mainColumn_content Like LCase("*bin1*binx*") Or LCase(ws_def.Cells(Row, other_voltage_start_point - 2).value) Like LCase("*bin1*binx*") Then
                        binNumStart = 1
                        binNumStop = 2
                    ElseIf str_mainColumn_content Like LCase("*bin1*") Or LCase(ws_def.Cells(Row, other_voltage_start_point - 2).value) Like LCase("*bin1*") Then
                        binNumStart = 1
                        binNumStop = 1
                    ElseIf str_mainColumn_content Like LCase("*binx*") Or LCase(ws_def.Cells(Row, other_voltage_start_point - 2).value) Like LCase("*binx*") Then
                        binNumStart = 2
                        binNumStop = 2
                    ElseIf str_mainColumn_content Like LCase("*biny*") Or LCase(ws_def.Cells(Row, other_voltage_start_point - 2).value) Like LCase("*biny*") Then
                        binNumStart = 3
                        binNumStop = 3
                    ElseIf LCase(ws_def.Cells(Row, other_voltage_start_point - 2).value) Like "" Then
                        binNumStart = 1
                        binNumStop = 3
                    Else
                        Call Print_Error_Message(Error_Info, "LIB_VDD_BINNING", "", "The Content of HVCC in the sheet " & sheetName & " is wrong. Error!!!")
                        'TheExec.ErrorLogMessage "The Content of HVCC in the sheet " & sheetName & " is wrong. Error!!!"    ''20230807: Removed by sanity check
                    End If
        
                    If binNumStop < binNumStart Then
                        Call Print_Error_Message(Error_Info, "LIB_VDD_BINNING", "", "sheet:" & sheetName & ", row:" & Row & ", the Content of HVCC bin number doesn't match Total_Bincut_Num of Vdd_Binning_Def sheets. Error!!!")
                        'TheExec.ErrorLogMessage "sheet:" & sheetName & ", row:" & Row & ", the Content of HVCC bin number doesn't match Total_Bincut_Num of Vdd_Binning_Def sheets. Error!!!"    ''20230807: Removed by sanity check
                    End If
                            
                '''//Check if columns of HBV BinCut powerDomains match columns of BV BinCut powerDomains.
                ElseIf str_mainColumn_content Like "vdd_*" Then '''ex: "VDD_PCPU".
                    For Col = other_voltage_start_point To col_jobBlock_end
                        strTemp = UCase(Trim(ws_def.Cells(Row, Col).value))
                        If Not strTemp Like "*COMMENT*" Then '''20220103
                        '''//Check if the tracking power exists in the column of the header.
                                If strTemp Like "*,*" Then
                                    strAry_trackpower = Split(strTemp, ",")
                                    powerDomain = UCase(Trim(strAry_trackpower(0)))
                                    trackpowerTemp = UCase(Trim(Replace(strTemp, (UCase(strAry_trackpower(0)) & ","), "")))
                                Else
                                    powerDomain = strTemp
                                    trackpowerTemp = vbNullString
                                End If
                        Else
                            powerDomain = strTemp
                        End If
                        If dict_BinCutFlow_Domain2Column.item(powerDomain) = Col Then
                            If dict_IsCorePower.Exists(powerDomain) Then
                                If trackpowerTemp = AllBinCut(VddBinStr2Enum(powerDomain)).TRACKINGPOWER Then
                                    '''Do nothing
                                Else
                                    Call Print_Error_Message(Error_Info, "LIB_VDD_BINNING", "", "sheet:" & sheetName & ", cell (row:" & Row & ",column:" & Col & "), content:" & strTemp & ", trackpower of powerDomain is different from cell (row:" & Row_of_Title & ", col:" & dict_BinCutFlow_Domain2Column.item(powerDomain) & ") in the header. Error!!!")
                                    'TheExec.ErrorLogMessage "sheet:" & sheetName & ", cell (row:" & Row & ",column:" & Col & "), content:" & strTemp & ", trackpower of powerDomain is different from cell (row:" & Row_of_Title & ", col:" & dict_BinCutFlow_Domain2Column.item(powerDomain) & ") in the header. Error!!!"    ''20230807: Removed by sanity check
                                End If
                            End If
                        Else
                            Call Print_Error_Message(Error_Info, "LIB_VDD_BINNING", "initVddBinCondition", "sheet:" & sheetName & ", cell (row:" & Row & ",column:" & Col & "), content:" & strTemp & ", it is different from the header VDD in row" & Row_of_Title & ". Error!!!")
                            'TheExec.ErrorLogMessage strTemp & " in row" & Row & " col" & Col & " of sheet " & sheetName & " is different from the header VDD in row" & Row_of_Title & ". Error!!!"    ''20230807: Removed by sanity check
                        End If
                    Next Col
                    
                    If HVCC_flag = 1 Then
                        strTemp = UCase(Trim(ws_def.Cells(Row, other_voltage_start_point - 1).value))
                        
                        If dict_BinCutFlow_Domain2Column.item(strTemp) = other_voltage_start_point - 1 Then
                            '''Do nothing
                        Else
                            Call Print_Error_Message(Error_Info, "LIB_VDD_BINNING", "initVddBinCondition", strTemp & " in row" & Row & " col" & other_voltage_start_point - 1 & " of sheet " & sheetName & " is different from the header VDD in row" & Row_of_Title & ". Error!!!")
                            'TheExec.Datalog.WriteComment strTemp & " in row" & Row & " col" & other_voltage_start_point - 1 & " of sheet " & sheetName & " is different from the header VDD in row" & Row_of_Title & ". Error!!!"
                            'TheExec.ErrorLogMessage strTemp & " in row" & Row & " col" & other_voltage_start_point - 1 & " of sheet " & sheetName & " is different from the header VDD in row" & Row_of_Title & ". Error!!!"    ''20230807: Removed by sanity check
                        End If
                    End If
                   
                '''//Check if any performance_mode exists in column "Performance Mode".
                ElseIf LCase(ws_def.Cells(Row, dict_BinCutFlow_Domain2Column.item(UCase("Performance Mode")))) Like "m*" Then '''ex: "MS001", "MS001_GPU".
                    '''//Get performance_mode
                    str_mode_temp = UCase(ws_def.Cells(Row, dict_BinCutFlow_Domain2Column.item(UCase("Performance Mode"))).value)
                    split_content = Split(str_mode_temp, "_")
                    performance_mode = UCase(split_content(0))
                    
                    '''//If with all empty conditions, it means performance_mode without any additional_mode.
                    If UBound(split_content) > 0 Then
                        additional_mode = UCase(Replace(UCase(str_mode_temp), (performance_mode & "_"), ""))
                    Else
                        additional_mode = vbNullString
                    End If
                    
                    '''//Check if the main performance_mode exists in the dictionary "VddbinPmodeDict".
                    If VddbinPmodeDict.Exists(performance_mode) Then
                        main_p_mode = VddBinStr2Enum(performance_mode)
                        powerDomain = AllBinCut(main_p_mode).powerPin
                        
                        If gb_bincut_power_list(VddBinStr2Enum(powerDomain)) <> "" Then
                            '''//Check if performance_mode exists in gb_bincut_power_list(VddBinStr2Enum(powerDomain).
                            If UCase("*," & gb_bincut_power_list(VddBinStr2Enum(powerDomain)) & ",*") Like UCase("*," & performance_mode & ",*") Then
                                '''pmode exists in the list, so that do nothing...
                            Else
                                '''=============================================================================================='''
                                '''//Check digit2-3 of the p_mode are same as p_modes in the power_list, ex: "MC" of "MC601" and "MC602".
                                '''20190704: As the discussion with SWLINZA, we should check digit1-2 of the p_mode are same as p_modes in the power_list.
                                '''=============================================================================================='''
                                split_content = Split(gb_bincut_power_list(VddBinStr2Enum(powerDomain)), ",")
                                
                                If mid(UCase(ws_def.Cells(Row, 2).value), 1, 2) = mid(UCase(split_content(0)), 1, 2) Then
                                    gb_bincut_power_list(VddBinStr2Enum(powerDomain)) = gb_bincut_power_list(VddBinStr2Enum(powerDomain)) & "," & performance_mode
                                Else
                                    Call Print_Error_Message(Error_Info, "LIB_VDD_BINNING", "", "sheet:" & sheetName & "," & performance_mode & " is incosistent with " & powerDomain & " power_seq " & split_content(0) & ". Please check Domain and Mode columns in Vdd_Binning_Def. initVddBinCondition has the incorrect keyword. Error!!!")
                                    'TheExec.ErrorLogMessage "sheet:" & sheetName & "," & performance_mode & " is incosistent with " & powerDomain & " power_seq " & split_content(0) & ". Please check Domain and Mode columns in Vdd_Binning_Def. initVddBinCondition has the incorrect keyword. Error!!!"    ''20230807: Removed by sanity check
                                End If
                            End If
                        Else
                            gb_bincut_power_list(VddBinStr2Enum(powerDomain)) = performance_mode
                        End If
                    
                        '''//Add the additional mode into the dictionary "AdditionalModeDict" for Additional Mode.
                        If additional_mode <> "" Then
                            If isParsingOutsideBinCutFlow = True Then
                                If dict_OutsideBinCut_additionalMode.Exists(additional_mode) = True Then
                                    If sheetName <> dict_OutsideBinCut_additionalMode.item(additional_mode) Then
                                        Call Print_Error_Message(Error_Info, "LIB_VDD_BINNING", "", "sheet:" & sheetName & ", performance mode:" & str_mode_temp & ", it has the duplicate string about additional mode:" & additional_mode & ". Error!!!")
                                        'TheExec.ErrorLogMessage "sheet:" & sheetName & ", performance mode:" & str_mode_temp & ", it has the duplicate string about additional mode:" & additional_mode & ". Error!!!"    ''20230807: Removed by sanity check
                                        Exit Function
                                    Else
                                        addi_mode = AdditionalModeDict.item(additional_mode)
                                    End If
                                Else
                                    If AdditionalModeDict.Exists(additional_mode) Then
                                        addi_mode = AdditionalModeDict.item(additional_mode)
                                    Else
                                        cntAdditionalMode = cntAdditionalMode + 1
                                        AdditionalModeDict.Add UCase(additional_mode), cntAdditionalMode
                                        addi_mode = cntAdditionalMode
                
                                        ReDim Preserve AdditionalModeName(cntAdditionalMode)
                                        AdditionalModeName(cntAdditionalMode) = additional_mode
                                        dict_OutsideBinCut_additionalMode.Add additional_mode, sheetName
                                    End If
                                End If
                            Else
                                If AdditionalModeDict.Exists(additional_mode) Then
                                    addi_mode = AdditionalModeDict.item(additional_mode)
                                Else
                                    cntAdditionalMode = cntAdditionalMode + 1
                                    AdditionalModeDict.Add UCase(additional_mode), cntAdditionalMode
                                    addi_mode = cntAdditionalMode
            
                                    ReDim Preserve AdditionalModeName(cntAdditionalMode)
                                    AdditionalModeName(cntAdditionalMode) = additional_mode
                                End If
                            End If
                        Else
                            addi_mode = -1
                        End If
                        
                        '''//Parsing testCondition into BinCut(p_mode, bin_number) array.
                        For i = 0 To cntVddbinPin - 1
                            For passBinCut = binNumStart To binNumStop
                                Col = other_voltage_start_point + i
                            
                                '''//Check if any testcondition contains "#REF!" or "#NAME?".
                                If IsError(ws_def.Cells(Row, Col).value) Then
                                    testCondition = vbNullString
                                    Call Print_Error_Message(Error_Info, "LIB_VDD_BINNING", "initVddBinCondition", "sheet:" & sheetName & ", cell (row:" & Row & ", column:" & Col & "). it has the incorrect content. Error!!!")
                                    'TheExec.ErrorLogMessage "sheet:" & sheetName & ", cell (row:" & Row & ", column:" & Col & "). it has the incorrect content. Error!!!"    ''20230807: Removed by sanity check
                                Else
                                    selected_powerDomain = UCase(dict_BinCutFlow_Column2Domain.item(Col))
                                    
                                    '''//Get and trim string of the testCondition from Non_Binning_Rail.
                                    testCondition = LCase(Trim(ws_def.Cells(Row, Col).value))
                                    
                                    '''//Check if testCondition contains any keyword about PassBin(Bin1/BinX/BinY)...
                                    If testCondition Like "*bin1*" Then '''Bin1
                                        bincutNum = 1
                                    ElseIf testCondition Like "*binx*" Then '''BinX
                                        bincutNum = 2
                                    ElseIf testCondition Like "*biny*" Then '''BinY
                                        bincutNum = 3
                                    Else
                                        bincutNum = 0
                                    End If
                                    
                                    '''//Check if bincutNum is greater than PassBinCut_ary(Ubound(PassBinCut_ary)).
                                    '''//PassBinCut_ary(Ubound(PassBinCut_ary)) is the highest Bin number of the BinCut voltage table(sheet "Vdd_Binning_Def").
                                    '''20210802: Modified to check if testCondition contains any keyword about PassBin(Bin1/BinX/BinY) greater than the highest bin number.
                                    If bincutNum > PassBinCut_ary(UBound(PassBinCut_ary)) Then
                                        Call Print_Error_Message(Error_Info, "LIB_VDD_BINNING", "initVddBinCondition", "sheet:" & sheetName & ",cell:(row " & Row & ", column " & Col & "),testCondition:" & testCondition & ", it contains the incorrect keyword about PassBin greater than the highest bin number for initVddBinCondition. Error!!!")
                                        'TheExec.ErrorLogMessage "sheet:" & sheetName & ",cell:(row " & Row & ", column " & Col & "),testCondition:" & testCondition & ", it contains the incorrect keyword about PassBin greater than the highest bin number for initVddBinCondition. Error!!!"    ''20230807: Removed by sanity check
                                    End If
                                    
                                    '''//Check if testCondition contains "(" but no ")".
                                    If testCondition Like "*(*" And Not (testCondition) Like "*)*" Then
                                        Call Print_Error_Message(Error_Info, "LIB_VDD_BINNING", "", "Please check the cell (row " & Row & ", column " & Col & ") of sheet " & sheetName & ". The cell contains the incorrect format. Error!!!")
                                        'TheExec.ErrorLogMessage "Please check the cell (row " & Row & ", column " & Col & ") of sheet " & sheetName & ". The cell contains the incorrect format. Error!!!"    ''20230807: Removed by sanity check
                                    End If
                                    
                                    '''==========================================================================================================================================='''
                                    '''//Check if testCondition contain keyword "*Evaluate*Bin*" to decide "is_BinCutJob_for_StepSearch" = True (BinCut stepSearch) while isParsingOutsideBinCutFlow = False.
                                    '''==========================================================================================================================================='''
                                    If testCondition Like LCase("*Evaluate*Bin*") And isParsingOutsideBinCutFlow = False Then
                                        strSplitted = Split(LCase(testCondition), LCase("Evaluate Bin"))
                                        
                                        '''//Check if testCondition with keyword "*Evaluate*Bin*" has the correct performance mode.
                                        '''20210414: Modified to add "is_for_BinSearch as Boolean" for AllBinCut(p_mode).
                                        If VddbinPmodeDict.Exists(UCase(Trim(strSplitted(0)))) = True And dict_IsCorePower.Exists(UCase(Trim(strSplitted(0)))) = False Then
                                            AllBinCut(VddBinStr2Enum(UCase(Trim(strSplitted(0))))).is_for_BinSearch = True
                                            is_BinCutJob_for_StepSearch = True
                                        Else
                                            Call Print_Error_Message(Error_Info, "LIB_VDD_BINNING", "initVddBinCondition", "sheet:" & sheetName & ",cell:(row " & Row & ", column " & Col & "),testCondition:" & testCondition & ". It doesn't contain any correct performance mode in testCondition, please check sheet " & sheetName & ". Error!!!")
                                            'TheExec.ErrorLogMessage "sheet:" & sheetName & ",cell:(row " & Row & ", column " & Col & "),testCondition:" & testCondition & ". It doesn't contain any correct performance mode in testCondition, please check sheet " & sheetName & ". Error!!!"    ''20230807: Removed by sanity check
                                        End If
                                    End If
                                    
                                    If testCondition Like LCase("Bincut_X_X_X*") Then 'WWWWW20220630 T0TX
                                        If HVCC_flag = 0 Then
                                            If addi_mode > 0 Then
                                                DC_Mapping(main_p_mode, passBinCut).Addtional_LV_DC_Value(addi_mode) = testCondition
                                            Else
                                                DC_Mapping(main_p_mode, passBinCut).LV_DC_Value = testCondition
                                            End If

                                        ElseIf HVCC_flag = 1 Then
                                            If addi_mode > 0 Then
                                                DC_Mapping(main_p_mode, passBinCut).Addtional_HVCC_DC_Value(addi_mode) = testCondition
                                            Else
                                                DC_Mapping(main_p_mode, passBinCut).HVCC_DC_Value = testCondition
                                            End If
                                        End If
                                        
                                     ElseIf isParsingOutsideBinCutFlow = False Then
                                        If HVCC_flag = 0 Then
                                            If addi_mode > 0 Then
                                                BinCut(main_p_mode, passBinCut).Addtional_OTHER_VOLTAGE(VddBinStr2Enum(selected_powerDomain), addi_mode) = testCondition
                                            Else
                                                BinCut(main_p_mode, passBinCut).OTHER_VOLTAGE(VddBinStr2Enum(selected_powerDomain)) = testCondition
                                            End If
                                        ElseIf HVCC_flag = 1 Then
                                            If addi_mode > 0 Then
                                                BinCut(main_p_mode, passBinCut).HVCC_Addtional_OTHER_VOLTAGE(VddBinStr2Enum(selected_powerDomain), addi_mode) = testCondition
                                            Else
                                                BinCut(main_p_mode, passBinCut).HVCC_OTHER_VOLTAGE(VddBinStr2Enum(selected_powerDomain)) = testCondition
                                            End If
                                        End If
                                    Else
                                        '''20201021: As per discussion with TSMC PCLINZG, he suggested us to use the same testCondition for outsideBinCutFlow BV and HBV.
                                        If addi_mode > 0 Then
                                            BinCut(main_p_mode, passBinCut).OutsideBinCut_Addtional_OTHER_VOLTAGE(VddBinStr2Enum(selected_powerDomain), addi_mode) = testCondition
                                            BinCut(main_p_mode, passBinCut).OutsideBinCut_HVCC_Addtional_OTHER_VOLTAGE(VddBinStr2Enum(selected_powerDomain), addi_mode) = testCondition
                                        Else
                                            BinCut(main_p_mode, passBinCut).OutsideBinCut_OTHER_VOLTAGE(VddBinStr2Enum(selected_powerDomain)) = testCondition
                                            BinCut(main_p_mode, passBinCut).OutsideBinCut_HVCC_OTHER_VOLTAGE(VddBinStr2Enum(selected_powerDomain)) = testCondition
                                        End If
                                        Flag_NonbinningrailOutsideBinCut_parsed = True
                                    End If
                                End If
                            Next passBinCut
                        Next i
                    Else '''If VddbinPmodeDict.Exists(UCase(replace_p_name(0)))=false
                        Call Print_Error_Message(Error_Info, "LIB_VDD_BINNING", "initVddBinCondition", "sheet:" & sheetName & ",row:" & Row & ",cell:" & str_mode_temp & ", it doesn't contain any correct performance mode. Error!!!")
                        'TheExec.ErrorLogMessage "sheet:" & sheetName & ",row:" & Row & ",cell:" & str_mode_temp & ", it doesn't contain any correct performance mode. Error!!!"    ''20230807: Removed by sanity check
                    End If
                End If '''If str_mainColumn_content Like "*hvcc*"
            End If '''If IsError(ws_def.Cells(row, other_voltage_start_point).Value) Then
            
            Row = Row + 1
        Wend
    End If '''If row_of_title > 0
    
    '''====================================================-====================================================================
    '''[Step3] Check if "cntAdditionalMode" should be "<=" with "MaxAdditionalModeCount".
    '''====================================================-====================================================================
    If cntAdditionalMode > MaxAdditionalModeCount Then
        Call Print_Error_Message(Error_Info, "LIB_VDD_BINNING", "initVddBinCondition", "sheet:" & sheetName & ",number of BinCut additional modes:" & cntAdditionalMode & ", it is greater than BinCut globalVariable MaxAdditionalModeCount=" & MaxAdditionalModeCount & ". Please check BinCut flow table and globalVariable MaxAdditionalModeCount. Error!!!")
        'TheExec.ErrorLogMessage "sheet:" & sheetName & ",number of BinCut additional modes:" & cntAdditionalMode & ", it is greater than BinCut globalVariable MaxAdditionalModeCount=" & MaxAdditionalModeCount & ". Please check BinCut flow table and globalVariable MaxAdditionalModeCount. Error!!!"    ''20230807: Removed by sanity check
    End If
    
    If isParsingOutsideBinCutFlow = False Then
        '''//Split pin_groups and get each powerDomain, then sort the sequence about p_mode for each powerDomain. BinCut PowerDomain consists of CorePower and OtherRail.
        If FullBinCutPowerinFlowSheet <> "" Then
            pinGroup_BinCut = Split(FullBinCutPowerinFlowSheet, ",")
            
            For i = 0 To UBound(pinGroup_BinCut)
                powerDomain = UCase(pinGroup_BinCut(i))
            
                If dict_IsCorePower.item(powerDomain) = True Then '''CorePower
                    If FullCorePowerinFlowSheet <> "" Then
                        FullCorePowerinFlowSheet = FullCorePowerinFlowSheet & "," & powerDomain
                    Else
                        FullCorePowerinFlowSheet = powerDomain
                    End If
                    
                    dict_IsCorePowerInBinCutFlowSheet.Add powerDomain, True
                Else '''OtherRail
                    If FullOtherRailinFlowSheet <> "" Then
                        FullOtherRailinFlowSheet = FullOtherRailinFlowSheet & "," & powerDomain
                    Else
                        FullOtherRailinFlowSheet = powerDomain
                    End If
                    
                    dict_IsCorePowerInBinCutFlowSheet.Add powerDomain, False
                End If
            Next i
        Else
            Call Print_Error_Message(Error_Info, "LIB_VDD_BINNING", "initVddBinCondition", "FullBinCutPowerinFlowSheet should not be empty. Please check Vdd_Binning_Def_appA and Non_Binning_Rail. Error!!!")
            'TheExec.ErrorLogMessage "FullBinCutPowerinFlowSheet should not be empty. Please check Vdd_Binning_Def_appA and Non_Binning_Rail. Error!!!"    ''20230807: Removed by sanity check
        End If
        
        '''//Split pin_groups and get powerDomains of BinCut CorePower and OtherRail.
        '''CorePower
        If FullCorePowerinFlowSheet <> "" Then
            pinGroup_CorePower = Split(FullCorePowerinFlowSheet, ",")
        Else
            Call Print_Error_Message(Error_Info, "LIB_VDD_BINNING", "initVddBinCondition", "FullCorePowerinFlowSheet should not be empty. Please check Vdd_Binning_Def_appA and Non_Binning_Rail. Error!!!")
            'TheExec.ErrorLogMessage "FullCorePowerinFlowSheet should not be empty. Please check Vdd_Binning_Def_appA and Non_Binning_Rail. Error!!!"    ''20230807: Removed by sanity check
        End If
        
        '''OtherRail
        If FullOtherRailinFlowSheet <> "" Then
            pinGroup_OtherRail = Split(FullOtherRailinFlowSheet, ",")
        Else
            Call Print_Error_Message(Error_Info, "LIB_VDD_BINNING", "initVddBinCondition", "FullOtherRailinFlowSheet should not be empty. Please check Vdd_Binning_Def_appA and Non_Binning_Rail. Error!!!")
            'TheExec.ErrorLogMessage "FullOtherRailinFlowSheet should not be empty. Please check Vdd_Binning_Def_appA and Non_Binning_Rail. Error!!!"    ''20230807: Removed by sanity check
        End If
    End If
Exit Function
errHandler:
    Call Print_Error_Message(Error_Info, "LIB_VDD_BINNING", "initVddBinCondition") 'Add ErrHandler 2023/08/01
    If AbortTest Then Exit Function Else Resume Next
End Function

Public Function Parsing_EquationVoltages(sheetName As String)
On Error GoTo errHandler 'Add ErrHandler 2023/08/01
    Dim sheetName_EquationVoltages As String
    Dim sheetName_temp As String
    Dim EquationVoltages_temp As String
    Dim ws_def As Worksheet
    Dim wb As Workbook
    Dim Row As Long, Col As Long
    Dim MaxRow As Long
    Dim maxcol As Long
    Dim split_content() As String
    Dim performance_mode As String
    Dim additional_mode As String
    Dim main_p_mode As Integer
    Dim p_mode As Integer
    Dim addi_mode As Integer '''For the additional mode
    Dim i As Long, j As Long, k As Long, L As Long
    Dim S As Long
    Dim other_voltage_start_point As Long
    Dim HVCC_flag As Long
    Dim passBinCut As Long
    '''for testjob mapping
    Dim Row_of_Title As Long
    Dim row_of_testJob As Long
    Dim job_keyword As String
    Dim row_jobBlock As Long
    '''
    Dim cnt_testJob As Long
    Dim idx_testJob As Long
    Dim col_testJob() As Long
    Dim col_jobBlock_start As Long
    Dim col_jobBlock_end As Long
    '''for Performance Mode
    Dim col_mode As Long
    '''for powerDomain
    Dim got_correct_header As Boolean
    Dim got_CorrectDomain As Boolean
    '''for trackPower
    Dim trackpowerTemp As String
    Dim strAry_trackpower() As String
    '''variables
    Dim COREPOWER As Long
    Dim powerDomain As String
    Dim selected_powerDomain As String
    Dim str_mode_temp As String
    Dim binNumStart As Long
    Dim binNumStop As Long
    Dim isSheetFound As Boolean
    Dim isParsingOutsideBinCutFlow As Boolean
    Dim isIgnoreColumn As Boolean
    Dim testCondition As String
    Dim str_mainColumn_content As String
    Dim strTemp As String
    Dim strSplitted() As String
    Dim bincutNum As Long
    Dim EquationVoltages_ALL() As String
    Dim EquationVoltages_step As String
'''//==================================================================================================================================================================================//'''
'''//Note:
'''1. Non_Binning_Rail should be parsed prior to sheet Non_Binning_Rail_Post_BinCut.
'''2. As per discussion with TSMC PCLINZG, he suggested us to use the same testCondition for outsideBinCutFlow BV and HBV.
'''3. Please check keyword "Non_Binning_Rail_Outside" of sheetName for the vbt functions "initVddBinCondition" and "parsing_OutsideBinCut_flow_table".
'''//==================================================================================================================================================================================//'''
    '''*****************************************************************'''
    '''//Check if the sheet exists
    'sheetName = "Non_Binning_Rail"
    sheetName_temp = sheetName
    sheetName = sheetName & "_" & TheExec.CurrentJob
    sheetName_EquationVoltages = "EquationVoltages" '''for EquationVoltages.
    dict_EquationVoltages_ALL.RemoveAll
    dict_EquationVoltages_Domain2Column.RemoveAll       ''20230613: Added for initialize
    dict_EquationVoltages_Column2Domain.RemoveAll       ''20230613: Added for initialize
    
    Set wb = Application.ActiveWorkbook
    Call check_Sheet_Range(sheetName, wb, ws_def, MaxRow, maxcol, isSheetFound, False)
    If isSheetFound = False Then                        ''' If "Non_Binning_Rail_JOBNAME" does not exist
        sheetName = sheetName_temp
        Call check_Sheet_Range(sheetName, wb, ws_def, MaxRow, maxcol, isSheetFound, False)
    End If
    If gS_Non_Binning_Rail_sheet = "" Then          '' Get sheetName for first time
        gS_Non_Binning_Rail_sheet = sheetName       '' 20211224: Added to Print Bincut Sheet Name in Print_Bincut_Config
    End If
    If gS_EquationVoltages_sheet = "" And LCase(sheetName) Like LCase("EquationVoltages") Then
        gS_EquationVoltages_sheet = sheetName
    End If
    ReDim EquationVoltages_ALL(cntVddbinPin)
    
    '''*****************************************************************'''
    
    If isSheetFound = True Then
        '''//init
        '''Since all col_XXX and row_XXX related variables with default values=0, no need to initialize them as 0.
        binNumStart = 1
        binNumStop = Total_Bincut_Num '3 ''20231102: Carter found incorrect usage
        HVCC_flag = 0
        other_voltage_start_point = 0
        trackpowerTemp = vbNullString
        Row_of_Title = 0
        row_of_testJob = 0
        cnt_testJob = 0
        idx_testJob = -1
        got_correct_header = True
        isIgnoreColumn = False
        got_CorrectDomain = False
        str_mainColumn_content = vbNullString
        str_mode_temp = vbNullString
        performance_mode = vbNullString
        additional_mode = vbNullString
        job_keyword = vbNullString
        col_jobBlock_start = 0
        col_jobBlock_end = 0
        col_mode = 0
        
        If LCase(sheetName) Like LCase("*" & sheetName_EquationVoltages & "*") Then
            If FullBinCutPowerinFlowSheet <> "" Then
                isParsingOutsideBinCutFlow = True '''Parsing outsideBinCutFlow...
            Else
                isParsingOutsideBinCutFlow = False '''Parsing BinCutFlow (sheet "Non_Binning_Rail")
                TheExec.Datalog.WriteComment "EquationVoltages should be parsed prior to sheet EquationVoltages"
                'TheExec.ErrorLogMessage "EquationVoltages should be parsed prior to sheet EquationVoltages"    ''20230807: Removed by sanity check
            End If
        Else
            isParsingOutsideBinCutFlow = False '''//Parsing BinCutFlow (sheet "Non_Binning_Rail")
        End If
        '''20210819: Modified to move the vbt code about resetting globalVariables of BinCut testCondition from the vbt function initVddBinCondition to the vbt function Reset_BinCut_GlobalVariable_for_initVddBinning.
        
        '''//Get keyword for BinCut testJob mapping.
        '''20210819: Modified to assemble job_keyword by bincutJobName according to the vbt function Mapping_TestJobName_to_BincutJobName.
        job_keyword = LCase("*" & bincutJobName & "*") '''ex: "*cp1*", "*cp2*", "*ft_room*", "*ft_hot*", "*qa*".
        If TheExec.enableWord("T0TX_ROOM") = True And Not TheExec.enableWord("HardIP_RMA") = True Then job_keyword = LCase("*T0TX_ROOM*") 'WWWWW20220631 T0TX
        If TheExec.enableWord("T0TX_HOT") = True And Not TheExec.enableWord("HardIP_RMA") = True Then job_keyword = LCase("*T0TX_HOT*")   'WWWWW20220631 T0TX

    Else
        Exit Function
    End If
    
    '''//Find the keyword of BinCut testjob, and find column of the selected testJob.
    '''20210819: As per discussion with Jeff, he suggested us to use the keyword "Binning Domain" for each testJob.
    If job_keyword <> "" Then
        For Row = 1 To MaxRow
            For Col = 1 To maxcol
                If LCase(ws_def.Cells(Row, Col).value) Like LCase("DC Spec") Then
                    ReDim Preserve col_testJob(cnt_testJob)
                    col_testJob(cnt_testJob) = Col
                    cnt_testJob = cnt_testJob + 1
                    
                    If row_of_testJob = 0 Then
                        row_of_testJob = Row
                    End If
                End If
            Next Col
            
            If row_of_testJob > 0 Then
                Exit For
            End If
        Next Row
    End If
    '''//Check if any matched block for IGXL Job.
    If cnt_testJob > 0 Then
        For i = 0 To cnt_testJob - 1
            If LCase(ws_def.Cells(row_of_testJob, col_testJob(i) + 1).value) Like LCase(job_keyword) Then
                idx_testJob = i
                other_voltage_start_point = col_testJob(i) + 1
                
                '''//Get start/stop columns for block of the selected testJob.
                col_jobBlock_start = col_testJob(idx_testJob) '''column of Domain
                If idx_testJob = UBound(col_testJob) Then
                    col_jobBlock_end = maxcol
                Else
                    col_jobBlock_end = col_testJob(idx_testJob + 1) - 1
                End If
                
                '''//Get column of "Performance Mode"
                For Col = col_jobBlock_start To col_jobBlock_end
                        col_mode = 1
                        Exit For
                Next Col
                Exit For
            End If
        Next i
    End If
    
    '''//Find the column of BinCut 1st powerDomain.
    If idx_testJob > -1 And col_mode > 0 And other_voltage_start_point > 0 Then
        '''Do nothing...
    Else
        other_voltage_start_point = 0
        Call Print_Error_Message(Error_Info, "LIB_VDD_BINNING", "Parsing_EquationVoltages", "sheet:" & sheetName & ", initVddBinCondition doesn't have the correct header for the current testJob:" & bincutJobName & ". Error!!!")
        'TheExec.ErrorLogMessage "sheet:" & sheetName & ", initVddBinCondition doesn't have the correct header for the current testJob:" & bincutJobName & ". Error!!!"    ''20230807: Removed by sanity check
        Exit Function
    End If
    
    '''====================================================-====================================================================
    '''[Step1] Find row of the Header with column of all BinCut powerDomains and pattern keywords.
    '''====================================================-====================================================================
    '''//If column of the selected testJob is found, start to parse the header VDD.
    For Row = row_of_testJob + 1 To MaxRow
        str_mainColumn_content = LCase(ws_def.Cells(Row, other_voltage_start_point).value)
        
        '''//Check if powerDomain is listed in VddbinPinDict (defined by sheet "Vdd_Binning_Def").
        For Col = col_jobBlock_start To col_jobBlock_end
            strTemp = UCase(Trim(ws_def.Cells(Row, Col).value))
        
            If strTemp <> "" Then
                If Col >= other_voltage_start_point And Col < other_voltage_start_point + cntVddbinPin Then '''powerDomain
                    '''//Check if the tracking power exists in the column of the header.
                    If strTemp Like "*,*" Then
                        strAry_trackpower = Split(strTemp, ",")
                        powerDomain = UCase(Trim(strAry_trackpower(0)))
                        trackpowerTemp = UCase(Trim(Replace(strTemp, (UCase(strAry_trackpower(0)) & ","), "")))
                    Else
                        powerDomain = strTemp
                        trackpowerTemp = vbNullString
                    End If
                    
                Else
                    powerDomain = vbNullString
                    trackpowerTemp = vbNullString
                End If
                
         
        
        
        '---------------------------------------------------------------------------------------------------
                If dict_IsCorePower.Exists(UCase(powerDomain)) = True Then
                        '''//Add column of the powerDomain into the dictionary "dict_BinCutFlow_Domain2Column".
                    If dict_EquationVoltages_Domain2Column.Exists(powerDomain) Then
                        If dict_EquationVoltages_Domain2Column.item(powerDomain) = Col Then
                            got_correct_header = got_correct_header And True
                        Else
                            got_correct_header = got_correct_header And False
                            Call Print_Error_Message(Error_Info, "LIB_VDD_BINNING", "Parsing_EquationVoltages", "sheet:" & sheetName & ", it has the duplicate powerdomain:" & powerDomain & " in row" & Row & " of the header VDD. Error!!!")
                            'TheExec.ErrorLogMessage "sheet:" & sheetName & ", it has the duplicate powerdomain:" & powerDomain & " in row" & Row & " of the header VDD. Error!!!"    ''20230807: Removed by sanity check
                        End If
                    Else
                        got_correct_header = got_correct_header And True
                        dict_EquationVoltages_Domain2Column.Add powerDomain, Col
                        dict_EquationVoltages_Column2Domain.Add Col, powerDomain
                    End If
                End If
                
                '''//Check columns of non-powerDomain, ex: "All Others".
                If powerDomain = "" Then
                    '''//Add column of the powerDomain into the dictionary "dict_BinCutFlow_Domain2Column".
                    If dict_EquationVoltages_Domain2Column.Exists(strTemp) Then
                        If dict_EquationVoltages_Domain2Column.item(strTemp) = Col Then
                            got_correct_header = got_correct_header And True
                        Else
                            got_correct_header = got_correct_header And False
                            Call Print_Error_Message(Error_Info, "LIB_VDD_BINNING", "Parsing_EquationVoltages", "sheet:" & sheetName & ", it has the duplicate powerdomain:" & powerDomain & " in row" & Row & " of the header VDD. Error!!!")
                            'TheExec.ErrorLogMessage "sheet:" & sheetName & ", it has the duplicate powerdomain:" & powerDomain & " in row" & Row & " of the header VDD. Error!!!"    ''20230807: Removed by sanity check
                        End If
                    Else
                        got_correct_header = got_correct_header And True
                        dict_EquationVoltages_Domain2Column.Add strTemp, Col
                        dict_EquationVoltages_Column2Domain.Add Col, strTemp
                    End If
                End If '''If powerDomain = vbNullString
        '---------------------------------------------------------------------------------------------------
            End If '''If LCase(ws_def.Cells(row, col).Value) <> ""
        Next Col
        
        
        
        If Not FullBinCutPowerinFlowSheet_outside <> "" And FullBinCutPowerinFlowSheet = FullBinCutPowerinFlowSheet_outside Then '''20220110
            'TheExec.ErrorLogMessage "EquationVoltages IS NOT MATCH THAN EquationVoltages Error!!!"    ''20230807: Removed by sanity check
            Call Print_Error_Message(Error_Info, "LIB_VDD_BINNING", "Parsing_EquationVoltages", "EquationVoltages IS NOT MATCH THAN EquationVoltages Error!!!")
        End If
        If got_correct_header = True Then
            Row_of_Title = Row
            Exit For
        End If
    Next Row
  
    '''====================================================-====================================================================
    '''[Step2] Parse each row to get testConditions of powerDomain.
    '''====================================================-====================================================================
    If Row_of_Title > 0 Then '''It means that columns of BinCut powerDomains are found.
        While LCase(ws_def.Cells(Row, 1).value) <> "end"
            If Row > 2 Then
                EquationVoltages_temp = UCase(ws_def.Cells(Row, 1).value)
                If Not UCase(EquationVoltages_temp) Like "*EQN*" Then
                    'TheExec.ErrorLogMessage "EquationVoltages don't have EQN Error!!!"    ''20230807: Removed by sanity check
                    Call Print_Error_Message(Error_Info, "LIB_VDD_BINNING", "", "EquationVoltages don't have EQN Error!!! ")
                End If
                
                If Not (VddbinPmodeDict.Exists(EquationVoltages_temp)) Then
                    cntVddbinPmode = cntVddbinPmode + 1
                    VddbinPmodeDict.Add EquationVoltages_temp, cntVddbinPmode
                    
                    ReDim Preserve VddBinName(cntVddbinPmode)
                    VddBinName(cntVddbinPmode) = EquationVoltages_temp
                End If
            End If
            '''//Check if any testcondition contains "#REF!" or "#NAME?".
            If IsError(ws_def.Cells(Row, other_voltage_start_point).value) Then
                str_mainColumn_content = vbNullString
                Call Print_Error_Message(Error_Info, "LIB_VDD_BINNING", "Parsing_EquationVoltages", "sheet:" & sheetName & ", cell (row:" & Row & ",column:" & other_voltage_start_point & "), content:" & ws_def.Cells(Row, 1).value & ". The cell contains the incorrect content. Error!!!")
                'TheExec.ErrorLogMessage "sheet:" & sheetName & ", cell (row:" & Row & ",column:" & other_voltage_start_point & "), content:" & ws_def.Cells(Row, 1).value & ". The cell contains the incorrect content. Error!!!"    ''20230807: Removed by sanity check
            Else
                str_mainColumn_content = LCase(ws_def.Cells(Row, other_voltage_start_point).value)
                
                '''//HVCC block (for HBV) in "Non_Binning_Rail" sheet
                '''If keyword of "HVCC" exists in the column of other_voltage_start_point or other_voltage_start_point-1, check which bin_number to use...
                
                '''//Check if columns of HBV BinCut powerDomains match columns of BV BinCut powerDomains.
                If str_mainColumn_content Like "vdd_*" Then '''ex: "VDD_PCPU".
                    For Col = other_voltage_start_point To col_jobBlock_end
                        strTemp = UCase(Trim(ws_def.Cells(Row, Col).value))
                        If Not strTemp Like "*COMMENT*" Then '''20220103

                            '''//Check if the tracking power exists in the column of the header.
                            If strTemp Like "*,*" Then
                                strAry_trackpower = Split(strTemp, ",")
                                powerDomain = UCase(Trim(strAry_trackpower(0)))
                                trackpowerTemp = UCase(Trim(Replace(strTemp, (UCase(strAry_trackpower(0)) & ","), "")))
                            Else
                                powerDomain = strTemp
                                trackpowerTemp = vbNullString
                            End If
                        Else
                            powerDomain = strTemp
                        End If
                        
                        If strTemp Like "*COMMENT*" Then Exit For
                        If dict_EquationVoltages_Domain2Column.item(powerDomain) = Col Then
                            If dict_IsCorePower.Exists(powerDomain) Then
                                If trackpowerTemp = AllBinCut(VddBinStr2Enum(powerDomain)).TRACKINGPOWER Then
                                    '''Do nothing
                                Else
                                    Call Print_Error_Message(Error_Info, "LIB_VDD_BINNING", "Parsing_EquationVoltages", "sheet:" & sheetName & ", cell (row:" & Row & ",column:" & Col & "), content:" & strTemp & ", trackpower of powerDomain is different from cell (row:" & Row_of_Title & ", col:" & dict_BinCutFlow_Domain2Column.item(powerDomain) & ") in the header. Error!!!")
                                    'TheExec.ErrorLogMessage "sheet:" & sheetName & ", cell (row:" & Row & ",column:" & Col & "), content:" & strTemp & ", trackpower of powerDomain is different from cell (row:" & Row_of_Title & ", col:" & dict_BinCutFlow_Domain2Column.item(powerDomain) & ") in the header. Error!!!"    ''20230807: Removed by sanity check
                                End If
                            End If
                        Else
                            Call Print_Error_Message(Error_Info, "LIB_VDD_BINNING", "Parsing_EquationVoltages", "sheet:" & sheetName & ", cell (row:" & Row & ",column:" & Col & "), content:" & strTemp & ", it is different from the header VDD in row" & Row_of_Title & ". Error!!!")
                            'TheExec.ErrorLogMessage strTemp & " in row" & Row & " col" & Col & " of sheet " & sheetName & " is different from the header VDD in row" & Row_of_Title & ". Error!!!"    ''20230807: Removed by sanity check
                        End If
                    Next Col
                
                Else
                    For Col = other_voltage_start_point To other_voltage_start_point + cntVddbinPin - 1
                        '''//Check if the main performance_mode exists in the dictionary "VddbinPmodeDict".
                        If VddbinPmodeDict.Exists(EquationVoltages_temp) Then
                            main_p_mode = VddBinStr2Enum(EquationVoltages_temp)
                            'powerDomain = AllBinCut(main_p_mode).powerPin
                        End If
                        
                        'For i = 0 To cntVddbinPin - 1
                            For passBinCut = binNumStart To binNumStop
                                'col = other_voltage_start_point + i
                                
                                    '''//Check if any testcondition contains "#REF!" or "#NAME?".
                                If IsError(ws_def.Cells(Row, Col).value) Then
                                    testCondition = vbNullString
                                    Call Print_Error_Message(Error_Info, "LIB_VDD_BINNING", "Parsing_EquationVoltages", "sheet:" & sheetName & ", cell (row:" & Row & ", column:" & Col & "). it has the incorrect content. Error!!!")
                                    'TheExec.ErrorLogMessage "sheet:" & sheetName & ", cell (row:" & Row & ", column:" & Col & "). it has the incorrect content. Error!!!"    ''20230807: Removed by sanity check
                                Else
                                    selected_powerDomain = UCase(dict_EquationVoltages_Column2Domain.item(Col))
                                        
                                        '''//Get and trim string of the testCondition from Non_Binning_Rail.
                                    testCondition = LCase(Trim(ws_def.Cells(Row, Col).value))
                                        
                                        '''//Check if testCondition contains any keyword about PassBin(Bin1/BinX/BinY)...
                                    If testCondition Like "*bin1*" Then '''Bin1
                                        bincutNum = 1
                                    ElseIf testCondition Like "*binx*" Then '''BinX
                                        bincutNum = 2
                                    ElseIf testCondition Like "*biny*" Then '''BinY
                                        bincutNum = 3
                                    Else
                                        bincutNum = 0
                                    End If
                                        
                                        '''//Check if bincutNum is greater than PassBinCut_ary(Ubound(PassBinCut_ary)).
                                        '''//PassBinCut_ary(Ubound(PassBinCut_ary)) is the highest Bin number of the BinCut voltage table(sheet "Vdd_Binning_Def").
                                        '''20210802: Modified to check if testCondition contains any keyword about PassBin(Bin1/BinX/BinY) greater than the highest bin number.
                                    If bincutNum > PassBinCut_ary(UBound(PassBinCut_ary)) Then
                                        Call Print_Error_Message(Error_Info, "LIB_VDD_BINNING", "Parsing_EquationVoltages", "sheet:" & sheetName & ",cell:(row " & Row & ", column " & Col & "),testCondition:" & testCondition & ", it contains the incorrect keyword about PassBin greater than the highest bin number for initVddBinCondition. Error!!!")
                                        'TheExec.ErrorLogMessage "sheet:" & sheetName & ",cell:(row " & Row & ", column " & Col & "),testCondition:" & testCondition & ", it contains the incorrect keyword about PassBin greater than the highest bin number for initVddBinCondition. Error!!!"    ''20230807: Removed by sanity check
                                    End If
                                    
                                    '''//Check if testCondition contains "(" but no ")".
                                    If testCondition Like "*(*" And Not (testCondition) Like "*)*" Then
                                        Call Print_Error_Message(Error_Info, "LIB_VDD_BINNING", "Parsing_EquationVoltages", "Please check the cell (row " & Row & ", column " & Col & ") of sheet " & sheetName & ". The cell contains the incorrect format. Error!!!")
                                        'TheExec.ErrorLogMessage "Please check the cell (row " & Row & ", column " & Col & ") of sheet " & sheetName & ". The cell contains the incorrect format. Error!!!"    ''20230807: Removed by sanity check
                                    End If
                                    
                                    '''==========================================================================================================================================='''
                                    '''//Check if testCondition contain keyword "*Evaluate*Bin*" to decide "is_BinCutJob_for_StepSearch" = True (BinCut stepSearch) while isParsingOutsideBinCutFlow = False.
                                    '''==========================================================================================================================================='''
                                    If testCondition Like LCase("*Evaluate*Bin*") And isParsingOutsideBinCutFlow = False Then
                                        strSplitted = Split(LCase(testCondition), LCase("Evaluate Bin"))
                                        
                                        '''//Check if testCondition with keyword "*Evaluate*Bin*" has the correct performance mode.
                                        '''20210414: Modified to add "is_for_BinSearch as Boolean" for AllBinCut(p_mode).
                                        If VddbinPmodeDict.Exists(UCase(Trim(strSplitted(0)))) = True And dict_IsCorePower.Exists(UCase(Trim(strSplitted(0)))) = False Then
                                            AllBinCut(VddBinStr2Enum(UCase(Trim(strSplitted(0))))).is_for_BinSearch = True
                                            is_BinCutJob_for_StepSearch = True
                                        Else
                                            Call Print_Error_Message(Error_Info, "LIB_VDD_BINNING", "Parsing_EquationVoltages", "sheet:" & sheetName & ",cell:(row " & Row & ", column " & Col & "),testCondition:" & testCondition & ". It doesn't contain any correct performance mode in testCondition, please check sheet " & sheetName & ". Error!!!")
                                            'TheExec.ErrorLogMessage "sheet:" & sheetName & ",cell:(row " & Row & ", column " & Col & "),testCondition:" & testCondition & ". It doesn't contain any correct performance mode in testCondition, please check sheet " & sheetName & ". Error!!!"    ''20230807: Removed by sanity check
                                        End If
                                    End If
                                        'selected_powerDomain = UCase(dict_BinCutFlow_Column2Domain.item(col))
                                    'If isParsingOutsideBinCutFlow = False Then
                                    
                                        BinCut(main_p_mode, passBinCut).OutsideBinCut_EquationVoltages(VddBinStr2Enum(selected_powerDomain)) = testCondition
      
                                        'Flag_NonbinningrailOutsideBinCut_parsed = True
                                    'End If
                                End If
                            Next passBinCut
                        'Next i
                    Next Col
                End If '''If str_mainColumn_content Like "*hvcc*"
            End If '''If IsError(ws_def.Cells(row, other_voltage_start_point).Value) Then
'            If Row > 2 Then
'                dict_PBC_ALL.Add PBC_temp, PBC_ALL
'            End If
            Row = Row + 1
        Wend
    End If '''If row_of_title > 0
    
    '''====================================================-====================================================================
    '''[Step3] Check if "cntAdditionalMode" should be "<=" with "MaxAdditionalModeCount".
    '''====================================================-====================================================================
    If cntAdditionalMode > MaxAdditionalModeCount Then
        Call Print_Error_Message(Error_Info, "LIB_VDD_BINNING", "", "sheet:" & sheetName & ",number of BinCut additional modes:" & cntAdditionalMode & ", it is greater than BinCut globalVariable MaxAdditionalModeCount=" & MaxAdditionalModeCount & ". Please check BinCut flow table and globalVariable MaxAdditionalModeCount. Error!!!")
        'TheExec.ErrorLogMessage "sheet:" & sheetName & ",number of BinCut additional modes:" & cntAdditionalMode & ", it is greater than BinCut globalVariable MaxAdditionalModeCount=" & MaxAdditionalModeCount & ". Please check BinCut flow table and globalVariable MaxAdditionalModeCount. Error!!!"    ''20230807: Removed by sanity check
    End If
    
    If isParsingOutsideBinCutFlow = False Then
        '''//Split pin_groups and get each powerDomain, then sort the sequence about p_mode for each powerDomain. BinCut PowerDomain consists of CorePower and OtherRail.
        If FullBinCutPowerinFlowSheet <> "" Then
            pinGroup_BinCut = Split(FullBinCutPowerinFlowSheet, ",")
            
            For i = 0 To UBound(pinGroup_BinCut)
                powerDomain = UCase(pinGroup_BinCut(i))
            
                If dict_IsCorePower.item(powerDomain) = True Then '''CorePower
                    If FullCorePowerinFlowSheet <> "" Then
                        FullCorePowerinFlowSheet = FullCorePowerinFlowSheet & "," & powerDomain
                    Else
                        FullCorePowerinFlowSheet = powerDomain
                    End If
                    
                    dict_IsCorePowerInBinCutFlowSheet.Add powerDomain, True
                Else '''OtherRail
                    If FullOtherRailinFlowSheet <> "" Then
                        FullOtherRailinFlowSheet = FullOtherRailinFlowSheet & "," & powerDomain
                    Else
                        FullOtherRailinFlowSheet = powerDomain
                    End If
                    
                    dict_IsCorePowerInBinCutFlowSheet.Add powerDomain, False
                End If
            Next i
        Else
            Call Print_Error_Message(Error_Info, "LIB_VDD_BINNING", "Parsing_EquationVoltages", "FullBinCutPowerinFlowSheet should not be empty. Please check Vdd_Binning_Def_appA and Non_Binning_Rail. Error!!!")
            'TheExec.ErrorLogMessage "FullBinCutPowerinFlowSheet should not be empty. Please check Vdd_Binning_Def_appA and Non_Binning_Rail. Error!!!"    ''20230807: Removed by sanity check
        End If
        
        '''//Split pin_groups and get powerDomains of BinCut CorePower and OtherRail.
        '''CorePower
        If FullCorePowerinFlowSheet <> "" Then
            pinGroup_CorePower = Split(FullCorePowerinFlowSheet, ",")
        Else
            Call Print_Error_Message(Error_Info, "LIB_VDD_BINNING", "Parsing_EquationVoltages", "FullCorePowerinFlowSheet should not be empty. Please check Vdd_Binning_Def_appA and Non_Binning_Rail. Error!!!")
            'TheExec.ErrorLogMessage "FullCorePowerinFlowSheet should not be empty. Please check Vdd_Binning_Def_appA and Non_Binning_Rail. Error!!!"    ''20230807: Removed by sanity check
        End If
        
        '''OtherRail
        If FullOtherRailinFlowSheet <> "" Then
            pinGroup_OtherRail = Split(FullOtherRailinFlowSheet, ",")
        Else
            Call Print_Error_Message(Error_Info, "LIB_VDD_BINNING", "Parsing_EquationVoltages", "FullOtherRailinFlowSheet should not be empty. Please check Vdd_Binning_Def_appA and Non_Binning_Rail. Error!!!")
            'TheExec.ErrorLogMessage "FullOtherRailinFlowSheet should not be empty. Please check Vdd_Binning_Def_appA and Non_Binning_Rail. Error!!!"    ''20230807: Removed by sanity check
        End If
    End If
Exit Function
errHandler:
    Call Print_Error_Message(Error_Info, "LIB_VDD_BINNING", "Parsing_EquationVoltages") 'Add ErrHandler 2023/08/01
    If AbortTest Then Exit Function Else Resume Next
End Function

'20201223: Modified to check if any sheet contains Keyword in sheetName.
'20201222: Modified to revise the vbt function "initVddBinCondition" for multiple "Non_Binning_Rail_Outside_BinCut" sheets.
'20201222: Created to parse multiple "Non_Binning_Rail_Outside_BinCut" sheets.
Public Function parsing_OutsideBinCut_flow_table(keyword_sheetName As String)
On Error GoTo errHandler 'Add ErrHandler 2023/08/01
    Dim i As Long
    Dim count_WorkSheet As Integer
    Dim outsideBincutSheetsArr() As String
    Dim idxSheet As Integer
    Dim maxSheetsNum As Integer
    Dim sheetName As String
'''//==================================================================================================================================================================================//'''
'''//Note:
'''Please check sheetName keyword "Non_Binning_Rail_Outside" of sheetName for the vbt functions "initVddBinCondition" and "parsing_OutsideBinCut_flow_table".
'''//==================================================================================================================================================================================//'''
    '''init
    idxSheet = -1

    '''//Check if the sheet exists.
    If keyword_sheetName = "Non_Binning_Rail_Outside" Then
        count_WorkSheet = Application.ActiveWorkbook.Worksheets.Count
        
        '''//Check if sheet name contains keyword_sheetName.
        For i = 1 To count_WorkSheet
            If ActiveWorkbook.Worksheets(i).name Like "*" & keyword_sheetName & "*" Then
                idxSheet = idxSheet + 1
                ReDim Preserve outsideBincutSheetsArr(idxSheet)
                outsideBincutSheetsArr(idxSheet) = ActiveWorkbook.Worksheets(i).name
            End If
        Next i
    ElseIf keyword_sheetName <> "" Then
        Call Print_Error_Message(Error_Info, "LIB_VDD_BINNING", "parsing_OutsideBinCut_flow_table", keyword_sheetName & " is not the correct keyword to find Outside BinCut flow table for parsing_OutsideBinCut_flow_table. Error!!!")
        'TheExec.Datalog.WriteComment keyword_sheetName & " is not the correct keyword to find Outside BinCut flow table for parsing_OutsideBinCut_flow_table. Error!!!"
        'TheExec.ErrorLogMessage keyword_sheetName & " is not the correct keyword to find Outside BinCut flow table for parsing_OutsideBinCut_flow_table. Error!!!"    ''20230807: Removed by sanity check
        Exit Function
    Else '''If keyword_sheetName is empty...
        Exit Function
    End If
    
    '''//Use sheet-loop to parse each Outside BinCut flow table.
    If idxSheet > -1 Then
        For i = 0 To idxSheet
            sheetName = outsideBincutSheetsArr(i)
            
            '''//Parsing each of sheets with Keyword in sheetName.
            initVddBinCondition sheetName
        Next i
    End If
Exit Function
errHandler:
    Call Print_Error_Message(Error_Info, "LIB_VDD_BINNING", "parsing_OutsideBinCut_flow_table") 'Add ErrHandler 2023/08/01
    If AbortTest Then Exit Function Else Resume Next
End Function

'20210901: Modified to rename "IndexLevelPerSite As New SiteLong" as "Step_Current As New SiteLong" for Public Type Instance_Info.
'20210812: C651 Toby updated the rules that Product voltage=BinCut voltage+binning_GuardBand.
'20210806: Modified to get and print Guardband(GB) according to the BinCut testjob.
'20210706: Modified to replace is_BinCutJob_for_StepSearch with inst_info.is_BinSearch.
'20210219: Modified to check the flag "Flag_Skip_Printing_Safe_Voltage" to skip printing BV strings of BinCut Safe Voltages.
'20201111: Modified to replace the vbt function name "print_bincut_power" with "print_bincut_voltage".
'20201104: Modified to replace "bincutJobName = "cp1" with inst_info.is_binsearch=True.
'20201102: Modified to update "inst_info.is_BV_Safe_Voltage_printed" and "inst_info.is_BV_Payload_Voltage_printed".
'20201029: Modified to remove the redundant arguments "str_dynamic_offset() As String" and "str_Selsrm_DSSC_Info() As String" from print_bincut_power.
'20201027: Modified to use "Public Type Instance_Info".
'20200925: Modified the branch for "indexstep_per_site".
'20200319: Modified for "Flag_PrintDcvsShadowVoltage".
'20200214: Modified to print dynamic_offset.
'20200214: Modified to print eqn information for payload voltages.
'20200211: Modified to get init voltage for DCVS shadow voltages.
'20200211: Modified to replace "FlowTestCondName" with "AdditionalModeName".
'20200206: Modified to replace "print_main_power_init" with "print_bincut_power".
'20200203: Created to merge the functions: print_main_power, print_alt_power, print_main_power_payload, print_alt_power_payload.
'20191219: Modified to use dictionaries of Domain2Pin and Pin2Domain.
'20191105: Modified to print offsetTestType.
'20191002: Modified to add BinCut voltageType.
'20180910: Modified to control print BinCut voltages by "remove_printing_voltage".
Public Function print_bincut_voltage(inst_info As Instance_Info, Optional passBinCut As SiteLong, Optional remove_printing_voltage As Boolean = False, Optional Flag_PrintDcvsShadowVoltage As Boolean = False, _
                                        Optional voltageType As Integer = BincutVoltageType.None, Optional DcSpecsCategoryForInitPat As String = vbNullString)
On Error GoTo errHandler 'Add ErrHandler 2023/08/01
    Dim site As Variant
    Dim i As Long, j As Long
    Dim powerDomain As String
    Dim powerPin As String
    Dim PowerPins As String
    Dim strTemp As String
    Dim strPrefix As String
    Dim strOutput As String
    Dim strOutputEQN As String
    Dim voltage_PowerDomain As Double
    Dim performance_mode As String
    Dim dbl_GB_BinCutJob As Double
    Dim vdd_voltage_val_tmp() As New SiteDouble         'Dim vdd_voltage_val_tmp(MaxBincutPowerdomainCount) As New SiteDouble
    Dim powerPinValue As New PinListData
'''//==================================================================================================================================================================================//'''
'''//Note:
'''str_Selsrm_DSSC_Info     : store info about SELSRM bits comparison.
'''str_Selsrm_DSSC_Bit      : store info about SELSRM bits sequence(LSB->MSB) of each site.
'''20210812: C651 Toby updated the rules that Product voltage=BinCut voltage+binning_GuardBand.
'''//==================================================================================================================================================================================//'''
    '''//Check the flag "Flag_Skip_Printing_Safe_Voltage" to skip printing BV strings of BinCut Safe Voltages.
    Dim vdd_voltage_val_PLD As New PinListData
    ReDim vdd_voltage_val_tmp(cntVddbinPin) As New SiteDouble
    If remove_printing_voltage = True Or (Flag_Skip_Printing_Safe_Voltage = True And voltageType = BincutVoltageType.SafeVoltage) Then
        If Flag_Skip_Printing_Safe_Voltage = True And voltageType = BincutVoltageType.SafeVoltage Then
            TheExec.Datalog.WriteComment "****************separated for BinCut step****************"
        End If
        inst_info.is_BV_Safe_Voltage_printed = True
'        inst_info.is_BV_Payload_Voltage_printed = True
        Exit Function
    Else
        '''//init
        strPrefix = vbNullString
        
        '''//Check if Flag_PrintDcvsShadowVoltage is enabled to print DCVS shadow voltages (calculation values of BinCut payload voltages).
        If Flag_PrintDcvsShadowVoltage = True Then
            TheExec.Datalog.WriteComment "Print DCVS shadow voltages (from calculation), not DCVS real values!!!"
        End If

        '''//Check the performance mode if it has the additional mode, ex: MS003_GPU.
        '''If special_voltae_setup = True, it means that the performance mode has the additional mode.
        If inst_info.Special_Voltage_setup = True Then
            performance_mode = VddBinName(inst_info.p_mode) & "_" & AdditionalModeName(inst_info.addi_mode)
        Else
            performance_mode = VddBinName(inst_info.p_mode)
        End If
        
        '''//Get the prefix of voltage string.
        If voltageType = BincutVoltageType.InitialVoltage Then
            strPrefix = "Initial_Voltage_" & performance_mode
        Else
            strPrefix = "BV_" & performance_mode
        End If
        
        If glb_TesterType = "UltraFLEXplus" Then
            For i = 0 To UBound(pinGroup_BinCut)
                powerDomain = pinGroup_BinCut(i)
                powerPin = Get1stPinFromPingroup(VddbinDomain2Pin(powerDomain))
                If PowerPins = "" Then
                    PowerPins = powerPin
                Else
                    PowerPins = PowerPins & "," & powerPin
                End If
'#If isUFP = True Then
'                vdd_voltage_val_tmp(VddBinStr2Enum(powerDomain)) = thehdw.DCVS.Pins(powerPin).Voltage.ValuePerSite
'#End If
            Next i
            TheHdw.DCVS.pins(PowerPins).Meter.mode = tlDCVSMeterVoltage
            powerPinValue = TheHdw.DCVS.pins(PowerPins).Meter.Read(tlStrobe, 10, 10000, tlDCVSMeterReadingFormatAverage)
        End If
        
        For Each site In TheExec.sites
            '''init
            strOutput = vbNullString
            strOutputEQN = vbNullString
            strTemp = vbNullString
            
            '''//Print strings of dynamic_offset or SELSRAM_DSSC prior to BV string of BinCut payload voltages.
            If (voltageType <> BincutVoltageType.InitialVoltage And voltageType <> BincutVoltageType.SafeVoltage And voltageType <> BincutVoltageType.None) And inst_info.offsetTestTypeIdx <> ldcbfd Then
                '''//dynamic_offset of binning p_mode.
                If inst_info.str_dynamic_offset(site) <> "" Then
                    TheExec.Datalog.WriteComment inst_info.str_dynamic_offset(site)
                End If
                
                '''//SELSRAM_DSSC
                If inst_info.str_Selsrm_DSSC_Info(site) <> "" Then
                    TheExec.Datalog.WriteComment inst_info.str_Selsrm_DSSC_Info(site)
                End If
            End If
            
            '''//In testjob "CP1", datalog always show EQN and passbin, C, M and CPGB.
            '''20190716: Modified to unify the unit for IDS. ids_current with unit mA.
            If inst_info.is_BinSearch = True And (voltageType <> BincutVoltageType.InitialVoltage And voltageType <> BincutVoltageType.SafeVoltage And voltageType <> BincutVoltageType.None) Then
                If inst_info.Step_Current(site) <> -1 Then
                    '''//Get the matched Guardband(GB) according to the BinCut testjob.
                    '''20210812: C651 Toby updated the rules that Product voltage=BinCut voltage+binning_GuardBand.
                    dbl_GB_BinCutJob = BinCut(inst_info.p_mode, DYNAMIC_VBIN_IDS_ZONE(inst_info.p_mode).passBinCut(inst_info.Step_Current)).CP_GB(DYNAMIC_VBIN_IDS_ZONE(inst_info.p_mode).EQ_Num(inst_info.Step_Current) - 1)
                    
                    strOutputEQN = strPrefix & "," & site & ", EQN = " & DYNAMIC_VBIN_IDS_ZONE(inst_info.p_mode).EQ_Num(inst_info.Step_Current) & _
                                    ", PASSBIN = " & DYNAMIC_VBIN_IDS_ZONE(inst_info.p_mode).passBinCut(inst_info.Step_Current) & _
                                    ", C = " & DYNAMIC_VBIN_IDS_ZONE(inst_info.p_mode).c(inst_info.Step_Current) & _
                                    ", M = " & DYNAMIC_VBIN_IDS_ZONE(inst_info.p_mode).m(inst_info.Step_Current) & _
                                    ", GB = " & dbl_GB_BinCutJob & _
                                    ", IDS = " & inst_info.ids_current & " mA"
                Else
                    strOutputEQN = vbNullString
                End If
            End If
            
            '''//Since tracking power is already added to pinGroup, no need to print tracking power...
            '''As per discussion with SWLINZA, ZHHUANG, and PCLIN. We decided to measure the voltage of 1st powerPin to print each power domain.
            '''We also merged the vbt code for online and offline tests. Only read the real voltage value from DCVS.
            For i = 0 To UBound(pinGroup_BinCut)
                powerDomain = pinGroup_BinCut(i)
                
                '''//Only payload voltages can use "DCVS_shadow_voltages".
                If Flag_PrintDcvsShadowVoltage = True And voltageType <> BincutVoltageType.InitialVoltage And voltageType <> BincutVoltageType.SafeVoltage And voltageType <> BincutVoltageType.None Then
                    voltage_PowerDomain = BinCut_Payload_Voltage(VddBinStr2Enum(powerDomain)) / 1000 '''DCVS should use unit: V
                Else
                    '''//Read the real voltage value from DCVS.
                    '''ToDo: For project with UltraFlexPlus, it can directly read voltage values from DCVS by using "ValuePerSite"...
                    powerPin = Get1stPinFromPingroup(VddbinDomain2Pin(powerDomain))
                    '''ToDo: Check if powerPin is DCVS or DCVI by checking VddbinPinDcvsType...
                    If glb_TesterType = "Jaguar" Then
                        voltage_PowerDomain = TheHdw.DCVS.pins(powerPin).Voltage.value
                    Else

'                        voltage_PowerDomain = vdd_voltage_val_tmp(VddBinStr2Enum(powerDomain))(site)

                        voltage_PowerDomain = powerPinValue.pins(powerPin).value
                    End If
                End If
                
                If strTemp <> "" Then
                    strTemp = strTemp & "," & powerDomain & "=" & Format(voltage_PowerDomain, "0.000")
                Else
                    strTemp = powerDomain & "=" & Format(voltage_PowerDomain, "0.000")
                End If
            Next i
            
            '''//Print testType and offsetType at the end of BV string.
            If voltageType <> BincutVoltageType.None Then
                If inst_info.offsetTestTypeIdx <> ldcbfd Then
                    strOutput = strPrefix & "," & site & "," & strTemp & "," & " (" & BincutVoltageTypeName(voltageType) & "," & TestTypeName(inst_info.offsetTestTypeIdx) & ")"
                Else
                    strOutput = strPrefix & "," & site & "," & strTemp & "," & " (" & BincutVoltageTypeName(voltageType) & ")"
                End If
            End If
            
            '''//Print out the string in the datalog.
            TheExec.Datalog.WriteComment strOutput
            
            '''//Update the status of "inst_info.is_BV_Safe_Voltage_printed" and "inst_info.is_BV_Payload_Voltage_printed" to control printing BV strings into the datalog.
            If voltageType = BincutVoltageType.SafeVoltage Then
                inst_info.is_BV_Safe_Voltage_printed = True
            ElseIf (voltageType <> BincutVoltageType.InitialVoltage And voltageType <> BincutVoltageType.SafeVoltage And voltageType <> BincutVoltageType.None) Then
                inst_info.is_BV_Payload_Voltage_printed = True
            End If
            
            '''//Print info about BinCut EQN for BinSearch.
            If inst_info.is_BinSearch = True And strOutputEQN <> "" Then
                strOutputEQN = strOutputEQN & "," & strTemp
            
                '''ex:BV_VDD_SOC_MS001,1,EQN = 7, PASSBIN = 1, C = 609.375, M = 70, CPGB = 78.125, _
                '''IDS = 27.4 mA,,VDD_PCPU=0.752,VDD_ECPU=0.752,VDD_GPU=0.752,VDD_SOC=0.752,VDD_DCS_DDR=0.752,VDD_AVE=0.752,VDD_DISP=0.752, _
                '''VDD_SRAM_CPU=0.752,VDD_SRAM_ANE=0.752,VDD_SRAM_GPU=0.752,VDD_SRAM_SOC=0.752,VDD_FIXED=0.800,VDD_LOW=0.735
                TheExec.Datalog.WriteComment strOutputEQN
            End If
        Next site
    End If
Exit Function
errHandler:
    Call Print_Error_Message(Error_Info, "LIB_VDD_BINNING", "print_bincut_voltage") 'Add ErrHandler 2023/08/01
    If AbortTest Then Exit Function Else Resume Next
End Function

'20210207: Modified to move the vbt code of switching DCVS Valt from print_voltage_info_before_FuncPat to GradeSearch_XXX_VT.
'20201111: Modified to replace the vbt function name "print_bincut_power" with "print_bincut_voltage".
'20201029: Modified to use inst_info.previousDcvsOutput and inst_info.currentDcvsOutput.
'20201029: Modified to use inst_info.is_BV_Safe_Voltage_printed and inst_info.is_BV_Payload_Voltage_printed.
'20201029: Modified to remove the redundant arguments "str_dynamic_offset() As String" and "str_Selsrm_DSSC_Info() As String" from print_voltage_info_before_FuncPat.
'20201027: Modified to use "Public Type Instance_Info".
'20201026: Modified to revise the vbt code for TD pattern burst proposed by C651 Toby.
'20201008: Modified to move the vbt code of printing payload voltages for Pattern Burst(not decompose pattern set) from "print_voltage_info_before_FuncPat" to "prepare_DCVS_Output_for_RailSwitch".
'20200921: Modified to check if "Test_Type = TestType.Mbist".
'20200918: Created to print BinCut voltage before running FuncPat.
'20200113: As per discussion with Leon/Jeff/PSYAO/Minder/PCLIN, we decided to print payload voltages for pattern bursted without decomposing pattern.
Public Function print_voltage_info_before_FuncPat(inst_info As Instance_Info)
On Error GoTo errHandler
'''//==================================================================================================================================================================================//'''
'''//Note:
'''<For projects without Rail-Switch>
'''   Print BinCut voltages of Func Pattern for TD or Mbist test instances
'''<For projects with Rail-Switch>
'''   Print BinCut voltages of Payload Pattern(Func Pattern) for Mbist test instances
'''   Print BinCut voltages of Init Pattern in Func Patsets for TD test instances
'''//==================================================================================================================================================================================//'''
    If inst_info.Test_Type = testType.Mbist Then '''ex: "*cpu*bist*", "*gfx*bist*", "*gpu*bist*", "*soc*bist*".
        '''//For Mbist test, it prints BinCut payload voltages before running Payload pattern. So that it can use DCVS shadow voltages.
        If inst_info.is_BV_Payload_Voltage_printed = False Then
            print_bincut_voltage inst_info, CurrentPassBinCutNum, Flag_Remove_Printing_BV_voltages, Flag_PrintDcvsShadowVoltage, BincutVoltageType.PayloadVoltage
        End If
    Else '''For TD/SA/SCAN test instances...
        '''//Use the flags "inst_info.is_BV_Safe_Voltage_printed" and "inst_info.is_BV_Payload_Voltage_printed = False" to avoid printing duplicate BV strings of safe voltages and payload voltages for TD init+pl+init+pl patset.
        If Flag_Enable_Rail_Switch Then '''For projects with Rail Switch
            If inst_info.is_BV_Safe_Voltage_printed = False Then
                print_bincut_voltage inst_info, CurrentPassBinCutNum, Flag_Remove_Printing_BV_voltages, Flag_PrintDcvsShadowVoltage, BincutVoltageType.SafeVoltage
            End If
        Else
            If inst_info.is_BV_Payload_Voltage_printed = False Then
                print_bincut_voltage inst_info, CurrentPassBinCutNum, Flag_Remove_Printing_BV_voltages, Flag_PrintDcvsShadowVoltage, BincutVoltageType.PayloadVoltage
            End If
        End If
    End If
Exit Function
errHandler:
    Call Print_Error_Message(Error_Info, "LIB_VDD_BINNING", "print_voltage_info_before_FuncPat") 'Add ErrHandler 2023/08/01
    If AbortTest Then Exit Function Else Resume Next
End Function

'20210528: Modified to add inst_info to call Calculate_Harvest_Core_DSSC_Source.
'20210513: Modified to use Calculate_Harvest_Core_DSSC_Source.
'20201111: Modified to replace the vbt function name "print_bincut_power" with "print_bincut_voltage".
'20201111: Modified to move "prepare_DCVS_Output_for_RailSwitch" from modudle LIB_Vdd_Binning_customer to LIB_VDD_BINNING.
'20201029: Modified to use inst_info.previousDcvsOutput and inst_info.currentDcvsOutput.
'20201029: Modified to use inst_info.is_BV_Safe_Voltage_printed and inst_info.is_BV_Payload_Voltage_printed.
'20201029: Modified to remove the redundant arguments "str_dynamic_offset() As String" and "str_Selsrm_DSSC_Info() As String" from prepare_DCVS_Output_for_RailSwitch.
'20201027: Modified to use "Public Type Instance_Info".
'20201014: Modified to use "Check_PayloadPattern_with_DCVS" for offline, requested by Leon Weng.
'20201014: Modified to check if powerDomain is not "PRESERVED" or "RESERVED", requested by Leon Weng.
'20201008: Modified to replace "PrintedBVinDatalog" with "is_BV_Payload_Voltage_printed"
'20201008: Modified to move the vbt code of printing payload voltages for Pattern Burst(not decompose pattern set) from "print_voltage_info_before_FuncPat" to "prepare_DCVS_Output_for_RailSwitch".
'20201008: Modified to force DCVS to Valt and skip Check_PayloadPattern_with_DCVS for offline.
'20201008: Modified to check if "Test_Type = testType.TD".
'20201008: Modified to add the vbt code of printing BinCut payload voltages for offline simulation.
'20201007: Modified to check vbump only for online tests. requested by Leon Weng.
'20200924: Modified to add the argument "idxBlock_Selsrm_Pattern" for SELSRM DSSC signal setup.
'20200918: Created to decide syncup DCVS output and print BinCut payload voltages.
'20191009: Modified For open socket or offline simulation.
'20190626: As the discussion with TSMC PSYAO, we found vbump didn't exist in all BinCut powerDomain of BinCut patterns.
'20190121: Modified by Oscar. Signal is reusable so we remove the Recalculation and Resend.
'20181119: Resend DSSC Selsram bits for the pattern.
Public Function prepare_DCVS_Output_for_RailSwitch(inst_info As Instance_Info, str_pattern As String, Optional idxBlock_Selsrm_Pattern As Integer = -1)
On Error GoTo errHandler 'Add ErrHandler 2023/08/01
    Dim i As Integer
    Dim selSramPat As String
'''//==================================================================================================================================================================================//'''
'''//Note:
'''//tlDCVSVoltageOutput: This property gets or sets the output DAC used to force voltage (main or alternate).
'''It can detect DCVS output Vmain/Valt of the binning CorePower powerDomains.
'''The returned information is one of the following values:
'''     tlDCVSVoltageMain   : 1 ==> Main output voltage DAC.
'''     tlDCVSVoltageAlt    : 2 ==> Alternate output voltage DAC.
'''//==================================================================================================================================================================================//'''
    If inst_info.Test_Type = testType.TD Then
        '''//Notice!!! for TD tests, FuncPat pattern group consists of INIT and Payload patterns.
        '''***************************************************************************'''
        '''//TD patt_group (init+payload1+init+payload2)
        '''//Since Payload patterns in the group use the same BinCut payload voltages, it can skip SelSram bits calculation for the second and following payload patterns.
        '''***************************************************************************'''
        If idxBlock_Selsrm_Pattern > -1 Then
            '''//Resend DSSC Selsram bits for the pattern.
            If LCase(str_pattern) Like LCase(SelsramMapping(idxBlock_Selsrm_Pattern).Pattern) Then  '''for DSSC patt_group(init+pl+init+pl), 20181115
                selSramPat = str_pattern
                TheHdw.DSSC.pins(inst_info.selsrm_DigSrc_Pin).Pattern(str_pattern).Source.Signals.DefaultSignal = inst_info.selsrm_DigSrc_SignalName
            End If
        End If
        
        '''//Calculate DSSC bit sequence of Harvest Core DSSC (FSTP and MultiFSTP), then do DSSC DigSrcWaveSetup with the pattern and DSSC bit sequence before running the pattern.
        '''Since PrePatt and FuncPat share this utility, we could use this as the common funciton for Harvest Core DSSC.
        '''Flag_Harvest_Core_DSSC_Ready = True if HarvPmodeTable and HARV_Pmode_Table.
        If inst_info.MultiFSTP_Enable = True And inst_info.idxDigSrcPattern <= UBound(inst_info.digSrcPatterns) Then
            If InStr(str_pattern, inst_info.digSrcPatterns(inst_info.idxDigSrcPattern)) > 1 Then
                Call Calculate_Harvest_Core_DSSC_Source_For_UserFunction(inst_info.digSrcPatterns(inst_info.idxDigSrcPattern), CStr(inst_info.digSrcLabel(inst_info.idxDigSrcPattern)), inst_info.Harvest_Core_DigSrc_Pin, inst_info.Harvest_Core_DigSrc_SignalName)
                inst_info.idxDigSrcPattern = inst_info.idxDigSrcPattern + 1
            End If
            'Call Calculate_Harvest_Core_DSSC_Source_New(inst_info.inst_name, VddBinName(inst_info.p_mode), str_pattern, inst_info.Harvest_Core_DigSrc_Pin, inst_info.Harvest_Core_DigSrc_SignalName)
        Else
             Call Calculate_Harvest_Core_DSSC_Source(inst_info.inst_name, VddBinName(inst_info.p_mode), str_pattern, inst_info.Harvest_Core_DigSrc_Pin, inst_info.Harvest_Core_DigSrc_SignalName, inst_info.Pattern_Pmode, inst_info.By_Mode)
        End If
        If Flag_Enable_Rail_Switch Then '''For projects with Rail Switch
            '''20190626: As the discussion with TSMC PSYAO, we found vbump didn't exist in all BinCut powerDomain of BinCut patterns.
            '''//It detects that Any of SELSRAM powerDomains is switched to Valt.
            '''//If that, all BinCut powerDomains will be switched to Valt.
            If inst_info.InitVmain_PayloadValt = True Then
                inst_info.currentDcvsOutput = inst_info.Pat2DCVSOutput(str_pattern)
                If inst_info.previousDcvsOutput <> inst_info.currentDcvsOutput Then
                    select_DCVS_output_for_powerDomain inst_info.Pat2DCVSOutput(str_pattern)
                    If inst_info.Pat2DCVSOutput(str_pattern) = tlDCVSVoltageMain Then
                        restore_core_power_vddbinning inst_info.InitValt_CorePowerStored
                        'print_bincut_voltage inst_info, CurrentPassBinCutNum, Flag_Remove_Printing_BV_voltages, Flag_PrintDcvsShadowVoltage, BincutVoltageType.InitialVoltage
                        TheExec.Datalog.WriteComment "Bincut voltage switch to Vmain!!"
                    Else
                        Set_PayloadVoltage_to_DCVS Flag_Enable_Rail_Switch, pinGroup_BinCut, BinCut_Payload_Voltage
                        TheHdw.Wait 0.001
                        print_bincut_voltage inst_info, CurrentPassBinCutNum, Flag_Remove_Printing_BV_voltages, Flag_PrintDcvsShadowVoltage, BincutVoltageType.PayloadVoltage
                    End If
                End If
                inst_info.previousDcvsOutput = inst_info.currentDcvsOutput
            Else
                inst_info.previousDcvsOutput = inst_info.currentDcvsOutput      ''20221128: Init DVCS output
                For i = 0 To UBound(selsramLogicPingroup)
                    '''20201014: Modified to check if powerDomain is not "PRESERVED" or "RESERVED", requested by Leon Weng.
                    If UCase(selsramLogicPingroup(i)) <> "PRESERVED" And UCase(selsramLogicPingroup(i)) <> "RESERVED" Then
                        If (TheHdw.DCVS.pins(selsramLogicPingroup(i)).Voltage.Output = tlDCVSVoltageAlt) Then
                            inst_info.currentDcvsOutput = tlDCVSVoltageAlt
                            ''Define: 1 selsrm pin Valt, inst_info -> Valt
                            Exit For
                        Else
                            ''Define: All selsrm pin Vmain, inst_info -> Vmain
                            inst_info.currentDcvsOutput = tlDCVSVoltageMain
                        End If
                    End If
                Next i
                
                '''//DCVS should be switched to Valt by Pattern with vbump prior to Payload pattern.
                '''//Check if vbump in patset before running the payload pattern, especially for project with rail-switch.
                '''20201007: Modified to check vbump only for online tests. requested by Leon Weng.
                If LCase(str_pattern) Like "*_pl*" Or Not (LCase(str_pattern) Like "*_in*") Then '''1st pattern of non-init patterns.
                    '''**********************************************************************************************************'''
                    '''Offline tests and OpenSocket can't detect Vmain or Valt of DCVS.
                    '''So that we use the keyword "*_pl*" to detect the payload patterns and switch DCVS to Valt for TD instance.
                    '''**********************************************************************************************************'''
                    If Flag_VDD_Binning_Offline = True Or TheExec.flow.enableWord("Vddbinning_OpenSocket") = True Then '''offline or OpenSocket
                        select_DCVS_output_for_powerDomain tlDCVSVoltageAlt
                        inst_info.currentDcvsOutput = tlDCVSVoltageAlt
                    End If
                    
                    If inst_info.enable_DecomposePatt = True Then
                        Call Check_PayloadPattern_with_DCVS(inst_info.inst_name, Flag_Enable_Rail_Switch, str_pattern, inst_info.currentDcvsOutput, inst_info.enable_DecomposePatt)
                    End If
                End If
                
                '''//Projects with Rail Switch might have the incomplete pin listed in the patterns, so that it detects the output status of binning power to sync up other pins.
                '''ToDo: Maybe it can skip SyncUp_DCVS_Output if "inst_info.enable_DecomposePatt = False"...
                If Flag_SyncUp_DCVS_Output_enable Then
                    If inst_info.previousDcvsOutput <> inst_info.currentDcvsOutput Then       ''20221128: Only do SyncUp when voltage state changed
                        Call SyncUp_DCVS_Output(inst_info.p_mode, inst_info.currentDcvsOutput, SyncUp_PowerPin_Group) '''This is to sync up logic powers and sram powers on the same DCVS output (for TD testing)
                    End If
                End If
                
                '''//Print BinCut payload voltages before running TD payload patterns of projects with Rail-Switch.
                If inst_info.is_BV_Payload_Voltage_printed = False Then
                    If inst_info.previousDcvsOutput = tlDCVSVoltageMain And inst_info.currentDcvsOutput = tlDCVSVoltageAlt Then
                        '''==============================================================================================='''
                        '''[Note]: For projects with Rail Switch, it can detect the first transition Vmain-> Valt of the binning CorePower and print BinCut voltages for payload.
                        '''Especially, for those pattern set (INIT-> PL1 -> INIT -> PL2), it only needs to print BinCut payload voltages once.
                        '''No need to print BinCut payload voltages for PL2 again.
                        '''==============================================================================================='''
                        print_bincut_voltage inst_info, CurrentPassBinCutNum, Flag_Remove_Printing_BV_voltages, Flag_PrintDcvsShadowVoltage, BincutVoltageType.PayloadVoltage
                    ElseIf inst_info.enable_DecomposePatt = False Then '''without decomposing pattern sets
                        '''==============================================================================================='''
                        '''//For TD/SCAN test, it prints BinCut payload voltages (values from calculation, not from DCVS) before running pattern bursted without decomposing pattern.
                        '''20200113: As per discussion with Leon/Jeff/PSYAO/Minder/PCLIN:
                        '''we decided to use shadow voltage (calculation value) for printing BinCut payload voltages for pattern bursted without decomposing pattern.
                        '''==============================================================================================='''
                        print_bincut_voltage inst_info, CurrentPassBinCutNum, Flag_Remove_Printing_BV_voltages, True, BincutVoltageType.PayloadVoltage
                    End If
                End If
                
                '''//Store current status to "PreviousDcvsOutput" for the control of printing payload voltage in the datalog.
                inst_info.previousDcvsOutput = inst_info.currentDcvsOutput
            End If
        End If
    End If
    
    '''//Check if any alarm exists.
    '''==============================================================================================='''
    '''This method forces an alarm check. It determines whether alarms are present and reports on them.
    '''This method clears alarms. For this reason, do not use it for monitoring alarms during debugging.
    '''20200106: As per discussion with SWLINZA, he suggested us to add this to check any alarm.
    '''==============================================================================================='''
If LCase(glb_TesterType) = "jaguar" Then TheHdw.Alarms.Check 'WWW20220928


Exit Function
errHandler:
    Call Print_Error_Message(Error_Info, "LIB_VDD_BINNING", "prepare_DCVS_Output_for_RailSwitch") 'Add ErrHandler 2023/08/01
    If AbortTest Then Exit Function Else Resume Next
End Function

'20210305: Modified to add the argument "siteResult" to the vbt function "StoreCapFailcycle".
'20210129: Modified to revise the vbt code for DevChar.
'20210125: Modified to remove "Optional voltage_Pmode_EQNbased As SiteDouble" from the arguments of the vbt function "run_patt_from_FuncPat_for_BinCut".
'20201210: Created to run pattern decomposed from FuncPat patset for BinCut online and offline.
'20201118: Modified to use "TheHdw.Digital.Patgen.PatternBurstPassedPerSite" to get siteResult of pattern pass/fail.
Public Function run_patt_from_FuncPat_for_BinCut(inst_info As Instance_Info, indexPatt As Long, str_pattern As String, funcPatPass As SiteBoolean, _
                                                    Optional idxBlock_Selsrm_Pattern As Integer = -1, Optional CaptureSize As Long, Optional failpins As String)
On Error GoTo errHandler 'Add ErrHandler 2023/08/01
    Dim siteResult As New SiteBoolean
    Dim HarvCnt As Long, ComCnt As Long
    Dim HarvFailResult() As New SiteBoolean, CommnFailResult() As New SiteBoolean   'this is fail flag result
    Dim HarvFailPinCnt() As New SiteLong, CommnFailPinCnt() As New SiteLong
    Dim strAry_PathSplit As Variant, strAry_PatNameSplit As Variant
    Dim TempDigitalPin() As String
    Dim chkloop As Long, TotalHarvFailPinCnt As Long, TotalCommnFailPinCnt As Long
    Dim HarvFailResultOr As Boolean, CommnFailResultOr As Boolean
    Dim PinGroup_disable_compare() As String, Disable_PinGrp_cnt As Long, PinGrp_idx As Integer   '''20220407, Disable compare harvPinGroup when doing regular BV search
    Dim site As Variant 'Carter, 20240304
    Dim F_PayloadVoltage_set As Boolean
'''//==================================================================================================================================================================================//'''
'''//Note:
'''20210531: Modified to update theExec.sites.Selected for MultiFSTP before running PrePatt in run_prepatt_decompose_VT.
'''It seemed that theExec.sites.Selected masked the failed site (not siteShutDown). But the site still ran pattern.test without updating test results.
'''Discussed this with Chihome. He saw this is ancient projects, and he suggested us to check if test results were correct.
'''We checked test results, and it seemed no error with PassBin and EQN.
'''//==================================================================================================================================================================================//'''
'''//==================================================================================================================================================================================//'''
'''//Run pattern decomposed from FuncPatt patset, and get siteResult of pattern pass/fail.
'''step1: Sync up DCVS output and print BinCut payload voltage for projects with Rail Switch for TD instance.
'''step2: Run the pattern decomposed from FuncPat, and get siteResult of pattern pass/fail.
'''step3: [Optional] Store Fail cycle of the pattern from Capture Memory(CMEM) for the current BinCut search step.
'''step4: Check alarmFail for pattern.
'''step5: [Optional] Save result about pattern Pass/Fail for COFInstance.
'''step6: Update pattern pass/fail to the patPass flag.
'''//==================================================================================================================================================================================//'''

    '''//step1: Sync up DCVS output and print BinCut payload voltage for projects with Rail Switch for TD instance.
    Call prepare_DCVS_Output_for_RailSwitch(inst_info, str_pattern, idxBlock_Selsrm_Pattern)

    '''//step2: Run the pattern decomposed from FuncPat, and get siteResult..
    If Flag_VDD_Binning_Offline = False Then '''If the test mode is Online.
        '''//Only CP1 uses CMEM.
        '''20220125,HarvPinGrp use full Cmem size
        If inst_info.enable_HarvPinGrp = True Then
            ''220610: Added before cmem initial for UFlexPlus
            If glb_TesterType = "UltraFLEXplus" Then
                TheHdw.Digital.Patgen.ScanBurstEnabled = True
                TheExec.Datalog.Setup.ScanSetup.EnableScanLogging = True
            End If
            ' Capture all failures (full 256K single or 512K dual).
            TheHdw.Digital.CMEM.SetCaptureConfig -1, CmemCaptFail
            Harvest_CMEM_InitSetup
        ElseIf inst_info.enable_CMEM_Collection = True Then
            ''220610: Added before cmem initial for UFlexPlus
            If glb_TesterType = "UltraFLEXplus" Then
                TheHdw.Digital.Patgen.ScanBurstEnabled = True
                TheExec.Datalog.Setup.ScanSetup.EnableScanLogging = True
            End If
            TheHdw.Digital.CMEM.SetCaptureConfig CaptureSize, CmemCaptFail, tlCMEMCaptureSourcePassFailData
            Harvest_CMEM_InitSetup CaptureSize  ''20230314: Bincut FFC only use 128 cycle
        End If
        
        '''20220407, Disable compare harvPinGroup when doing regular BV search
        If inst_info.patCnt_HarvPinGrp <> 0 And inst_info.enable_HarvPinGrp = False And Flag_HarvPinFlag_Mapping_Table_Parsed = True Then
            Disable_PinGrp_cnt = 0
            strAry_PathSplit = Split(LCase(str_pattern), "\")       '' Pattern name
            strAry_PatNameSplit = Split(LCase(strAry_PathSplit(UBound(strAry_PathSplit))), "_")     '' "in" or "pl"

            'Only payload pat need to check harv pin
            If strAry_PatNameSplit(3) Like "*pl*" Then
                For chkloop = 0 To UBound(inst_info.HarvPinGrpCheckData)        '' Matched pattern keyword, Ubound should be 1
                    If LCase(strAry_PathSplit(UBound(strAry_PathSplit))) Like LCase(inst_info.HarvPinGrpCheckData(chkloop).keyword) Then    '' ex: "*CCC0*" matched
                        'Harvest pattern found
                        For HarvCnt = 0 To UBound(HarvPinFlagMapping(inst_info.HarvPinGrpCheckData(chkloop).MappingIndex).HarvPinGrpConditionArr)       '' Harvest flags
                            For Each site In TheExec.sites
                                If TheExec.sites.item(site).FlagState(HarvPinFlagMapping(inst_info.HarvPinGrpCheckData(chkloop).MappingIndex).HarvPinGrpCondFailFlagArr(HarvCnt)) = logicTrue Then
                                    ' When Harvest Flag = True, disable compare this pin group
                                    
                                    ReDim Preserve PinGroup_disable_compare(Disable_PinGrp_cnt)
                                    '' save pingroups all site together, then enable them after pattern.test
                                    PinGroup_disable_compare(UBound(PinGroup_disable_compare)) = HarvPinFlagMapping(inst_info.HarvPinGrpCheckData(chkloop).MappingIndex).HarvPinGrpConditionArr(HarvCnt)
                                    
                                    TheHdw.Digital.pins(PinGroup_disable_compare(UBound(PinGroup_disable_compare))).DisableCompare = True
                                    Disable_PinGrp_cnt = Disable_PinGrp_cnt + 1     '' count pingroups from all site together
                                    ''20221130: Modified PinMask Format, align with SA harvest
                                    TheExec.Datalog.WriteComment "--(Enable Pin mask Feature)--" & "site = " & site & ", PinGroup = " & PinGroup_disable_compare(UBound(PinGroup_disable_compare))
                                    PinGroup_disable_compare(UBound(PinGroup_disable_compare)) = CStr(site) & "," & PinGroup_disable_compare(UBound(PinGroup_disable_compare))
                                    ''save site info, PinGroup_disable_compare(idx) -> "1,pg_PCPU_Core_0"
                                End If
                            Next site
                        Next HarvCnt
                    End If
                Next chkloop
            End If
        End If

        
        '''//Run the pattern decomposed from FuncPat.
        Call TheHdw.patterns(str_pattern).test(pfAlways, 0, inst_info.result_mode)
        
        ''20221223: For print voltage output(Vmain/Valt) in debug print mode
        If TheExec.flow.enableWord("DebugPrintFlag") = True And Flag_isPrinted_BincutVoltageOutput = False Then
            Call Print_BincutVoltageOutput_AfterPattern
        End If
        
        '''//Get siteResult of pattern pass/fail.
        '''//Warning!!! currently "TheHdw.Digital.Patgen.PatternBurstPassedPerSite" doesn't support "result_mode=tlResultModeModule" with PatternBurst=Yes and DecomposePatt=No.
        'T-Col TTR approve by Si -- 230413
        ''20230926: Modified to get siteResult right after pattern.test
        siteResult = TheHdw.Digital.Patgen.PatternBurstPassedPerSite
        
        '''Turn off Harvest PinGroup (executed for all site at the same time)
        If Disable_PinGrp_cnt > 0 And Flag_HarvPinFlag_Mapping_Table_Parsed = True Then
            For PinGrp_idx = 0 To UBound(PinGroup_disable_compare)
                Dim str_site As String
                Dim split_content() As String
                split_content = Split(PinGroup_disable_compare(PinGrp_idx), ",", 2)
                str_site = split_content(0)
                PinGroup_disable_compare(PinGrp_idx) = split_content(1)
                TheHdw.Digital.pins(PinGroup_disable_compare(PinGrp_idx)).DisableCompare = False
                ''20221130: Modified PinMask Format, align with SA harvest
                TheExec.Datalog.WriteComment "--(Disable Pin mask Feature)--" & "site = " & str_site & ", PinGroup = " & PinGroup_disable_compare(PinGrp_idx)
            Next PinGrp_idx
        End If
        
        '''20220125,HarvPinGrp check result
        If inst_info.enable_HarvPinGrp = True Then
            strAry_PathSplit = Split(LCase(str_pattern), "\")
            strAry_PatNameSplit = Split(LCase(strAry_PathSplit(UBound(strAry_PathSplit))), "_")

            'Only payload pat need to check harv pin
            If strAry_PatNameSplit(3) Like "*pl*" Then
                ''' no matter the pattern is harv or not, we should always get the pass/fail result.
                'siteResult = TheHdw.Digital.Patgen.PatternBurstPassedPerSite
                
                ''20230926: Added to save non-harvest pattern result in harvest instance
                Dim isHarvPat As Boolean
                isHarvPat = False
                
                For chkloop = 0 To UBound(inst_info.HarvPinGrpCheckData)
                    If LCase(strAry_PathSplit(UBound(strAry_PathSplit))) Like LCase(inst_info.HarvPinGrpCheckData(chkloop).keyword) Then
                        ''20230926: Added to save non-harvest pattern result in harvest instance
                        isHarvPat = True
                        
                        'init temp buffer
                        ReDim HarvFailResult(UBound(HarvPinFlagMapping(inst_info.HarvPinGrpCheckData(chkloop).MappingIndex).HarvPinGrpConditionArr))
                        ReDim HarvFailPinCnt(UBound(HarvPinFlagMapping(inst_info.HarvPinGrpCheckData(chkloop).MappingIndex).HarvPinGrpConditionArr))
                        If HarvPinFlagMapping(chkloop).Enable_otherPinGrp = True Then
                            ReDim CommnFailResult(UBound(HarvPinFlagMapping(inst_info.HarvPinGrpCheckData(chkloop).MappingIndex).HarvPinGrpComFailArr))
                            ReDim CommnFailPinCnt(UBound(HarvPinFlagMapping(inst_info.HarvPinGrpCheckData(chkloop).MappingIndex).HarvPinGrpComFailArr))
                        End If
                        
                        'judge Harv
                        For HarvCnt = 0 To UBound(HarvPinFlagMapping(inst_info.HarvPinGrpCheckData(chkloop).MappingIndex).HarvPinGrpConditionArr)
                            HarvFailResult(HarvCnt) = False
                            HarvFailPinCnt(HarvCnt) = 0
                            If glb_TesterType = "UltraFLEXplus" Then
                                Call update_HarvPinGrp_Result_UFP(HarvPinFlagMapping(inst_info.HarvPinGrpCheckData(chkloop).MappingIndex).HarvPinGrpConditionArr(HarvCnt), HarvFailResult(HarvCnt), HarvFailPinCnt(HarvCnt))
                            Else
                                Call update_HarvPinGrp_Result_UF(HarvPinFlagMapping(inst_info.HarvPinGrpCheckData(chkloop).MappingIndex).HarvPinGrpConditionArr(HarvCnt), HarvFailResult(HarvCnt), HarvFailPinCnt(HarvCnt))
                            End If
                        Next HarvCnt
        
                        'judge Commn
                        If HarvPinFlagMapping(chkloop).Enable_otherPinGrp = True Then
                            For ComCnt = 0 To UBound(HarvPinFlagMapping(inst_info.HarvPinGrpCheckData(chkloop).MappingIndex).HarvPinGrpComFailArr)
                                CommnFailResult(ComCnt) = False
                                CommnFailPinCnt(ComCnt) = 0
                                If glb_TesterType = "UltraFLEXplus" Then
                                    Call update_HarvPinGrp_Result_UFP(HarvPinFlagMapping(inst_info.HarvPinGrpCheckData(chkloop).MappingIndex).HarvPinGrpComFailArr(ComCnt), CommnFailResult(ComCnt), CommnFailPinCnt(ComCnt))
                                Else
                                    Call update_HarvPinGrp_Result_UF(HarvPinFlagMapping(inst_info.HarvPinGrpCheckData(chkloop).MappingIndex).HarvPinGrpComFailArr(ComCnt), CommnFailResult(ComCnt), CommnFailPinCnt(ComCnt))
                                End If
                            Next ComCnt
                        End If
                        
                        For Each site In TheExec.sites
                            'only recode bin1 final result
                            If CurrentPassBinCutNum(site) = 1 And inst_info.Step_Current <= DYNAMIC_VBIN_IDS_ZONE(inst_info.p_mode).step_Mapping(1, AllBinCut(inst_info.p_mode).Last_Bin1_EQN) Then
                                TotalHarvFailPinCnt = 0
                                TotalCommnFailPinCnt = 0
                                HarvFailResultOr = False
                                CommnFailResultOr = False
                                
                                For HarvCnt = 0 To UBound(HarvPinFlagMapping(inst_info.HarvPinGrpCheckData(chkloop).MappingIndex).HarvPinGrpConditionArr)
                                    If inst_info.HarvPinGrpCheckData(chkloop).HarvFailResult(HarvCnt)(site) = False Then
                                        inst_info.HarvPinGrpCheckData(chkloop).HarvFailResult(HarvCnt)(site) = HarvFailResult(HarvCnt)(site)
                                    End If
                                    If HarvFailResult(HarvCnt)(site) = True Then
                                        HarvFailResultOr = True
                                        TotalHarvFailPinCnt = TotalHarvFailPinCnt + HarvFailPinCnt(HarvCnt)(site)
                                    End If
                                Next HarvCnt
                                
                                '' If other pin group header exists in the table
                                If HarvPinFlagMapping(chkloop).Enable_otherPinGrp = True Then
                                    For ComCnt = 0 To UBound(HarvPinFlagMapping(inst_info.HarvPinGrpCheckData(chkloop).MappingIndex).HarvPinGrpComFailArr)
                                        If inst_info.HarvPinGrpCommnFailResult(site) = False Then
                                            inst_info.HarvPinGrpCommnFailResult(site) = CommnFailResult(ComCnt)(site)
                                        End If
                                        If CommnFailResult(ComCnt)(site) = True Then
                                            CommnFailResultOr = True
                                            TotalCommnFailPinCnt = TotalCommnFailPinCnt + CommnFailPinCnt(ComCnt)(site)
                                        End If
                                    Next ComCnt
                                End If
                                
                                
                                ' if pat fail but no fail result, set com fail to true
                                If siteResult(site) = False And HarvFailResultOr = False And CommnFailResultOr = False Then
                                    'CommnFailResult(site) = True
                                    inst_info.HarvPinGrpCommnFailResult(site) = True
                                End If
                                
                                ' check all pattern fail pin cnt is same as sum of harv and com, or set com fail to true
                                TempDigitalPin = TheHdw.Digital.FailedPins(site)
                                If UBound(TempDigitalPin) + 1 > (TotalHarvFailPinCnt + TotalCommnFailPinCnt) Then
                                    If inst_info.HarvPinGrpCommnFailResult(site) = False Then
                                        inst_info.HarvPinGrpCommnFailResult(site) = True
                                    End If
                                End If
                                TheExec.Datalog.WriteComment "HRAM Fail Pin:" & CStr(UBound(TempDigitalPin) + 1) & ",Harv Fail Pin:" & CStr(TotalHarvFailPinCnt) & ",Com Fail Pin:" & TotalCommnFailPinCnt
                            End If
                            '''20240313 judge if BIN1 E1 harvest fail, will force VBINRESULT to BIN1 E1
                            If inst_info.HarvPinGrpCommnFailResult(site) = False And HarvFailResultOr = True And inst_info.Step_Current = DYNAMIC_VBIN_IDS_ZONE(inst_info.p_mode).step_Mapping(1, AllBinCut(inst_info.p_mode).Last_Bin1_EQN) And inst_info.Harvest_With_BIN1E1 = True Then
                                inst_info.All_Site_Mask = inst_info.All_Site_Mask - 2 ^ site
                                inst_info.grade_found(site) = False
                                TheExec.Datalog.WriteComment "Site: " & site & " harvest on EQN = " & AllBinCut(inst_info.p_mode).Last_Bin1_EQN & ", so force result to EQN = " & AllBinCut(inst_info.p_mode).Last_Bin1_EQN & "!!!"
                            End If
                        Next
                    End If
                Next chkloop
                
                ''20230926: Added to save non-harvest pattern result in harvest instance
                If isHarvPat = False Then
                    For Each site In TheExec.sites
                        If siteResult(site) = False And CurrentPassBinCutNum(site) = 1 Then
                            If inst_info.HarvPinGrp_NonHarvFail(site) = False Then
                                inst_info.HarvPinGrp_NonHarvFail(site) = True
                            End If
                        End If
                    Next
                End If
            Else
                ''20230926: Modified to get siteResult right after pattern.test
                'siteResult = TheHdw.Digital.Patgen.PatternBurstPassedPerSite
                'if bin1 initial pat fail turn com fail result to true
                For Each site In TheExec.sites
                    If siteResult(site) = False And CurrentPassBinCutNum(site) = 1 Then
                        '''20230926: Added to save non-harvest pattern result in harvest instance
                        If inst_info.HarvPinGrp_NonHarvFail(site) = False Then
                            inst_info.HarvPinGrp_NonHarvFail(site) = True
                        End If
'                        If inst_info.HarvPinGrpCommnFailResult(site) = False Then
'                            inst_info.HarvPinGrpCommnFailResult(site) = True
'                        End If
                    End If
                Next
            End If
        Else
            '''//Get siteResult of pattern pass/fail.
            '''//Warning!!! currently "TheHdw.Digital.Patgen.PatternBurstPassedPerSite" doesn't support "result_mode=tlResultModeModule" with PatternBurst=Yes and DecomposePatt=No.
            ''20230926: Modified to get siteResult right after pattern.test
            'siteResult = TheHdw.Digital.Patgen.PatternBurstPassedPerSite
        End If

        '''//step3: [Optional] Store Fail cycle of the pattern from Capture Memory(CMEM) for the current BinCut search step.
        '''20210305: Modified to add the argument "siteResult" to the vbt function "StoreCapFailcycle".
        '''20221125: 20221125: Use pattern pins for bincut first fail cycle, "TheHdw.Digital.Patterns.Files().DigitalPins.List" for UFP only.
        If siteResult.Any(False) And inst_info.enable_CMEM_Collection = True Then
            If glb_TesterType = "UltraFLEXplus" Then        ''UFP use pattern pin
                ''Get pattern pins, save in inst_info
                'Call GetPatternPinStr(inst_info.ary_FuncPat_decomposed_FFC, inst_info.str_patternPins(indexPatt))       '''20230201: Added for ffc on UFP, save pattern name
                Call GetPatternPinStr(str_pattern, inst_info.str_patternPins(indexPatt))
                ''Use pattern pins for CMEM capture
                Call StoreCapFailcycle_UFP(siteResult, inst_info.str_patternPins(indexPatt), indexPatt, CaptureSize, inst_info.Step_CMEM_Data)
            Else    ''UltraFlex uses failpins = "All_Digital" for CMEM capture
                Call StoreCapFailcycle_UF(siteResult, failpins, indexPatt, CaptureSize, inst_info.Step_CMEM_Data)
            End If
        End If
        
        '20220512, clean cmem after each pattern run
        If inst_info.enable_HarvPinGrp = True Or inst_info.enable_CMEM_Collection = True Then
            TheHdw.Digital.CMEM.SetCaptureConfig 0, CmemCaptNone '''CmemCaptNone: Capture no cycles.
            TheHdw.Digital.CMEM.CentralFields = tlCMEMNone  '20220512, add for turn off cmem capture    20230201
        End If
        
        ''' store pattern pass/fail for auto shmoo
        If autoShmoo.IsEnableAutoShmoo Then
            autoShmoo.ShmooItem.SetPatternPassFail UCase(str_pattern), siteResult
            autoShmoo.ShmooItem.powerPin = "VDD_SOC,VDD_LOW" 'AllBinCut(inst_info.p_mode).powerPin
        End If
    Else '''Offline
        Call run_patt_offline_simulation(str_pattern, inst_info.result_mode, siteResult)
    End If
    
    '''//step4: Check alarmFail for pattern.
    Call check_alarmFail_for_pattern(siteResult)
    
    '''//step5: [Optional] Save result about pattern Pass/Fail for COFInstance.
    '''Use "update_patt_result_for_COFInstance" to record per pattern pass/fail and save EQN-based BinCut payload voltage of per site for "COFInstance".
    If inst_info.enable_COFInstance = True Then
        Call update_patt_result_for_COFInstance(inst_info, indexPatt, str_pattern, siteResult)
    End If

    '''//step6: Update pattern pass/fail to the patPass flag.
    '''20210129: Modified to revise the vbt code for DevChar.
    If inst_info.is_DevChar_Running = False Then
        Call update_Pattern_result_to_PattPass(siteResult, funcPatPass)
    End If
Exit Function
errHandler:
    Call Print_Error_Message(Error_Info, "LIB_VDD_BINNING", "run_patt_from_FuncPat_for_BinCut") 'Add ErrHandler 2023/08/01
    If AbortTest Then Exit Function Else Resume Next
End Function

'20200207: Created to merge set_core_power_main and set_core_power_alt.
'20191210: Modified to use selsramPin domainGroup.
Public Function select_DCVS_output_for_powerDomain(selected_DCVS_output As Integer, Optional domainGroup As String)
On Error GoTo errHandler 'Add ErrHandler 2023/08/01
    Dim split_content() As String
    Dim i As Long
    Dim PinGroup As String
'''//==================================================================================================================================================================================//'''
'''//Note:
'''//tlDCVSVoltageOutput: This property gets or sets the output DAC used to force voltage (main or alternate).
'''It can detect DCVS output Vmain/Valt of the binning CorePower powerDomains
'''The returned information is one of the following values:
'''     tlDCVSVoltageMain   : 1 ==> Main output voltage DAC.
'''     tlDCVSVoltageAlt    : 2 ==> Alternate output voltage DAC.
'''//==================================================================================================================================================================================//'''
    '''//Check if the input argument "selected_DCVS_output" is correct.
    If selected_DCVS_output = tlDCVSVoltageMain Or selected_DCVS_output = tlDCVSVoltageAlt Then
        '''FullBinCutPowerinFlowSheet contains BinCut corePower and otherRail powerDomains after parsing the sheet "Non_Binning_Rail"(initVddBinCondition).
        '''20191210: Modified to use selsramPin domainGroup.
        If domainGroup = "" Then
            If selsramPin <> "" Then
                domainGroup = selsramPin
            Else
                domainGroup = FullBinCutPowerinFlowSheet
            End If
        End If
    
        split_content = Split(domainGroup, ",")
       
        '''init
        PinGroup = vbNullString
    
        If UBound(split_content) < 0 Then
            Call Print_Error_Message(Error_Info, "LIB_VDD_BINNING", "select_DCVS_output_for_powerDomain", "pin_group is incorrect for set_core_power_main. Error!!!")
            'TheExec.ErrorLogMessage "pin_group is incorrect for set_core_power_main. Error!!!"    ''20230807: Removed by sanity check
        Else
            For i = 0 To UBound(split_content)
                If PinGroup = "" Then
                    PinGroup = VddbinDomain2Pin(split_content(i))
                Else
                    PinGroup = PinGroup & "," & VddbinDomain2Pin(split_content(i))
                End If
            Next i
        End If
    
        '''//Switch all BinCut powerDomains to the selected DCVS output.
        '''ToDo: Check if it needs to separate DCVS into HexVS, UVS256 and VSM groups.
        TheHdw.DCVS.pins(PinGroup).Voltage.Output = selected_DCVS_output
    Else
        Call Print_Error_Message(Error_Info, "LIB_VDD_BINNING", "select_DCVS_output_for_powerDomain", "The input argument of select_DCVS_output_for_powerDomain should be tlDCVSVoltageMain or tlDCVSVoltageAlt. Please check the argument selected_DCVS_output. Error!!!")
        'TheExec.ErrorLogMessage "The input argument of select_DCVS_output_for_powerDomain should be tlDCVSVoltageMain or tlDCVSVoltageAlt. Please check the argument selected_DCVS_output. Error!!!"    ''20230807: Removed by sanity check
    End If
Exit Function
errHandler:
    Call Print_Error_Message(Error_Info, "LIB_VDD_BINNING", "select_DCVS_output_for_powerDomain") 'Add ErrHandler 2023/08/01
    If AbortTest Then Exit Function Else Resume Next
End Function

'20210113: Modified to check if cnt_DecomposedPinList>0.
'20201111: Modified to move "SyncUp_DCVS_Output" from modudle LIB_Vdd_Binning_customer to LIB_VDD_BINNING.
'20200106: Modified to remove the ErrorLogMessage.
'20200103: Modified to check the argument "powerPin" is N/C or not.
'20191219: Modified to use dictionaries of Domain2Pin and Pin2Domain.
'20191127: Modified for the revised InitVddBinTable.
'20181224: Modified to add "Flag_SyncUp_DCVS_Output_enable" in LIB_Vdd_Binning_GlobalVariable to control SyncUp on/off.
'20181120: Modified for GradeSearch_VT, added CurrentDcvsOutput to check Vmain/Valt for powerpin.
'20180921: SyncUp is added for switching OtherRail to Valt when detecting CorePower in Valt.
Public Function SyncUp_DCVS_Output(p_mode As Integer, currentDcvsOutput As Integer, powerGroup As String)
On Error GoTo errHandler 'Add ErrHandler 2023/08/01
    Dim PinGroup As String
    Dim split_powerGroup() As String
    Dim i As Integer
    Dim j As Integer
    Dim powerDomain As String
    Dim strAry_pinSyncup() As String
    Dim cnt_DecomposedPinList As Long
    Dim domainTemp As String
'''//==================================================================================================================================================================================//'''
'''//Note:
'''This is to detect voltage source (DCVS Vmain or Valt) and sync up logic powers and sram powers on the same DCVS output (for TD tests).
'''//==================================================================================================================================================================================//'''
    '''init
    PinGroup = vbNullString
    
    If powerGroup <> "" Then
        split_powerGroup = Split(powerGroup, ",")
        
        '''//Check if the powerGroup exists.
        For i = 0 To UBound(split_powerGroup)
            powerDomain = split_powerGroup(i)
            
            '''//Check if powerDomain belongs to BinCut CorePower or OtherRail (listed in globalVariable "FullBinCutPowerinFlowSheet").
            If UCase("*," & FullBinCutPowerinFlowSheet & ",*") Like UCase("*," & powerDomain & ",*") Then
                '''VddbinDomain2Pin
                If PinGroup <> "" Then
                    PinGroup = PinGroup & "," & VddbinDomain2Pin(powerDomain)
                Else
                    PinGroup = VddbinDomain2Pin(powerDomain)
                End If
            Else
                '''//Decompose powerDomain and check each powerPin to check the argument "powerGroup" is N/C or not.
                Call TheExec.DataManager.DecomposePinList(split_powerGroup(i), strAry_pinSyncup, cnt_DecomposedPinList)
                
                If cnt_DecomposedPinList > 0 Then
                    For j = 0 To cnt_DecomposedPinList - 1
                        If TheExec.DataManager.NumberChannelTypesForPin(strAry_pinSyncup(j)) > 0 Then
                            '''//If the powerGroup is connected to DCVS, re-assembly the pinGroup
                            If PinGroup <> "" Then
                                PinGroup = PinGroup & "," & strAry_pinSyncup(j)
                            Else
                                PinGroup = strAry_pinSyncup(j)
                            End If
                        Else
                            Call Print_Error_Message(Error_Info, "LIB_VDD_BINNING", "SyncUp_DCVS_Output", strAry_pinSyncup(j) & " of " & powerDomain & " in " & powerGroup & " doesn't exist in PinMap or ChannelMap. Error!!!")
                            'TheExec.Datalog.WriteComment strAry_pinSyncup(j) & " of " & powerDomain & " in " & powerGroup & " doesn't exist in PinMap or ChannelMap. Error!!!"
                            'TheExec.ErrorLogMessage strAry_pinSyncup(j) & " of " & powerDomain & " in " & powerGroup & " doesn't exist in PinMap or ChannelMap. Error!!!"
                        End If
                    Next j
                Else
                    Call Print_Error_Message(Error_Info, "LIB_VDD_BINNING", "SyncUp_DCVS_Output", "Domain: " & powerDomain & "," & " isn't defined in PinMap for SyncUp_DCVS_Output. Error!!!")
                    'TheExec.ErrorLogMessage "Domain: " & powerDomain & "," & " isn't defined in PinMap for SyncUp_DCVS_Output. Error!!!"    ''20230807: Removed by sanity check
                End If
            End If
        Next i
    Else
        Call Print_Error_Message(Error_Info, "LIB_VDD_BINNING", "SyncUp_DCVS_Output", "The argument powerGroup for SyncUp_DCVS_Output is empty. Error!!!")
        'TheExec.ErrorLogMessage "The argument powerGroup for SyncUp_DCVS_Output is empty. Error!!!"
    End If
    
    '''***********************************************************************************************************'''
    '''//tlDCVSVoltageOutput: This property gets or sets the output DAC used to force voltage (main or alternate).//
    '''//It can detect DCVS output Vmain/Valt of the binning CorePower powerDomains
    ''' The returned information is one of the following values:
    ''' tlDCVSVoltageMain: 1 ==> Main output voltage DAC.
    ''' tlDCVSVoltageAlt: 2 ==> Alternate output voltage DAC.
    '''***********************************************************************************************************'''
    If PinGroup <> "" Then
        If currentDcvsOutput = tlDCVSVoltageAlt Then
            TheHdw.DCVS.pins(PinGroup).Voltage.Output = tlDCVSVoltageAlt
            TheExec.Datalog.WriteComment "Bincut voltage switch to Valt!!"
        Else
            TheHdw.DCVS.pins(PinGroup).Voltage.Output = tlDCVSVoltageMain
            TheExec.Datalog.WriteComment "Bincut voltage switch to Vmain!!"
        End If
    End If
Exit Function
errHandler:
    Call Print_Error_Message(Error_Info, "LIB_VDD_BINNING", "SyncUp_DCVS_Output") 'Add ErrHandler 2023/08/01
    If AbortTest Then Exit Function Else Resume Next
End Function

'20210712: Modified to check Flag_Enable_Rail_Switch because C651 put PFF pattern(with vbump) to Prepatt patset in Mbist test instances.
'20201230: Modified to check powerDomain and powerPin by dictionaries of Domain2Pin and Pin2Domain.
'20191219: Modified to use dictionaries of Domain2Pin and Pin2Domain.
'20190625: Modified to replace the hard-code with the global variable "pinGroup_BinCut".
'20190617: Modified to use siteDouble "CorePowerStored" to save/restore voltages for BinCut powerPins.
Public Function save_core_power_vddbinning(CorePowerStored() As SiteDouble)
On Error GoTo errHandler 'Add ErrHandler 2023/08/01
    Dim i As Long
    Dim powerPin As String
    
    Dim TempString_Pinlist As String

    Dim TempPinListData As New PinListData
    Dim pinArray() As String
    ReDim pinArray(UBound(pinGroup_BinCut)) As String
    '''//pinGroup_BinCut is created after initVddBinCondition (parsing "Non_Binning_Rail")
    '''It contains the pin names and sequence of corePower and otherRail.
    
    
    ' If glb_TesterType = "UltraFLEXplus" Then
    '     For i = 0 To UBound(pinGroup_BinCut)
        
    '         '''init
    '         powerPin = vbNullString
            
    '         '''//As per discussion with SWLINZA, ZHHUANG, and PCLIN. We decided to measure the voltage of 1st powerPin of each power domain.
    '         '''//Check powerDomain and powerPin by dictionaries of Domain2Pin and Pin2Domain.
    '         If domain2pinDict.Exists(UCase(pinGroup_BinCut(i))) = True Then
    '             powerPin = Get1stPinFromPingroup(VddbinDomain2Pin(pinGroup_BinCut(i)))
    '         ElseIf pin2domainDict.Exists(UCase(pinGroup_BinCut(i))) = True Then
    '             powerPin = UCase(pinGroup_BinCut(i))
    '         Else
    '             powerPin = vbNullString
    '             theexec.Datalog.WriteComment pinGroup_BinCut(i) & ", it is not BinCut powerDomain or pinGroup for save_core_power_vddbinning. Error!!!"
    '             theexec.ErrorLogMessage pinGroup_BinCut(i) & ", it is not BinCut powerDomain or pinGroup for save_core_power_vddbinning. Error!!!"
    '         End If
    '         PinArray(i) = powerPin
    '         If powerPin <> "" Then
    '             If TempString_Pinlist = "" Then
    '                 TempString_Pinlist = powerPin
    '             Else
    '                 TempString_Pinlist = TempString_Pinlist & "," & powerPin
    '             End If
    '             TempPinListData.AddPin powerPin
    '         End If

    '     Next i
    '     If Flag_Enable_Rail_Switch = True Then
    '         TheHdw.DCVS.Pins(TempString_Pinlist).Voltage.Output = tlDCVSVoltageAlt
    '         TheHdw.DCVS.Pins(TempString_Pinlist).Meter.mode = tlDCVSMeterVoltage
    '         TempPinListData = TheHdw.DCVS.Pins(TempString_Pinlist).Meter.Read(tlStrobe, 10, 10000, tlDCVSMeterReadingFormatAverage)
    '     Else
    '         TheHdw.DCVS.Pins(TempString_Pinlist).Voltage.Output = tlDCVSVoltageMain
    '         TheHdw.DCVS.Pins(TempString_Pinlist).Meter.mode = tlDCVSMeterVoltage
    '         TempPinListData = TheHdw.DCVS.Pins(TempString_Pinlist).Meter.Read(tlStrobe, 10, 10000, tlDCVSMeterReadingFormatAverage)
    '     End If
        
    '     For i = 0 To UBound(pinGroup_BinCut)
    '         If PinArray(i) <> "" Then
    '             For Each site In theexec.sites
    '                 CorePowerStored(i) = TempPinListData.Pins(PinArray(i)).value
    '             Next site
    '         End If
    '     Next i
        
    ' Else
    
    ''20230703: Removed meter read to avoid unnecessary voltage switch Vmain/Valt for Mbist init pattern.
    For i = 0 To UBound(pinGroup_BinCut)
        '''init
        powerPin = vbNullString
    
        '''//As per discussion with SWLINZA, ZHHUANG, and PCLIN. We decided to measure the voltage of 1st powerPin of each power domain.
        '''//Check powerDomain and powerPin by dictionaries of Domain2Pin and Pin2Domain.
        If domain2pinDict.Exists(UCase(pinGroup_BinCut(i))) = True Then
            powerPin = Get1stPinFromPingroup(VddbinDomain2Pin(pinGroup_BinCut(i)))
        ElseIf pin2domainDict.Exists(UCase(pinGroup_BinCut(i))) = True Then
            powerPin = UCase(pinGroup_BinCut(i))
        Else
            powerPin = vbNullString
            TheExec.Datalog.WriteComment pinGroup_BinCut(i) & ", it is not BinCut powerDomain or pinGroup for save_core_power_vddbinning. Error!!!"
            TheExec.ErrorLogMessage pinGroup_BinCut(i) & ", it is not BinCut powerDomain or pinGroup for save_core_power_vddbinning. Error!!!"
        End If
        
        '''//If powerPin exists, read DCVS Vmain value.
        pinArray(i) = powerPin
        If powerPin <> "" Then
            If TempString_Pinlist = "" Then
                TempString_Pinlist = powerPin
            Else
                TempString_Pinlist = TempString_Pinlist & "," & powerPin
            End If
        End If
    Next i
                '''20210712: Modified to check Flag_Enable_Rail_Switch because C651 put PFF pattern(with vbump) to Prepatt patset in Mbist test instances.
    For i = 0 To UBound(pinGroup_BinCut)
        If pinArray(i) <> "" Then
                If Flag_Enable_Rail_Switch = True Then
               CorePowerStored(i) = TheHdw.DCVS.pins(pinArray(i)).Voltage.Alt.ValuePerSite
                Else
                CorePowerStored(i) = TheHdw.DCVS.pins(pinArray(i)).Voltage.Main.ValuePerSite ' 20230718ValuePerSite
                End If
        End If
    Next i
    ' End If
Exit Function
errHandler:
    Call Print_Error_Message(Error_Info, "LIB_VDD_BINNING", "save_core_power_vddbinning") 'Add ErrHandler 2023/08/01
    If AbortTest Then Exit Function Else Resume Next
End Function

'20210712: Modified to check Flag_Enable_Rail_Switch because C651 put PFF pattern(with vbump) to Prepatt patset in Mbist test instances.
'20210123: Modified to replace "SiteAwareValue" with "ValuePerSite" for UltraFlex with IGXL10.
'20200731: Modified to reset hexvsPingroup and nonhexvsPingroup for each powerDomain.
'20200210: Modified to check UltraFlex and UltraFlexPlus.
'20191219: Modified to use dictionaries of Domain2Pin and Pin2Domain.
'20190624: Modified to replace split_power_group with the public variable "pinGroup_BinCut".
'20190617: Modified to use siteDouble "CorePowerStored()" to save/restore voltages for BinCut powerPins.
Public Function restore_core_power_vddbinning(CorePowerStored() As SiteDouble)
On Error GoTo errHandler 'Add ErrHandler 2023/08/01
    Dim i As Long, j As Long
    Dim split_content() As String
    Dim hexvsPingroup As String
    Dim nonhexvsPingroup As String
    Dim PinGroup As String
    Dim powerDomain As String
    Dim powerPin As String
    Dim site As Variant 'Carter, 20240304
    '''//pinGroup_BinCut is created after initVddBinCondition (parsing "Non_Binning_Rail")
    '''It contains the pin names and sequence of core_power and other rail.
    For i = 0 To UBound(pinGroup_BinCut)
        '''//init
        hexvsPingroup = vbNullString
        nonhexvsPingroup = vbNullString
        PinGroup = vbNullString
        powerDomain = vbNullString
        powerPin = vbNullString
        
        '''//Get powerPins from powerDomain.
        powerDomain = UCase(pinGroup_BinCut(i))
        PinGroup = VddbinDomain2Pin(powerDomain)
        split_content = Split(PinGroup, ",")
        
        '''//Assembly temporary pinGroup for HexVs
        For j = 0 To UBound(split_content)
            powerPin = UCase(Trim(split_content(j)))
            
            'If LCase(VddbinPinDcvsType(powerPin)) Like "hexvs" Then '''for HexVs
            If (VddbinPinDcvsType(powerPin) = "HEXVS") Or (VddbinPinDcvsType(powerPin) = "VHDVS") Or (VddbinPinDcvsType(powerPin) = "VSM") Or _
                    (VddbinPinDcvsType(powerPin) = "VS-800MA") Or (VddbinPinDcvsType(powerPin) = "VS-5A") Then              '''for HexVs 20221228: Modified for UFP UVS ("VS-5A")
                If hexvsPingroup = "" Then
                    hexvsPingroup = powerPin
                Else
                    hexvsPingroup = hexvsPingroup & "," & powerPin
                End If
            Else '''for non-HexVs
                If nonhexvsPingroup = "" Then
                    nonhexvsPingroup = powerPin
                Else
                    nonhexvsPingroup = nonhexvsPingroup & "," & powerPin
                End If
            End If
        Next j
        
        If hexvsPingroup <> "" Then
            '''20210123: Modified to replace "SiteAwareValue" with "ValuePerSite" for UltraFlex with IGXL10.
            '''20210712: Modified to check Flag_Enable_Rail_Switch because C651 put PFF pattern(with vbump) to Prepatt patset in Mbist test instances.
            If Flag_Enable_Rail_Switch = True Then
                TheHdw.DCVS.pins(hexvsPingroup).Voltage.Alt.ValuePerSite = CorePowerStored(i) '''unit:V
            Else
                TheHdw.DCVS.pins(hexvsPingroup).Voltage.Main.ValuePerSite = CorePowerStored(i) '''unit:V
            End If
        End If
        
        'T-Col TTR approve by Si -- 230413
        If nonhexvsPingroup <> "" Then
            For Each site In TheExec.sites
                '''20210712: Modified to check Flag_Enable_Rail_Switch because C651 put PFF pattern(with vbump) to Prepatt patset in Mbist test instances.
                If Flag_Enable_Rail_Switch = True Then
                    TheHdw.DCVS.pins(nonhexvsPingroup).Voltage.Alt.value = CorePowerStored(i)
                Else
                    TheHdw.DCVS.pins(nonhexvsPingroup).Voltage.Main.value = CorePowerStored(i)
                End If
            Next site
        End If
    Next i
Exit Function
errHandler:
    Call Print_Error_Message(Error_Info, "LIB_VDD_BINNING", "restore_core_power_vddbinning") 'Add ErrHandler 2023/08/01
    If AbortTest Then Exit Function Else Resume Next
End Function

'20210123: Modified to replace "SiteAwareValue" with "ValuePerSite" for UltraFlex with IGXL10.
'20200327: Modified to use the flag "Flag_Skip_ReApplyInitVolageToDCVS" to skip "set_core_power_vddbinning_VT".
'20200210: Modified to use siteAwareValue for HexVS.
'20200210: Modified to reduce code-complexity.
'20200114: Modified to check pinGroup by "VddbinDomain2Pin".
'20191219: Modified to use dictionaries of Domain2Pin and Pin2Domain.
'20191007: Modified to remove the undefined coniditions of "LV" and "HV".
'20190624: Modified to replace split_power_group with the public variable "pinGroup_BinCut".
'20190619: Modified to get voltages for BinCut Init pattern with the dedicated DC category and selectors.
'20190606: Modified to add the argument "DcSpecsCategoryForInitPat as string" for Init patterns with the new test setting DC Specs.
'20190606: Modified to use "Flag_Enable_Rail_Switch" for reading safe voltages "_VRS_GLB" or "GLB" from "Global Specs".
'20190524: Modified to use "Flag_Read_SafeVoltage_from_DCspecs" for reading safe voltages from "DC Specs".
'20190521: Modified for getting VRS from DC category of the new "DC_Specs" sheets.
'20180523: Modified for getting BinCut INIT voltages (safe voltages) from sheet "Global Specs" for projects with Rail Switch.
Public Function set_core_power_vddbinning_VT(DC_Level As String, Optional DcSpecsCategoryForInitPat As String)
On Error GoTo errHandler 'Add ErrHandler 2023/08/01
    Dim site As Variant
    Dim i As Long
    Dim j As Long
    Dim main_power As String
    Dim gb As Double
    Dim powerDomain As String
    Dim powerPin As String
    Dim split_content() As String
    Dim voltage_Temp As New SiteDouble
    Dim PinGroup As String
    Dim PinTemp As String
    Dim hexvsPingroup As String
    Dim nonhexvsPingroup As String
    Dim EnableApplyInitVolageToDCVS As Boolean
    '''For instance context
    Dim DCCategory As String
    Dim DCSelector As String
    Dim ACCategory As String
    Dim ACSelector As String
    Dim TimeSetSheet As String
    Dim EdgeSetSheet As String
    Dim LevelsSheet As String
    Dim Overlay As String
'''//==================================================================================================================================================================================//'''
'''//Note:
'''//"Flag_Read_SafeVoltage_from_DCspecs" for safe voltages. True: from "DC Specs"; False: from "Global Specs".
'''Discussed this with T-autogen Jeff, we decided to follow the rule of assemblying voltage names (powerDomain + "_VOP_VAR") in sheet "DC Specs".
'''"VDD_XXX_VAR" is applied to Vmain for Init patterns.
'''If DcSpecsCategoryForInitPat has the string, it uses the DcSpecsCategoryForInitPat for Init patterns.
'''If DcSpecsCategoryForInitPat is empty, it just follows the default DC Specs category.
'''//==================================================================================================================================================================================//'''
    Select Case DC_Level
        Case "NV": gb = 1
        Case Else:
            '''20191007: Modified to remove the undefined coniditions of "LV" and "HV".
            Call Print_Error_Message(Error_Info, "LIB_VDD_BINNING", "set_core_power_vddbinning_VT", "DC_Level:" & DC_Level & ", it is not defined for set_core_power_vddbinning_VT!!! Error!!!")
            'TheExec.ErrorLogMessage "DC_Level:" & DC_Level & ", it is not defined for set_core_power_vddbinning_VT!!! Error!!!"    ''20230807: Removed by sanity check
    End Select
      
    '''**********************************************************************************************************'''
    '''Note: Safe voltages for init Patt usually use the same DC category.
    If Flag_Skip_ReApplyInitVolageToDCVS = True Then
        Call TheExec.DataManager.GetInstanceContext(DCCategory, DCSelector, ACCategory, ACSelector, TimeSetSheet, EdgeSetSheet, LevelsSheet, Overlay)
        
        If DcSpecsCategoryForInitPat <> "" Then
            If LCase(DCCategory) = LCase(DcSpecsCategoryForInitPat) Then
                EnableApplyInitVolageToDCVS = False
            Else
                EnableApplyInitVolageToDCVS = True
            End If
        Else '''DcSpecsCategoryForInitPat is empty.
            '''//If initial voltages and safe voltage(init voltage) use the same DC category, it can skip "set_core_power_vddbinning_VT" after initial voltages...
            '''Note: For PTE/TTR, it can use the flag "Flag_Skip_ReApplyInitVolageToDCVS" to skip "set_core_power_vddbinning_VT".
            If IsLevelLoadedForApplyLevelsTiming = True Then
                EnableApplyInitVolageToDCVS = False
            Else
                EnableApplyInitVolageToDCVS = True
            End If
        End If
    Else
        EnableApplyInitVolageToDCVS = True
    End If
    '''**********************************************************************************************************'''
    
    If EnableApplyInitVolageToDCVS = True Then
        '''//pinGroup_BinCut is created after initVddBinCondition (parsing "Non_Binning_Rail")
        '''//It contains the pin names and sequence of CorePower and OtherRail.
        For i = 0 To UBound(pinGroup_BinCut)
            powerDomain = UCase(Trim(pinGroup_BinCut(i)))
            
            If (DC_Level = "NV") Then
                '''//init
                hexvsPingroup = vbNullString
                nonhexvsPingroup = vbNullString
                
                '''//Get powerPins from powerDomain
                PinGroup = VddbinDomain2Pin(powerDomain)
                split_content = Split(PinGroup, ",")
                
                '''//Assembly temporary pinGroup for HexVs
                For j = 0 To UBound(split_content)
                    PinTemp = UCase(Trim(split_content(j)))
                    
                    '''20200210: Modified to use siteAwareValue for HexVS.
                    'If LCase(VddbinPinDcvsType(PinTEMP)) Like "hexvs" Then '''for HexVs
                    If (VddbinPinDcvsType(PinTemp) = "HexVS") Or (VddbinPinDcvsType(PinTemp) = "VHDVS") Or (VddbinPinDcvsType(PinTemp) = "VSM") Or _
                            (VddbinPinDcvsType(PinTemp) = "VS-800mA") Or (VddbinPinDcvsType(PinTemp) = "VS-5A") Then              '''for HexVs 20221228: Modified for UFP UVS ("VS-5A")
                        If hexvsPingroup = "" Then
                            hexvsPingroup = PinTemp
                        Else
                            hexvsPingroup = hexvsPingroup & "," & PinTemp
                        End If
                    Else '''for non-HexVs
                        If nonhexvsPingroup = "" Then
                            nonhexvsPingroup = PinTemp
                        Else
                            nonhexvsPingroup = nonhexvsPingroup & "," & PinTemp
                        End If
                    End If
                Next j

                '''//Get 1st powerPin from powerDomain
                powerPin = Get1stPinFromPingroup(VddbinDomain2Pin(pinGroup_BinCut(i)))
                
                If Flag_Read_SafeVoltage_from_DCspecs Then '''Get safe voltages " "VDD_***_VAR" " from "DC Specs".
                    If DcSpecsCategoryForInitPat <> "" Then '''for Mbist
                        voltage_Temp = Floor(TheExec.Specs.DC.item(powerPin & "_VAR").Categories.item(DcSpecsCategoryForInitPat).Selectors.item("typ").ContextValue * gb * 1000) / 1000
                    Else
                        voltage_Temp = Floor(TheExec.Specs.DC.item(powerPin & "_VAR").ContextValue * gb * 1000) / 1000
                    End If
                Else '''Get safe voltages "VDD_***_VRS_GLB" from "Global Specs".
                    If Flag_Enable_Rail_Switch Then
                        voltage_Temp = Floor(TheExec.Specs.Globals(powerPin & "_VRS" & "_GLB").ContextValue * gb * 1000) / 1000
                    Else
                        voltage_Temp = Floor(TheExec.Specs.Globals(powerPin & "_GLB").ContextValue * gb * 1000) / 1000
                    End If
                End If
                    
                '''20200210: Modified to use siteAwareValue for HexVS.
                '''20210123: Modified to replace "SiteAwareValue" with "ValuePerSite" for UltraFlex with IGXL10.
                If hexvsPingroup <> "" Then
                    TheHdw.DCVS.pins(hexvsPingroup).Voltage.Main.ValuePerSite = voltage_Temp
                End If
                
                If nonhexvsPingroup <> "" Then
                    For Each site In TheExec.sites
                        TheHdw.DCVS.pins(nonhexvsPingroup).Voltage.Main.value = voltage_Temp
                    Next site
                End If
            Else
                '''20191007: Modified to remove the undefined coniditions of "LV" and "HV".
                Call Print_Error_Message(Error_Info, "LIB_VDD_BINNING", "set_core_power_vddbinning_VT", "DC_Level:" & DC_Level & ", it is not defined to get DC Levels for set_core_power_vddbinning_VT!!! Error!!!")
                'TheExec.ErrorLogMessage "DC_Level:" & DC_Level & ", it is not defined to get DC Levels for set_core_power_vddbinning_VT!!! Error!!!"    ''20230807: Removed by sanity check
            End If
        Next i
    End If
Exit Function
errHandler:
    Call Print_Error_Message(Error_Info, "LIB_VDD_BINNING", "set_core_power_vddbinning_VT") 'Add ErrHandler 2023/08/01
    If AbortTest Then Exit Function Else Resume Next
End Function

'20210906: Modified to merge the branches of the vbt function Parsing_Instance_Pmode.
'20210809: Modified to revise the vbt code to get main_performance_mode and additional_mode.
'20201230: Modified to revise the vbt code for the new naming rule of OutsideBinCut performance mode, ex: "MS004_TD_SOC_MHV".
'20200211: Modified to replace the function name "FlowTestCondStr2Enum" with "AdditionalModeStr2Enum".
'20200205: Modified to check if the input "Performance_mode" exists in the dictionary.
'20191202: Modified for the revised initVddBinCondition.
'20191127: Modified for the revised InitVddBinTable.
'20190510: Modified to merge "powerDomain = AllBinCut(p_mode).powerPin" into Parse_Performance_Mode
'20190508: Created for parsing the performance mode and getting addi_mode.
Public Function Parsing_Instance_Pmode(performance_mode As String, p_mode As Integer, addi_mode As Integer, Special_Voltage_setup As Boolean)
On Error GoTo errHandler 'Add ErrHandler 2023/08/01
    Dim i As Integer
    Dim j As Integer
    Dim split_content() As String
    Dim strTemp As String
    Dim main_performance_mode As String
    Dim additional_mode As String
    '''init
    p_mode = 0
    addi_mode = 0
    Special_Voltage_setup = False
    main_performance_mode = vbNullString
    additional_mode = vbNullString
    
    If LCase(performance_mode) Like "*lv" Or LCase(performance_mode) Like "*hv" Then ''20220321
        main_performance_mode = performance_mode
    Else
        '''//OutsideBinCut might use the new naming rules for performance modes, ex: "MS004_TD_SOC_MHV".
        split_content = Split(UCase(Trim(performance_mode)), "_")
        
        For i = 0 To UBound(split_content)
            If Trim(split_content(i)) Like UCase("m*##*") Then
                If VddbinPmodeDict.Exists(UCase(Trim(split_content(i)))) Then
                    main_performance_mode = UCase(Trim(split_content(i)))
                    
                    '''//Check if additional_mode exists...
                    If i = UBound(split_content) Then '''without additional_mode
                        Special_Voltage_setup = False
                    Else '''with additional_mode
                        Special_Voltage_setup = True
                        
                        '''//Assembly the string for additional_mode.
                        additional_mode = UCase(Trim(split_content(i + 1)))
                        
                        If (i + 1) < UBound(split_content) Then
                            For j = i + 2 To UBound(split_content)
                                additional_mode = additional_mode & "_" & UCase(Trim(split_content(j)))
                            Next j
                        End If
                    End If
                    
                    '''//Once if it gets 1st keyword with performance mode, exit loop.
                    Exit For
                End If
            End If
        Next i
    End If
    
    '''//Check if it gets the performance mode.
    '''20210906: Modified to merge the branches of the vbt function Parsing_Instance_Pmode.
    If main_performance_mode <> "" Then
        '''//Check p_mode for the main_performance_mode.
        If VddbinPmodeDict.Exists(UCase(main_performance_mode)) Then
            p_mode = VddBinStr2Enum(main_performance_mode)
        Else
            Call Print_Error_Message(Error_Info, "LIB_VDD_BINNING", "Parsing_Instance_Pmode", performance_mode & ", it wasn't the performance mode defined in BinCut voltage table for Parsing_Instance_Pmode. Error!!!")
            'TheExec.Datalog.WriteComment performance_mode & ", it wasn't the performance mode defined in BinCut voltage table for Parsing_Instance_Pmode. Error!!!"
            'TheExec.ErrorLogMessage performance_mode & ", it wasn't the performance mode defined in BinCut voltage table for Parsing_Instance_Pmode. Error!!!"    ''20230807: Removed by sanity check
            Exit Function
        End If
    Else '''If main_performance_mode = vbNullString Then
        Call Print_Error_Message(Error_Info, "LIB_VDD_BINNING", "Parsing_Instance_Pmode", performance_mode & ", it doesn't have the correct format of the performance mode for Parsing_Instance_Pmode. Error!!!")
        'TheExec.Datalog.WriteComment performance_mode & ", it doesn't have the correct format of the performance mode for Parsing_Instance_Pmode. Error!!!"
        'TheExec.ErrorLogMessage performance_mode & ", it doesn't have the correct format of the performance mode for Parsing_Instance_Pmode. Error!!!"    ''20230807: Removed by sanity check
        Exit Function
    End If
    
    '''//Check addi_mode for the additional_mode.
    If Special_Voltage_setup = True Then
        If AdditionalModeDict.Exists(additional_mode) Then
            addi_mode = AdditionalModeStr2Enum(additional_mode)
        Else
            Call Print_Error_Message(Error_Info, "LIB_VDD_BINNING", "Parsing_Instance_Pmode", performance_mode & ", it doesn't have the correct additional mode for Parsing_Instance_Pmode. Error!!!")
            'TheExec.Datalog.WriteComment performance_mode & ", it doesn't have the correct additional mode for Parsing_Instance_Pmode. Error!!!"
            'TheExec.ErrorLogMessage performance_mode & ", it doesn't have the correct additional mode for Parsing_Instance_Pmode. Error!!!"    ''20230807: Removed by sanity check
        End If
    End If
Exit Function
errHandler:
    Call Print_Error_Message(Error_Info, "LIB_VDD_BINNING", "Parsing_Instance_Pmode") 'Add ErrHandler 2023/08/01
    If AbortTest Then Exit Function Else Resume Next
End Function

'20210730: Modified to show the error message to users if the current testJob is for BinCut search without any Efuse category "Product_Identifier", as requested by C651 Toby.
'20210707: Modified to merge the branches of checking the keyword in Efuse category for the vbt function Parsing_IDSname_from_BDF_Table.
'20210707: Modified to add the special case "product_identifier_cp1" of Efuse_BitDef_Table for the vbt function Parsing_IDSname_from_BDF_Table.
'20210707: Modified to update the string array to dict_strPmode2EfuseCategory.
'20210705: Modified to revise the vbt code to parse Efuse category with the prefix "ids_...".
'20210703: Modified to use dict_strPmode2EfuseCategory as the dictionary of p_mode and array of the related Efuse category.
'20210703: Modified to use dict_EfuseCategory2BinCutTestJob as the dictionary of Efuse category and the matched programming state in Efuse.
'20210703: Modified to revise the vbt code of parsing Efuse_BitDef_Table.
'20210702: Modified to check column "Default or Real" in Efuse_BitDef_Table.
'20210701: Modified to revise the vbt code for BinCut search in FT.
'20210701: Modified to update AllBinCut(p_mode).listed_in_Efuse_BDF in the vbt function Parsing_IDSname_from_BDF_Table.
'20210617: Check_IDS and judge_IDS are dedicated to Efuse processed IDS for BinCut search, as requested by C651 Si.
'20210121: Modified to check Harvest 10 cores.
'20201210: Modified to use the flag "is_BinCutJob_for_StepSearch" for "check_bincutJob_for_StepSearch" to check if the test program is binSearch or functional test.
'20200817: Modified to use the variable "str_alg_temp".
'20200731: Modified to get Efuse IDS name for CP1 only.
'20200731: Modified to merge MappingBincutJobName and Mapping_TPJobName_to_BincutJobName into Mapping_TestJobName_to_BincutJobName.
'20200730: Modified to check IDS_VDD_XXX_BINCHECK for Harvest powerPin.
'20200730: Modified to check if VDD_XXX__M*### is the correct BinCut performance mode.
'20200729: Modified to check if IDS name from the column "bank_config eFuse Bit Def" belongs to BinCut powerDomain by BinCut testJob.
'20200703: Modiifed to use "check_Sheet_Range".
'20200410: Modified to check "Base Voltage" in Efuse_BitDef_Table and Vdd_Binning_Def.
'20200114: Modified to check if powerDomain exists in domain2pinDict or pin2domainDict.
'20191015: Modified to use MaxSiteCount array for ids_name.
'20190710: Modified to check if "VDD_***_M*###" is listed in BinCut power_seq.
'20190627: Modified to use the global variable "pinGroup_BinCut" for BinCut powerPins.
'20190521: Modified for mapping IDS names to current BinCut testjob.
'20190514: Created for parsing the sheet "EFUSE_BitDef_Table" to get IDS names, especially for testjobs.
Public Function Parsing_IDSname_from_BDF_Table()
On Error GoTo errHandler 'Add ErrHandler 2023/08/01
    Dim site As Long
    '''
    Dim wb As Workbook
    Dim ws_def As Worksheet
    Dim sheetName As String
    Dim MaxRow As Long, maxcol As Long
    Dim isSheetFound As Boolean
    '''
    Dim Row As Long
    Dim Col As Long
    Dim Row_of_Title As Long
    Dim col_Efuse_category As Long
    Dim col_stage As Long
    Dim col_alg As Long
    Dim col_defaultValue As Long
    Dim col_default_Real As Long
    Dim enableRowParsing As Boolean
    '''
    Dim k As Long
    Dim idx_powerDomain As Integer
    Dim powerDomain As String
    Dim str_category_temp As String
    Dim str_stage_temp As String
    Dim str_alg_temp As String
    Dim str_pmode_temp As String
    Dim str_default_Real As String
    Dim str_ids_name As String
    Dim str_testJob As String '''Mapping_TestJobName_to_BincutJobName
    Dim idx_testJob As Long
    Dim BDF_tastJob As String
    Dim strAry_EfuseCategory() As String
    Dim idx_EfuseCategory As Long
    Dim str_Efuse_write_ProductIdentifier As String
'''//==================================================================================================================================================================================//'''
'''//Caution!!!
'''1. Remember to check the column "Programming Stage" for BinCut testJob mapping.(check the function "MappingBincutJobName")
'''2. Remember to check core number of the Harvest powerPin in TestPlan and "EFUSE_BitDef_Table"!!! => ex: ids_vdd_sram_gpu_7, ids_vdd_sram_gpu_8. "7" and "8" are core numbers of Harvest powerPin.
'''3. Remember to check the format of test temperature. => ex: ids_vdd_soc_105. "105" is the test temperature.
'''4. Check_IDS and judge_IDS are dedicated to Efuse processed IDS for BinCut search, as requested by C651 Si, 20210617.
'''//==================================================================================================================================================================================//'''
    '''*****************************************************************'''
    '''//Check if the sheet exists
    sheetName = "EFUSE_BitDef_Table"
    Set wb = Application.ActiveWorkbook
    Call check_Sheet_Range(sheetName, wb, ws_def, MaxRow, maxcol, isSheetFound)
    '''*****************************************************************'''
    If isSheetFound = True Then
        '''//Init
        dict_strPmode2EfuseCategory.RemoveAll
        dict_EfuseCategory2BinCutTestJob.RemoveAll
        dict_EfuseCategory2BDF.RemoveAll
        
        '''//Clear the content of the globalVariable "IDS_for_BinCut" before parsing the sheet "EFUSE_BitDef_Table"
        For idx_powerDomain = 0 To UBound(pinGroup_BinCut)
            powerDomain = pinGroup_BinCut(idx_powerDomain)
            For site = 0 To MaxSiteCount - 1
                IDS_for_BinCut(VddBinStr2Enum(powerDomain)).ids_name(site) = vbNullString
            Next site
        Next idx_powerDomain
    Else
        Call Print_Error_Message(Error_Info, "LIB_VDD_BINNING", "Parsing_IDSname_from_BDF_Table", "sheet:" & sheetName & ", it doesn't exist in this IGXL workbook. Error!!!")
        'TheExec.ErrorLogMessage "sheet:" & sheetName & ", it doesn't exist in this IGXL workbook. Error!!!"    ''20230807: Removed by sanity check
        Exit Function
    End If '''If isSheetFound = True
    
    '''//Parse the header of the table.
    For Row = 1 To MaxRow
        For Col = 1 To maxcol
            '''//The header should include "bank_config eFuse Bit Def", "Programming Stage", and "Algorithm".
            If LCase(Trim(ws_def.Cells(Row, Col).value)) Like LCase("bank_config eFuse Bit Def") Then
                col_Efuse_category = Col
                Row_of_Title = Row
            End If
            
            If Row_of_Title > 0 Then
                '''20210702: Modified to check column "Default or Real" in Efuse_BitDef_Table.
                If LCase(Trim(ws_def.Cells(Row_of_Title, Col).value)) Like LCase("Programming Stage") Then '''ex: "CP1", "FT1", "WLFT1".
                    col_stage = Col
                ElseIf LCase(Trim(ws_def.Cells(Row_of_Title, Col).value)) Like LCase("Algorithm") Then '''ex: "app", "ids", "vddbin".
                    col_alg = Col
                ElseIf LCase(Trim(ws_def.Cells(Row_of_Title, Col).value)) Like LCase("Default or Real") Then '''ex: "bincut", "Default", "Real".
                    col_default_Real = Col
                ElseIf LCase(Trim(ws_def.Cells(Row_of_Title, Col).value)) Like LCase("Default Value") Then
                    col_defaultValue = Col
                End If
            End If
        Next Col
        
        '''//If row the header is found, parse each row of the table.
        If Row_of_Title > 0 Then
            If col_Efuse_category > 0 And col_stage > 0 And col_alg > 0 Then
                enableRowParsing = True
                Exit For
            Else
                enableRowParsing = False
                Call Print_Error_Message(Error_Info, "LIB_VDD_BINNING", "Parsing_IDSname_from_BDF_Table", "sheet:" & sheetName & ", it doesn't have the correct header for Parsing_IDSname_from_BDF_Table. Error!!!")
                'TheExec.ErrorLogMessage "sheet:" & sheetName & ", it doesn't have the correct header for Parsing_IDSname_from_BDF_Table. Error!!!"    ''20230807: Removed by sanity check
                Exit For
            End If
        End If
    Next Row

    If enableRowParsing = True Then
        For Row = Row_of_Title + 1 To MaxRow
            '''*************************************************************************************************************************************'''
            '''//Get Efuse category, algorithm, Default/Real, and Programming Stage for each row from Efuse_BitDef_Table.
            '''*************************************************************************************************************************************'''
            '''//Efuse category.
            str_category_temp = UCase(Trim(ws_def.Cells(Row, col_Efuse_category).value))
            
            '''//Algorithm.
            str_alg_temp = UCase(Trim(ws_def.Cells(Row, col_alg).value))
            
            '''//Default or Real.
            str_default_Real = UCase(Trim(ws_def.Cells(Row, col_default_Real).value))
            
            '''//Get testJob from the column "Programming Stage".
            '''Mapping_TestJobName_to_BincutJobName
            str_stage_temp = UCase(Trim(ws_def.Cells(Row, col_stage).value))
            
            '''*****************************************************************************************************************************************'''
            '''//Check the Efuse category from the column "bank_config eFuse Bit Def" in Efuse_BitDef_Table.
            '''//If "Algorithm" = "vddbin" and Efuse category with the keyword about BinCut performance mode "VDD_XXX__M*###", the item must be Efuse product voltage.
            '''//If "Algorithm" = "ids" and Efuse category with the prefix "ids_", check if it contains the keyword about BinCut powerDomain.
            '''//If "Algorithm" = "app" and Efuse category "product_identifier", the item must be Efuse product_identifier.
            '''//If "Algorithm" = "app" and Efuse category with the keyword "power_binning", the item must be Efuse power_binning.
            '''//If "Algorithm" = "base" and Efuse category with the keyword "*base*voltage*", the item must be Efuse base voltage.
            '''*****************************************************************************************************************************************'''
            If (LCase(str_category_temp) Like "vdd*" And LCase(str_alg_temp) = "vddbin" And LCase(str_default_Real) = "bincut") _
            Or (((LCase(str_category_temp) Like "product_identifier*" Or LCase(str_category_temp) Like "*power_binning*") And LCase(str_alg_temp) = "app")) Then
                '''init
                str_pmode_temp = vbNullString
                powerDomain = vbNullString
                               
                '''//Check if the Efuse category for Efuse product voltages with keyword “*_shadow*”, ex: "vdd_gpu_mg001_shadow".
                If LCase(str_category_temp) Like LCase("*_shadow") Then
                    str_pmode_temp = UCase(Replace(UCase(str_category_temp), UCase("_shadow"), ""))
                ElseIf LCase(str_category_temp) Like LCase("TE_MISC_POWER_BINNING_CP1") Then '''ex: "TE_MISC_POWER_BINNING_CP1".
                    str_pmode_temp = UCase(str_category_temp)
                ElseIf LCase(str_category_temp) Like LCase("*_cp1") Then '''ex: "product_identifier_cp1".
                    str_pmode_temp = UCase(Replace(UCase(str_category_temp), UCase("_cp1"), ""))
                ElseIf LCase(str_category_temp) Like LCase("*_wlft1") Then '''ex: "product_identifier_wlft1".
                    str_pmode_temp = UCase(Replace(UCase(str_category_temp), UCase("_wlft1"), ""))
                Else
                    str_pmode_temp = UCase(str_category_temp)
                End If
                
                '''//Check if the keyword "VDD_XXX__M*###" in Efuse category is the correct BinCut performance mode.
                If LCase(str_pmode_temp) Like "vdd*" Then
                    If VddbinPmodeDict.Exists(str_pmode_temp) = True Then
                        powerDomain = AllBinCut(VddBinStr2Enum(str_pmode_temp)).powerPin
                        
                        '''//Remove the prefix "vdd_" from str_pmode_temp.
                        If LCase(str_pmode_temp) Like LCase("vdd_*") Then '''ex: "vdd_gpu_mg001".
                            str_pmode_temp = UCase(Replace(UCase(str_pmode_temp), UCase(powerDomain & "_"), ""))
                        End If
                        
                        '''//Check if str_pmode_temp exists in gb_bincut_power_list.
                        If UCase("*," & gb_bincut_power_list(VddBinStr2Enum(powerDomain)) & ",*") Like UCase("*," & str_pmode_temp & ",*") Then
                            str_pmode_temp = VddBinName(VddBinStr2Enum(str_pmode_temp)) '''ex: "VDD_PCPU_MP001"
                        Else
                            '20211022: Added to support SRAM Voltage calculate with equation
                            If dict_IsCorePower(powerDomain) = True And UCase(powerDomain) Like "*VDD*_SRAM*" Then
                                str_pmode_temp = VddBinName(VddBinStr2Enum(str_pmode_temp))
                            Else
                                str_pmode_temp = vbNullString
                                Call Print_Error_Message(Error_Info, "LIB_VDD_BINNING", "", "sheet:" & sheetName & ",row:" & Row & ",Efuse category:" & str_category_temp & ", it doesn't contain any performance mode listed in BinCut flow table. Error!!!")
                                'TheExec.ErrorLogMessage "sheet:" & sheetName & ",row:" & Row & ",Efuse category:" & str_category_temp & ", it doesn't contain any performance mode listed in BinCut flow table. Error!!!"    ''20230807: Removed by sanity check
                            End If
                        End If
                    Else
                        str_pmode_temp = vbNullString
                    End If '''If VddbinPmodeDict.Exists(str_pmode_temp) = True
                End If
                
                '''//Store Efuse category for BinCut p_mode into the dictionary dict_strPmode2EfuseCategory.
                '''20210701: Modified to update AllBinCut(p_mode).listed_in_Efuse_BDF in the vbt function Parsing_IDSname_from_BDF_Table.
                '''20210703: Modified to use dict_strPmode2EfuseCategory as the dictionary of p_mode and array of the related Efuse category.
                '''20210707: Modified to update the string array to dict_strPmode2EfuseCategory.
                If str_pmode_temp <> "" Then
                    If dict_strPmode2EfuseCategory.Exists(str_pmode_temp) = True Then
                        '''init
                        strAry_EfuseCategory = dict_strPmode2EfuseCategory.item(str_pmode_temp)
                        
                        '''//Check if Efuse category exists in dict_strPmode2EfuseCategory for BinCut p_mode.
                        For idx_EfuseCategory = 0 To UBound(strAry_EfuseCategory)
                            '''//Check if the Efuse category has the duplicate item in Efuse_BitDef_Table.
                            If LCase(str_category_temp) = LCase(strAry_EfuseCategory(idx_EfuseCategory)) Then
                                str_category_temp = vbNullString
                                Exit For
                            End If
                        Next idx_EfuseCategory
                        
                        '''//Efuse category has no conflict, then store this into dict_strPmode2EfuseCategory.
                        If str_category_temp <> "" Then
                            idx_EfuseCategory = UBound(strAry_EfuseCategory) + 1
                            ReDim Preserve strAry_EfuseCategory(idx_EfuseCategory) As String
                            '''ToDo: Maybe we can do bubble sorting for the sequence of Efuse category from dict_strPmode2EfuseCategory...
                            strAry_EfuseCategory(idx_EfuseCategory) = str_category_temp
                            dict_strPmode2EfuseCategory.item(str_pmode_temp) = strAry_EfuseCategory
                            '''//Updated the property "AllBinCut(p_mode).listed_in_Efuse_BDF" for BinCut p_mode.
                            If VddbinPmodeDict.Exists(str_pmode_temp) = True Then
                                AllBinCut(VddBinStr2Enum(str_pmode_temp)).listed_in_Efuse_BDF = True
                            End If
                        Else
                            str_category_temp = vbNullString
                            Call Print_Error_Message(Error_Info, "LIB_VDD_BINNING", "Parsing_IDSname_from_BDF_Table", "sheet:" & sheetName & ",row:" & Row & ",Efuse category:" & str_category_temp & ", it has the duplicate Efuse category in sheet:" & sheetName & ". Error!!!")
                            'TheExec.ErrorLogMessage "sheet:" & sheetName & ",row:" & Row & ",Efuse category:" & str_category_temp & ", it has the duplicate Efuse category in sheet:" & sheetName & ". Error!!!"    ''20230807: Removed by sanity check
                        End If
                    Else
                        ReDim strAry_EfuseCategory(0) As String
                        strAry_EfuseCategory(0) = str_category_temp
                        dict_strPmode2EfuseCategory.Add str_pmode_temp, strAry_EfuseCategory
                        '''//Updated the property "AllBinCut(p_mode).listed_in_Efuse_BDF" for BinCut p_mode.
                        If VddbinPmodeDict.Exists(str_pmode_temp) = True Then
                            AllBinCut(VddBinStr2Enum(str_pmode_temp)).listed_in_Efuse_BDF = True
                        End If
                    End If
                Else
                    str_category_temp = vbNullString
                End If '''If str_pmode_temp <> ""
                
            '''//If "Algorithm" = "ids" and Efuse category with the prefix "ids_", check if it contains the keyword about BinCut powerDomain.
            ElseIf (LCase(str_category_temp) Like "ids*" And LCase(str_alg_temp) = "ids") And LCase(str_default_Real) = "real" Then '''ex: "ids_vdd_pcpu", "ids_vdd_ecpu", "ids_vdd_cpu_sram", "ids_vdd_gpu_5", "ids_vdd_sram_gpu_5".
                '''init
                str_ids_name = vbNullString
                powerDomain = vbNullString
            
                '''//Check the column "Programming Stage" for BinCut testJob mapping.
                '''*****************************************************************************************************************************************'''
                '''20210617: Check_IDS and judge_IDS are dedicated to Efuse processed IDS for BinCut search, as requested by C651 Si.
                '''So that all BinCut testJobs use Efuse IDS values that are fused in CP1!!!
                '''ToDo: Please discuss this with project Efuse owner to see if rules about "Programming Stage" in Efuse_BitDef_Table are changed.
                '''*****************************************************************************************************************************************'''
                '''20210707: Modified to parse Efuse category with the keyword "ids_*" while the programming stage is "cp1".
                If LCase(str_stage_temp) = "cp1" Then
                    str_ids_name = UCase(Replace(LCase(str_category_temp), "ids_", ""))
                Else
                    str_ids_name = vbNullString
                End If
                
                '''//Check if str_category_temp about IDS name from Efuse category is not empty.
                If str_ids_name <> "" Then
                    '''//Remove string about core number of the Harvest powerPin.
                    '''*****************************************************************************************************************************************'''
                    '''For Harvest powerPin, IDS_name will be defined in "VBT_LIB_DC_IDS\IDS_eFuse_Write_"
                    '''Note: Remember to check core number of the Harvest powerPin in TestPlan and "EFUSE_BitDef_Table"!!!
                    '''ex: ids_vdd_sram_gpu_7, ids_vdd_sram_gpu_8. "7" and "8" are core numbers of Harvest powerPin.
                    '''*****************************************************************************************************************************************'''
                    '''ToDo: Remember to check string about the core number of Harvest powerPin...
                    If LCase(str_ids_name) Like "*_3" Or LCase(str_ids_name) Like "*_4" Or LCase(str_ids_name) Like "*_5" _
                    Or LCase(str_ids_name) Like "*_7" Or LCase(str_ids_name) Like "*_8" Or LCase(str_ids_name) Like "*_9" Or LCase(str_ids_name) Like "*_10" Then
                        str_ids_name = UCase(mid(str_ids_name, 1, Len(str_ids_name) - 2))
                    ElseIf LCase(str_ids_name) Like "*_bincheck" Then
                        str_ids_name = UCase(Replace(Replace(LCase(str_ids_name), "ids_", ""), "_bincheck", ""))
                    End If
                    
                    '''//Remove string of the test temperature.
                    '''*****************************************************************************************************************************************'''
                    '''Remember to check the format of test temperature, ex: ids_vdd_soc_105. "105" is the test temperature.
                    '''*****************************************************************************************************************************************'''
                    If LCase(str_ids_name) Like "*_25" Or LCase(str_ids_name) Like "*_25_*" Then
                        str_ids_name = UCase(Replace(LCase(str_ids_name), "_25", "")) '''25
                    ElseIf LCase(str_ids_name) Like "*_25c*" Then
                        str_ids_name = UCase(Replace(LCase(str_ids_name), "_25c", "")) '''25C
                    ElseIf LCase(str_ids_name) Like "*_85" Or LCase(str_ids_name) Like "*_85_*" Then
                        str_ids_name = UCase(Replace(LCase(str_ids_name), "_85", "")) '''85
                    ElseIf LCase(str_ids_name) Like "*_85c*" Then
                        str_ids_name = UCase(Replace(LCase(str_ids_name), "_85c", "")) '''85C
                    ElseIf LCase(str_ids_name) Like "*_105" Or LCase(str_ids_name) Like "*_105_*" Then
                        str_ids_name = UCase(Replace(LCase(str_ids_name), "_105", "")) '''105
                    ElseIf LCase(str_ids_name) Like "*_105c*" Then
                        str_ids_name = UCase(Replace(LCase(str_ids_name), "_105c", "")) '''105C
                    ElseIf LCase(str_ids_name) Like "*_hot*" Then
                        str_ids_name = UCase(Replace(LCase(str_ids_name), "_hot", "")) '''hot  20211225 Example ids_vdd_low_hot =>ids_vdd_low
                    End If
                    
                    '''//Check if the string contains the keyword about powerDomain.
                    If domain2pinDict.Exists(UCase(str_ids_name)) = True Then
                        powerDomain = UCase(str_ids_name)
                    ElseIf pin2domainDict.Exists(UCase(str_ids_name)) = True Then
                        powerDomain = UCase(VddbinPin2Domain(str_ids_name))
                    Else
                        powerDomain = vbNullString
                        '''ToDo: Maybe we can add the utility to show warning here...
                    End If
                Else
                    powerDomain = vbNullString
                End If
                
                '''//If str_domain_temp of IDS name is not empty, check if powerDomain belongs to BinCut powerDomains.
                If powerDomain <> "" Then
                    If VddbinPmodeDict.Exists(powerDomain) = True Then
                        idx_powerDomain = VddBinStr2Enum(powerDomain)

                        '''***********************************************************************************************'''
                        '''//Check if IDS_for_BinCut(idx_PowerDomain).IDS_name of BinCut powerDomain is already defined...
                        '''***********************************************************************************************'''
                        For site = 0 To MaxSiteCount - 1
                            If IDS_for_BinCut(idx_powerDomain).ids_name(site) <> "" Then
                                '''//Check IDS_VDD_XXX_BINCHECK for Harvest powerPin.
                                If LCase(IDS_for_BinCut(idx_powerDomain).ids_name(site)) Like "*bincheck*" Then
                                    '''Do nothing, keep IDS_VDD_XXX_BINCHECK as IDS name for Harvest powerPin.
                                Else
                                    IDS_for_BinCut(idx_powerDomain).ids_name(site) = str_category_temp
                                End If
                            Else '''If IDS_for_BinCut(idx_PowerDomain).IDS_name is empty...
                                IDS_for_BinCut(idx_powerDomain).ids_name(site) = str_category_temp
                            End If
                        Next site
                    Else
                        str_category_temp = vbNullString
                        Call Print_Error_Message(Error_Info, "LIB_VDD_BINNING", "Parsing_IDSname_from_BDF_Table", "sheet:" & sheetName & ",row:" & Row & ",Efuse category:" & str_category_temp & ", it doesn't get any correct IDS name of BinCut powerPin while Parsing_IDSname_from_BDF_Table. Error!!!")
                        'TheExec.ErrorLogMessage "sheet:" & sheetName & ",row:" & Row & ",Efuse category:" & str_category_temp & ", it doesn't get any correct IDS name of BinCut powerPin while Parsing_IDSname_from_BDF_Table. Error!!!"    ''20230807: Removed by sanity check
                    End If
                    
                    '''//Since str_category_temp with the keyword IDS is useful, this can be stored with "Programming stage" into dict_EfuseCategory2BinCutTestJob
                End If '''If str_domain_temp <> ""
                
            '''//If "Algorithm" = "base" and Efuse category with the keyword "*base*voltage*", check if the default value is same as "Base Voltage" in sheet Vdd_Binning_Def.
            ElseIf LCase(str_category_temp) Like "*base*voltage" And LCase(str_alg_temp) = "base" Then
                If LCase(Trim(ws_def.Cells(Row, col_alg).value)) = "base" Then
                    If ws_def.Cells(Row, col_defaultValue).value <> "" Then
                        BaseVoltageFromEfuseBDF = CDbl(ws_def.Cells(Row, col_defaultValue).value)

                        '''//Compare BaseVoltage values from from "Vdd_Binning_Def" and "EFUSE_BitDef_Table".
                        If BaseVoltageFromEfuseBDF <> VddbinningBaseVoltage Then
                            Call Print_Error_Message(Error_Info, "LIB_VDD_BINNING", "Parsing_IDSname_from_BDF_Table", "sheet:" & sheetName & ",row:" & Row & ",Efuse category:" & str_category_temp & ", Base_Voltage=" & BaseVoltageFromEfuseBDF & " is inconsistent with the value in header of Vdd_Binning_Def tables. Error!!!")
                            'TheExec.ErrorLogMessage "sheet:" & sheetName & ",row:" & Row & ",Efuse category:" & str_category_temp & ", Base_Voltage=" & BaseVoltageFromEfuseBDF & " is inconsistent with the value in header of Vdd_Binning_Def tables. Error!!!"    ''20230807: Removed by sanity check
                        End If
                    Else
                        Call Print_Error_Message(Error_Info, "LIB_VDD_BINNING", "Parsing_IDSname_from_BDF_Table", "sheet:" & sheetName & ",row:" & Row & ", category:" & str_category_temp & ", Base_Voltage is undefined in Efuse_BitDef_Table. Error!!!")
                        'TheExec.ErrorLogMessage "sheet:" & sheetName & ",row:" & Row & ",category:" & str_category_temp & ", Base_Voltage is undefined in Efuse_BitDef_Table. Error!!!"    ''20230807: Removed by sanity check
                    End If
                    
                    '''//No need to store Efuse category with the keyword "*base*voltage*" into dict_EfuseCategory2BinCutTestJob.
                    str_category_temp = vbNullString
                End If
            Else
                str_category_temp = vbNullString
            End If '''If LCase(str_category_temp) Like...
            
            '''*****************************************************************************************************************************************'''
            '''//Store the programming stage by getBinCutJobDefinition if Efuse category is correct.
            '''*****************************************************************************************************************************************'''
            '''20210703: Modified to use dict_EfuseCategory2BinCutTestJob as the dictionary of Efuse category and the matched programming state in Efuse.
            If str_category_temp <> "" Then
                '''//Mapping Efuse "Programming stage" to BinCut testJob.
               ' If UCase(str_stage_temp) = UCase("CP2") Then
               '20220711
                str_testJob = Mapping_TestJobName_to_BincutJobName(str_stage_temp)
                idx_testJob = getBinCutJobDefinition(str_testJob)
                BDF_tastJob = str_stage_temp
                
                'End If
                If dict_EfuseCategory2BinCutTestJob.Exists(str_category_temp) = True And dict_EfuseCategory2BDF.Exists(str_category_temp) = True Then
                    Call Print_Error_Message(Error_Info, "LIB_VDD_BINNING", "Parsing_IDSname_from_BDF_Table", "sheet:" & sheetName & ",row:" & Row & ",Efuse category:" & str_category_temp & ", it has the duplicate programming stage in sheet:" & sheetName & ". Error!!!")
                    'TheExec.ErrorLogMessage "sheet:" & sheetName & ",row:" & Row & ",Efuse category:" & str_category_temp & ", it has the duplicate programming stage in sheet:" & sheetName & ". Error!!!"    ''20230807: Removed by sanity check
                Else
                    dict_EfuseCategory2BinCutTestJob.Add str_category_temp, idx_testJob
                    dict_EfuseCategory2BDF.Add str_category_temp, BDF_tastJob
                End If
            End If
        Next Row
    End If '''If enableRowParsing = True Then
        
    '''//If the testJob is for BinCut search, it should have the dedicated "Product_Identifier", as commented by C651 Si and Toby, 20210727.
    '''//If the current testJob is for BinCut search without any Efuse category "Product_Identifier", show the error message to users, as requested by C651 Toby.
    '''20210730: Modified to show the error message to users if the current testJob is for BinCut search without any Efuse category "Product_Identifier", as requested by C651 Toby.
    If is_BinCutJob_for_StepSearch = True Then
        str_Efuse_write_ProductIdentifier = get_Efuse_category_by_BinCut_testJob("write", "Product_Identifier")
        
        If str_Efuse_write_ProductIdentifier = "" And Not (LCase(TheExec.CurrentJob) = "wlft1" Or LCase(TheExec.CurrentJob) = "wlft2") Then
            Call Print_Error_Message(Error_Info, "LIB_VDD_BINNING", "Parsing_IDSname_from_BDF_Table", "testJob:" & TheExec.CurrentJob & ", it is for BinCut search, but it doesn't have any Efuse category about Product_Identifier. Please check Efuse_BitDef_Table and BinCut flow table. Error!!!")
            'TheExec.ErrorLogMessage "testJob:" & TheExec.CurrentJob & ", it is for BinCut search, but it doesn't have any Efuse category about Product_Identifier. Please check Efuse_BitDef_Table and BinCut flow table. Error!!!"    ''20230807: Removed by sanity check
        End If
    End If
Exit Function
errHandler:
    Call Print_Error_Message(Error_Info, "LIB_VDD_BINNING", "Parsing_IDSname_from_BDF_Table") 'Add ErrHandler 2023/08/01
    If AbortTest Then Exit Function Else Resume Next
End Function

'20210727: C651 Toby provided the BinCut flow with testCondition "M*### E1 voltage" for non BinCut search in CP1, so that judge_stored_IDS(check_IDS) should be compatible with all BinCut testJobs.
'20210707: Modified to check if ids_name (Efuse category) exists in dict_EfuseCategory2BinCutTestJob.
'20201228: Patty asked us to adapt Efuse object vbt code.
'20201210: Modified to use the flag "is_BinCutJob_for_StepSearch" for "check_bincutJob_for_StepSearch" to check if the test program is binSearch or functional test.
'20200807: Modified to merge the redundant site-loop.
'20190813: Modified to use different IDS lo_limit by BinCut testjobs.
'20190630: Modified to show the error message when str_IDS_PowerDomain is empty.
'20190523: Created for the new data type "IDS_for_BinCut".
'20170810: SWLINZA modified the vbt code to Get Resolution and set it for low limit for OtherRail.
Public Function get_lo_limit_for_IDS(powerDomain As String, Lo_Limit As SiteDouble)
On Error GoTo errHandler 'Add ErrHandler 2023/08/01
    Dim site As Variant
    Dim str_IDS_PowerDomain As String
    '''for Efuse obj vbt.
    Dim field As eFuseBdfField
'''//==================================================================================================================================================================================//'''
'''//Note:
'''1. The flag "is_BinCutJob_for_StepSearch" is True if any testCondition from the table "Non_Binning_Rail" has the keyword "*Evaluate*Bin*".
'''2. As per discussion with TSMC SWLINZA, for powerPin group, it should use 1st powerPin to check IDS limit of powerPin group, 20210707.
'''ex: powerGroup: VDD_FIXED_GRP, and its 1st powerPin: VDD_FIXED, so that compare IDS value of VDD_FIXED with IDS_limit of VDD_FIXED_GRP. It must have Efuse category in Efuse_BitDef_Table to store IDS for VDD_FIXED.
'''3. C651 Toby provided the BinCut flow with testCondition "M*### E1 voltage" for non BinCut search in CP1, so that judge_stored_IDS(check_IDS) should be compatible with all BinCut testJobs, 20210727.
'''//==================================================================================================================================================================================//'''
    For Each site In TheExec.sites
        str_IDS_PowerDomain = IDS_for_BinCut(VddBinStr2Enum(powerDomain)).ids_name(site)
        
        '''//If Efuse IDS name of powerDomain exists in Efuse_BitDef_Table, take Efuse IDS Resolution as IDS low limit for OtherRail.
        If dict_EfuseCategory2BinCutTestJob.Exists(UCase(str_IDS_PowerDomain)) = True Then
            '''For project with Efuse obj vbt code.
            '''20201228: Patty asked us to adapt Efuse object vbt code.
            Set field = BdfDataBase.Bank_Cfg.Fields(str_IDS_PowerDomain)
            Lo_Limit(site) = 1# * field.Resolution
        ElseIf Flag_VDD_Binning_Offline = True Or TheExec.flow.enableWord("Vddbinning_OpenSocket") = True Then '''If the tester is offline or opensocket.
            Lo_Limit(site) = 0
        Else
            Lo_Limit(site) = 0
            Call Print_Error_Message(Error_Info, "LIB_VDD_BINNING", "get_lo_limit_for_IDS", powerDomain & " has no definition of the Efuse IDS resolution as IDS lo_limit. get_lo_limit_for_IDS has the error. Error!!!")
            'TheExec.Datalog.WriteComment powerDomain & " has no definition of the Efuse IDS resolution as IDS lo_limit. get_lo_limit_for_IDS has the error. Error!!!"
            'TheExec.ErrorLogMessage powerDomain & " has no definition of the Efuse IDS resolution as IDS lo_limit. get_lo_limit_for_IDS has the error. Error!!!"    ''20230807: Removed by sanity check
        End If
    Next site
Exit Function
errHandler:
    Call Print_Error_Message(Error_Info, "LIB_VDD_BINNING", "get_lo_limit_for_IDS") 'Add ErrHandler 2023/08/01
    If AbortTest Then Exit Function Else Resume Next
End Function

'20191002: Created to init BincutVoltageTypeName for printing BV strings.
Public Function initBincutVoltageType()
On Error GoTo errHandler
'''//==================================================================================================================================================================================//'''
'''//Note:
'''Please update Enum BincutVoltageType and MaxBincutVoltageType in GlobalVariable once if you want insert the new type!!!
'''//==================================================================================================================================================================================//'''
    BincutVoltageTypeName(BincutVoltageType.None) = vbNullString
    BincutVoltageTypeName(BincutVoltageType.InitialVoltage) = "Initial Voltage"
    BincutVoltageTypeName(BincutVoltageType.SafeVoltage) = "Safe Voltage"
    BincutVoltageTypeName(BincutVoltageType.PayloadVoltage) = "Payload Voltage"
    BincutVoltageTypeName(BincutVoltageType.PostbincutBinningpower) = "Postbincut BinningPower_BinResult Voltage"
    BincutVoltageTypeName(BincutVoltageType.PostbincutAllpower) = "Payload AllPower_BinResult Voltage"
Exit Function
errHandler:
    Call Print_Error_Message(Error_Info, "LIB_VDD_BINNING", "initBincutVoltageType") 'Add ErrHandler 2023/08/01
    If AbortTest Then Exit Function Else Resume Next
End Function

'20191105: Created to init TestType names.
Public Function initTestTypeName()
On Error GoTo errHandler
    TestTypeName(testType.TD) = "TD"
    TestTypeName(testType.Mbist) = "Mbist"
    TestTypeName(testType.Func) = "Func"
    TestTypeName(testType.RTOS) = "RTOS"
    TestTypeName(testType.ldcbfd) = vbNullString
Exit Function
errHandler:
    Call Print_Error_Message(Error_Info, "LIB_VDD_BINNING", "initTestTypeName") 'Add ErrHandler 2023/08/01
    If AbortTest Then Exit Function Else Resume Next
End Function

'20210113: Modified to check if cnt_DecomposedPinList>0.
'20200423: Modified to move "Dim dictPin2Dcspec As New Dictionary" into GlobalVariable.
'20191231: Modified to use TestJob of test program for DC Specs.
'20191219: Created for Domain2Pin and Pin2Domain.
Public Function initDomain2Pin(domainList As String, dictDomain2Pin As Dictionary, dictPin2Domain As Dictionary)
On Error GoTo errHandler 'Add ErrHandler 2023/08/01
    Dim strAry_powerDomain() As String
    Dim i As Integer
    Dim j As Integer
    Dim strAry_pinVddbin() As String
    Dim cnt_DecomposedPinList As Long
    Dim PinGroup As String
    Dim domainTemp As String
    Dim domainTrackpower As String
    Dim PinTemp As String
'''//==================================================================================================================================================================================//'''
'''//Note:
'''//Check if powerDomain or powerPin is connected to DCVS.
'''1. BinCut powerDdomains are the pinGroup, and powerDomains include pins.
'''2. domain2pinDict, pin2domainDict are the dictionaries in GlobalVarible to store domains and pins.
'''//==================================================================================================================================================================================//'''
    '''init
    PinGroup = vbNullString

    If domainList <> "" Then
        strAry_powerDomain = Split(domainList, ",")
        
        '''//Use TestJob of test program to find the matched DC Specs.
        ParsingDCspec UCase(TheExec.CurrentJob), dictPin2Dcspec
        
        For i = 0 To UBound(strAry_powerDomain)
            domainTemp = UCase(Trim(strAry_powerDomain(i)))
            domainTrackpower = vbNullString
            
            If VddbinPmodeDict.Exists(domainTemp) = True Then
                PinGroup = vbNullString
                
                '''//For main Domain
                Call TheExec.DataManager.DecomposePinList(domainTemp, strAry_pinVddbin, cnt_DecomposedPinList)
                
                If cnt_DecomposedPinList > 0 Then
                    If dictDomain2Pin.Exists(domainTemp) Then
                        '''Do nothing
                    Else
                        '''//For powerPin in main domain
                        For j = 0 To cnt_DecomposedPinList - 1
                            PinTemp = UCase(Trim(strAry_pinVddbin(j)))
                        
                            If TheExec.DataManager.ChannelType(PinTemp) = "N/C" Then
                                '''Do nothing
                            Else
                                '''//Check if DCVS is connected to the powerPin.
                                If IsDcvsConnected(PinTemp) = True Then
                                    If dictPin2Domain.Exists(PinTemp) Then
                                        '''Do nothing
                                    Else
                                        dictPin2Domain.Add PinTemp, domainTemp
                                    End If
                                
                                    If PinGroup = "" Then
                                        PinGroup = PinTemp
                                    Else
                                        PinGroup = PinGroup & "," & PinTemp
                                    End If
                                End If
                            End If
                        Next j
                        
                        If PinGroup = "" Then
                            Call Print_Error_Message(Error_Info, "LIB_VDD_BINNING", "initDomain2Pin", "Domain: " & domainTemp & " for initDomain2Pin contains no DCVS powerPin. Error!!!")
                            'TheExec.ErrorLogMessage "Domain: " & domainTemp & " for initDomain2Pin contains no DCVS powerPin. Error!!!"    ''20230807: Removed by sanity check
                        End If
                        
                        '''//Check if tracking power domain of main domain exists.
                        If AllBinCut(VddBinStr2Enum(domainTemp)).TRACKINGPOWER <> "" Then
                            domainTrackpower = AllBinCut(VddBinStr2Enum(domainTemp)).TRACKINGPOWER
                            
                            Call TheExec.DataManager.DecomposePinList(domainTrackpower, strAry_pinVddbin, cnt_DecomposedPinList)
                            
                            '''//Check if tracking power has cnt_DecomposedPinList>0.
                            If cnt_DecomposedPinList > 0 Then
                                For j = 0 To cnt_DecomposedPinList - 1
                                    PinTemp = UCase(Trim(strAry_pinVddbin(j)))
                                    
                                    '''//Check if DCVS is connected to the powerPin.
                                    If IsDcvsConnected(PinTemp) = True Then
                                        If dictPin2Domain.Exists(PinTemp) Then
                                            '''Do nothing
                                        Else
                                            '''//Update dictionary of Domain2Pin
                                            dictPin2Domain.Add PinTemp, domainTemp
                                        End If
                                        
                                        If PinGroup = "" Then
                                            Call Print_Error_Message(Error_Info, "LIB_VDD_BINNING", "initDomain2Pin", "TrackingPower:" & domainTrackpower & " of Domain:" & domainTemp & " contains no DCVS powerPin. Error!!!")
                                            'TheExec.ErrorLogMessage "TrackingPower:" & domainTrackpower & " of Domain:" & domainTemp & " contains no DCVS powerPin. Error!!!"    ''20230807: Removed by sanity check
                                        Else
                                            PinGroup = PinGroup & "," & PinTemp
                                        End If
                                    End If
                                Next j
                            Else
                                Call Print_Error_Message(Error_Info, "LIB_VDD_BINNING", "initDomain2Pin", "TrackingPower:" & domainTrackpower & " of Domain:" & domainTemp & " isn't defined in PinMap for initDomain2Pin. Error!!!")
                                'TheExec.ErrorLogMessage "TrackingPower:" & domainTrackpower & " of Domain:" & domainTemp & " isn't defined in PinMap for initDomain2Pin. Error!!!"    ''20230807: Removed by sanity check
                            End If
                        End If
                        
                        '''//Update dictionary of Domain2Pin.
                        dictDomain2Pin.Add domainTemp, PinGroup
                        
                        If Not CheckDomainPinsDCspec(domainTemp, dictDomain2Pin, dictPin2Dcspec) Then
                            Call Print_Error_Message(Error_Info, "LIB_VDD_BINNING", "initDomain2Pin", "Domain: " & UCase(Trim(domainTemp)) & "," & " has different dcspec pins. Error!!!")
                            'TheExec.ErrorLogMessage "Domain: " & UCase(Trim(domainTemp)) & "," & " has different dcspec pins. Error!!!"    ''20230807: Removed by sanity check
                        End If
                    End If
                Else
                    Call Print_Error_Message(Error_Info, "LIB_VDD_BINNING", "initDomain2Pin", "Domain: " & UCase(Trim(domainTemp)) & "," & " isn't defined in PinMap for initDomain2Pin. Error!!!")
                    'TheExec.ErrorLogMessage "Domain: " & UCase(Trim(domainTemp)) & "," & " isn't defined in PinMap for initDomain2Pin. Error!!!"    ''20230807: Removed by sanity check
                End If
            Else
                Call Print_Error_Message(Error_Info, "LIB_VDD_BINNING", "initDomain2Pin", "Domain: " & UCase(Trim(domainTemp)) & "," & " doesn't enumerate in Vdd_Binning_Def. Error!!!")
                'TheExec.ErrorLogMessage "Domain: " & UCase(Trim(domainTemp)) & "," & " doesn't enumerate in Vdd_Binning_Def. Error!!!"    ''20230807: Removed by sanity check
            End If
        Next i
    Else
        Call Print_Error_Message(Error_Info, "LIB_VDD_BINNING", "initDomain2Pin", "Argument of domainList for initDomain2Pin should not be empty. Error!!!")
        'TheExec.ErrorLogMessage "Argument of domainList for initDomain2Pin should not be empty. Error!!!"    ''20230807: Removed by sanity check
    End If
Exit Function
errHandler:
    Call Print_Error_Message(Error_Info, "LIB_VDD_BINNING", "initDomain2Pin") 'Add ErrHandler 2023/08/01
    If AbortTest Then Exit Function Else Resume Next
End Function

'20200106: Modified to remove the ErrorLogMessage.
'20191219: Created for Domain2Pin.
Public Function VddbinDomain2Pin(vddbinDomain As String) As String
On Error GoTo errHandler 'Add ErrHandler 2023/08/01
    Dim strTemp As String
    strTemp = UCase(vddbinDomain)

    If domain2pinDict.Exists(strTemp) Then
        VddbinDomain2Pin = UCase(domain2pinDict.item(strTemp))
    Else
        VddbinDomain2Pin = "Domain_Error"
        Call Print_Error_Message(Error_Info, "LIB_VDD_BINNING", "VddbinDomain2Pin", "Vddbin Domain=" & vddbinDomain & ", but it doesn't exist in VddbinDomain2Pin. Error!!!")
        'TheExec.ErrorLogMessage "Vddbin Domain=" & vddbinDomain & ", but it doesn't exist in VddbinDomain2Pin. Error!!!"
    End If
Exit Function
errHandler:
    Call Print_Error_Message(Error_Info, "LIB_VDD_BINNING", "VddbinDomain2Pin") 'Add ErrHandler 2023/08/01
    If AbortTest Then Exit Function Else Resume Next
End Function

'20200106: Modified to remove the ErrorLogMessage.
'20191219: Created for Pin2Domain.
Public Function VddbinPin2Domain(vddbinPin As String) As String
On Error GoTo errHandler 'Add ErrHandler 2023/08/01
    Dim strTemp As String
    strTemp = UCase(vddbinPin)

    If pin2domainDict.Exists(strTemp) Then
        VddbinPin2Domain = UCase(pin2domainDict.item(strTemp))
    Else
        VddbinPin2Domain = "Pin_Error"
        Call Print_Error_Message(Error_Info, "LIB_VDD_BINNING", "VddbinPin2Domain", "Pin:" & vddbinPin & ", but it doesn't exist in the dictionary pin2domainDict. Please check BinCut voltage table and DC specs sheets for VddbinPin2Domain. Error!!!")
        'TheExec.ErrorLogMessage "Pin:" & vddbinPin & ", but it doesn't exist in the dictionary pin2domainDict. Please check BinCut voltage table and DC specs sheets for VddbinPin2Domain. Error!!!"
    End If
Exit Function
errHandler:
    Call Print_Error_Message(Error_Info, "LIB_VDD_BINNING", "VddbinPin2Domain") 'Add ErrHandler 2023/08/01
    If AbortTest Then Exit Function Else Resume Next
End Function

'20210113: Modified to check if cnt_DecomposedPinList>0.
'20200106: Modified to remove the ErrorLogMessage.
'20191219: Created to check DCVS type for powerPin.
Public Function IsDcvsConnected(powerPin As String) As Boolean
On Error GoTo errHandler 'Add ErrHandler 2023/08/01
    Dim strAry_PinName() As String
    Dim cnt_DecomposedPinList As Long
    Dim typesCount As Long
    Dim strAry_InstrumentTypes() As String
    Call TheExec.DataManager.DecomposePinList(powerPin, strAry_PinName(), cnt_DecomposedPinList)
    
    '''//If NumberPins>1, it means that powerPin might be a powerDomain(pinGroup) or an incorrect powerPin.
    If cnt_DecomposedPinList = 1 Then
        Call TheExec.DataManager.GetChannelTypes(powerPin, typesCount, strAry_InstrumentTypes())
        
        If LCase(strAry_InstrumentTypes(0)) Like "dcvs*" Then
            IsDcvsConnected = True
            If VddbinPinDcvstypeDict.Exists(UCase(powerPin)) = True Then
                '''Do nothing
            Else
                VddbinPinDcvstypeDict.Add UCase(powerPin), UCase(GetInstrument_BV(powerPin, 0))
            End If
        Else
            IsDcvsConnected = False
        End If
    Else
        IsDcvsConnected = False
        Call Print_Error_Message(Error_Info, "LIB_VDD_BINNING", "IsDcvsConnected", "Pin:" & powerPin & ", it is not a correct powerPin for IsDcvsConnected. Error!!!")
        'TheExec.ErrorLogMessage "Pin:" & powerPin & ", it is not a correct powerPin for IsDcvsConnected. Error!!!"
    End If
Exit Function
errHandler:
    Call Print_Error_Message(Error_Info, "LIB_VDD_BINNING", "IsDcvsConnected") 'Add ErrHandler 2023/08/01
    If AbortTest Then Exit Function Else Resume Next
End Function

'20200106: Modified to remove the ErrorLogMessage.
'20191219: Created for Pin2Domain
Public Function VddbinPinDcvsType(vddbinPin As String) As String
On Error GoTo errHandler 'Add ErrHandler 2023/08/01
    Dim strTemp As String
    strTemp = UCase(vddbinPin)

    If VddbinPinDcvstypeDict.Exists(strTemp) Then
        VddbinPinDcvsType = UCase(VddbinPinDcvstypeDict.item(strTemp))
    Else
        VddbinPinDcvsType = "Pin_Error"
        Call Print_Error_Message(Error_Info, "LIB_VDD_BINNING", "VddbinPinDcvsType", "Pin:" & vddbinPin & ", it doesn't exist in VddbinPinDcvsType. Error!!!")
        'TheExec.ErrorLogMessage "Pin:" & vddbinPin & ", it doesn't exist in VddbinPinDcvsType. Error!!!"
    End If
Exit Function
errHandler:
    Call Print_Error_Message(Error_Info, "LIB_VDD_BINNING", "VddbinPinDcvsType") 'Add ErrHandler 2023/08/01
    If AbortTest Then Exit Function Else Resume Next
End Function

'20200106: Modified to remove the ErrorLogMessage.
'20191219: Created to get 1st Pin from PinGroup.
Public Function Get1stPinFromPingroup(powerDomain As String) As String
On Error GoTo errHandler 'Add ErrHandler 2023/08/01
    Dim split_content() As String
    Dim i As Long
    split_content = Split(powerDomain, ",")
    
    If UBound(split_content) > -1 Then
        For i = 0 To UBound(split_content)
            If TheExec.DataManager.ChannelType(split_content(i)) = "N/C" Then
                ' Do nothing
            Else
                Get1stPinFromPingroup = split_content(i)
                Exit For
            End If
        Next i
    Else
        Call Print_Error_Message(Error_Info, "LIB_VDD_BINNING", "Get1stPinFromPingroup", "powerDomain: " & powerDomain & ", it is incorrect for Get1stPinFromPingroup. Error!!!")
        'TheExec.ErrorLogMessage "powerDomain: " & powerDomain & ", it is incorrect for Get1stPinFromPingroup. Error!!!"
    End If
Exit Function
errHandler:
    Call Print_Error_Message(Error_Info, "LIB_VDD_BINNING", "Get1stPinFromPingroup") 'Add ErrHandler 2023/08/01
    If AbortTest Then Exit Function Else Resume Next
End Function
 
'20200618: Modified to move globalVariable "Public pin2dcspecstatusDict As New dictionary" to local variable of the function "CheckDomainPinsDCspec".
'20200106: Modified to remove the ErrorLogMessage.
'20191220: Created for checking group pins DCspec
Public Function CheckDomainPinsDCspec(powerDomain As String, dictDomain2Pin As Variant, DCspecTable As Variant) As Boolean
On Error GoTo errHandler 'Add ErrHandler 2023/08/01
    Dim strTemp As String, DCspecTmp As String, CompareBase As String
    Dim split_content() As String
    Dim isCompResultDiff As Boolean
    Dim i As Long
    Dim AllPinExist As Boolean
    Dim pin2dcspecstatusDict As New Dictionary '''Note: "same" means group-pins has same dc spec, "diff" means group-pins has different dc spec, "none" means isn't group-pin.
    AllPinExist = True
    isCompResultDiff = False
    strTemp = dictDomain2Pin.item(powerDomain)
    split_content = Split(strTemp, ",")
    pin2dcspecstatusDict.RemoveAll
    
    For i = 0 To UBound(split_content)
        If UCase(split_content(i)) Like "*_FT" Then split_content(i) = Replace(split_content(i), "_FT", "", compare:=vbTextCompare)
        If Not DCspecTable.Exists(split_content(i) & "_VAR") Then
            AllPinExist = False
            Call Print_Error_Message(Error_Info, "LIB_VDD_BINNING", "CheckDomainPinsDCspec", "Pin:" & split_content(i) & "_VAR" & ", it doesn't exist in DCspec. Error!!!")
            'TheExec.ErrorLogMessage "Pin:" & split_content(i) & "_VAR" & ", it doesn't exist in DCspec. Error!!!"
        End If
    Next i
    
    If AllPinExist = True Then
        CompareBase = DCspecTable.item(split_content(0) & "_VAR")
    
        If UBound(split_content) = 0 Then
            If Not pin2dcspecstatusDict.Exists(split_content(0)) Then
                pin2dcspecstatusDict.Add split_content(0), "none"
            End If
            CheckDomainPinsDCspec = True
        Else
            For i = 1 To UBound(split_content)
                DCspecTmp = DCspecTable.item(split_content(i) & "_VAR")
                If Not (DCspecTmp = CompareBase) Then
                    isCompResultDiff = True
                    Exit For
                End If
            Next i
            
            If isCompResultDiff = True Then
                For i = 0 To UBound(split_content)
                    If Not pin2dcspecstatusDict.Exists(split_content(i)) Then
                        pin2dcspecstatusDict.Add split_content(i), "diff"
                    End If
                Next i
                CheckDomainPinsDCspec = False
            Else
                For i = 0 To UBound(split_content)
                    If Not pin2dcspecstatusDict.Exists(split_content(i)) Then
                        pin2dcspecstatusDict.Add split_content(i), "same"
                    End If
                Next i
                CheckDomainPinsDCspec = True
            End If
        End If
    End If
Exit Function
errHandler:
    Call Print_Error_Message(Error_Info, "LIB_VDD_BINNING", "CheckDomainPinsDCspec") 'Add ErrHandler 2023/08/01
    If AbortTest Then Exit Function Else Resume Next
End Function

'20200807: Modified to merge the branches.
'20200731: Modified to check if Spec name starts with the keyword "VDD".
'20200710: Modified to parse multiple DC_Spec sheets.
'20200703: Modiifed to use "check_Sheet_Range".
'20200508: Modified to merge "Find_Spec_Sheet" and "Find_JobList_Sheet" into "Find_Sheet".
'20191231: Modified to use TestJob of test program for DC Specs.
'20191220: Created for Parsing DCspec.
Public Function ParsingDCspec(testJob As String, DCspecTable As Variant)
On Error GoTo errHandler 'Add ErrHandler 2023/08/01
    Dim wb As Workbook
    Dim ws_def As Worksheet
    Dim sheetName As String
    Dim MaxRow As Long
    Dim maxcol As Long
    Dim i As Long, j As Long, k As Long
    Dim bincut_catlog_pos() As Integer
    Dim bincut_catlog_cnt As Integer
    Dim Row_of_Title As Integer
    Dim col_symbol As Integer
    Dim PinTemp As String
    Dim valueTemp As String
    Dim enableRowParsing As Boolean
    Dim isSheetFound As Boolean
    Dim strAry_DCSheetName() As String
    '''ToDo: Maybe we can check if testJob is emulated in "Enum BinCutJobDefinition"...
    If testJob <> "" Then
        '''//Find the matched DC specs sheet name for testJob from "GeneratedJobListSheet"
        Parsing_GeneratedJobList_Sheet testJob, "GeneratedJobListSheet", sheetName
        
        '''//Parse multiple DC_Spec sheets.
        strAry_DCSheetName = Split(sheetName, ",")
        
        For k = 0 To UBound(strAry_DCSheetName)
            '''*****************************************************************'''
            '''//Check if testJob has the matched DC Specs sheet.
            Set wb = Application.ActiveWorkbook
            Call check_Sheet_Range(strAry_DCSheetName(k), wb, ws_def, MaxRow, maxcol, isSheetFound)
            '''*****************************************************************'''
            If isSheetFound = True Then
                '''//Init
                '''Since all col_XXX and row_XXX related variables with default values=0, no need to initialize them as 0.
                bincut_catlog_cnt = 0
                Row_of_Title = 0
                enableRowParsing = False
                
                '''//Check the header of the table.
                '''Get the columns for the diverse coefficient.
                For i = 1 To MaxRow
                    For j = 1 To maxcol
                        If UCase(ws_def.Cells(i, j).value) = "SYMBOL" Then
                            col_symbol = j
                            Row_of_Title = i
                        End If
                        
                        If Row_of_Title > 0 Then
                            If UCase(ws_def.Cells(Row_of_Title - 1, j).value) Like "BINCUT_*" Then
                                ReDim Preserve bincut_catlog_pos(bincut_catlog_cnt)
                                bincut_catlog_pos(bincut_catlog_cnt) = j
                                bincut_catlog_cnt = bincut_catlog_cnt + 1
                            End If
                        End If
                    Next j
                    
                    If Row_of_Title > 0 Then
                        If bincut_catlog_cnt > 0 Then
                            enableRowParsing = True
                            Exit For
                        Else
                            enableRowParsing = False
                            If Not UCase(strAry_DCSheetName(k)) Like "*_BI" And Not UCase(strAry_DCSheetName(k)) Like "*_SC" Then
                                Call Print_Error_Message(Error_Info, "LIB_VDD_BINNING", "", strAry_DCSheetName(k) & " doesn't contain any DC Category name with keyword BINCUT. Error!!!")
                                'TheExec.Datalog.WriteComment strAry_DCSheetName(k) & " doesn't contain any DC Category name with keyword BINCUT. Error!!!"
                                'TheExec.ErrorLogMessage strAry_DCSheetName(k) & " doesn't contain any DC Category name with keyword BINCUT. Error!!!"    ''20230807: Removed by sanity check
                            End If
                            Exit For
                        End If
                    End If
                Next i
                
                '''//Start parsing the cells
                If enableRowParsing = True Then
                    For i = Row_of_Title + 1 To MaxRow
                        valueTemp = vbNullString
                        PinTemp = ws_def.Cells(i, col_symbol).value
                        
                        '''//Check if Spec name starts with the keyword "VDD".
                        If (UCase(PinTemp) Like "VDD*") Then
                            For j = 0 To UBound(bincut_catlog_pos)
                                If valueTemp = "" Then
                                    valueTemp = valueTemp & ws_def.Cells(i, bincut_catlog_pos(j)).value & "," & ws_def.Cells(i, bincut_catlog_pos(j) + 1).value & "," & ws_def.Cells(i, bincut_catlog_pos(j) + 2).value
                                Else
                                    valueTemp = valueTemp & "," & ws_def.Cells(i, bincut_catlog_pos(j)).value & "," & ws_def.Cells(i, bincut_catlog_pos(j) + 1).value & "," & ws_def.Cells(i, bincut_catlog_pos(j) + 2).value
                                End If
                            Next j
                        
                            If DCspecTable.Exists(PinTemp) Then
                                Call Print_Error_Message(Error_Info, "LIB_VDD_BINNING", "ParsingDCspec", "DCspec pins repeat. Error!!!")
                                'TheExec.ErrorLogMessage "DCspec pins repeat. Error!!!"    ''20230807: Removed by sanity check
                            Else
                               DCspecTable.Add PinTemp, valueTemp
                            End If
                        End If
                    Next i
                Else
                    If Not UCase(strAry_DCSheetName(k)) Like "*_BI" And Not UCase(strAry_DCSheetName(k)) Like "*_SC" Then
                        Call Print_Error_Message(Error_Info, "LIB_VDD_BINNING", "ParsingDCspec", strAry_DCSheetName(k) & " doesn't have the correct header for ParsingDCspec. Error!!!")
                        'TheExec.Datalog.WriteComment strAry_DCSheetName(k) & " doesn't have the correct header for ParsingDCspec. Error!!!"
                        'TheExec.ErrorLogMessage strAry_DCSheetName(k) & " doesn't have the correct header for ParsingDCspec. Error!!!"    ''20230807: Removed by sanity check
                    End If
                End If
            End If '''If isSheetFound = True
        Next k '''For k = 0 To UBound(strAry_DCSheetName)
    Else
        Call Print_Error_Message(Error_Info, "LIB_VDD_BINNING", "ParsingDCspec", "The argument testJob of ParsingDCspec is incorrect for ParsingDCspec. Error!!!")
        'TheExec.ErrorLogMessage "The argument testJob of ParsingDCspec is incorrect for ParsingDCspec. Error!!!"    ''20230807: Removed by sanity check
    End If
Exit Function
errHandler:
    Call Print_Error_Message(Error_Info, "LIB_VDD_BINNING", "ParsingDCspec") 'Add ErrHandler 2023/08/01
    If AbortTest Then Exit Function Else Resume Next
End Function

'20210903: Modified to merge properties of "Public Type Instance_Step_Control" into Public Type Instance_Info "Public Type Instance_Info".
'20210809: Modified to remove the redundant property "FoundLevel As New SiteDouble" from Public Type Instance_Step_Control.
'20210809: Modified to check AllBinCut(p_mode).is_for_BinSearch to decide if it has to reset VBIN_Result for p_mode after MultiFSTP.
'20210806: Modified to remove the redundant property "IndexLevelIncDec As New SiteLong" from Public Type Instance_Step_Control.
'20210803: Modified to remove the redundant properties "IDS_current_fail As New SiteLong" and "IDS_current_Min As Double" from Public Type Instance_Step_Control.
'20210422: Modified to use Reset_VBinResult.
'20210422: Modified to remove PassBinCutByDomain and PassBinCutByPmode since the unused vbt code of GradeSearchMethod was removed.
'20210325: Modified to use Flag_Vddbin_DoAll_DebugCollection for TheExec.EnableWord("Vddbin_DoAll_DebugCollection").
'20201209: Modified to use the arguments "inst_info As Instance_Info" and "step_control As Instance_Step_Control" for ResetPmodePowerforBincut.
'20200807: Modified to reduce the redundant site-loop.
'20200525: Modified to remove the redundant site-loop for siteVariants.
'20200427: Modified to reset "PassBinCutByPmode".
'20200423: Modified to replace "BinCut(p_mode, bincutNum(site)).tested = True" with "VBIN_RESULT(p_mode).tested=True".
'20191219: Created to init BinCut binning p_mode power.
Public Function ResetPmodePowerforBincut(inst_info As Instance_Info)
On Error GoTo errHandler 'Add ErrHandler 2023/08/01
    Dim site As Variant
'''//==================================================================================================================================================================================//'''
'''//Caution!!!
'''//If EnableWord "Vddbin_DoAll_DebugCollection" is enabled, initial performance mode result for Char. BinCut search voltage(Grade) and efuse product voltage(GradeVdd) as 0.
'''//==================================================================================================================================================================================//'''
    '''init
    inst_info.grade_found = False
    inst_info.AnySiteGradeFound = False

    For Each site In TheExec.sites
        inst_info.All_Site_Mask = inst_info.All_Site_Mask + 2 ^ site
        inst_info.IDS_ZONE_NUMBER = DYNAMIC_VBIN_IDS_ZONE(inst_info.p_mode).IDS_ZONE_NUMBER
    Next site
    
    '''20210325: Modified to use Flag_Vddbin_DoAll_DebugCollection for TheExec.EnableWord("Vddbin_DoAll_DebugCollection").
    If Flag_Vddbin_DoAll_DebugCollection = True Or TheExec.enableWord("Vddbin_PTE_Debug") = True Then
        CurrentPassBinCutNum = 1 '''set the bincut number to 1.
    
        '''//Initialize bincut search voltage and efuse product voltage of the performance mode for BinCut doAll Char.
        '''20210422: Modified to use Reset_VBinResult.
        '''20210809: Modified to check AllBinCut(p_mode).is_for_BinSearch to decide if it has to reset VBIN_Result for p_mode after MultiFSTP.
        If AllBinCut(inst_info.p_mode).is_for_BinSearch = True Then
            'VBIN_RESULT(AllBinCut(inst_info.p_mode).PREVIOUS_Performance_Mode).tested = False        'set the previous performance mode to false
            Call Reset_VBinResult(inst_info.p_mode)
        End If
    '20230605: Modified for new COF_StepInheritence rule, do not set E1 when there's no grade. Requested by C651 Oscar and TSMC ZYLINI
    'VBIN_RESULT(inst_info.p_mode).step_1stPass_in_IDS_Zone = -1 -> no real step was found
    ElseIf Flag_Vddbin_COF_StepInheritance = True Then
        For Each site In TheExec.sites
            If VBIN_RESULT(inst_info.p_mode).step_1stPass_in_IDS_Zone(site) = -1 Then
                '''//Initialize bincut search voltage and efuse product voltage of the performance mode
                If AllBinCut(inst_info.p_mode).is_for_BinSearch = True Then
                    Call Reset_VBinResult(inst_info.p_mode, site)
                End If
            End If
        Next site
    End If
Exit Function
errHandler:
    Call Print_Error_Message(Error_Info, "LIB_VDD_BINNING", "ResetPmodePowerforBincut") 'Add ErrHandler 2023/08/01
    If AbortTest Then Exit Function Else Resume Next
End Function

'20210422: Modified to use Reset_VBinResult.
'20210422: Modified to remove PassBinCutByDomain and PassBinCutByPmode since the unused vbt code of GradeSearchMethod was removed.
'20210120: Modified to use VBIN_RESULT(p_mode).step_1stPass_in_IDS_Zone to store the first pass step in Dynamic IDS Zone and find the correspondent PassBinCut number.
'20201203: Modified to use "idx_PowerDomain = VddBinStr2Enum(pinGroup_BinCut(i))".
'20200423: Modified to replace "BinCut(p_mode, bincutNum(site)).tested = True" with "VBIN_RESULT(p_mode).tested=True".
'20200320: Modified to use the flag "Flag_Skip_ReApplyPayloadVoltageToDCVS".
'20200317: Modified for SearchByPmode.
'20200130: Modified to init DomainPassBinCutNum, BinCut_Init_Voltage, and BinCut_Payload_Voltage.
'20191230: Created for init VBIN_RESULT array.
Public Function initVbinTest()
On Error GoTo errHandler 'Add ErrHandler 2023/08/01
    Dim p_mode As Integer
    Dim i As Long
    Dim idx_powerDomain As Long
    '''//Initialize VBIN_RESULT for each p_mode.
    For p_mode = 0 To MaxPerformanceModeCount
        Call Reset_VBinResult(p_mode)
    Next p_mode
    
    '''//Init BinCut_Init_Voltage, and BinCut_Payload_Voltage.
    For i = 0 To UBound(pinGroup_BinCut)
        idx_powerDomain = VddBinStr2Enum(pinGroup_BinCut(i))
        BinCut_Init_Voltage(idx_powerDomain) = 0
        BinCut_Payload_Voltage(idx_powerDomain) = 0
        Previous_Payload_Voltage(idx_powerDomain) = 0
    Next i
    
    CurrentPassBinCutNum = 1
Exit Function
errHandler:
    Call Print_Error_Message(Error_Info, "LIB_VDD_BINNING", "initVbinTest") 'Add ErrHandler 2023/08/01
    If AbortTest Then Exit Function Else Resume Next
End Function

'20210819: Modified to set VBIN_RESULT(p_mode).GRADE and VBIN_RESULT(p_mode).GRADEVDD as 0.
'20210809: Modified to remove the redundant property "ALL_SITE_MIN As New SiteDouble" from Public Type VBIN_RESULT_TYPE.
'20210422: Created to reset VBin_Result(p_mode).
Public Function Reset_VBinResult(p_mode As Integer, Optional ByVal site As Long = -1)
On Error GoTo errHandler
    If site = -1 Then
        VBIN_RESULT(p_mode).GRADE = 0                          '''set the lvcc result = 0
        VBIN_RESULT(p_mode).GRADEVDD = 0
        VBIN_RESULT(p_mode).tested = False                      '''set the flag "Tested" to false
        VBIN_RESULT(p_mode).FLAGFAIL = False                    '''set the fail flag to false, avoid the result always set to fail
        VBIN_RESULT(p_mode).step_in_IDS_Zone = 0                '''default is from step0
        VBIN_RESULT(p_mode).step_1stPass_in_IDS_Zone = -1
        VBIN_RESULT(p_mode).step_in_BinCut = Max_IDS_Step
        VBIN_RESULT(p_mode).passBinCut = 1
        VBIN_RESULT(p_mode).DSSC_Dec = -1
    Else
        ''20230605: Modified for new COF_StepInheritence rule, do not set E1 when there's no grade. Requested by C651 Oscar and TSMC ZYLINI
        VBIN_RESULT(p_mode).GRADE(site) = 0                          '''set the lvcc result = 0
        VBIN_RESULT(p_mode).GRADEVDD(site) = 0
        VBIN_RESULT(p_mode).tested(site) = False                      '''set the flag "Tested" to false
        VBIN_RESULT(p_mode).FLAGFAIL(site) = False                    '''set the fail flag to false, avoid the result always set to fail
        VBIN_RESULT(p_mode).step_in_IDS_Zone(site) = 0                '''default is from step0
        VBIN_RESULT(p_mode).step_1stPass_in_IDS_Zone(site) = -1
        VBIN_RESULT(p_mode).step_in_BinCut(site) = Max_IDS_Step
        VBIN_RESULT(p_mode).passBinCut(site) = 1
        VBIN_RESULT(p_mode).DSSC_Dec(site) = -1
    End If
Exit Function
errHandler:
    Call Print_Error_Message(Error_Info, "LIB_VDD_BINNING", "Reset_VBinResult") 'Add ErrHandler 2023/08/01
    If AbortTest Then Exit Function Else Resume Next
End Function

'20201027: Modified to use "Public Type Instance_Info".
'20200130: Created to get p_mode, addi_mode, and Testtype from the test instance.
Public Function Get_Pmode_Addimode_Testtype_fromInstance(inst_info As Instance_Info)
On Error GoTo errHandler
    '''//Check if any input argument is empty.
    If inst_info.inst_name <> "" And inst_info.performance_mode <> "" Then
        '''//Init
        inst_info.p_mode = 0
        inst_info.Special_Voltage_setup = False
        inst_info.addi_mode = 0
        inst_info.Test_Type = testType.ldcbfd
        inst_info.offsetTestTypeIdx = testType.Func
        inst_info.jobIdx = BinCutJobDefinition.COND_ERROR
        
        '''//According to the instance keyword to decided "Test_Type", ex: elb, spi, rtos
        decide_test_type inst_info.Test_Type, inst_info.inst_name
        
        '''//According to the instance keyword to decide test type, then get dynamic_offset...
        inst_info.offsetTestTypeIdx = decide_offset_testType_byInstName(inst_info.inst_name)
        
        '''//Get BinCut test job name
        inst_info.jobIdx = getBinCutJobDefinition(bincutJobName)
        
        '=================================================================================
        ' Identify if the Performance Mode has the special test Condition (additional mode)
        '=================================================================================
        '''//Init the parameters for parsing the performance mode
        '''//Split the performance mode to get the main performance mode and the additional mode, and get powerpin from the performance mode.
        Call Parsing_Instance_Pmode(inst_info.performance_mode, inst_info.p_mode, inst_info.addi_mode, inst_info.Special_Voltage_setup)
        
        '''//Get powerDomain from the binning p_mode.
        inst_info.powerDomain = AllBinCut(inst_info.p_mode).powerPin
    Else
        Call Print_Error_Message(Error_Info, "LIB_VDD_BINNING", "Get_Pmode_Addimode_Testtype_fromInstance", "Argument 'performance mode' of the instance:" & inst_info.inst_name & " is incorrect for Get_Pmode_Addimode_Testtype_fromInstance. Error!!!")
        'TheExec.ErrorLogMessage "Argument 'performance mode' of the instance:" & inst_info.inst_name & " is incorrect for Get_Pmode_Addimode_Testtype_fromInstance. Error!!!"    ''20230807: Removed by sanity check
    End If
Exit Function
errHandler:
    Call Print_Error_Message(Error_Info, "LIB_VDD_BINNING", "Get_Pmode_Addimode_Testtype_fromInstance") 'Add ErrHandler 2023/08/01
    If AbortTest Then Exit Function Else Resume Next
End Function

'20200320: Modified to save BinCut Payload Voltages of the previous instance into globalvariable "Previous_Payload_Voltage".
'20200130: Created to init BinCut_Payload_Voltage.
Public Function Init_BinCut_Voltage_Array()
On Error GoTo errHandler 'Add ErrHandler 2023/08/01
    Dim i As Long
    For i = 0 To UBound(pinGroup_BinCut)
        Previous_Payload_Voltage(VddBinStr2Enum(pinGroup_BinCut(i))) = BinCut_Payload_Voltage(VddBinStr2Enum(pinGroup_BinCut(i)))
        'BinCut_Init_Voltage(VddBinStr2Enum(pinGroup_BinCut(i))) = 0
        BinCut_Payload_Voltage(VddBinStr2Enum(pinGroup_BinCut(i))) = 0
    Next i
Exit Function
errHandler:
    Call Print_Error_Message(Error_Info, "LIB_VDD_BINNING", "Init_BinCut_Voltage_Array") 'Add ErrHandler 2023/08/01
    If AbortTest Then Exit Function Else Resume Next
End Function

'20210104: Modified to replace "SiteAwareValue" with "ValuePerSite" for UltraFlex with IGXL10.
'20201215: Modified to reduce the redundant site-loop.
'20200320: Modified to use the flag "Flag_Skip_ReApplyPayloadVoltageToDCVS".
'20191210: Modified to check if powerPin belongs to selsramPin pinGroup.
Public Function Set_PayloadVoltage_to_DCVS(Enable_Rail_Switch As Boolean, powerDomainGroup() As String, voltagePayload() As SiteDouble)
On Error GoTo errHandler 'Add ErrHandler 2023/08/01
    Dim site As Variant
    Dim powerDomain As String
    Dim PinGroup As String
    Dim split_content() As String
    Dim i As Long
    Dim j As Long
    Dim hexvsPingroup As String
    Dim nonhexvsPingroup As String
    Dim PinTemp As String
    Dim anySiteSelected As Boolean
'''//==================================================================================================================================================================================//'''
'''//Note:
'''For projects with Rail-Switch, BinCut payload voltage values are applied to DCVS Valt.
'''//==================================================================================================================================================================================//'''
    For i = 0 To UBound(powerDomainGroup)
        '''//init
        hexvsPingroup = vbNullString
        nonhexvsPingroup = vbNullString
        PinTemp = vbNullString
        anySiteSelected = False
    
        '''//Get powerPins from powerDomain
        powerDomain = powerDomainGroup(i)
        PinGroup = VddbinDomain2Pin(powerDomain)
        split_content = Split(PinGroup, ",")
        
        '''//Check if any site needs to update BinCut Payload Voltages to DCVS.
        '''If BinCut_Payload_Voltage is same as Previous_Payload_Voltage, it could skip Re-applly payload voltages to DCVS.
        If Flag_Skip_ReApplyPayloadVoltageToDCVS = True Then
            If IsLevelLoadedForApplyLevelsTiming = False Then
                For Each site In TheExec.sites
                    If voltagePayload(VddBinStr2Enum(powerDomain))(site) <> Previous_Payload_Voltage(VddBinStr2Enum(powerDomain))(site) Then
                        anySiteSelected = True
                        Exit For
                    End If
                Next site
            Else
                anySiteSelected = True
            End If
        Else
            anySiteSelected = True
        End If
        
        If anySiteSelected = True Then
            '''//Assembly temporary pinGroup for HexVs
            For j = 0 To UBound(split_content)
                PinTemp = UCase(Trim(split_content(j)))
                
                'If LCase(VddbinPinDcvsType(PinTEMP)) Like "*vs*" Then '''for HexVs  20221223: Modified for UFP UVS ("VS-5A")
                If (VddbinPinDcvsType(PinTemp) = "HEXVS") Or (VddbinPinDcvsType(PinTemp) = "VHDVS") Or (VddbinPinDcvsType(PinTemp) = "VSM") Or _
                        (VddbinPinDcvsType(PinTemp) = "VS-800MA") Or (VddbinPinDcvsType(PinTemp) = "VS-5A") Then              '''for HexVs 20221228: Modified for UFP UVS ("VS-5A")
                    If hexvsPingroup = "" Then
                        hexvsPingroup = PinTemp
                    Else
                        hexvsPingroup = hexvsPingroup & "," & PinTemp
                    End If
                Else '''for non-HexVs
                    If nonhexvsPingroup = "" Then
                        nonhexvsPingroup = PinTemp
                    Else
                        nonhexvsPingroup = nonhexvsPingroup & "," & PinTemp
                    End If
                End If
            Next j
        End If
        
        '''input: voltagePayload, scale & unit: mV.
        '''DCVS applies unit:V. So that we need to convert mV into V for DCVS.
        If hexvsPingroup <> "" Then
            If Enable_Rail_Switch = True Then
                '''***********************************************************************'''
                '''//Check if powerPin belongs to selsramPin pinGroup.
                '''selsram powerPin     : Set the payload voltage to Valt.
                '''non-selsram powerPin : Set the payload voltage to Vmain.
                '''***********************************************************************'''
                '''ToDo: Maybe we can create the dictionary for selsramPin when parsing the table "SELSRM_Mapping_Table"...
                '''20210104: Modified to replace "SiteAwareValue" with "ValuePerSite" for UltraFlex with IGXL10.
                 If LCase("*," & selsramPin & ",*") Like LCase("*," & powerDomain & ",*") Then '''selsram powerPin
                    TheHdw.DCVS.pins(hexvsPingroup).Voltage.Alt.ValuePerSite = voltagePayload(VddBinStr2Enum(powerDomain)).divide(1000)
                Else
                    TheHdw.DCVS.pins(hexvsPingroup).Voltage.ValuePerSite = voltagePayload(VddBinStr2Enum(powerDomain)).divide(1000)
                End If
            Else '''project without rail-switch.
                TheHdw.DCVS.pins(hexvsPingroup).Voltage.ValuePerSite = voltagePayload(VddBinStr2Enum(powerDomain)).divide(1000)
            End If
        End If
        
        If nonhexvsPingroup <> "" Then
            If Enable_Rail_Switch = True Then
                If LCase("*," & selsramPin & ",*") Like LCase("*," & powerDomain & ",*") Then '''selsram powerPin
                    For Each site In TheExec.sites
                        TheHdw.DCVS.pins(nonhexvsPingroup).Voltage.Alt.value = voltagePayload(VddBinStr2Enum(powerDomain)) / 1000
                    Next site
                Else '''non-selsram powerPin
                    For Each site In TheExec.sites
                        TheHdw.DCVS.pins(nonhexvsPingroup).Voltage.value = voltagePayload(VddBinStr2Enum(powerDomain)) / 1000
                    Next site
                End If
            Else '''project without rail-switch.
                For Each site In TheExec.sites
                    TheHdw.DCVS.pins(nonhexvsPingroup).Voltage.value = voltagePayload(VddBinStr2Enum(powerDomain)) / 1000
                Next site
            End If
        End If
    Next i
Exit Function
errHandler:
    Call Print_Error_Message(Error_Info, "LIB_VDD_BINNING", "Set_PayloadVoltage_to_DCVS") 'Add ErrHandler 2023/08/01
    If AbortTest Then Exit Function Else Resume Next
End Function

'20200717: Modified to split specGrp and put voltagePayload into each spec of Overlay.
'20200618: Modified to use VddbinDomain2DcSpecGrp.
'20200615: Created for "Call Instance".
Public Function Set_PayloadVoltage_to_Overlay(Enable_Rail_Switch As Boolean, powerDomainGroup() As String, voltagePayload() As SiteDouble, OverlayName As String)
On Error GoTo errHandler 'Add ErrHandler 2023/08/01
    Dim site As Variant
    Dim i As Long
    Dim j As Long
    Dim powerDomain As String
    Dim specGrp As String
    Dim split_content() As String
'''//==================================================================================================================================================================================//'''
'''//Note:
'''1. Use Overlay in test instance sheet to overwrite voltages defined in DC Specs avoid HardIP/RTOS applyLevelsTiming and ForceCondition to overwrite BinCut payload voltages.
'''2. Remember to use Remove_PayloadVoltage_from_Overlay to remove Overlay after the test.
'''3. IGXL doesn't support one input with multiple specs names, so that we have to split specGrp and setup spec by spec-loop.
'''//==================================================================================================================================================================================//'''
    With TheExec.Overlays
        If (.Contains(OverlayName) <> False) Then .Remove OverlayName
        .Add (OverlayName)
    End With
    
    For i = 0 To UBound(powerDomainGroup)
        '''//init
        powerDomain = powerDomainGroup(i)
        
        '''//Check if powerDomain belongs to powerPin or pinGroup...
        If VddbinPinDict.Exists(UCase(powerDomain)) = True Then
            specGrp = VddbinDomain2DcSpecGrp(powerDomain)

            '''//Split specGrp and put payload voltage into each spec of Overlay.
            split_content = Split(specGrp, ",")
            
            For j = 0 To UBound(split_content)
                '''//If one of them exists, create the spec item.
                If specGrp <> "" Then
                    With TheExec.Overlays(OverlayName)
                        .Specs.Add (split_content(j))
                        .Specs.item(split_content(j)).value = voltagePayload(VddBinStr2Enum(powerDomain)).divide(1000)
                    End With
                End If
            Next j
        Else
            Call Print_Error_Message(Error_Info, "LIB_VDD_BINNING", "Set_PayloadVoltage_to_Overlay", "powerDomain:" & powerDomain & ", it is not BinCut powerPin or pinGroup. It is incorrect for Set_PayloadVoltage_to_Overlay. Error!!!")
            'TheExec.ErrorLogMessage "powerDomain:" & powerDomain & ", it is not BinCut powerPin or pinGroup. It is incorrect for Set_PayloadVoltage_to_Overlay. Error!!!"    ''20230807: Removed by sanity check
        End If
    Next i
Exit Function
errHandler:
    Call Print_Error_Message(Error_Info, "LIB_VDD_BINNING", "Set_PayloadVoltage_to_Overlay") 'Add ErrHandler 2023/08/01
    If AbortTest Then Exit Function Else Resume Next
End Function

'20200616: Created to remove BinCut Payload voltages from Overlay for HardIP instance..
Public Function Remove_PayloadVoltage_from_Overlay(OverlayName As String)
On Error GoTo errHandler 'Add ErrHandler 2023/08/01
    Dim strAry_OverlayName() As String
    Dim i As Double
    strAry_OverlayName = Split(OverlayName, ",")

    For i = 0 To UBound(strAry_OverlayName)
        With TheExec.Overlays
            If (.Contains(strAry_OverlayName(i)) <> False) Then
                .Remove strAry_OverlayName(i)
                TheExec.Datalog.WriteComment "Remove Overlay Name: " & strAry_OverlayName(i)
            End If
        End With
    Next i
Exit Function
errHandler:
    Call Print_Error_Message(Error_Info, "LIB_VDD_BINNING", "Remove_PayloadVoltage_from_Overlay") 'Add ErrHandler 2023/08/01
    If AbortTest Then Exit Function Else Resume Next
End Function

'20200320: Created to get instance context and assembly all information into the string.
Public Function GetInstanceContextIntoString() As String
On Error GoTo errHandler 'Add ErrHandler 2023/08/01
    Dim DCCategory As String
    Dim DCSelector As String
    Dim ACCategory As String
    Dim ACSelector As String
    Dim TimeSetSheet As String
    Dim EdgeSetSheet As String
    Dim LevelsSheet As String
    Dim Overlay As String
    '''//TheExec.DataManager.GetInstanceContext(DCCategory As String, DCSelector As String, ACCategory As String, ACSelector As String, TimeSetSheet As String, EdgeSetSheet As String, LevelsSheet As String, Overlay As String, [MemberNumber As Long = -1])
    Call TheExec.DataManager.GetInstanceContext(DCCategory, DCSelector, ACCategory, ACSelector, TimeSetSheet, EdgeSetSheet, LevelsSheet, Overlay)
    
    GetInstanceContextIntoString = DCCategory & "," & DCSelector & "," & ACCategory & "," & ACSelector & "," & TimeSetSheet & "," & LevelsSheet & "," & Overlay
Exit Function
errHandler:
    Call Print_Error_Message(Error_Info, "LIB_VDD_BINNING", "GetInstanceContextIntoString") 'Add ErrHandler 2023/08/01
    If AbortTest Then Exit Function Else Resume Next
End Function

'20200424: Created to check Adjust_Max and Adjust_Min for adjust_VddBinning.
'20191023: Modified to check if "MaxPV(pmode0/pmode1)" is in the column "Comment" of "Vdd_Binning_Def" or not.
Public Function Check_Adjust_Max_Min(Adjust_Max_Enable As Boolean, Adjust_Min_Enable As Boolean, Optional Adjust_Power_Max_list As String, Optional Adjust_Power_Min_list As String)
On Error GoTo errHandler 'Add ErrHandler 2023/08/01
    Dim site As Variant
    Dim i As Integer
    Dim j As Integer
    Dim group_array() As String
    Dim Max_array() As String
    Dim Max_value As Double
    Dim Min_array() As String
    Dim Min_value As Double
'''//==================================================================================================================================================================================//'''
'''//Example:
'''Adjust_Power_Max_list = "P_mode1,P_mode2 + P_mode3,P_mode4" => P_mode1 and P_mode2 will fuse the max value1. P_mode3 and P_mode4 will fuse the max value2.
'''ex: MC607=500mv, MC608=550mv. If both p_modes are set as "allowequal", MC607 will be adjusted to 550mv. Then MC607 and MC608 will be 550mv.
'''//==================================================================================================================================================================================//'''
    If (Adjust_Max_Enable <> Flag_Adjust_Max_Enable) Or (Adjust_Min_Enable <> Flag_Adjust_Min_Enable) Then
        Call Print_Error_Message(Error_Info, "LIB_VDD_BINNING", "", "The instance arguments Adjust_Max_Enable or Adjust_Min_Enable of adjust_VddBinning might be inconsistent with column Comment of Vdd_Binning_Def_appA. Error!!!")
        'TheExec.ErrorLogMessage "The instance arguments Adjust_Max_Enable or Adjust_Min_Enable of adjust_VddBinning might be inconsistent with column Comment of Vdd_Binning_Def_appA. Error!!!"    ''20230807: Removed by sanity check
    Else
        '''//Check if "MaxPV(pmode0/pmode1)" is in the column "Comment" of "Vdd_Binning_Def" or not.
        If Adjust_Max_Enable = True Then
            If Adjust_Power_Max_list <> Adjust_Power_Max_pmode Then
                Call Print_Error_Message(Error_Info, "LIB_VDD_BINNING", "Check_Adjust_Max_Min", "Argument Adjust_Power_Max_list: " & Adjust_Power_Max_list & " of adjust_VddBinning is inconsistent with  " & Adjust_Power_Max_pmode & " from column Comment of Vdd_Binning_Def_appA. Error!!!")
                'TheExec.ErrorLogMessage "Argument Adjust_Power_Max_list: " & Adjust_Power_Max_list & " of adjust_VddBinning is inconsistent with  " & Adjust_Power_Max_pmode & " from column Comment of Vdd_Binning_Def_appA. Error!!!"    ''20230807: Removed by sanity check
            End If
        End If
        
        If Adjust_Min_Enable = True Then
            If Adjust_Power_Min_list <> Adjust_Power_Min_pmode Then
                Call Print_Error_Message(Error_Info, "LIB_VDD_BINNING", "Check_Adjust_Max_Min", "Argument Adjust_Power_Min_list: " & Adjust_Power_Min_list & " of adjust_VddBinning is inconsistent with  " & Adjust_Power_Min_pmode & " from column Comment of Vdd_Binning_Def_appA. Error!!!")
                'TheExec.ErrorLogMessage "Argument Adjust_Power_Min_list: " & Adjust_Power_Min_list & " of adjust_VddBinning is inconsistent with  " & Adjust_Power_Min_pmode & " from column Comment of Vdd_Binning_Def_appA. Error!!!"    ''20230807: Removed by sanity check
            End If
        End If
    End If
    
    '''//Adjust_Max_Enable is related to "Allowequal"
    If Adjust_Max_Enable = True Then
        group_array = Split(Adjust_Power_Max_list, "+")
        
        For j = 0 To UBound(group_array)
            Max_array = Split(group_array(j), ",")
            
            If UBound(Max_array) < 1 Then
                TheExec.Datalog.WriteComment "Adjust_Power_Max pin list of adjust_VddBinning is less than 2"
                'TheExec.ErrorLogMessage "Adjust_Power_Max pin list of adjust_VddBinning is less than 2"
            End If
            
            For Each site In TheExec.sites
                Max_value = 0
                For i = 0 To UBound(Max_array)
                    If AllBinCut(VddBinStr2Enum(Max_array(i))).Used = True Then
                        If CDec(Max_value) < CDec(VBIN_RESULT(VddBinStr2Enum(Max_array(i))).GRADEVDD) Then
                             Max_value = VBIN_RESULT(VddBinStr2Enum(Max_array(i))).GRADEVDD
                        End If
                    Else
                        TheExec.Datalog.WriteComment "The Performance Mode " & Max_array(i) & " is not used for adjust_VddBinning!!!"
                    End If
                Next i
                For i = 0 To UBound(Max_array)
                    If CDec(VBIN_RESULT(VddBinStr2Enum(Max_array(i))).GRADEVDD) <> CDec(Max_value) Then
                        VBIN_RESULT(VddBinStr2Enum(Max_array(i))).GRADEVDD = Max_value
                        TheExec.Datalog.WriteComment "site:" & site & "," & Max_array(i) & ",BinCut Product voltage is adjusted to Max Value: " & Max_value
                    End If
                Next i
            Next site
        Next j
    End If
    
    '''//Adjust_Min_Enable is related to "Allowequal".
    '''ex: MC607=500mv, MC608=550mv. If both p_modes are set as "allowequal", MC608 will be adjusted to 500mv. Then MC607 and MC608 will be 500mv.
    If Adjust_Min_Enable = True Then
        group_array = Split(Adjust_Power_Min_list, "+")
        
        For j = 0 To UBound(group_array)
            Min_array = Split(group_array(j), ",")
            
            If UBound(Min_array) < 1 Then
                TheExec.ErrorLogMessage "Adjust_Power_Min pin list of adjust_VddBinning is less than 2"
            End If
            
            For Each site In TheExec.sites
                Min_value = 9999
                For i = 0 To UBound(Min_array)
                    If AllBinCut(VddBinStr2Enum(Min_array(i))).Used = True Then
                        If CDec(Min_value) > CDec(VBIN_RESULT(VddBinStr2Enum(Min_array(i))).GRADEVDD) Then
                             Min_value = VBIN_RESULT(VddBinStr2Enum(Min_array(i))).GRADEVDD
                        End If
                    Else
                        Call Print_Error_Message(Error_Info, "LIB_VDD_BINNING", "Check_Adjust_Max_Min", "The Performance Mode " & Min_array(i) & " Doesn't Exist. Error!!!")
                        'TheExec.ErrorLogMessage "The Performance Mode " & Min_array(i) & " Doesn't Exist. Error!!!"    ''20230807: Removed by sanity check
                    End If
                Next i
                For i = 0 To UBound(Min_array)
                    If VBIN_RESULT(VddBinStr2Enum(Min_array(i))).GRADEVDD <> Min_value Then
                        VBIN_RESULT(VddBinStr2Enum(Min_array(i))).GRADEVDD = Min_value
                        Call Print_Error_Message(Error_Info, "LIB_VDD_BINNING", "Check_Adjust_Max_Min", "site:" & site & "," & Min_array(i) & ",BinCut Product voltage is adjusted to Min Value: " & Min_value)
                        'TheExec.Datalog.WriteComment "site:" & site & "," & Min_array(i) & ",BinCut Product voltage is adjusted to Min Value: " & Min_value
                    End If
                Next i
            Next site
        Next j
    End If
Exit Function
errHandler:
    Call Print_Error_Message(Error_Info, "LIB_VDD_BINNING", "Check_Adjust_Max_Min") 'Add ErrHandler 2023/08/01
    If AbortTest Then Exit Function Else Resume Next
End Function

'20210302: Modified to remove the If condition for DevChar.
'20210126: Modified to revise the vbt code for DevChar.
'20201111: Modified to replace the vbt function name "print_bincut_power" with "print_bincut_voltage".
'20201027: Modified to use "Public Type Instance_Info".
'20200924: Modified to remove the redundant argument "IndexLevelPerSite As SiteLong" from Set_BinCut_Initial_by_ApplyLevelsTiming.
'20200921: Discussed "RTOS_bootup_relay" / "KeepAliveFlag" / "spi_ttr_flag" with SWLINZA and PCLINZG. We decided to remove these SPI/RTOS branches because RTOS didn't use pattern test since Cebu/Sicily/Tonga/JC-Chop/Ellis/Bora.
'20200425: Modified to adjust the flow for "print_bincut_power".
'20200424: Modified to use "Set_BinCut_Initial_by_ApplyLevelsTiming" to set BinCut initial voltage by ApplyLevelsTiming.
'20200324: Modified to skip ApplyLevelsTiming when current instance has the same level/timing as previous instance for project with rail-switch.
'20200320: Modified to check instance contexts of current instance and previous instance.
'20200206: Modified to replace "print_main_power_init" with "print_bincut_power".
Public Function Set_BinCut_Initial_by_ApplyLevelsTiming(inst_info As Instance_Info)
On Error GoTo errHandler
    Dim DCCategory As String
    Dim DCSelector As String
    Dim ACCategory As String
    Dim ACSelector As String
    Dim TimeSetSheet As String
    Dim EdgeSetSheet As String
    Dim LevelsSheet As String
    Dim Overlay As String
'''//==================================================================================================================================================================================//'''
'''//Note:
'''Set initial voltages from category "Bincut_X_X_X" in DC_Specs sheet by ApplyLevelsTiming.
'''Print the initial voltages, and applies them to DCVS Vmain and Valt by ApplyLevelsTiming (DCVS voltage source will be switched to Vmain).
'''Skip ApplyLevelsTiming when current instance has the same level/timing as previous instance for project with rail-switch.
'''//==================================================================================================================================================================================//'''
    If Flag_Skip_ReApplyPayloadVoltageToDCVS = True Then
        CurrentBinCutInstanceContext = GetInstanceContextIntoString
        
        If CurrentBinCutInstanceContext = PreviousBinCutInstanceContext And Flag_Enable_Rail_Switch = True Then
            select_DCVS_output_for_powerDomain tlDCVSVoltageMain
            inst_info.currentDcvsOutput = tlDCVSVoltageMain
            IsLevelLoadedForApplyLevelsTiming = False
        Else
            Call TheHdw.Digital.ApplyLevelsTiming(True, True, True, tlPowered)
            inst_info.currentDcvsOutput = tlDCVSVoltageMain
            IsLevelLoadedForApplyLevelsTiming = True
        End If
    ElseIf False And (TheExec.enableWord("T0TX_ROOM") = True Or TheExec.enableWord("T0TX_HOT") = True) Then 'WWWWW
        Call TheExec.DataManager.GetInstanceContext(DCCategory, DCSelector, ACCategory, ACSelector, TimeSetSheet, EdgeSetSheet, LevelsSheet, Overlay)

        If (DC_Mapping(inst_info.p_mode, 1).Addtional_LV_DC_Value(inst_info.addi_mode) Like LCase("*lv") And inst_info.addi_mode <> 0) Or _
            (DC_Mapping(inst_info.p_mode, 1).LV_DC_Value Like LCase("*lv") And inst_info.addi_mode = 0) Or _
            (DC_Mapping(inst_info.p_mode, 1).Addtional_HVCC_DC_Value(inst_info.addi_mode) Like LCase("*lv") And inst_info.addi_mode <> 0) Or _
            (DC_Mapping(inst_info.p_mode, 1).HVCC_DC_Value Like LCase("*lv") And inst_info.addi_mode = 0) Then
            Call TheHdw.Digital.ApplyLevelsTiming(True, True, True, tlPowered, , , , LevelsSheet, "Bincut_X_X_X", "min", TimeSetSheet, ACCategory, ACSelector, EdgeSetSheet)
        
        ElseIf (DC_Mapping(inst_info.p_mode, 1).Addtional_LV_DC_Value(inst_info.addi_mode) Like LCase("*nv") And inst_info.addi_mode <> 0) Or _
            (DC_Mapping(inst_info.p_mode, 1).LV_DC_Value Like LCase("*nv") And inst_info.addi_mode = 0) Or _
            (DC_Mapping(inst_info.p_mode, 1).Addtional_HVCC_DC_Value(inst_info.addi_mode) Like LCase("*nv") And inst_info.addi_mode <> 0) Or _
            (DC_Mapping(inst_info.p_mode, 1).HVCC_DC_Value Like LCase("*nv") And inst_info.addi_mode = 0) Then
            Call TheHdw.Digital.ApplyLevelsTiming(True, True, True, tlPowered, , , , LevelsSheet, "Bincut_X_X_X", "typ", TimeSetSheet, ACCategory, ACSelector, EdgeSetSheet)
        
        Else
            Call TheHdw.Digital.ApplyLevelsTiming(True, True, True, tlPowered, , , , LevelsSheet, "Bincut_X_X_X", "max", TimeSetSheet, ACCategory, ACSelector, EdgeSetSheet)
        
        End If
    Else
        Call TheHdw.Digital.ApplyLevelsTiming(True, True, True, tlPowered)
        inst_info.currentDcvsOutput = tlDCVSVoltageMain
        IsLevelLoadedForApplyLevelsTiming = True
    End If
    
    '''//Print initial voltages
    '''Ex: Initial_Voltage_VDD_SOC_MS003,0,VDD_PCPU=0.752,VDD_ECPU=0.752,VDD_GPU=0.752,VDD_SOC=0.752, ...
    print_bincut_voltage inst_info, , Flag_Remove_Printing_BV_voltages, False, BincutVoltageType.InitialVoltage
Exit Function
errHandler:
    Call Print_Error_Message(Error_Info, "LIB_VDD_BINNING", "Set_BinCut_Initial_by_ApplyLevelsTiming") 'Add ErrHandler 2023/08/01
    If AbortTest Then Exit Function Else Resume Next
End Function

'20200506: Created to decide testType for string.
Public Function decide_test_type_for_string(strInput As String) As testType
On Error GoTo errHandler 'Add ErrHandler 2023/08/01
    Dim strTemp As String
    strTemp = LCase(strInput)
    
    '''//Please check "Enum TestType" and "MaxTestType" in GlobalVariable.
    If strTemp Like "*td*" Then
        decide_test_type_for_string = testType.TD
    ElseIf strTemp Like "*bist*" Then
        decide_test_type_for_string = testType.Mbist
    ElseIf strTemp Like "*tmps*" Then
        decide_test_type_for_string = testType.TMPS
    ElseIf strTemp Like "*spi*" Then
        decide_test_type_for_string = testType.SPI
    ElseIf strTemp Like "*rtos*" Then
        decide_test_type_for_string = testType.RTOS
    ElseIf strTemp Like "*ldcbfd*" Then
        decide_test_type_for_string = testType.ldcbfd
    Else
        Call Print_Error_Message(Error_Info, "LIB_VDD_BINNING", "decide_test_type_for_string", "decide_test_type_for_string can't decide testType for input:" & strInput & ". Error!!!")
        'TheExec.ErrorLogMessage "decide_test_type_for_string can't decide testType for input:" & strInput & ". Error!!!"    ''20230807: Removed by sanity check
    End If
Exit Function
errHandler:
    Call Print_Error_Message(Error_Info, "LIB_VDD_BINNING", "decide_test_type_for_string") 'Add ErrHandler 2023/08/01
    If AbortTest Then Exit Function Else Resume Next
End Function

'20210903: Modified to merge properties of "Public Type Instance_Step_Control" into Public Type Instance_Info "Public Type Instance_Info".
'20210901: Modified to rename "StepCount As Long" as "count_Step As New SiteLong" for Public Type Instance_Info.
'20210901: Modified to move "Step_GradeFound As New SiteLong" from Public Type Instance_Step_Control to the vbt function Update_VBinResult_by_Step.
'20210901: Modified to rename "IndexFoundLevel As New SiteLong" as "Step_GradeFound As New SiteLong" for Public Type Instance_Step_Control.
'20210901: Modified to rename "IndexLevelPerSite As New SiteLong" as "Step_Current As New SiteLong" for Public Type Instance_Info.
'20210813: Modified to use Set_VBinResult_by_Step for updating PassBin, Pass step, and voltage to VBIN_Result.
'20210810: Modified to merge the vbt function Check_anySite_GradeFound into the vbt function Update_VBinResult_by_Step.
'20210810: Modified to use gotPassStep as flag to determine if update VBIN_Result or not.
'20210809: Modified to remove the redundant property "FoundLevel As New SiteDouble" from Public Type Instance_Step_Control.
'20210809: Modified to revise the branches of the vbt code for Linear algorithm and IDS algorithm of BinCut gradesearch.
'20210809: Modified to revise the vbt code with VBIN_RESULT(inst_info.p_mode).step_in_IDS_Zone(site).
'20210806: Modified to update VBIN_RESULT(inst_info.p_mode).GRADEVDD = DYNAMIC_VBIN_IDS_ZONE(inst_info.p_mode).Product_Voltage(step_control.IndexFoundLevel(site)) for IDS mode.
'20210722: Modified to use VBIN_IDS_ZONE(p_mode).Product_Voltage(Max_IDS_Zone, Max_IDS_Step) and DYNAMIC_VBIN_IDS_ZONE(p_mode).Product_Voltage(Max_IDS_Step) for GradeVDD.
'20210720: Modified to revise the vbt function Update_VBinResult_by_Step for BinCut search in FT.
'20210420: Modified to remove the unused vbt code of GradeSearchMethod. Once if C651 provided any new definition of GradeSearchMethod, we will revise the vbt code.
'20210302: Modified to use "step_control.FoundLevel(site) = DYNAMIC_VBIN_IDS_ZONE(inst_info.p_mode).Voltage(inst_info.IndexLevelPerSite(site))".
'20210226: Modified to use step_Start and step_Stop to get startVoltage and StopVoltage.
'20210125: Modified to remove "voltage_Pmode_EQNbased As SiteDouble" from the arguments of the vbt function "Update_VBinResult_by_Step".
'20210120: Modified to use VBIN_RESULT(p_mode).step_1stPass_in_IDS_Zone to store the first pass step in Dynamic IDS Zone and find the correspondent PassBinCut number.
'20201209: Modified to use the arguments "inst_info As Instance_Info" and "step_control As Instance_Step_Control" for Update_VBinResult_by_Step.
'20201207: Modified to add the argument "flag_All_Patt_Pass".
'20200602: Modified to remove the condition "IndexLevelPerSite(site) = 0" for "gradeAlg = GradeSearchAlgorithm.ids".
'20200317: Modified for SearchByPmode.
'20190313: Modified to add "FIRSTPASSBINCUT(p_mode)" for storing the first passbinnum of P_mode.
Public Function Update_VBinResult_by_Step(inst_info As Instance_Info)
On Error GoTo errHandler 'Add ErrHandler 2023/08/01
    Dim site As Variant
    Dim gotPassStep As Boolean
    Dim step_GradeFound As New SiteLong
'''//==================================================================================================================================================================================//'''
'''//Note:
'''//For step control:
'''count_Step: step index for DYNAMIC_VBIN_IDS_ZONE(p_mode).
'''Step_Current: step has beed tested.
'''//==================================================================================================================================================================================//'''
    '''init
    inst_info.Grade_Not_Found_Mask = 0
    inst_info.On_StopVoltage_Mask = 0
    inst_info.Grade_Found_Mask = 0

    For Each site In TheExec.sites
        '''init
        '''20210810: Modified to use gotPassStep as flag to determine if update VBIN_Result or not.
        gotPassStep = False
        
        '''//If IDS mode fails at the first step, it has to switch BinCut search algorithm from IDS mode to Linear mode immediately.
        If inst_info.gradeAlg(site) = GradeSearchAlgorithm.ids And inst_info.All_Patt_Pass(site) = False And inst_info.count_Step = 0 Then
            '''*********************************************************************************************************************************************************************************'''
            '''//If the algorithm is IDS mode and this site is failed in first step, we will change the algorithm to linear mode and change the direction for stopvoltage.
            '''
            '''  Site0   Site1   Site2     => if the algo = IDS, means that there is lower EQ number can be tested and the IDS start EQ is must not last EQ number.
            '''  EQ2(F)  EQ2(P)  EQ2(P)    => All sites start in EQ2, and only site 0 fail in first step, the site 0 will change to linear mode.
            '''  EQ1(P)  EQ3(P)  EQ3(F)    => only site 0 run the different direction , Although site 2 is also failed, it is not failed in first step. We do not need to change the algorithm.
            '''  EQ1(P)  EQ4(P)  EQ4(F)
            '''
            '''  Grade   Grade   Grade
            '''  EQ1     EQ4     EQ2
            '''*********************************************************************************************************************************************************************************'''
            inst_info.gradeAlg(site) = GradeSearchAlgorithm.linear
            inst_info.step_Stop(site) = DYNAMIC_VBIN_IDS_ZONE(inst_info.p_mode).Max_Step - 1
        End If
        
        '''//grade found at linear algorithm.
        If (inst_info.gradeAlg = GradeSearchAlgorithm.linear And inst_info.All_Patt_Pass(site) = True) And inst_info.grade_found(site) = False _
        Or ((inst_info.gradeAlg = GradeSearchAlgorithm.ids And inst_info.All_Patt_Pass(site) = True) And inst_info.grade_found(site) = False And inst_info.Step_Current(site) = inst_info.step_Stop(site)) Then
            '''*********************************************************************************************************************************************************************************'''
            '''//When the patterns are passed in this step, there are some conditions will be identified to the grade had been found.
            ''' 1. Algorithm = IDS, And the grade was not been found in this test Instance. But the voltage of the step reach the stopvoltage (pattern never failed in this site).
            ''' 2. Algorithm = Linear, And the grade was not been found in this test Instance.
            '''
            '''  Linear                                                      IDS
            '''  site0            site1           site2                      site0   site1            site2
            '''  EQ4(F)           EQ4(F)          EQ4(P)                     EQ3(F)  EQ3(P)           EQ3(P)
            '''  EQ3(F)           EQ3(P) => found EQ3(P)                     EQ2(F)  EQ4(P) => found  EQ4(F)
            '''  EQ2(P) => found  EQ2(P)          EQ2(P)                     EQ1(P)  EQ4(P)           EQ4(F)
            '''
            '''  Grade            Grade           Grade                      Grade   Grade            Grade
            '''  EQ2              EQ3             EQ4                        EQ1     EQ4              EQ3
            '''*********************************************************************************************************************************************************************************'''
            '''//Determine the pass Step in Dynamic_IDS_zone of the p_mode.
            step_GradeFound(site) = inst_info.Step_Current(site)
            gotPassStep = True
        End If
        
        '''//grade found at IDS algorithm.
        If ((inst_info.gradeAlg = GradeSearchAlgorithm.ids And inst_info.All_Patt_Pass(site) = False) And inst_info.grade_found(site) = False) Then
            '''*********************************************************************************************************************************************************************************'''
            '''//When the patterns are failed in this step, there are some conditions will be identified to the grade had been found.
            ''' 1. Algorithm = IDS, And the grade was not been found in this test Instance.
            '''
            ''' IDS
            ''' site0   site1   site2
            ''' EQ3(F)  EQ3(P)  EQ3(P)
            ''' EQ2(F)  EQ4(P)  EQ4(F)  => found
            ''' EQ1(P)  EQ4(P)  EQ4(F)
            '''
            ''' Grade   Grade   Grade
            ''' EQ1     EQ4     EQ3
            '''*********************************************************************************************************************************************************************************'''
            '''//Determine the pass Step in Dynamic_IDS_zone of the p_mode.
            step_GradeFound(site) = inst_info.Step_Current(site) + 1
            gotPassStep = True
        End If
        
        '''//If the pass step in Dynamic_IDS_zone of p_mode is found, update VBIN_Result for p_mode.
        '''20210809: Modified to revise the branches of the vbt code for Linear algorithm and IDS algorithm of BinCut gradesearch.
        '''20210813: Modified to use Set_VBinResult_by_Step for updating PassBin, Pass step, and voltage to VBIN_Result.
        If gotPassStep = True Then
            '''//Update PassBin, Pass step, flag"VBIN_Result(p_mode).tested", and voltage to VBIN_Result by the step in Dynamic_IDS_Zone.
            Call Set_VBinResult_by_Step(site, inst_info.p_mode, step_GradeFound(site))
            
            '''//Update the flag about grade_found.
            If inst_info.All_Site_Mask And 2 ^ site = 0 Then '20240313 prevent each site BIN1 E1 harvest, but will run with another site due to another site common fail on BIN1 E1
            Else
                inst_info.grade_found(site) = True
            End If
            inst_info.AnySiteGradeFound = True
            
            '''//Update PassBin to the globalVariable CurrentPassBinCutNum.
            CurrentPassBinCutNum = VBIN_RESULT(inst_info.p_mode).passBinCut
        End If
        
        '''//Check if any site passes or fails on the current step.
        '''20210810: Modified to merge the vbt function Check_anySite_GradeFound into the vbt function Update_VBinResult_by_Step.
        '''*********************************************************************************************************************************************************************************'''
        '''//For all sites, we need to record some flags to identify if need to exit the loop when there is no reason to seach.
        '''      On_StopVoltage_Mask => record how many sites had reached the stopvoltage.
        '''      Grade_Found_Mask    => record how many sites had reached the stopvoltage.
        '''*********************************************************************************************************************************************************************************'''
        If inst_info.grade_found(site) = False Then
            inst_info.Grade_Not_Found_Mask = inst_info.Grade_Not_Found_Mask + 2 ^ site
            
            If inst_info.Step_Current(site) = inst_info.step_Stop(site) Then
                inst_info.On_StopVoltage_Mask = inst_info.On_StopVoltage_Mask + 2 ^ site
            End If
            
            If inst_info.HarvestBinningFlag <> "" And inst_info.Harvest_With_BIN1E1 = True And TheExec.sites.item(site).FlagState(strGlb_Flag_HarvestBinningFlag_SingleFSTP) = logicTrue Then
                inst_info.Grade_Found_Mask = inst_info.Grade_Found_Mask Or inst_info.Grade_Not_Found_Mask '20240403 just run BIN1 EQN1 on SingleFSTP
            End If
        Else
            inst_info.Grade_Found_Mask = inst_info.Grade_Found_Mask + 2 ^ site
        End If
    Next site
Exit Function
errHandler:
    Call Print_Error_Message(Error_Info, "LIB_VDD_BINNING", "Update_VBinResult_by_Step") 'Add ErrHandler 2023/08/01
    If AbortTest Then Exit Function Else Resume Next
End Function

'20210813: Created to set VBIN_Result by the step.
Public Function Set_VBinResult_by_Step(site As Variant, p_mode As Integer, lng_step_selected As Long)
On Error GoTo errHandler
    If DYNAMIC_VBIN_IDS_ZONE(p_mode).Used(site) = True Then
        '''//Update the pass step in Dynamic_IDS_zone of p_mode to VBIN_Result for p_mode.
        VBIN_RESULT(p_mode).step_in_IDS_Zone = lng_step_selected
        
        '''//Store the first pass step in Dynamic IDS Zone and find the correspondent PassBinCut number.
        VBIN_RESULT(p_mode).step_1stPass_in_IDS_Zone = VBIN_RESULT(p_mode).step_in_IDS_Zone
        
        '''//step_in_BinCut = EQN-1.
        VBIN_RESULT(p_mode).step_in_BinCut = DYNAMIC_VBIN_IDS_ZONE(p_mode).EQ_Num(VBIN_RESULT(p_mode).step_in_IDS_Zone(site)) - 1
        
        '''//Get the current BinCut number.
        VBIN_RESULT(p_mode).passBinCut = DYNAMIC_VBIN_IDS_ZONE(p_mode).passBinCut(VBIN_RESULT(p_mode).step_in_IDS_Zone(site))
        
        '''//Update BinCut voltage(Grade) of p_mode according to the pass Step in Dynamic_IDS_zone.
        '''20210809: Modified to remove the redundant property "FoundLevel As New SiteDouble" from Public Type Instance_Step_Control.
        VBIN_RESULT(p_mode).GRADE = DYNAMIC_VBIN_IDS_ZONE(p_mode).Voltage(VBIN_RESULT(p_mode).step_in_IDS_Zone(site))
        
        '''//Update Efuse product voltage(GradeVDD) of p_mode. => Efuse product voltage(GradeVDD) = BinCut voltage(Grade) + Guardband.
        VBIN_RESULT(p_mode).GRADEVDD = DYNAMIC_VBIN_IDS_ZONE(p_mode).Product_Voltage(VBIN_RESULT(p_mode).step_in_IDS_Zone(site))
        
        VBIN_RESULT(p_mode).tested = True
        
        If Flag_BinX_Info_Parsed = True And TheExec.enableWord("VddBinning_Pre_E1") = False Then
            '20240201 check if EQN is BinX
            If Pmode2BinX_EQN_Dict.Exists(p_mode) And Check_EQN_UpToBinX(p_mode, site) = True Then
                Bincut_Domain2BinX(VddBinStr2Enum(AllBinCut(p_mode).powerPin))(site) = True
            End If
        End If
    Else
        Call Print_Error_Message(Error_Info, "LIB_VDD_BINNING", "Set_VBinResult_by_Step", "site:" & site & "," & VddBinName(p_mode) & ", it doesn't have any correct Dynamic_IDS_zone for Set_VBinResult_by_Step. Error!!!")
        'TheExec.ErrorLogMessage "site:" & site & "," & VddBinName(p_mode) & ", it doesn't have any correct Dynamic_IDS_zone for Set_VBinResult_by_Step. Error!!!"    ''20230807: Removed by sanity check
    End If
Exit Function
errHandler:
    Call Print_Error_Message(Error_Info, "LIB_VDD_BINNING", "Set_VBinResult_by_Step") 'Add ErrHandler 2023/08/01
    If AbortTest Then Exit Function Else Resume Next
End Function

'20210910: Modified to revise the format of FirstChangeBinInfo, as requested by C651 Si and TSMC ZYLINI.
'20210909: Modified to revise the vbt code for FirstChangeBinInfo requested by C651 Si, as discussed with TSMC ZYLINI and ZQLIN.
'20210903: Modified to merge properties of "Public Type Instance_Step_Control" into Public Type Instance_Info "Public Type Instance_Info".
'20210901: Modified to rename "IndexLevelPerSite As New SiteLong" as "Step_Current As New SiteLong" for Public Type Instance_Info.
'20210812: Modified to rename the property "step_lowest As New SiteLong" as "step_inherit As New SiteLong".
'20210810: Modified to add the property "step_Lowest As New SiteLong" to Public Type DYNAMIC_VBIN_IDS_ZONE.
'20210806: Modified to remove the redundant property "IndexLevelIncDec As New SiteLong" from Public Type Instance_Step_Control.
'20210629: Modified to check the EnableWord("Vddbin_PTE_Debug")=False to determine "Print Bincut Fail Info", as suggested by Chihome.
'20210420: Modified to remove the unused vbt code of GradeSearchMethod. Once if C651 provided any new definition of GradeSearchMethod, we will revise the vbt code.
'20210107: Modified to add for recording first changed binnum mode data, requested by C651 Si.
'20201209: Modified to use the arguments "inst_info As Instance_Info" and "step_control As Instance_Step_Control" for Decide_NextStep_for_GradeSearch.
'20200508: Created to decide Next Step for GradeSearch.
'20200317: Modified for SearchByPmode.
Public Function Decide_NextStep_for_GradeSearch(inst_info As Instance_Info)
On Error GoTo errHandler 'Add ErrHandler 2023/08/01
    Dim site As Variant
    Dim sitelng_next_step As New SiteLong
    Dim EQN_UpToBINX As New SiteBoolean: EQN_UpToBINX = False
'''//==================================================================================================================================================================================//'''
'''//Note:
'''1. IndexLevelPerSite: step in Dynamic_IDS_Zone has been tested.
'''2. Define the direction of each site in the IDS zone:
'''     BinCut Search Algorithm = IDS => move to small step (step3 -> step2 -> step1 -> step0 -> step0 -> step0)
'''     BinCut Search Algorithm = Linear => move to large step (step1 -> step2 -> step3 -> step3 -> step3)
'''//==================================================================================================================================================================================//'''
    '''On_StopVoltage_Mask = False
    For Each site In TheExec.sites 'decide next StepCount
        If inst_info.gradeAlg(site) = GradeSearchAlgorithm.ids Then
            sitelng_next_step(site) = -1
        ElseIf inst_info.gradeAlg(site) = GradeSearchAlgorithm.linear Then
            sitelng_next_step(site) = 1
        End If
        
        inst_info.Step_Current(site) = inst_info.Step_Current(site) + sitelng_next_step(site)
        
        '''//Check if the next step is within the correct steps of p_mode.
        '''step_inherit is the step with the lowest BinCut voltage in Dynamic_IDS_zone.
        If inst_info.Step_Current(site) > DYNAMIC_VBIN_IDS_ZONE(inst_info.p_mode).Max_Step - 1 Then
            inst_info.Step_Current(site) = DYNAMIC_VBIN_IDS_ZONE(inst_info.p_mode).Max_Step - 1
        ElseIf inst_info.Step_Current(site) < DYNAMIC_VBIN_IDS_ZONE(inst_info.p_mode).step_inherit Then
            inst_info.Step_Current(site) = DYNAMIC_VBIN_IDS_ZONE(inst_info.p_mode).step_inherit
        End If
        
        '20240129 Becuase Bin1 BinX merge, passBinCut won't > 1
        If inst_info.grade_found = False And Flag_BinX_Info_Parsed = True And FirstChangeBinInfo.FirstChangeBinMode(site) = 999 And Flag_Fuse_BinX(site) = False Then
            If Pmode2BinX_EQN_Dict.Exists(inst_info.p_mode) Then
                Dim binX_step_ary As Variant, binX_step As Variant, binX_step_str As Variant
                binX_step_str = Pmode2BinX_EQN_Dict(inst_info.p_mode)     ''Get EQN number from sheet
                binX_step_ary = Split(binX_step_str, ",")       ''Split EQN number by ','
                For Each binX_step In binX_step_ary
                    If DYNAMIC_VBIN_IDS_ZONE(inst_info.p_mode).EQ_Num(inst_info.Step_Current) = binX_step And CurrentPassBinCutNum(site) = 1 Then
                        EQN_UpToBINX(site) = True
                        Exit For
                    End If
                Next binX_step
            End If
            If TheExec.enableWord("Vddbin_PTE_Debug") = False And EQN_UpToBINX(site) = True Then
                FirstChangeBinInfo.FirstChangeBinMode(site) = inst_info.p_mode
                FirstChangeBinInfo.FirstChangeBinType(site) = inst_info.offsetTestTypeIdx
                '''20210910: Modified to revise the format of FirstChangeBinInfo, as requested by C651 Si and TSMC ZYLINI.
                FirstChangeBinInfo.str_Pmode_Test(site) = inst_info.inst_name
            End If
        ElseIf inst_info.grade_found = False And DYNAMIC_VBIN_IDS_ZONE(inst_info.p_mode).passBinCut(inst_info.Step_Current) <> CurrentPassBinCutNum Then
            '''********************************************************************************'''
            '''20210107: Modified to add for recording first changed binnum mode data, requested by C651 Si.
            '''20210629: Modified to check the EnableWord("Vddbin_PTE_Debug")=False to determine "Print Bincut Fail Info", as suggested by Chihome.
            '''20210909: Modified to revise the vbt code for FirstChangeBinInfo requested by C651 Si, as discussed with TSMC ZYLINI and ZQLIN.
            If TheExec.enableWord("Vddbin_PTE_Debug") = False Then
                If DYNAMIC_VBIN_IDS_ZONE(inst_info.p_mode).passBinCut(inst_info.Step_Current) > CurrentPassBinCutNum Then
                    FirstChangeBinInfo.FirstChangeBinMode(site) = inst_info.p_mode
                    FirstChangeBinInfo.FirstChangeBinType(site) = inst_info.offsetTestTypeIdx
                    '''20210910: Modified to revise the format of FirstChangeBinInfo, as requested by C651 Si and TSMC ZYLINI.
                    FirstChangeBinInfo.str_Pmode_Test(site) = inst_info.inst_name
                End If
            End If
            '''********************************************************************************'''
            CurrentPassBinCutNum = DYNAMIC_VBIN_IDS_ZONE(inst_info.p_mode).passBinCut(inst_info.Step_Current)
        End If
    Next site
Exit Function
errHandler:
    Call Print_Error_Message(Error_Info, "LIB_VDD_BINNING", "Decide_NextStep_for_GradeSearch") 'Add ErrHandler 2023/08/01
    If AbortTest Then Exit Function Else Resume Next
End Function

'20210903: Modified to merge properties of "Public Type Instance_Step_Control" into Public Type Instance_Info "Public Type Instance_Info".
'20210901: Modified to rename "IndexLevelPerSite As New SiteLong" as "Step_Current As New SiteLong" for Public Type Instance_Info.
'20210830: Modified to merge the branches of the vbt function Update_PassBinCut_for_GradeNotFound.
'20210813: Modified to move "VBIN_RESULT(inst_info.p_mode).tested" from the vbt function Update_PassBinCut_for_GradeNotFound to the vbt function Set_VBinResult_by_Step.
'20210420: Modified to remove the unused vbt code of GradeSearchMethod. Once if C651 provided any new definition of GradeSearchMethod, we will revise the vbt code.
'20201209: Modified to use the arguments "inst_info As Instance_Info" and "step_control As Instance_Step_Control" for Update_PassBinCut_for_GradeNotFound.
'20200511: Created to update CurrentPassBinCutNum for DUT "grade_found=false".
'20200317: Modified for SearchByPmode.
Public Function Update_PassBinCut_for_GradeNotFound(inst_info As Instance_Info)
On Error GoTo errHandler 'Add ErrHandler 2023/08/01
    Dim site As Variant
    Dim HarvCnt As Long, HarvPatCnt As Long
    For Each site In TheExec.sites
        If inst_info.grade_found(site) = False Then
            '''20220125,judge bin1 eq1 harv result. if only harv fail use bin1eq1 result and turn on harv flag
            '''20230926: Added to save non-harvest pattern result in harvest instance
            If inst_info.enable_HarvPinGrp = True And inst_info.HarvStartWithBin1(site) = True Then
                If inst_info.HarvPinGrpCommnFailResult(site) = True Or inst_info.HarvPinGrp_NonHarvFail(site) = True Then
                    VBIN_RESULT(inst_info.p_mode).step_in_IDS_Zone = inst_info.Step_Current
                    VBIN_RESULT(inst_info.p_mode).step_in_BinCut = DYNAMIC_VBIN_IDS_ZONE(inst_info.p_mode).EQ_Num(VBIN_RESULT(inst_info.p_mode).step_in_IDS_Zone) - 1
                    VBIN_RESULT(inst_info.p_mode).passBinCut = DYNAMIC_VBIN_IDS_ZONE(inst_info.p_mode).passBinCut(VBIN_RESULT(inst_info.p_mode).step_in_IDS_Zone)
                    CurrentPassBinCutNum = VBIN_RESULT(inst_info.p_mode).passBinCut
                    
                    '''//If the performance mode is tested first time, set the BinCut(P_mode, CurrentPassBinCutNum).Tested = True.
                    '''ToDo: Maybe it needs to be modified for the failed DUT...
                    VBIN_RESULT(inst_info.p_mode).tested = True
                    
                    '''20210830: Modified to merge the branches of the vbt function Update_PassBinCut_for_GradeNotFound.
                    '''//If the grade is not found, set the result to fail.
                    If VBIN_RESULT(inst_info.p_mode).FLAGFAIL = False Then '''one of the instance fails
                        VBIN_RESULT(inst_info.p_mode).FLAGFAIL = True
                        VBIN_RESULT(inst_info.p_mode).GRADE = 0
                        VBIN_RESULT(inst_info.p_mode).GRADEVDD = 0
                        VBIN_RESULT(inst_info.p_mode).step_in_BinCut = Max_IDS_Step
                    End If
                    
                    '''20230821: Modified to support multi harvest flag, other pin group
                    If inst_info.HarvPinGrpCommnFailResult(site) = True Then
                        Call Set_MultiHarvestFlag_ForBV(OtherPin, site, , inst_info.performance_mode)
                    End If
                    If inst_info.HarvPinGrp_NonHarvFail(site) = True Then
                        Call Set_MultiHarvestFlag_ForBV(NonHarvest, site, , inst_info.performance_mode)
                    End If
                Else
                    ''20230821: Modified to support multi harvest flag, other pin group
                    Dim isHarvFail As Boolean
                    isHarvFail = False
                    
                    For HarvPatCnt = 0 To UBound(inst_info.HarvPinGrpCheckData)
                        For HarvCnt = 0 To UBound(HarvPinFlagMapping(inst_info.HarvPinGrpCheckData(HarvPatCnt).MappingIndex).HarvPinGrpConditionArr)
                            If inst_info.HarvPinGrpCheckData(HarvPatCnt).HarvFailResult(HarvCnt) = True Then        '' Check fail result from inst_info 20220304
                                TheExec.sites.item(site).FlagState(HarvPinFlagMapping(inst_info.HarvPinGrpCheckData(HarvPatCnt).MappingIndex).HarvPinGrpCondFailFlagArr(HarvCnt)) = logicTrue
                                ''20220609: Print harvest flag state for datalog
                                ''Flags: HarvestMappingTable((0th matched pattern).row).flags(which flag)
                                TheExec.flow.TestLimit 1, 0, 0, _
                                                Tname:=HarvPinFlagMapping(inst_info.HarvPinGrpCheckData(HarvPatCnt).MappingIndex).HarvPinGrpConditionArr(HarvCnt)
                                TheExec.flow.TestLimit 1, 0, 0, _
                                                Tname:=HarvPinFlagMapping(inst_info.HarvPinGrpCheckData(HarvPatCnt).MappingIndex).HarvPinGrpCondFailFlagArr(HarvCnt)
                                                        '' Pin Group
                                TheExec.Datalog.WriteComment "Site: " & site & " Failed in Bincut Bin1EQ1, so harvest flag is assigned true!"
                                '''20230821: Modified to support multi harvest flag
                                Call Set_MultiHarvestFlag_ForBV(PinHarvest, site, HarvPinFlagMapping(inst_info.HarvPinGrpCheckData(HarvPatCnt).MappingIndex).HarvPinGrpCondFailFlagArr(HarvCnt))
                                
                                isHarvFail = True
                            End If
                        Next HarvCnt
                    Next HarvPatCnt
                    
                    If isHarvFail = True Then
                        If Flag_BinX_Info_Parsed = True Then
                            '20230901: Bin1eq1 only harv fail, set result to bin1 last eqn (normally EQ2)
                            inst_info.grade_found(site) = True
                            VBIN_RESULT(inst_info.p_mode).step_in_IDS_Zone = DYNAMIC_VBIN_IDS_ZONE(inst_info.p_mode).step_Mapping(1, AllBinCut(inst_info.p_mode).Last_Bin1_EQN)
                            VBIN_RESULT(inst_info.p_mode).step_in_BinCut = DYNAMIC_VBIN_IDS_ZONE(inst_info.p_mode).EQ_Num(VBIN_RESULT(inst_info.p_mode).step_in_IDS_Zone) - 1
                            VBIN_RESULT(inst_info.p_mode).GRADE = DYNAMIC_VBIN_IDS_ZONE(inst_info.p_mode).Voltage(VBIN_RESULT(inst_info.p_mode).step_in_IDS_Zone)
                            VBIN_RESULT(inst_info.p_mode).GRADEVDD = DYNAMIC_VBIN_IDS_ZONE(inst_info.p_mode).Product_Voltage(VBIN_RESULT(inst_info.p_mode).step_in_IDS_Zone)
                            CurrentPassBinCutNum(site) = VBIN_RESULT(inst_info.p_mode).passBinCut
                            VBIN_RESULT(inst_info.p_mode).tested = True
                        Else
                            'Bin1eq1 only harv fail, set result to bin1 eq1
                            inst_info.grade_found(site) = True
                            VBIN_RESULT(inst_info.p_mode).step_in_IDS_Zone = DYNAMIC_VBIN_IDS_ZONE(inst_info.p_mode).step_Mapping(VBIN_RESULT(inst_info.p_mode).passBinCut, 1)
                            VBIN_RESULT(inst_info.p_mode).step_in_BinCut = DYNAMIC_VBIN_IDS_ZONE(inst_info.p_mode).EQ_Num(VBIN_RESULT(inst_info.p_mode).step_in_IDS_Zone) - 1
                            VBIN_RESULT(inst_info.p_mode).GRADE = DYNAMIC_VBIN_IDS_ZONE(inst_info.p_mode).Voltage(VBIN_RESULT(inst_info.p_mode).step_in_IDS_Zone)
                            VBIN_RESULT(inst_info.p_mode).GRADEVDD = DYNAMIC_VBIN_IDS_ZONE(inst_info.p_mode).Product_Voltage(VBIN_RESULT(inst_info.p_mode).step_in_IDS_Zone)
                            CurrentPassBinCutNum(site) = VBIN_RESULT(inst_info.p_mode).passBinCut
                            VBIN_RESULT(inst_info.p_mode).tested = True
                        End If
                    Else
                        '''20230821: Modified to support multi harvest flag, other pin group
                        Call Set_MultiHarvestFlag_ForBV(NonHarvest, site, , inst_info.performance_mode)
                        
                        VBIN_RESULT(inst_info.p_mode).step_in_IDS_Zone = inst_info.Step_Current
                        VBIN_RESULT(inst_info.p_mode).step_in_BinCut = DYNAMIC_VBIN_IDS_ZONE(inst_info.p_mode).EQ_Num(VBIN_RESULT(inst_info.p_mode).step_in_IDS_Zone) - 1
                        VBIN_RESULT(inst_info.p_mode).passBinCut = DYNAMIC_VBIN_IDS_ZONE(inst_info.p_mode).passBinCut(VBIN_RESULT(inst_info.p_mode).step_in_IDS_Zone)
                        CurrentPassBinCutNum = VBIN_RESULT(inst_info.p_mode).passBinCut
                        
                        '''//If the performance mode is tested first time, set the BinCut(P_mode, CurrentPassBinCutNum).Tested = True.
                        '''ToDo: Maybe it needs to be modified for the failed DUT...
                        VBIN_RESULT(inst_info.p_mode).tested = True
                    
                        '''20210830: Modified to merge the branches of the vbt function Update_PassBinCut_for_GradeNotFound.
                        '''//If the grade is not found, set the result to fail.
                        If VBIN_RESULT(inst_info.p_mode).FLAGFAIL = False Then '''one of the instance fails
                            VBIN_RESULT(inst_info.p_mode).FLAGFAIL = True
                            VBIN_RESULT(inst_info.p_mode).GRADE = 0
                            VBIN_RESULT(inst_info.p_mode).GRADEVDD = 0
                            VBIN_RESULT(inst_info.p_mode).step_in_BinCut = Max_IDS_Step
                        End If
                    End If
                End If
            Else
                VBIN_RESULT(inst_info.p_mode).step_in_IDS_Zone = inst_info.Step_Current
                VBIN_RESULT(inst_info.p_mode).step_in_BinCut = DYNAMIC_VBIN_IDS_ZONE(inst_info.p_mode).EQ_Num(VBIN_RESULT(inst_info.p_mode).step_in_IDS_Zone) - 1
                VBIN_RESULT(inst_info.p_mode).passBinCut = DYNAMIC_VBIN_IDS_ZONE(inst_info.p_mode).passBinCut(VBIN_RESULT(inst_info.p_mode).step_in_IDS_Zone)
                CurrentPassBinCutNum = VBIN_RESULT(inst_info.p_mode).passBinCut
            
                '''//If the performance mode is tested first time, set the BinCut(P_mode, CurrentPassBinCutNum).Tested = True.
                '''ToDo: Maybe it needs to be modified for the failed DUT...
                VBIN_RESULT(inst_info.p_mode).tested = True
            
                '''20210830: Modified to merge the branches of the vbt function Update_PassBinCut_for_GradeNotFound.
                '''//If the grade is not found, set the result to fail.
                If VBIN_RESULT(inst_info.p_mode).FLAGFAIL = False Then '''one of the instance fails
                    VBIN_RESULT(inst_info.p_mode).FLAGFAIL = True
                    VBIN_RESULT(inst_info.p_mode).GRADE = 0
                    VBIN_RESULT(inst_info.p_mode).GRADEVDD = 0
                    VBIN_RESULT(inst_info.p_mode).step_in_BinCut = Max_IDS_Step
                End If
            End If
        End If
    Next site
Exit Function
errHandler:
    Call Print_Error_Message(Error_Info, "LIB_VDD_BINNING", "Update_PassBinCut_for_GradeNotFound") 'Add ErrHandler 2023/08/01
    If AbortTest Then Exit Function Else Resume Next
End Function

'20210416: Modified to check if DSSC DigSrc patterns of SELSRM and Harvest Core exist.
'20201117: Modified to use "tlResultModeDomain" for pattern burst=Yes and decomposePatt=No. Requested by Leon Weng.
'20201029: Modified to remove the argument "result_mode As tlResultMode" and use inst_info.result_mode.
'20201029: Modified to use "Public Type Instance_Info".
'20201029: Modified to check if idxBlock_Selsrm_PrePatt = idxBlock_Selsrm_FuncPat.
'20201027: Modified to check SELSRM DSSC digsrc pattern for pattern burst requested by C651 Toby.
'20201016: Modified to adjust the sequence of arguments.
'20200918: Modified to add the argument "result_mode" for the vbt function "Check_and_Decompose_PrePatt_FuncPat".
'20200915: Modified to rename the argument "Flag_DecomposePatt_from_InstanceArg" to "str_Set_DecomposePatt"
'20200520: Created to check and decompose patsets PrePatt and FuncPat, and find SELSRAM DSSC pattern for DSSC digSrc.
Public Function Check_and_Decompose_PrePatt_FuncPat(inst_info As Instance_Info, result_mode As tlResultMode, str_Set_DecomposePatt As String, PrePatt As String, FuncPat As String)
On Error GoTo errHandler 'Add ErrHandler 2023/08/01
    Dim strTemp As String
'''//==================================================================================================================================================================================//'''
'''//Note:
'''No matter DecomposePatt = "YES" or "NO", it needs to decompose the pattern set to find the DSSC digsrc pattern.
'''//==================================================================================================================================================================================//'''
    '''//Init
    inst_info.enable_DecomposePatt = True
    '''SelSRM
    inst_info.patt_SelsrmDigSrc_decomposed_from_PrePatt = vbNullString
    inst_info.patt_SelsrmDigSrc_decomposed_from_FuncPat = vbNullString
    inst_info.patt_SelsrmDigSrc_single = vbNullString
    inst_info.idxBlock_Selsrm_PrePatt = -1
    inst_info.idxBlock_Selsrm_FuncPat = -1
    inst_info.idxBlock_Selsrm_singlePatt = -1
    '''str_Set_DecomposePatt comes from arguments of the instance.
    strTemp = LCase(str_Set_DecomposePatt)
    
    '''//Check if decomposing pattern is needed.
    If strTemp = "" Or strTemp = "yes" Or strTemp = "true" Then
        inst_info.enable_DecomposePatt = True
    ElseIf strTemp = "no" Or strTemp = "false" Then
        inst_info.enable_DecomposePatt = False
    Else
        inst_info.enable_DecomposePatt = True
        Call Print_Error_Message(Error_Info, "LIB_VDD_BINNING", "Check_and_Decompose_PrePatt_FuncPat", "Argument: " & str_Set_DecomposePatt & " from this instance doesn't have the correct format to decide DecomposePat. Error!!!")
        'TheExec.ErrorLogMessage "Argument: " & str_Set_DecomposePatt & " from this instance doesn't have the correct format to decide DecomposePat. Error!!!"    ''20230807: Removed by sanity check
    End If
    
    '''//Decompose the pattern set to check if any SELSRM DSSC digsrc pattern exists in the pattern set. Requested by C651 Toby.
    '''PrePatt
    If PrePatt <> "" Then
        inst_info.PrePatt = PrePatt
        '''//Decompose PrePatt and FuncPatt to find DSSC pattern by BlockType / Pattern keyword for SELSRAM DSSC bit array according to SELSRM_Mapping_Table.
        Find_DsscPatt_fromPattSet inst_info, inst_info.PrePatt, inst_info.ary_PrePatt_decomposed, inst_info.count_PrePatt_decomposed, inst_info.patt_SelsrmDigSrc_decomposed_from_PrePatt, inst_info.idxBlock_Selsrm_PrePatt
    End If
    
    '''FuncPat
    If FuncPat <> "" Then
        inst_info.FuncPat = FuncPat
        '''//Decompose PrePatt and FuncPatt to find DSSC pattern by BlockType / Pattern keyword for SELSRAM DSSC bit array according to SELSRM_Mapping_Table.
        Find_DsscPatt_fromPattSet inst_info, inst_info.FuncPat, inst_info.ary_FuncPat_decomposed, inst_info.count_FuncPat_decomposed, inst_info.patt_SelsrmDigSrc_decomposed_from_FuncPat, inst_info.idxBlock_Selsrm_FuncPat
    End If
    
    '''//For pattern without SelSram DSSC keyword, pattern count "idxPatt_Selsrm_PrePatt" keeps as -1.
    If inst_info.patt_SelsrmDigSrc_decomposed_from_PrePatt <> "" And inst_info.patt_SelsrmDigSrc_decomposed_from_FuncPat <> "" Then
        If inst_info.idxBlock_Selsrm_PrePatt = inst_info.idxBlock_Selsrm_FuncPat Then
            inst_info.patt_SelsrmDigSrc_single = inst_info.patt_SelsrmDigSrc_decomposed_from_PrePatt
            inst_info.idxBlock_Selsrm_singlePatt = inst_info.idxBlock_Selsrm_PrePatt
        Else
            inst_info.patt_SelsrmDigSrc_single = inst_info.patt_SelsrmDigSrc_decomposed_from_PrePatt
            inst_info.idxBlock_Selsrm_singlePatt = inst_info.idxBlock_Selsrm_PrePatt
            Call Print_Error_Message(Error_Info, "LIB_VDD_BINNING", "", "idxBlock_Selsrm_PrePatt and idxBlock_Selsrm_FuncPat are different. Please check SelsrmPat for PrePatt and FuncPat. Error!!!")
            'TheExec.ErrorLogMessage "idxBlock_Selsrm_PrePatt and idxBlock_Selsrm_FuncPat are different. Please check SelsrmPat for PrePatt and FuncPat. Error!!!"    ''20230807: Removed by sanity check
        End If
    ElseIf inst_info.patt_SelsrmDigSrc_decomposed_from_PrePatt <> "" Then
        inst_info.patt_SelsrmDigSrc_single = inst_info.patt_SelsrmDigSrc_decomposed_from_PrePatt
        inst_info.idxBlock_Selsrm_singlePatt = inst_info.idxBlock_Selsrm_PrePatt
    ElseIf inst_info.patt_SelsrmDigSrc_decomposed_from_FuncPat <> "" Then
        inst_info.patt_SelsrmDigSrc_single = inst_info.patt_SelsrmDigSrc_decomposed_from_FuncPat
        inst_info.idxBlock_Selsrm_singlePatt = inst_info.idxBlock_Selsrm_FuncPat
    ElseIf UCase(VddBinName(inst_info.p_mode)) Like "*MBIST*" Then
        inst_info.patt_SelsrmDigSrc_single = vbNullString
    Else
        inst_info.patt_SelsrmDigSrc_single = vbNullString
        TheExec.Datalog.WriteComment "PrePatt and FuncPat patsets don't contain any DSSC pattern."
    End If
    
    '''*********************************************************************************************************************************************'''
    '''inst_info.enable_DecomposePatt is refreshed as "False" at the beginning of each GradeSearch_XXX_VT instance.
    '''*********************************************************************************************************************************************'''
    '''//If flag of DecomposePatt = False, redim array of ary_PrePatt_decomposed and ary_FuncPat_decomposed as 0.
    If inst_info.enable_DecomposePatt = False Then '''without decomposing pattern sets
                ''' Leon Weng Comment Prepat can't burst during Coll
'        If PrePatt <> "" And inst_info.count_PrePatt_decomposed > 0 Then
'            ReDim inst_info.ary_PrePatt_decomposed(0)
'            inst_info.ary_PrePatt_decomposed(0) = PrePatt
'            inst_info.count_PrePatt_decomposed = 1
'        End If
        
        If FuncPat <> "" And inst_info.count_FuncPat_decomposed > 0 Then
            ReDim inst_info.ary_FuncPat_decomposed(0)
            inst_info.ary_FuncPat_decomposed(0) = FuncPat
            inst_info.count_FuncPat_decomposed = 1
        End If
        
        '''//Set "result_mode = tlResultModeDomain" (return a unique pass/fail result for each module and time domain) if pattern bursted without decomposing pattern.
        '''20201117: Modified to use "tlResultModeDomain" for pattern burst=Yes and decomposePatt=No. Requested by Leon Weng.
        If glb_TesterType = "UltraFLEXplus" Then
            inst_info.result_mode = result_mode
        Else
            If result_mode <> tlResultModeDomain Then
                Call Print_Error_Message(Error_Info, "LIB_VDD_BINNING", "Check_and_Decompose_PrePatt_FuncPat", "Should not use module mode on UltraFLEX!!!")
            End If
        inst_info.result_mode = tlResultModeDomain
        result_mode = tlResultModeDomain
        End If
    End If
    
        '''20221109: Added to redim pattern pin string array
    'If inst_info.enable_CMEM_Collection = True Then
        If inst_info.count_FuncPat_decomposed > 0 And glb_TesterType = "UltraFLEXplus" Then
            ReDim inst_info.str_patternPins(inst_info.count_FuncPat_decomposed - 1)
        Else
            ''Default value
            ReDim inst_info.str_patternPins(0)
            inst_info.str_patternPins(0) = vbNullString
        End If
            
    'End If
Exit Function
errHandler:
    Call Print_Error_Message(Error_Info, "LIB_VDD_BINNING", "Check_and_Decompose_PrePatt_FuncPat") 'Add ErrHandler 2023/08/01
    If AbortTest Then Exit Function Else Resume Next
End Function

'20210416: Modified to check if DSSC DigSrc patterns of SELSRM and Harvest Core exist.
'20201102: Modified to add the argument "inst_info As Instance_Info".
'20201102: Modified to use "Public Type Instance_Info".
'20201027: Modified to check SELSRM DSSC digsrc pattern for pattern burst requested by C651 Toby.
'20200915: Modified to update the status of "inst_info.enable_DecomposePatt".
'20200915: Modified to rename the argument "Flag_DecomposePatt_from_InstanceArg" to "str_Set_DecomposePatt".
'20191205: Created for finding DSSC pattern in the pattern set.
Public Function Find_DsscPatt_fromPattSet(inst_info As Instance_Info, pattSet As String, ary_patt_decompose() As String, count_patt_decompose As Long, _
                                            patt_Selsrm_digsrc As String, idxBlock_Selsrm_Pattern As Integer)
On Error GoTo errHandler 'Add ErrHandler 2023/08/01
    Dim i As Integer
    Dim idxBlock As Integer
    Dim strTemp As String
    Dim FlagReplace_DSRAM_Pat As Boolean
    Dim newKeyword As String
    Dim KeywordList As Variant
    Dim pat_ary() As String
    Dim pat_cnt As Long
    Dim found As Boolean
    '''//Init
    patt_Selsrm_digsrc = vbNullString
    idxBlock_Selsrm_Pattern = -1
    strTemp = LCase(inst_info.inst_name)
    FlagReplace_DSRAM_Pat = False
    found = False
    
    If pattSet <> "" Then
        '''//Decompose pattSet into the arrary "ary_patt_decompose".
        'GetPatFromPatternSet CStr(pattSet), ary_patt_decompose, count_patt_decompose
        GetPatsFromPatSets CStr(pattSet), ary_patt_decompose, count_patt_decompose, False
        GetPatsFromPatSets CStr(pattSet), ary_FuncPat_decomposed_FFC, count_patt_decompose, True        '''20230201: Added for ffc on UFP, save pattern name
        
        '''//If pattern count>0, use pattern-loop to find DSSC DigSrc pattern.
        If count_patt_decompose > 0 Then
            '''//Use pattern-loop.
            For i = 0 To count_patt_decompose - 1
                '''//Use pattern-loop to find DSSC DigSrc pattern of SELSRM by Pattern keyword for SELSRAM DSSC bit array according to SELSRM_Mapping_Table.
                For idxBlock = 0 To UBound(SelsramMapping)
                    If LCase(SelsramMapping(idxBlock).Pattern) <> "*" And LCase(ary_patt_decompose(i)) Like LCase(SelsramMapping(idxBlock).Pattern) Then
                        pat_ary = TheExec.DataManager.Raw.GetPatternsInSet(ary_patt_decompose(i), pat_cnt)
                        '''20220418, Static_MappingTable Parsing, now only support use bin1 keyword
                        If strTemp Like "*_hbv" Then 'hbv test instance
                            If inst_info.Special_Voltage_setup = True Then
                                If Static_Mapping(inst_info.p_mode, 1).HVCC_Addtional_Selsrm_Value(inst_info.addi_mode) <> "" Then
                                    KeywordList = Split(ary_patt_decompose(i), "_")
                                    If UCase(KeywordList(4)) <> "FU" Then
                                        newKeyword = "SRM" & Static_Mapping(inst_info.p_mode, 1).HVCC_Addtional_Selsrm_Value(inst_info.addi_mode)
                                        FlagReplace_DSRAM_Pat = True
                                    End If
                                End If
                            Else
                                If Static_Mapping(inst_info.p_mode, 1).HVCC_Selsrm_Value <> "" Then
                                    KeywordList = Split(ary_patt_decompose(i), "_")
                                    If UCase(KeywordList(4)) <> "FU" Then
                                        newKeyword = "SRM" & Static_Mapping(inst_info.p_mode, 1).HVCC_Selsrm_Value
                                        FlagReplace_DSRAM_Pat = True
                                    End If
                                End If
                            End If
                        ElseIf strTemp Like "*_bv" Then 'bv test instance
                            If inst_info.Special_Voltage_setup = True Then
                                If Static_Mapping(inst_info.p_mode, 1).Addtional_Selsrm_Value(inst_info.addi_mode) <> "" Then
                                    KeywordList = Split(ary_patt_decompose(i), "_")
                                    If UCase(KeywordList(4)) <> "FU" Then
                                        newKeyword = "SRM" & Static_Mapping(inst_info.p_mode, 1).Addtional_Selsrm_Value(inst_info.addi_mode)
                                        FlagReplace_DSRAM_Pat = True
                                    End If
                                End If
                            Else
                                If Static_Mapping(inst_info.p_mode, 1).Selsrm_Value <> "" Then
                                    KeywordList = Split(ary_patt_decompose(i), "_")
                                    If UCase(KeywordList(4)) <> "FU" Then
                                        newKeyword = "SRM" & Static_Mapping(inst_info.p_mode, 1).Selsrm_Value
                                        FlagReplace_DSRAM_Pat = True
                                    End If
                                End If
                            End If
                            
                        ElseIf strTemp Like "*_hv" Or strTemp Like "*_lv" Then  'SAA
                            FlagReplace_DSRAM_Pat = False
                            
                        Else
                            TheExec.Datalog.WriteComment "Instance: " & inst_info.inst_name & " doesn't end with any _bv or _hbv keyword. Error!!!"
                        End If
                        
                        If FlagReplace_DSRAM_Pat = True Then
                            ary_patt_decompose(i) = Replace(ary_patt_decompose(i), "SRMDSSC", newKeyword)
                            If UCase(ary_patt_decompose(i)) Like "*_L_*_BI_*_SRM1[EF]*" Then
                                ary_patt_decompose(i) = Replace(ary_patt_decompose(i), "THTA0", "THTB0", compare:=vbTextCompare)
                            End If
                        Else
                            ary_patt_decompose(i) = pat_ary(0)
                            patt_Selsrm_digsrc = ary_patt_decompose(i)
                            idxBlock_Selsrm_Pattern = idxBlock
                        End If
                        
                        Exit For
                    End If
                Next idxBlock
                
'''ToDo: Maybe it can use pattern-loop to find DSSC DigSrc pattern of Harvest Core here...
                '''If 1st DSSC pattern is found, exit for-loop.
'                If patt_Selsrm_digsrc <> "" Then
'                    Exit For
'                End If
                If LCase(ary_patt_decompose(i)) Like "*_xordssc*" Then
                    pat_ary = TheExec.DataManager.Raw.GetPatternsInSet(ary_patt_decompose(i), pat_cnt)
                    If pat_cnt <> 0 Then ary_patt_decompose(i) = pat_ary(0)
                End If
'                If LCase(TheExec.CurrentJob) = "wlft1" And LCase(ary_patt_decompose(i)) Like "*_pl2#_*" And Not TheExec.Flow.EnableWord("HTOL_TX_ROOM") Then
'                    ary_patt_decompose(i) = vbNullString
'                    found = True
'                End If
            Next i
            
'            If LCase(TheExec.CurrentJob) = "wlft1" And found Then
'                Dim PatStr As String
'                Dim CNT As Integer
'                PatStr = Join(ary_patt_decompose, ",")
'                While (InStr(1, PatStr, ",,") <> 0)
'                    PatStr = Replace(PatStr, ",,", ",")
'                Wend
'                Erase ary_patt_decompose
'                ary_patt_decompose = Split(PatStr, ",")
'                CNT = UBound(ary_patt_decompose)
'                While (ary_patt_decompose(CNT) = "")
'                    ReDim Preserve ary_patt_decompose(CNT - 1)
'                    CNT = CNT - 1
'                Wend
'                count_patt_decompose = CNT + 1
'            End If
        Else
            TheExec.Datalog.WriteComment "PatternSet: " & CStr(pattSet) & " contains no pattern. Error!!!"
            TheExec.ErrorLogMessage "PatternSet: " & CStr(pattSet) & " contains no pattern. Error!!!"
        End If
    End If
Exit Function
errHandler:
    Call Print_Error_Message(Error_Info, "LIB_VDD_BINNING", "Find_DsscPatt_fromPattSet") 'Add ErrHandler 2023/08/01
    If AbortTest Then Exit Function Else Resume Next
End Function

'20201230: Modified to check if FuncPat <> "".
'20201102: Modified to add the argument "enable_DecomposePatt as boolean".
'20200520: Created to show the errorLogMessage if "burst=no" and "Decompose_Pattern=false".
Public Function Check_Pattern_NoBurst_NoDecompose(FuncPat As String, count_decompose_FuncPat As Long, enable_DecomposePatt As Boolean)
On Error GoTo errHandler 'Add ErrHandler 2023/08/01
    Dim lastBurstPat As New SiteVariant
    Dim isGrp As New SiteBoolean
    Dim lastLabel As New SiteVariant
    Dim site As Variant 'Carter, 20240304
'''//==================================================================================================================================================================================//'''
'''//Note:
'''lastBurstPat: String or Variant. The pattern and a separated list of pattern files and groups or a pattern set name.
'''     This method returns the pattern set name to this parameter when the burst mode is enabled in the Pattern Sets sheet (Burst column is yes).
'''     When the burst mode is disabled (Burst column is no), this method returns the last executed pattern name to this parameter.
'''
'''isGrp: Boolean or Variant. Whether the last burst involved a pattern group, where:
'''     isGrp=True: The last burst involved a pattern group.
'''     isGrp=False: The last burst did not involve a pattern group.
'''//==================================================================================================================================================================================//'''
    If FuncPat <> "" And enable_DecomposePatt = False Then '''without decomposing pattern sets
        TheHdw.Digital.Patgen.ReadLastStart lastBurstPat, isGrp, lastLabel
        
        For Each site In TheExec.sites
            If count_decompose_FuncPat = 1 And lastBurstPat(site) <> FuncPat Then
                Call Print_Error_Message(Error_Info, "LIB_VDD_BINNING", "Check_Pattern_NoBurst_NoDecompose", "site:" & site & ", FuncPat:" & FuncPat & ", it is not pattern burst, but it is not decomposed for the instance. Error!!!")
                'TheExec.ErrorLogMessage "site:" & site & ", FuncPat:" & FuncPat & ", it is not pattern burst, but it is not decomposed for the instance. Error!!!"    ''20230807: Removed by sanity check
                Exit For
            End If
        Next site
    End If
Exit Function
errHandler:
    Call Print_Error_Message(Error_Info, "LIB_VDD_BINNING", "Check_Pattern_NoBurst_NoDecompose") 'Add ErrHandler 2023/08/01
    If AbortTest Then Exit Function Else Resume Next
End Function

'20210901: Modified to rename "IndexLevelPerSite As New SiteLong" as "Step_Current As New SiteLong" for Public Type Instance_Info.
'20201210: Modified to remove "IndexLevelPerSite As SiteLong" from the argument of the vbt function Get_PassBinNum_by_Step.
'20201030: Modified to use "Public Type Instance_Info".
'20200713: Modified to add the argument "PASSBINCUT as siteLong" for CurrentPassBinCutNum.
'20200525: Created to get PassBinNum by BinCut GradeSearch Step.
Public Function Get_PassBinNum_by_Step(inst_info As Instance_Info, PassBinCutCurrent As SiteLong, PassBinNum As SiteLong)
On Error GoTo errHandler 'Add ErrHandler 2023/08/01
    Dim site As Variant
'''//==================================================================================================================================================================================//'''
'''//Note:
'''BinCut search (is_BinSearch = True)  : get PassBinNum from DYNAMIC_VBIN_IDS_ZONE(p_mode).PASSBINCUT(IndexLevelPerSite).
'''BinCut check  (is_BinSearch = False) : get PassBinNum from CurrentPassBinCutNum.
'''//==================================================================================================================================================================================//'''
    If inst_info.is_BinSearch = True Then '''BinCut search.
        For Each site In TheExec.sites
            PassBinNum(site) = DYNAMIC_VBIN_IDS_ZONE(inst_info.p_mode).passBinCut(inst_info.Step_Current)
        Next site
    Else '''BinCut check.
        PassBinNum = PassBinCutCurrent
    End If
Exit Function
errHandler:
    Call Print_Error_Message(Error_Info, "LIB_VDD_BINNING", "Get_PassBinNum_by_Step") 'Add ErrHandler 2023/08/01
    If AbortTest Then Exit Function Else Resume Next
End Function

'20200618: Created to "check powerDomain--> powerPin --> DC Spec specName".
'20200617: Modified to check if VDD_XXX_VOP_VAR or VDD_XXX_VAR of powerPin exists.
Public Function initDomain2DcSpecGrp(domainList As String, dict_Domain2DcSpecGrp As Dictionary)
On Error GoTo errHandler 'Add ErrHandler 2023/08/01
    Dim i As Long
    Dim j As Long
    Dim powerDomain As String
    Dim powerPin As String
    Dim PinGroup As String
    Dim specName As String
    Dim specGrp As String
    Dim split_domainlist() As String
    Dim split_content() As String
'''//==================================================================================================================================================================================//'''
'''//Note:
'''Check powerDomain--> powerPin --> DC Spec specName.
'''//==================================================================================================================================================================================//'''
    dictDomain2DcSpecGrp.RemoveAll
    split_domainlist = Split(domainList, ",")

    For i = 0 To UBound(split_domainlist)
        powerDomain = split_domainlist(i)
        PinGroup = vbNullString
    
        '''//Check if powerDomain belongs to powerPin or pinGroup...
        If domain2pinDict.Exists(UCase(powerDomain)) = True Then
            PinGroup = VddbinDomain2Pin(powerDomain)
        ElseIf pin2domainDict.Exists(UCase(powerDomain)) = True Then
            PinGroup = powerDomain
        Else
            Call Print_Error_Message(Error_Info, "LIB_VDD_BINNING", powerDomain & " is not BinCut powerPin or pinGroup. It is incorrect for initDomain2DcSpecGrp. Error!!!")
            'TheExec.Datalog.WriteComment powerDomain & " is not BinCut powerPin or pinGroup. It is incorrect for initDomain2DcSpecGrp. Error!!!"
            'TheExec.ErrorLogMessage powerDomain & " is not BinCut powerPin or pinGroup. It is incorrect for initDomain2DcSpecGrp. Error!!!"    ''20230807: Removed by sanity check
        End If
        
        If PinGroup <> "" Then
            '''//Get powerPins from powerDomain
            split_content = Split(PinGroup, ",")
            specGrp = vbNullString
            
            For j = 0 To UBound(split_content)
                powerPin = split_content(j)
                specName = vbNullString
                If UCase(powerPin) Like "*_FT" Then powerPin = Replace(powerPin, "_FT", "", compare:=vbTextCompare)
                '''//Check if VDD_XXX_VOP_VAR or VDD_XXX_VAR exist...
                If dictPin2Dcspec.Exists(UCase(powerPin & "_VOP_VAR")) = True Then
                    specName = powerPin & "_VOP_VAR"
                ElseIf dictPin2Dcspec.Exists(UCase(powerPin & "_VAR")) = True Then
                    specName = powerPin & "_VAR"
                Else
                    specName = vbNullString
                     Call Print_Error_Message(Error_Info, "LIB_VDD_BINNING", powerDomain & " is not BinCut powerPin or pinGroup. It is incorrect for initDomain2DcSpecGrp. Error!!!")
                    'TheExec.Datalog.WriteComment powerDomain & " is not BinCut powerPin or pinGroup. It is incorrect for initDomain2DcSpecGrp. Error!!!"
                    'TheExec.ErrorLogMessage powerDomain & " is not BinCut powerPin or pinGroup. It is incorrect for initDomain2DcSpecGrp. Error!!!"    ''20230807: Removed by sanity check
                End If
                
                If specName <> "" Then
                   If specGrp <> "" Then
                       specGrp = specGrp & "," & specName
                   Else
                       specGrp = specName
                   End If
                End If
            Next j
            
            '''dictDomain2DcspecGrp
            If dictDomain2DcSpecGrp.Exists(powerDomain) Then
                '''Do nothing
            Else
                '''//Update dictionary of Domain2Pin
                dictDomain2DcSpecGrp.Add powerDomain, specGrp
            End If
        End If
    Next i
Exit Function
errHandler:
    Call Print_Error_Message(Error_Info, "LIB_VDD_BINNING", "initDomain2DcSpecGrp") 'Add ErrHandler 2023/08/01
    If AbortTest Then Exit Function Else Resume Next
End Function

'20200618: Created for Domain2DcSpecGrp.
Public Function VddbinDomain2DcSpecGrp(vddbinDomain As String) As String
On Error GoTo errHandler 'Add ErrHandler 2023/08/01
    Dim powerDomain As String
    powerDomain = UCase(vddbinDomain)

    If dictDomain2DcSpecGrp.Exists(powerDomain) Then
        VddbinDomain2DcSpecGrp = UCase(dictDomain2DcSpecGrp.item(powerDomain))
    Else
        VddbinDomain2DcSpecGrp = "Domain_Error"
        Call Print_Error_Message(Error_Info, "LIB_VDD_BINNING", "VddbinDomain2DcSpecGrp", "Vddbin Domain=" & vddbinDomain & ", " & vddbinDomain & " doesn't exist in dictDomain2DcSpecGrp. Error!!!")
        'TheExec.ErrorLogMessage "Vddbin Domain=" & vddbinDomain & ", " & vddbinDomain & " doesn't exist in dictDomain2DcSpecGrp. Error!!!"
    End If
Exit Function
errHandler:
    Call Print_Error_Message(Error_Info, "LIB_VDD_BINNING", "VddbinDomain2DcSpecGrp") 'Add ErrHandler 2023/08/01
    If AbortTest Then Exit Function Else Resume Next
End Function

'20210618: Modified to mask TheExec.sites(site).LastTestResultRaw because HardIP of Crete used pattern(pat).test(pfAlways).
'20201012: Modified to check if alarmFail(site) is triggered or not.
'20200901: TER factory thought that pfAlways didn't cause "TheExec.sites(Site).LastTestResultRaw" issue...
'20200819: Discussed "pfAlways" issue with Chihome, he found the same case in the offline simulation. He suggested us to ask TER factory for patch or .dll to fix this.
'20200817: Modified to check if "TheExec.sites(site).LastTestResultRaw=tlResultFail".
'20200815: Modified to remove "BV_Pass(site)".
'20200815: Modified to use "TheExec.sites(site).LastTestResultRaw" to get testResult of Call Instance.
'20200812: Modified to use BinCut globalVariable to check HardIP pattern result for BinCut call instance.
'20200622: Created for Decide_PattPass_by_failFlag.
Public Function Decide_PattPass_by_failFlag(flagName As String, pattPass As SiteBoolean)
On Error GoTo errHandler 'Add ErrHandler 2023/08/01
    Dim site As Variant
    Dim mySiteResult As Long
'''//==================================================================================================================================================================================//'''
'''//Note:
'''1. Check if "TheExec.sites(site).LastTestResultRaw=tlResultFail" for HardIP ELB vbt function "Meas_FreqVoltCurr_Universal_func" with "Call TheHdw.Patterns(Pat).Test(pfNever, 0)".
'''2. Check if flagState("F_BV_CALLINST") for HardIP ELB vbt function with "Call TheHdw.Patterns(Pat).Test(pfAlways, 0)".
'''
'''<LastTestResultRaw>
'''This property gets the ungated pass/fail status of the last test executed on the specified site based on the result column.
'''Read-only tlResultType.
'''tlResultFail: the test failed; tlResultNoTest: no test was available; tlResultPass: the test passed.
'''
'''<!!!Warning!!!>
'''1. "TheExec.Flow.LastFlowStepResult" has issues with "TheHdw.Patterns(Pat).test(pfAlways, 0)".
'''Please contact Teradyne factory/software team for this issue.
'''ToDo: Check if "TheExec.sites(site).LastTestResultRaw" and "TheExec.Flow.LastFlowStepResult" ready for "TheHdw.Patterns(Pat).test(pfAlways, 0)"...
'''2. For instance with pfAlways, it cau use failFlag or BV_Pass to get testResult about Pass/Fail.
'''3. Decide results of Pattern pass/fail and use-limit by failFlag of the instance and use-limit.
'''4. Remember to check if BV_Pass is used in LIB_HardIP\HardIP_WriteFuncResult.
'''//==================================================================================================================================================================================//'''
    
        '''//Check if alarmFail(site) is triggered or not.
    If glb_TesterType = "UltraFLEXplus" Then
    Dim alarmOccurred As New SiteBoolean
    alarmOccurred = False
         alarmOccurred = TheHdw.Alarms.GetAlarmingSites(True)        ''GetAlarmingSites(clearAlarm = True), only plus or > 10_30_90 can use
    
    For Each site In TheExec.sites
        If alarmOccurred(site) = True Then
                alarmFail(site) = True      ''Update alarm info to alarmFail array for UFP
            End If
        Next site
    End If
    For Each site In TheExec.sites
        If alarmFail(site) = True Then
            TheExec.Datalog.WriteComment "Site:" & site & ", alarmFail!!!"
            pattPass(site) = False
        Else
            '''//*****************************************************************************************************************************************************************************//'''
            '''//Note:
            '''Check if "TheExec.sites(site).LastTestResultRaw=tlResultFail" for HardIP ELB vbt function "Meas_FreqVoltCurr_Universal_func" with "Call TheHdw.Patterns(Pat).Test(pfNever, 0)".
            '''Check if flagState("F_BV_CALLINST") for HardIP ELB vbt function with "Call TheHdw.Patterns(Pat).Test(pfAlways, 0)".
            '''//*****************************************************************************************************************************************************************************//'''
            '''//Get siteResult.
            '''20210618: Modified to mask TheExec.sites(site).LastTestResultRaw because HardIP of Crete used pattern(pat).test(pfAlways).
            '''So that we can directly get results of pattern Pass/Fail and testLimit by checking flagstate of FailFlag "F_BV_CALLINST".
            'mySiteResult = TheExec.sites(site).LastTestResultRaw
            
            '''//According to siteResult, update pattPass for each site.
            '''20210618: Modified to mask TheExec.sites(site).LastTestResultRaw because HardIP of Crete used pattern(pat).test(pfAlways).
            If TheExec.sites.item(site).FlagState(flagName) = logicFalse And BV_Pass(site) = True Then 'And (mySiteResult = tlResultPass Or mySiteResult = tlResultNoTest) Then
                pattPass(site) = True
            ElseIf TheExec.sites.item(site).FlagState(flagName) = logicTrue Or BV_Pass(site) = False Then 'Or mySiteResult = tlResultFail Then
                If Flag_VDD_Binning_Offline = True Or TheExec.flow.enableWord("Vddbinning_OpenSocket") = True Then '''If the tester is offline or opensocket.
                    TheExec.Datalog.WriteComment "site:" & site & "," & flagName & "=" & CStr(CBool(TheExec.sites.item(site).FlagState(flagName) = logicFalse)) & ",but it is forced to pass for offline or opensocket!!!"
                    pattPass(site) = True
                    TheExec.sites.item(site).FlagState(flagName) = logicFalse
                Else
                    pattPass(site) = False
                End If
            End If
        End If
    Next site
Exit Function
errHandler:
    Call Print_Error_Message(Error_Info, "LIB_VDD_BINNING", "Decide_PattPass_by_failFlag") 'Add ErrHandler 2023/08/01
    If AbortTest Then Exit Function Else Resume Next
End Function

'20210703: Modified to use dict_strPmode2EfuseCategory as the dictionary of p_mode and array of the related Efuse category.
'20210703: Modified to use dict_EfuseCategory2BinCutTestJob as the dictionary of Efuse category and the matched programming state in Efuse.
'20210701: Modified to reset gb_bincut_power_list().
'20210526: Modified to add "Flag_Get_column_Monotonicity_Offset" for Monotonicity_Offset check because C651 Si revised the check rules.
'20201222: Modified to reset the dictionary "dict_OutsideBinCut_additionalMode".
'20201112: Modified to reset the dictionary "dict_IsCorePowerInBinCutFlowSheet".
'20201110: Modified to reset dictionaries of powerBinning.
'20201103: Modified to set cntAdditionalMode = 0 for parsing Non_Binning_Rail.
'20201023: Modified to reset "dict_BinCutFlow_Domain2Column" and "dict_BinCutFlow_Column2Domain".
'20201021: Modified to reset "dict_IsCorePower".
'20200702: Modified to add "Flag_NonbinningrailOutsideBinCut_parsed=false".
'20200622: Created to reset BinCut globalVariable for initVddBinning.
'20191219: Modified to reset dictionaries of Domain2Pin and Pin2Domain.
Public Function Reset_BinCut_GlobalVariable_for_initVddBinning()
On Error GoTo errHandler 'Add ErrHandler 2023/08/01
    Dim idx_powerDomain As Long
    Dim p_mode As Integer
    Dim addi_mode As Integer '''For the additional mode
    Dim passBinCut As Long
    Dim COREPOWER As Long
    '''//Init variables
    Flag_PowerBinningTable_Parsed = False
    Flag_Interpolation_enable = False
    is_BinCutJob_for_StepSearch = False
    Flag_Harvest_Pmode_Table_Parsed = False
    Flag_Harvest_Mapping_Table_Parsed = False
    Flag_Harvest_Core_DSSC_Ready = False
    Flag_Get_column_Monotonicity_Offset = False
    Flag_IDS_Mapping_enable = False
    gb_str_EfuseCategory_for_powerbinning = vbNullString
    '''//T0TX format WWWWW20220630 T0TX
    Flag_T0TX_ROOM_format = False
    Flag_T0TX_HOT_format = False
    If TheExec.enableWord("T0TX_ROOM") = True Then Flag_T0TX_ROOM_format = True
    If TheExec.enableWord("T0TX_HOT") = True Then Flag_T0TX_HOT_format = True
    Total_Bincut_Num = 0
    MaxEqnNum = 0
    MaxPerformanceModeCount = 0
    MaxAdditionalModeCount = 0
    MaxSiteCount = TheExec.sites.Existing.Count
    Max_IDS_Zone = 0
    Max_IDS_Dis_Zone = 0

    '''//Reset sheetName for Print Bincut Config
    gS_EquationVoltages_sheet = vbNullString '20220321
    gS_Non_Binning_Rail_sheet = vbNullString
    gS_Vdd_Binning_Def_appA_1 = vbNullString
    gS_Vdd_Binning_Def_appA_2 = vbNullString
    gS_Vdd_Binning_Def_appA_3 = vbNullString
    gS_PwrSeq_ShtName = vbNullString
    
    '''//Init power_list
    Power_List_All = vbNullString
    
    '''//Init the flag for parsing Selsrm_Mapping_Table
    Flag_SelsrmMappingTable_Parsed = False
    
    '''//Init the flag for parsing Power_Binning_Harvest table
    Flag_Enable_PowerBinning_Harvest = False
    
    '''//Init the flag for parsing Non_Binning_Rail_Outside_BV
    Flag_NonbinningrailOutsideBinCut_parsed = False
    
    '''//Clear the dictionary of Pmode2enum, additional_mode2enum.
    VddbinPinDict.RemoveAll
    VddbinPmodeDict.RemoveAll
    AdditionalModeDict.RemoveAll
    dict_OutsideBinCut_additionalMode.RemoveAll
    
    '''//Init dictionaries of Domain2Pin and Pin2Domain.
    domain2pinDict.RemoveAll
    pin2domainDict.RemoveAll
    VddbinPinDcvstypeDict.RemoveAll
    dictPin2Dcspec.RemoveAll
    dictDomain2DcSpecGrp.RemoveAll
    dict_IsCorePower.RemoveAll
    dict_IsCorePowerInBinCutFlowSheet.RemoveAll
    
    '''//Init the pin_Groups
    dict_BinCutFlow_Domain2Column.RemoveAll
    dict_BinCutFlow_Column2Domain.RemoveAll
    FullCorePowerinFlowSheet = vbNullString
    FullOtherRailinFlowSheet = vbNullString
    FullBinCutPowerinFlowSheet = vbNullString
    selsramLogicPin = vbNullString
    selsramSramPin = vbNullString
    selsramPin = vbNullString
    
    '''//Define how many kinds of performance modes
    '''init the array of VddBinName
    ReDim VddBinName(0)
    ReDim AdditionalModeName(0)
    cntVddbinPin = -1
    cntVddbinPmode = -1
    cntAdditionalMode = 0
    VddBinName(0) = "None"
    AdditionalModeName(0) = "None"
    
    '''//Init the sheet dictionary for PowerBinning
    PwrBin_SheetnameDict.RemoveAll
    PwrBin_SpecIdx2SpecNameDict.RemoveAll
    dict_Binned_Mode_Column2Ratio.RemoveAll
    dict_Binned_Mode_Ratio2Column.RemoveAll
    dict_Binned_Mode_Ratio2Idx.RemoveAll
    dict_Binned_Mode_Column2Ratio.RemoveAll
    dict_Other_Mode_Ratio2Column.RemoveAll
    dict_Other_Mode_Ratio2Idx.RemoveAll
    
    '''//Clear the array gb_bincut_power_list to reset the list of all performance_modes in each powerDomain.
    '''20210701: Modified to reset gb_bincut_power_list().
    '''20240520 dynamic will redim array size behind this function, move to initDynamicArray
'    For idx_powerDomain = 0 To UBound(gb_bincut_power_list)
'        gb_bincut_power_list(idx_powerDomain) = vbNullString
'    Next idx_powerDomain
    
    '''20210703: Modified to use dict_strPmode2EfuseCategory as the dictionary of p_mode and array of the related Efuse category.
    dict_strPmode2EfuseCategory.RemoveAll
    '''20210703: Modified to use dict_EfuseCategory2BinCutTestJob as the dictionary of Efuse category and the matched programming state in Efuse.
    dict_EfuseCategory2BinCutTestJob.RemoveAll
    '''20220711: Modified to use dict_EfuseCategory2BDF
    dict_EfuseCategory2BDF.RemoveAll
    
    '''//Initialize the array of BV and HBV testConditions by empty string "".
    '''//The vbt function "initVddBinCondition" supported multiple "Non_Binning_Rail_Outside_BinCut" sheets.
    '''20210819: Modified to move the vbt code about resetting globalVariables of BinCut testCondition from the vbt function initVddBinCondition to the vbt function Reset_BinCut_GlobalVariable_for_initVddBinning.
    '''20240520 dynamic will redim array size behind this function, move to initDynamicArray
'    For p_mode = 0 To MaxPerformanceModeCount   '0~60
'        For passBinCut = 0 To MaxPassBinCut         '0~3
'            For corePower = 0 To MaxBincutPowerdomainCount
'                BinCut(p_mode, passBinCut).OTHER_VOLTAGE(corePower) = vbNullString
'                BinCut(p_mode, passBinCut).HVCC_OTHER_VOLTAGE(corePower) = vbNullString
'                '''for OutsideBinCut sheet
'                BinCut(p_mode, passBinCut).OutsideBinCut_OTHER_VOLTAGE(corePower) = vbNullString
'                BinCut(p_mode, passBinCut).OutsideBinCut_HVCC_OTHER_VOLTAGE(corePower) = vbNullString
'
'                For addi_mode = 0 To MaxAdditionalModeCount
'                    BinCut(p_mode, passBinCut).Addtional_OTHER_VOLTAGE(corePower, addi_mode) = vbNullString
'                    BinCut(p_mode, passBinCut).HVCC_Addtional_OTHER_VOLTAGE(corePower, addi_mode) = vbNullString
'                    '''for OutsideBinCut sheet
'                    BinCut(p_mode, passBinCut).OutsideBinCut_Addtional_OTHER_VOLTAGE(corePower, addi_mode) = vbNullString
'                    BinCut(p_mode, passBinCut).OutsideBinCut_HVCC_Addtional_OTHER_VOLTAGE(corePower, addi_mode) = vbNullString
'                Next addi_mode
'            Next corePower
'        Next passBinCut
'    Next p_mode
'''ToDo: Reset the siteDouble array of Previous_Payload_Voltage by powerDomain-loop.
Exit Function
errHandler:
    Call Print_Error_Message(Error_Info, "LIB_VDD_BINNING", "Reset_BinCut_GlobalVariable_for_initVddBinning") 'Add ErrHandler 2023/08/01
    If AbortTest Then Exit Function Else Resume Next
End Function

'20200703: Created to check range of row and column for the sheet.
'Public Function check_Sheet_Range(sheetName As String, wb As Workbook, ws_def As Worksheet, MaxRow As Long, maxcol As Long, isSheetFound As Boolean, Optional enablePrint As Boolean = True)
'On Error GoTo errHandler
'    If Find_Sheet(sheetName) = True Then
'        wb.Sheets(sheetName).Unprotect
'        Set ws_def = wb.Sheets(sheetName)
'        ws_def.Select
'
'        '''//Check ranges of row and column
'        MaxRow = ws_def.Cells.SpecialCells(xlCellTypeLastCell).row
'        maxcol = ws_def.Cells.SpecialCells(xlCellTypeLastCell).Column
'
'        If MaxRow > 0 And maxcol > 0 Then
'            isSheetFound = True
'        Else
'            isSheetFound = False
'            MaxRow = 0
'            maxcol = 0
'            If enablePrint = True Then
'                TheExec.Datalog.WriteComment "Content of " & sheetName & " is empty or incorrect. Error!!!"
'                TheExec.ErrorLogMessage "Content of " & sheetName & " is empty or incorrect. Error!!!"
'            End If
'        End If
'    Else
'        isSheetFound = False
'        MaxRow = 0
'        maxcol = 0
'        If enablePrint = True Then
'            TheExec.Datalog.WriteComment sheetName & " doesn't exist in this workbook. Error!!!"
'            TheExec.ErrorLogMessage sheetName & " doesn't exist in this workbook. Error!!!"
'        End If
'    End If
'Exit Function
'errHandler:
'    TheExec.Datalog.WriteComment "Error encountered in VBT Function of check_Sheet_Range"
'    TheExec.ErrorLogMessage "Error encountered in VBT Function of check_Sheet_Range"
'    If AbortTest Then Exit Function Else Resume Next
'End Function

'20200709: Created to check if p_mode is tested.
Public Function isPmodeTested(p_mode As Integer) As Boolean
On Error GoTo errHandler 'Add ErrHandler 2023/08/01
    Dim site As Variant
    '''init
    isPmodeTested = False
    
'''ToDo: VBIN_RESULT(p_mode).tested is siteBoolean. Maybe we can use any other method to check if p_mode is tested...
    For Each site In TheExec.sites
        If VBIN_RESULT(p_mode).tested(site) = True Then
            isPmodeTested = True
            Exit For
        End If
    Next site
Exit Function
errHandler:
    Call Print_Error_Message(Error_Info, "LIB_VDD_BINNING", "isPmodeTested") 'Add ErrHandler 2023/08/01
    If AbortTest Then Exit Function Else Resume Next
End Function

'20201218: Modified to check if count_Pattern_decomposed > 0.
'20201208: Modified to merge the branches for inst_info.enable_PerEqnLog.
'20201030: Modified to use "Public Type Instance_Info".
'20201020: Modified to add the variables "COFInstance" and "PerEqnLog" for COFInstance.
'20201016: Created to decide the flag "Flag_Vddbin_COF_Instance".
Public Function decide_flag_for_COFInstance(inst_info As Instance_Info, count_pattern_decomposed As Long)
On Error GoTo errHandler
    If Flag_Vddbin_COF_Instance = True And inst_info.is_BinSearch = True = True And inst_info.enable_DecomposePatt = True Then
        If Flag_IDS_Distribution_enable = True Then
            inst_info.enable_COFInstance = False
            Call Print_Error_Message(Error_Info, "LIB_VDD_BINNING", "decide_flag_for_COFInstance", "COFInstance isn't compatible with IDS_Distribution_mode for GradeSearch_VT. Error!!!")
            'TheExec.ErrorLogMessage "COFInstance isn't compatible with IDS_Distribution_mode for GradeSearch_VT. Error!!!"    ''20230807: Removed by sanity check
        Else
            If count_pattern_decomposed > 0 Then
                inst_info.enable_COFInstance = True
                ReDim Info_COFInstance(count_pattern_decomposed - 1) '''use array size of the decoposed FuncPat.
            Else
                inst_info.enable_COFInstance = False
            End If
        End If
    Else
        inst_info.enable_COFInstance = False
    End If
    
    If inst_info.enable_COFInstance = True And Flag_Vddbin_COF_Instance_with_PerEqnLog = True Then
        inst_info.enable_PerEqnLog = True
    Else
        inst_info.enable_PerEqnLog = False
    End If
Exit Function
errHandler:
    Call Print_Error_Message(Error_Info, "LIB_VDD_BINNING", "decide_flag_for_COFInstance") 'Add ErrHandler 2023/08/01
    If AbortTest Then Exit Function Else Resume Next
End Function

'20210901: Modified to rename "StepCount As Long" as "count_Step As New SiteLong" for Public Type Instance_Info.
'20210901: Modified to rename "IndexLevelPerSite As New SiteLong" as "Step_Current As New SiteLong" for Public Type Instance_Info.
'20210125: Modified to remove "voltage_Binning_Pmode As SiteDouble" from the arguments of the vbt function "update_patt_result_for_COFInstance".
'20201210: Modified to use the arguments "inst_info As Instance_Info" and "step_control As Instance_Step_Control" for update_patt_result_for_COFInstance.
'20201016: Modfied to save EQN-based BinCut Payload voltage of binning P_mode. Requested by C651 Si Li.
'20201015: Created to save result about pattern Pass/Fail for COFInstance.
Public Function update_patt_result_for_COFInstance(inst_info As Instance_Info, indexPatt As Long, Pattern As String, pattPass As SiteBoolean)
On Error GoTo errHandler 'Add ErrHandler 2023/08/01
    Dim site As Variant
    Dim pattNameTemp As String
    Dim split_content() As String
    '''//Get pattern name while step_count=0.
    If inst_info.count_Step = 0 Then
        If Pattern Like "*\*" Then
            split_content = Split(UCase(Pattern), "\")
            pattNameTemp = split_content(UBound(split_content))
        Else
            pattNameTemp = UCase(Pattern)
        End If
        
        '20230609: Modified for COF instance
        'If split_content(UBound(split_content)) Like "*:*" Then
        If pattNameTemp Like "*:*" Then
            split_content = Split(pattNameTemp, ":")
            pattNameTemp = split_content(0)
        End If
        
        If UCase(pattNameTemp) Like "*.PAT" Then
            split_content = Split(UCase(pattNameTemp), ".PAT")
            pattNameTemp = split_content(0)
        End If
        
        If inst_info.Test_Type = testType.Mbist Then '''Mbist instance records all patterns.
            Info_COFInstance(indexPatt).is_payload_pattern = True
            Info_COFInstance(indexPatt).Pattern = pattNameTemp
        Else '''TD/Scan instance only records patterns with keywords "*_pllp*", "*_fulp*", and "*_pl*".
            If LCase(pattNameTemp) Like "*_pllp*" Or LCase(pattNameTemp) Like "*_fulp*" Or LCase(pattNameTemp) Like "*_pl*" Then
                Info_COFInstance(indexPatt).is_payload_pattern = True
                Info_COFInstance(indexPatt).Pattern = pattNameTemp
            End If
        End If
        
        Info_COFInstance(indexPatt).grade_found = False
    End If

    If Info_COFInstance(indexPatt).is_payload_pattern = True Then
        For Each site In TheExec.sites
            If pattPass = True Then
                If Info_COFInstance(indexPatt).grade_found(site) = False Then
                    Info_COFInstance(indexPatt).grade_found(site) = True
                    Info_COFInstance(indexPatt).PASSBIN = DYNAMIC_VBIN_IDS_ZONE(inst_info.p_mode).passBinCut(inst_info.Step_Current(site))
                    Info_COFInstance(indexPatt).EQN = DYNAMIC_VBIN_IDS_ZONE(inst_info.p_mode).EQ_Num(inst_info.Step_Current(site))
                    Info_COFInstance(indexPatt).Voltage = DYNAMIC_VBIN_IDS_ZONE(inst_info.p_mode).Voltage(inst_info.Step_Current(site))
                End If
            Else
                If Info_COFInstance(indexPatt).grade_found(site) = True Then
                    Info_COFInstance(indexPatt).grade_found(site) = False
                    Info_COFInstance(indexPatt).PASSBIN = -1
                    Info_COFInstance(indexPatt).EQN = -1
                End If
            End If
        Next site
    End If
Exit Function
errHandler:
    Call Print_Error_Message(Error_Info, "LIB_VDD_BINNING", "update_patt_result_for_COFInstance") 'Add ErrHandler 2023/08/01
    If AbortTest Then Exit Function Else Resume Next
End Function

'20201102: Modified to use "Public Type Instance_Info".
'20201016: Created to print info for COFInstance into the block "Judge_PF" in the datalog.
Public Function print_info_for_COFInstance(inst_info As Instance_Info)
On Error GoTo errHandler 'Add ErrHandler 2023/08/01
    Dim site As Variant
    Dim powerDomain As String
    Dim indexPatt As Long
    Dim channel As String
    If inst_info.enable_COFInstance = True Then
        powerDomain = AllBinCut(inst_info.p_mode).powerPin
    
        For indexPatt = 0 To UBound(Info_COFInstance)
            If Info_COFInstance(indexPatt).is_payload_pattern = True Then
                For Each site In TheExec.sites
                    '''//PassBin
                    If Info_COFInstance(indexPatt).grade_found(site) = True Then
                        TheExec.Datalog.WriteParametricResult site, TheExec.sites.item(site).TestNumber, logTestPass, parmTestLim, powerDomain, channel, 1, Info_COFInstance(indexPatt).PASSBIN, PassBinCut_ary(UBound(PassBinCut_ary)), _
                                                    unitNone, 0, unitNone, 0, , , Info_COFInstance(indexPatt).Pattern & "_PASSBIN", scaleNone
                    Else
                        TheExec.Datalog.WriteParametricResult site, TheExec.sites.item(site).TestNumber, logTestFail, parmTestLim, powerDomain, channel, 1, 0, PassBinCut_ary(UBound(PassBinCut_ary)), _
                                                    unitNone, 0, unitNone, 0, , , Info_COFInstance(indexPatt).Pattern & "_PASSBIN", scaleNone
                    End If
                    TheExec.sites.item(site).IncrementTestNumber
                    
                    '''//EQN
                    If Info_COFInstance(indexPatt).grade_found(site) = True Then
                        TheExec.Datalog.WriteParametricResult site, TheExec.sites.item(site).TestNumber, logTestPass, parmTestLim, powerDomain, channel, 1, Info_COFInstance(indexPatt).EQN, BinCut(inst_info.p_mode, VBIN_RESULT(inst_info.p_mode).passBinCut).Mode_Step + 1, _
                                                    unitNone, 0, unitNone, 0, , , Info_COFInstance(indexPatt).Pattern & "_EQN", scaleNone
                    Else
                        TheExec.Datalog.WriteParametricResult site, TheExec.sites.item(site).TestNumber, logTestFail, parmTestLim, powerDomain, channel, 1, 0, 1, _
                                                    unitNone, 0, unitNone, 0, , , Info_COFInstance(indexPatt).Pattern & "_EQN", scaleNone
                    End If
                    
                    '''//BinCut voltage
'''ToDo: Maybe we can replace "_CP" with the keyword about the current testJob...
                    If Info_COFInstance(indexPatt).grade_found(site) = True Then
                        TheExec.Datalog.WriteParametricResult site, TheExec.sites.item(site).TestNumber, logTestPass, parmTestLim, powerDomain, channel, 1, Info_COFInstance(indexPatt).Voltage / 1000, BinCut(inst_info.p_mode, VBIN_RESULT(inst_info.p_mode).passBinCut).CP_Vmax(0) / 1000, _
                                                    unitVolt, 0, unitVolt, 0, , , Info_COFInstance(indexPatt).Pattern & "_CP", scaleMilli, "%.4f"
                    Else
                        TheExec.Datalog.WriteParametricResult site, TheExec.sites.item(site).TestNumber, logTestFail, parmTestLim, powerDomain, channel, 1, 0, 1, _
                                                    unitVolt, 0, unitVolt, 0, , , Info_COFInstance(indexPatt).Pattern & "_CP", scaleMilli, "%.4f"
                    End If
                    TheExec.sites.item(site).IncrementTestNumber
                Next site
            End If
        Next indexPatt
    End If
Exit Function
errHandler:
    Call Print_Error_Message(Error_Info, "LIB_VDD_BINNING", "print_info_for_COFInstance") 'Add ErrHandler 2023/08/01
    If AbortTest Then Exit Function Else Resume Next
End Function

'20210901: Modified to rename "stepcountMax As Long" as "maxStep As New SiteLong" for Public Type Instance_Info.
'20210901: Modified to rename "IndexLevelPerSite As New SiteLong" as "Step_Current As New SiteLong" for Public Type Instance_Info.
'20210813: Modified to revise the vbt code for postBinCut or outsideBinCut instance names with keyword "*_binresult_" for the vbt initialize_inst_info and Apply_testcondition_InFlowSheet.
'20210809: Modified to check if Flag_Remove_Printing_BV_voltages = False for the vbt function initialize_inst_info.
'20210806: Modified to print the info about that the test instance is for BinCut search or check.
'20210805: Modified to update inst_info.is_BinSearch=True if testCondition for powerDomain of the binning p_mode contains the keyword "*evaluate*bin*".
'20210728: Modified to move "Dim Sram_Vth(MaxBincutPowerdomainCount) As New SiteDouble" into "Public Type Instance_Info".
'20210603: Modified to move inst_info.Pattern_Pmode and inst_info.By_Mode from initialize_inst_info to GradeSearch_XXX_VT.
'20210528: Modified to initalize inst_info.Pattern_Pmode and inst_info.By_Mode for Calculate_Harvest_Core_DSSC_Source.
'20210513: Modified to set inst_info.Harvest_Core_DigSrc_Pin and inst_info.Harvest_Core_DigSrc_SignalName.
'20210126: Modified to revise the vbt code for DevChar.
'20201217: Modified to initialize "inst_info.count_PrePatt_decomposed" and "inst_info.count_FuncPat_decomposed" as -1.
'20201204: Modified to initialize "inst_info.IndexLevelPerSite = -1" in the vbt function initialize_inst_info.
'20201111: Modified to initialize the siteDouble array "voltage_SelsrmBitCalc".
'20201102: Modified to add "enable_DecomposePatt" for DecomposePat.
'20201102: Modified to check if performance_mode<>"".
'20201030: Modified to move "Call Get_Pmode_Addimode_Testtype_fromInstance(inst_info)" into initialize_inst_info.
'20201029: Created to initialize inst_info for the instance.
Public Function initialize_inst_info(inst_info As Instance_Info, performance_mode As String)
On Error GoTo errHandler 'Add ErrHandler 2023/08/01
    Dim site As Variant
    Dim idx_powerDomain As Integer
    Dim str_testCondition As String
    Dim split_content() As String
    inst_info.inst_name = TheExec.DataManager.instancename

    If performance_mode <> "" Then
        inst_info.performance_mode = performance_mode
        inst_info.previousDcvsOutput = tlDCVSVoltageMain
        inst_info.currentDcvsOutput = tlDCVSVoltageMain
        inst_info.is_BV_Safe_Voltage_printed = False
        inst_info.is_BV_Payload_Voltage_printed = False
        inst_info.is_BinSearch = False                      '''True: BinSearch; False: Functional Test (Pass/Fail only).
        inst_info.enable_CMEM_Collection = False
        inst_info.enable_COFInstance = False
        inst_info.enable_PerEqnLog = False
        inst_info.enable_DecomposePatt = False
        inst_info.maxStep = 0
        inst_info.ids_current = 0
        inst_info.PrePattPass = True
        inst_info.funcPatPass = True
        inst_info.sitePatPass = True
        inst_info.count_PrePatt_decomposed = -1
        inst_info.count_FuncPat_decomposed = -1
        '''//Since Step_Current=0 is 1st step in DYNAMIC_IDS_Zone, it should initialize "Step_Current = -1" prior to "finde_start_voltage".
        inst_info.Step_Current = -1
        '''for DevChar.
        inst_info.is_DevChar_Running = TheExec.DevChar.Setups.IsRunning
        inst_info.Pattern_Pmode = vbNullString
        inst_info.By_Mode = vbNullString
        
        '''20220125,Add for check arg about harvPinGrp
        inst_info.enable_HarvPinGrp = False
        inst_info.HarvPinGrpCommnFailResult = False
        inst_info.HarvStartWithBin1 = False
        inst_info.StartBin1E1_Flag = False
        '''20230926: Added to save non-harvest pattern result in harvest instance
        inst_info.HarvPinGrp_NonHarvFail = False
        ReDim inst_info.sram_Vth(cntVddbinPin) As New SiteDouble
        ReDim inst_info.str_Selsrm_DSSC_Info(MaxSiteCount - 1) As String
        ReDim inst_info.str_Selsrm_DSSC_Bit(MaxSiteCount - 1) As String
        ReDim inst_info.str_dynamic_offset(MaxSiteCount - 1) As String
        ReDim inst_info.str_Selsrm_DSSC_Info(MaxSiteCount - 1) As String 'str_Selsrm_DSSC_Info(MaxSiteCount - 1) As String 20230526
        ReDim inst_info.str_Selsrm_DSSC_Bit(MaxSiteCount - 1) As String 'str_Selsrm_DSSC_Bit(MaxSiteCount - 1) As String 20230526
        ReDim inst_info.voltage_SelsrmBitCalc(cntVddbinPin) As New SiteDouble '''added for Selsrm Bit calculation, 20201111. 'MaxBincutPowerdomainCount
        
        If inst_info.is_DevChar_Running = True Then
            inst_info.DevChar_Setup = TheExec.DevChar.Setups.ActiveSetupName
            
            '''get_DevChar_Precondition
            If TheExec.DevChar.results(inst_info.DevChar_Setup).StartTime Like "1/1/0001*" Or TheExec.DevChar.results(inst_info.DevChar_Setup).StartTime Like "0001/1/1*" Then
                inst_info.get_DevChar_Precondition = False
            Else
                inst_info.get_DevChar_Precondition = True
            End If
        Else
            inst_info.DevChar_Setup = vbNullString
            inst_info.get_DevChar_Precondition = False
        End If
        
        '''20210728: Modified to move "Dim Sram_Vth(MaxBincutPowerdomainCount) As New SiteDouble" into "Public Type Instance_Info".
        For idx_powerDomain = 0 To cntVddbinPin
            inst_info.voltage_SelsrmBitCalc(idx_powerDomain) = 0
            inst_info.sram_Vth(idx_powerDomain) = 0
        Next idx_powerDomain
    
        For Each site In TheExec.sites
            inst_info.str_dynamic_offset(site) = vbNullString
            inst_info.str_Selsrm_DSSC_Bit(site) = vbNullString
            inst_info.str_Selsrm_DSSC_Info(site) = vbNullString
        Next site
        
        '''//Get p_mode, addi_mode, testtype, and offsettestype from test instance and its argument.
        Call Get_Pmode_Addimode_Testtype_fromInstance(inst_info)
        
        '''//Check if testCondition of the binning powerDomain contains the keyword "*evaluate*bin*".
        '''20210805: Modified to update inst_info.is_BinSearch=True if testCondition for powerDomain of the binning p_mode contains the keyword "*evaluate*bin*".
        For Each site In TheExec.sites.Active
            If inst_info.powerDomain <> "" Then
                str_testCondition = LCase(Trim(Get_BinCut_TestCondition(inst_info, VddBinStr2Enum(inst_info.powerDomain), CurrentPassBinCutNum(site))))
                            
                '''20210813: Modified to revise the vbt code for postBinCut or outsideBinCut instance names with keyword "*_binresult_" for the vbt initialize_inst_info and Apply_testcondition_InFlowSheet.
                '''*************************************************************************************'''
                '''//Keyword replacement of BinCut test condition of p_mode.
                '''//The flag "is_BinCutJob_for_StepSearch" is True if any testCondition from the table "Non_Binning_Rail" has the keyword "*Evaluate*Bin*".
                '''Since PostBinCut_Voltage_Set_VT support BV and HBV tests, it should replace keyword of testCondition from BV with "bin result".
                '''//For the special case, ex: "900mV (MS003)", do not replace the keyword of testCondtion with with "*Bin*Result*".
                '''20181009: As the request from KTCHAN, he defined that postbincut instances must have the keyword "*_binresult_*".
                '''20210126: Modified to revise the vbt code for DevChar.
                '''20210302: Modified to optimize the keyword replacement to "M*### Bin Result".
                If is_BinCutJob_for_StepSearch = True Then '''only for postBinCut or outsideBincut in CP1.
                    If (LCase(inst_info.inst_name) Like "*_binresult_*" And Not (str_testCondition) Like "*bin*result*" And Not (str_testCondition) Like "*#mv*") Or inst_info.is_DevChar_Running = True Then
                        split_content = Split(str_testCondition, " ")
                        
                        '''//Check if any correct keyword of performance_mode exists...
                        If VddbinPmodeDict.Exists(UCase(Trim(split_content(0)))) Then
                            '''//p_mode of non_binning CorePower.
                            If VBIN_RESULT(VddBinStr2Enum(UCase(split_content(0)))).tested = True Then
                                str_testCondition = split_content(0) & " " & "bin result"
                            End If
                        End If
                    End If
                End If
                '''*************************************************************************************'''
                
                If str_testCondition <> "" Then
                    If str_testCondition Like "*evaluate*bin*" Then
                        inst_info.is_BinSearch = True
                    Else
                        inst_info.is_BinSearch = False
                    End If
                    Exit For
                Else
                    inst_info.is_BinSearch = False
                    TheExec.sites.item(site).FlagState(strGlb_Flag_Vddbinning_Fail_Stop) = logicTrue
                    Call Print_Error_Message(Error_Info, "LIB_VDD_BINNING", "initialize_inst_info", "Instance: " & inst_info.inst_name & "," & inst_info.performance_mode & ",powerDomain:" & inst_info.powerDomain & ", testCondition is incorrect to determine if the instance is for BinCut search or check. Please check argument of the instance. Error!!!")
                    'TheExec.ErrorLogMessage "Instance: " & inst_info.inst_name & "," & inst_info.performance_mode & ",powerDomain:" & inst_info.powerDomain & ", testCondition is incorrect to determine if the instance is for BinCut search or check. Please check argument of the instance. Error!!!"    ''20230807: Removed by sanity check
                End If
            Else
                TheExec.Datalog.WriteComment "This instance is SAA test instance without binsearch!!! "
                inst_info.is_BinSearch = False
                Exit For
            End If
        Next site
        
        '''//Check if the test instance is for BinCut search or check.
        '''20210806: Modified to print the info about that the test instance is for BinCut search or check.
        '''20210809: Modified to check if Flag_Remove_Printing_BV_voltages = False for the vbt function initialize_inst_info.
        If Flag_Remove_Printing_BV_voltages = False Then
            If inst_info.is_BinSearch = True Then
                TheExec.Datalog.WriteComment "instance:" & inst_info.inst_name & "," & inst_info.performance_mode & ", the instance is for BinCut search"
            Else
                TheExec.Datalog.WriteComment "instance:" & inst_info.inst_name & "," & inst_info.performance_mode & ", the instance is for BinCut check"
            End If
        End If
    Else
        Call Print_Error_Message(Error_Info, "LIB_VDD_BINNING", "initialize_inst_info", "Instance: " & inst_info.inst_name & " doesn't have the correct performance_mode. Please check argument of the instance. Error!!!")
        'TheExec.ErrorLogMessage "Instance: " & inst_info.inst_name & " doesn't have the correct performance_mode. Please check argument of the instance. Error!!!"    ''20230807: Removed by sanity check
    End If
Exit Function
errHandler:
    Call Print_Error_Message(Error_Info, "LIB_VDD_BINNING", "initialize_inst_info") 'Add ErrHandler 2023/08/01
    If AbortTest Then Exit Function Else Resume Next
End Function

'20210903: Modified to merge properties of "Public Type Instance_Step_Control" into Public Type Instance_Info "Public Type Instance_Info".
'20210805: Modified to remove the redundant vbt function initialize_step_control since it initialized step_control.All_Site_Mask = 0 in the vbt function decide_binSearch_and_start_voltage.
'20201211: Created to initialize control flags from "inst_info" and "step_control" at the beginning of each step in step-loop.
Public Function initialize_control_flag_for_step_loop(inst_info As Instance_Info)
On Error GoTo errHandler
    inst_info.PrePattPass = True                    'initail the flag for init pattern
    inst_info.funcPatPass = True
    inst_info.sitePatPass = True
    inst_info.is_BV_Safe_Voltage_printed = False
    inst_info.is_BV_Payload_Voltage_printed = False
    inst_info.Grade_Not_Found_Mask = 0           'grade not found flag for all site
    inst_info.On_StopVoltage_Mask = 0            'already on stop voltage flag for all site
    inst_info.All_Patt_Pass = True               'initialize the flag for all sites.
    inst_info.AllSiteFailPatt = 0
Exit Function
errHandler:
    Call Print_Error_Message(Error_Info, "LIB_VDD_BINNING", "initialize_control_flag_for_step_loop") 'Add ErrHandler 2023/08/01
    If AbortTest Then Exit Function Else Resume Next
End Function

'20210901: Modified to check inst_info.HarvestBinningFlag for HarvestBinning.
'20201218: Modified to remove "enable_CMEM_collection As Boolean" from the arguments of the vbt function "check_flag_to_enable_CMEM_collection".
'20201218: Modified to add "count_FuncPat_decomposed As Long" to the arguments of the vbt function "check_flag_to_enable_CMEM_collection".
'20201218: Modified to move "resize_CMEM_Data_by_pattern_number" from "decide_bincut_feature_for_stepsearch" to "check_flag_to_enable_CMEM_collection".
'20201217: Created to decide if BinCut features are OK to be enabled for BinCut stepSearch.
Public Function decide_bincut_feature_for_stepsearch(inst_info As Instance_Info, count_FuncPat_decomposed As Long, Optional CaptureSize As Long, Optional failpins As String, Optional inst_CallInstance As String = vbNullString)
On Error GoTo errHandler
    '''//Check the flag "Flag_Enable_CMEM_Collection" to enable CMEM collection if tester is online.
    '''If inst_info.enable_CMEM_Collection = True, check and decide CaptureSize, failpins, and PrintSize for CMEM.
    If inst_CallInstance = "" Then
        Call check_flag_to_enable_CMEM_collection(inst_info, Flag_Enable_CMEM_Collection, count_FuncPat_decomposed, CaptureSize, failpins)
    Else
        ''20221220: If callinst, do not collect CMEM
        inst_info.enable_CMEM_Collection = False
    End If
    
    '''20220125,Add for check arg about harvPinGrp
    If inst_CallInstance = "" Then  ''220831: Skip for callinst instance
        Call check_arg_to_enable_HarvPinGrp(inst_info)
    End If
    
    '''//Decide if it's OK to enable COFInstance. If that, redim array size to store payload patterns pass/fail.
    Call decide_flag_for_COFInstance(inst_info, count_FuncPat_decomposed)
    
    '''//Checkscript uses this info to check if BinCut new features are activated or not.
    '''20210901: Modified to check inst_info.HarvestBinningFlag for HarvestBinning.
    TheExec.Datalog.WriteComment "Instance_Condition" & ", COFInstance:" & CStr(inst_info.enable_COFInstance) & ", PerEqnLog:" & CStr(inst_info.enable_PerEqnLog) & _
                                    ", Enable_CMEM_Collection:" & CStr(inst_info.enable_CMEM_Collection) & ", HarvestBinningFlag:" & inst_info.HarvestBinningFlag & _
                                    ", Harvest_With_BIN1E1:" & inst_info.Harvest_With_BIN1E1
Exit Function
errHandler:
    Call Print_Error_Message(Error_Info, "LIB_VDD_BINNING", "decide_bincut_feature_for_stepsearch") 'Add ErrHandler 2023/08/01
    If AbortTest Then Exit Function Else Resume Next
End Function

'20210730: Modified to show the error message to users if p_mode is for BinCut search without any Efuse category, as requested by C651 Toby.
'20210701: Created to update AllBinCut(p_mode).used after parsing BinCut flow table and Efuse_BitDef_Table.
Public Function update_bincut_pmode_list()
On Error GoTo errHandler 'Add ErrHandler 2023/08/01
    Dim idx_powerDomain As Long
    Dim powerDomain As String
    Dim split_content() As String
    Dim str_pmodeGroup_list As String
    Dim idx_pmode As Long
    Dim str_pmode_temp As String
    Dim str_Efuse_write_pmode As String
    Dim p_mode As Long, sram_mode_cnt As Long
    Dim Power_Seq() As String
    Dim Power_Seq_Tmp As String

    BinCut_Sram_Power_Seq.RemoveAll
    BinCut_Sram_Power_KeyMapping.RemoveAll
'''//==================================================================================================================================================================================//'''
'''//Note:
'''allbincut(p_mode).used is decided after parsing BinCut flow table and Efuse_BitDef_Table, and it can check if p_mode can be tested and fused for BinCut...
'''//==================================================================================================================================================================================//'''
    For idx_powerDomain = 0 To UBound(pinGroup_BinCut)
        '''//Get the BinCut powerDomain.
        powerDomain = pinGroup_BinCut(idx_powerDomain)
        
        If gb_bincut_power_list(VddBinStr2Enum(powerDomain)) <> "" Then
            '''init
            str_pmodeGroup_list = vbNullString
        
            '''//Get array of performance modes for powerDomain.
            split_content = Split(gb_bincut_power_list(VddBinStr2Enum(powerDomain)), ",")
            
            '''//Check AllBinCut(p_mode).listed_in_Efuse_BDF and update AllBinCut(p_mode).used for all performance modes from BinCut flow table.
            '''This step can make sure that the performances exists in BinCut flow table and Efuse_BitDef_Table definitely.
            For idx_pmode = 0 To UBound(split_content)
                str_pmode_temp = split_content(idx_pmode)
                
                '''//Check if the performance mode is listed in Efuse_BitDef_Table.
                '''ToDo: Maybe we can add the option here to skip checking AllBinCut(VddBinStr2Enum(str_pmode_temp)).listed_in_Efuse_BDF...
                If AllBinCut(VddBinStr2Enum(str_pmode_temp)).listed_in_Efuse_BDF = True Then
                    '''//If that, Update AllBinCut(p_mode).used for the performance mode.
                    AllBinCut(VddBinStr2Enum(str_pmode_temp)).Used = True
                    
                    '''//Add the performance mode to the new pmode_list.
                    If str_pmodeGroup_list <> "" Then
                        str_pmodeGroup_list = str_pmodeGroup_list & "," & str_pmode_temp
                    Else
                        str_pmodeGroup_list = str_pmode_temp
                    End If
                End If
                
                '''//Check if p_mode for BinCut search has the dedicated Efuse category in the current testJob.
                ''20210730: Modified to show the error message to users if p_mode is for BinCut search without any Efuse category, as requested by C651 Toby.
                If AllBinCut(VddBinStr2Enum(str_pmode_temp)).is_for_BinSearch = True And Not LCase(TheExec.CurrentJob) = "cp1" Then
                    str_Efuse_write_pmode = get_Efuse_category_by_BinCut_testJob("write", VddBinName(VddBinStr2Enum(str_pmode_temp)))
                    
                    If str_Efuse_write_pmode = "" Then
                        Call Print_Error_Message(Error_Info, "LIB_VDD_BINNING", "update_bincut_pmode_list", str_pmode_temp & ", it is for BinCut search, but it doesn't have any Efuse category about the performance mode. Please check Efuse_BitDef_Table and BinCut flow table. Error!!!")
                        'TheExec.Datalog.WriteComment str_pmode_temp & ", it is for BinCut search, but it doesn't have any Efuse category about the performance mode. Please check Efuse_BitDef_Table and BinCut flow table. Error!!!"
                        'TheExec.ErrorLogMessage str_pmode_temp & ", it is for BinCut search, but it doesn't have any Efuse category about the performance mode. Please check Efuse_BitDef_Table and BinCut flow table. Error!!!"    ''20230807: Removed by sanity check
                    End If
                End If
            Next idx_pmode
            
            '''//Update the new pmodeGroup_list to gb_bincut_power_list for the BinCut powerDomain.
            gb_bincut_power_list(VddBinStr2Enum(powerDomain)) = str_pmodeGroup_list
            
            '''*******************************************************************************************************************'''
            '''//Sort the Performance mode by MAX_ID to define the inherit sequence for different PowerDomain.
            '''//Enable the Performance mode by the Flow Table.
            '''*******************************************************************************************************************'''
            If gb_bincut_power_list(VddBinStr2Enum(powerDomain)) <> "" Then
                sort_power_seqence gb_bincut_power_list(VddBinStr2Enum(powerDomain)), BinCut_Power_Seq(VddBinStr2Enum(powerDomain)).Power_Seq
            Else
                BinCut_Power_Seq(VddBinStr2Enum(powerDomain)).Power_Seq = Split(gb_bincut_power_list(VddBinStr2Enum(powerDomain)), ",")
            End If
            
            '''*******************************************************************************************************************'''
            '''Check the performance_mode, and determine its previous performance_mode from the power_seq for voltage inheritance.
            '''*******************************************************************************************************************'''
            InitVddBinInherit BinCut_Power_Seq(VddBinStr2Enum(powerDomain)).Power_Seq
        Else
            '20211022: Added to support SRAM Voltage calculate with equation
            If dict_IsCorePower(powerDomain) = True And UCase(powerDomain) Like "*VDD*_SRAM*" Then
                sram_mode_cnt = 0
                Power_Seq_Tmp = vbNullString

                For p_mode = 0 To MaxPerformanceModeCount - 1
                    If AllBinCut(p_mode).powerPin = powerDomain Then
                        If sram_mode_cnt > 0 Then
                            If Power_Seq_Tmp <> mid(Replace(VddBinName(p_mode), powerDomain & "_", ""), 1, 3) Then
                                If Not BinCut_Sram_Power_Seq.Exists(Power_Seq_Tmp) Then
                                    BinCut_Sram_Power_Seq.Add Power_Seq_Tmp, Power_Seq
                                End If
                                sram_mode_cnt = 0
                                ReDim Power_Seq(sram_mode_cnt)
                            End If
                        End If
                        ReDim Preserve Power_Seq(sram_mode_cnt)
                        Power_Seq(sram_mode_cnt) = Replace(VddBinName(p_mode), powerDomain & "_", "")
                        Power_Seq_Tmp = mid(Power_Seq(sram_mode_cnt), 1, 3)
                        sram_mode_cnt = sram_mode_cnt + 1
                    End If
                Next p_mode
                If Not BinCut_Sram_Power_Seq.Exists(Power_Seq_Tmp) Then
                    BinCut_Sram_Power_Seq.Add Power_Seq_Tmp, Power_Seq
                    If Not BinCut_Sram_Power_KeyMapping.Exists(powerDomain) Then
                        BinCut_Sram_Power_KeyMapping.Add powerDomain, Power_Seq_Tmp
                    End If
                End If
            End If
        End If
    Next idx_powerDomain
Exit Function
errHandler:
    Call Print_Error_Message(Error_Info, "LIB_VDD_BINNING", "update_bincut_pmode_list") 'Add ErrHandler 2023/08/01
    If AbortTest Then Exit Function Else Resume Next
End Function

'20210831: Modified to remove the unused vbt code.
'20210720: Modified to revise the branches.
'20210706: Created to get Efuse category by BinCut testJob.
Public Function get_Efuse_category_by_BinCut_testJob(str_selector As String, str_keyword_EfuseCategory As String) As String
On Error GoTo errHandler 'Add ErrHandler 2023/08/01
    Dim i As Long
    Dim idx_CurrentBinCutJob As Long
    Dim strAry_temp_EfuseCategory() As String
    Dim str_Efuse_temp As String
    Dim gotCorrectSelector As Boolean
    '''init
    str_Efuse_temp = vbNullString
    
    '''//Check if str_selector is "read" or "write".
    If LCase(str_selector) = "read" Or LCase(str_selector) = "write" Then
        '''//Get BinCutJob definition for the current BinCut testJob.
        idx_CurrentBinCutJob = getBinCutJobDefinition(bincutJobName)
    
        '''//Get Efuse category for str_keyword_EfuseCategory.
        If dict_strPmode2EfuseCategory.Exists(UCase(str_keyword_EfuseCategory)) = True Then
            strAry_temp_EfuseCategory = dict_strPmode2EfuseCategory.item(UCase(str_keyword_EfuseCategory))
            
            '''//Check the array of Efuse category related to str_keyword_EfuseCategory.
            For i = 0 To UBound(strAry_temp_EfuseCategory)
                '''//Check if any Efuse category is fused prior to the current BinCut testJob.
                If dict_EfuseCategory2BinCutTestJob.Exists(strAry_temp_EfuseCategory(i)) = True Then
                    If LCase(str_selector) = "read" Then
                        If idx_CurrentBinCutJob > dict_EfuseCategory2BinCutTestJob.item(strAry_temp_EfuseCategory(i)) Then
                            If str_Efuse_temp <> "" Then
                                If dict_EfuseCategory2BinCutTestJob.item(strAry_temp_EfuseCategory(i)) > dict_EfuseCategory2BinCutTestJob.item(str_Efuse_temp) Then
                                    str_Efuse_temp = strAry_temp_EfuseCategory(i)
                                End If
                            Else '''If str_selected_EfuseCategory is empty...
                                str_Efuse_temp = strAry_temp_EfuseCategory(i)
                            End If
                        End If
                    ElseIf LCase(str_selector) = "write" Then '20220711
                        If idx_CurrentBinCutJob = dict_EfuseCategory2BinCutTestJob.item(strAry_temp_EfuseCategory(i)) And LCase(dict_EfuseCategory2BDF.item(strAry_temp_EfuseCategory(i))) = LCase(TheExec.CurrentJob) Then
                            str_Efuse_temp = strAry_temp_EfuseCategory(i)
                        End If
                    End If
                Else
                    str_Efuse_temp = vbNullString
                    'TheExec.Datalog.WriteComment "BinCut_testJob:" & bincutJobName & ",keyword_EfuseCategory:" & str_keyword_EfuseCategory & ", it doesn't have any correct programming stage to get the Efuse category for get_Efuse_category_by_BinCut_testJob. Please check Efuse_BitDef_Table. Error!!!"
                    'TheExec.ErrorLogMessage "BinCut_testJob:" & bincutJobName & ",keyword_EfuseCategory:" & str_keyword_EfuseCategory & ", it doesn't have any correct programming stage to get the Efuse category for get_Efuse_category_by_BinCut_testJob. Please check Efuse_BitDef_Table. Error!!!"
                End If
            Next i
        Else
            str_Efuse_temp = vbNullString
            Call Print_Error_Message(Error_Info, "LIB_VDD_BINNING", "get_Efuse_category_by_BinCut_testJob", "BinCut_testJob:" & bincutJobName & ",keyword_EfuseCategory:" & str_keyword_EfuseCategory & ", it isn't a correct keyword to find the Efuse category for get_Efuse_category_by_BinCut_testJob. Please check Efuse_BitDef_Table. Error!!!")
            'TheExec.ErrorLogMessage "BinCut_testJob:" & bincutJobName & ",keyword_EfuseCategory:" & str_keyword_EfuseCategory & ", , it isn't a correct keyword to find the Efuse category for get_Efuse_category_by_BinCut_testJob. Please check Efuse_BitDef_Table. Error!!!"    ''20230807: Removed by sanity check
        End If
    Else
        str_Efuse_temp = vbNullString
        Call Print_Error_Message(Error_Info, "LIB_VDD_BINNING", "get_Efuse_category_by_BinCut_testJob", "str_selector:" & str_selector & ", it isn't 'read' or 'write' to get the Efuse category for get_Efuse_category_by_BinCut_testJob. Please check Efuse_BitDef_Table. Error!!!")
        'TheExec.ErrorLogMessage "str_selector:" & str_selector & ", it isn't 'read' or 'write' to get the Efuse category for get_Efuse_category_by_BinCut_testJob. Please check Efuse_BitDef_Table. Error!!!"    ''20230807: Removed by sanity check
    End If
    
    '''//Output the string.
    get_Efuse_category_by_BinCut_testJob = str_Efuse_temp
Exit Function
errHandler:
    Call Print_Error_Message(Error_Info, "LIB_VDD_BINNING", "get_Efuse_category_by_BinCut_testJob") 'Add ErrHandler 2023/08/01
    If AbortTest Then Exit Function Else Resume Next
End Function


'20210118 Add for preadjustvddbinning binout only when no step could use
Public Function find_next_bin_eq_interpolation_fullsteps(p_mode As Integer, site As Variant, gradevdd_last As Variant, grade_last As Variant)
On Error GoTo errHandler 'Add ErrHandler 2023/08/01
    Dim idx_step As Long
    Dim next_bin_flag As Boolean
    Dim voltage_in_dynamic_ids_zone As Double
    Dim exit_while_flag As Boolean
    Dim printflag As Boolean
    '''init
    next_bin_flag = False
    printflag = False

    If CDec(VBIN_RESULT(p_mode).GRADE) > 0 Then
        idx_step = VBIN_RESULT(p_mode).step_in_IDS_Zone
        exit_while_flag = False
                
        If AllBinCut(p_mode).PREVIOUS_Performance_Mode = AllBinCut(p_mode).Allow_Equal And AllBinCut(p_mode).Allow_Equal <> 0 Then '''for AllowEqual
            '''//Note: If the vbt of checking GRADE is masked, please set globalVariable "Public Const Flag_Only_Check_PV_for_VoltageHeritage As Boolean = True".
            '''//Print the status of VBIN_RESULT(p_mode).is_Monotonicity_Offset_triggered(site) with PTR format in Adjust_Binning for datalogs.
            '''20210526: C651 Si revised the check rules to ensure that: (GradeVDD(P_mode)-GradeVDD(previous perfromance_mode))> Monotonicity_Offset.
            If Flag_Get_column_Monotonicity_Offset = True Then
                If (CDec(VBIN_RESULT(p_mode).GRADEVDD - VBIN_RESULT(AllBinCut(p_mode).PREVIOUS_Performance_Mode).GRADEVDD) < CDec(BinCut(p_mode, DYNAMIC_VBIN_IDS_ZONE(p_mode).passBinCut(idx_step)).Monotonicity_Offset(DYNAMIC_VBIN_IDS_ZONE(p_mode).EQ_Num(idx_step) - 1))) Then
                    VBIN_RESULT(p_mode).is_Monotonicity_Offset_triggered(site) = True
                    TheExec.Datalog.WriteComment "site:" & site & "," & VddBinName(p_mode) & "=" & VBIN_RESULT(p_mode).GRADEVDD
                End If
            End If
            
            While (CDec(VBIN_RESULT(p_mode).GRADEVDD) < CDec(gradevdd_last + BinCut(p_mode, DYNAMIC_VBIN_IDS_ZONE(p_mode).passBinCut(idx_step)).Monotonicity_Offset(DYNAMIC_VBIN_IDS_ZONE(p_mode).EQ_Num(idx_step) - 1)) And exit_while_flag = False)
            'Or (CDec(VBIN_RESULT(p_mode).GRADE) < CDec(grade_last) And exit_while_flag = False)
                idx_step = idx_step + 1
                If idx_step > AllBinCut(p_mode).Mode_Step Then
                    '''20210610: Modified to revise the message for the vbt function find_next_bin_eq_interpolation if it can't find any step to define the product voltage.
'                    TheExec.Datalog.WriteComment "The Performance Mode " & VddBinName(p_mode) & " EQ 1 Product Value is wrong. Error!!!"
                    Call Print_Error_Message(Error_Info, "LIB_VDD_BINNING", "", "site:" & site & "," & VddBinName(p_mode) & ",bin=" & CurrentPassBinCutNum(site) & ",but it can't find any step to define the product voltage. Error!!!")
'                    TheExec.Flow.TestLimit resultVal:=999, lowVal:=1, hiVal:=1, Tname:="find_next_bin_eq_interpolation Error"
                    '''20240303: Use modular function to bin out site
                    Call BinoutSite(site, VddbinFailStopSortNumber, VddbinFailStopBinNumber, strGlb_Flag_Vddbinning_Fail_Stop)

                    exit_while_flag = True
                    Exit Function
                Else
                    If CurrentPassBinCutNum < DYNAMIC_VBIN_IDS_ZONE(p_mode).passBinCut(idx_step) Then
                        CurrentPassBinCutNum(site) = DYNAMIC_VBIN_IDS_ZONE(p_mode).passBinCut(idx_step)
                    End If
                    voltage_in_dynamic_ids_zone = DYNAMIC_VBIN_IDS_ZONE(p_mode).Voltage(idx_step)
                    VBIN_RESULT(p_mode).GRADE = voltage_in_dynamic_ids_zone
                    VBIN_RESULT(p_mode).GRADEVDD = voltage_in_dynamic_ids_zone + BinCut(p_mode, VBIN_RESULT(p_mode).passBinCut).CP_GB(DYNAMIC_VBIN_IDS_ZONE(p_mode).EQ_Num(idx_step) - 1)
                    next_bin_flag = True
                    printflag = True
                End If
            Wend
        Else
            '''//Note: If the vbt of checking GRADE is masked, please set globalVariable "Public Const Flag_Only_Check_PV_for_VoltageHeritage As Boolean = True".
            '''20210526: C651 Si revised the check rules to ensure that: (GradeVDD(P_mode)-GradeVDD(previous perfromance_mode))> Monotonicity_Offset.
            '''20210610: Modified for the rule: (GradeVDD(P_mode)-GradeVDD(previous perfromance_mode))> Monotonicity_Offset if without Allow_Equal.
            If Flag_Get_column_Monotonicity_Offset = True Then
                If (CDec(VBIN_RESULT(p_mode).GRADEVDD - VBIN_RESULT(AllBinCut(p_mode).PREVIOUS_Performance_Mode).GRADEVDD) <= CDec(BinCut(p_mode, DYNAMIC_VBIN_IDS_ZONE(p_mode).passBinCut(idx_step)).Monotonicity_Offset(DYNAMIC_VBIN_IDS_ZONE(p_mode).EQ_Num(idx_step) - 1))) Then
                    VBIN_RESULT(p_mode).is_Monotonicity_Offset_triggered(site) = True
                    TheExec.Datalog.WriteComment "site:" & site & "," & VddBinName(p_mode) & "=" & VBIN_RESULT(p_mode).GRADEVDD
                End If
            End If
            
            'PC update for allow grade = previous grade 1012
            While (CDec(VBIN_RESULT(p_mode).GRADEVDD) <= CDec(gradevdd_last + BinCut(p_mode, DYNAMIC_VBIN_IDS_ZONE(p_mode).passBinCut(idx_step)).Monotonicity_Offset(DYNAMIC_VBIN_IDS_ZONE(p_mode).EQ_Num(idx_step) - 1)) And exit_while_flag = False) _
            'Or (CDec(VBIN_RESULT(p_mode).GRADE) < CDec(grade_last) And exit_while_flag = False)
                idx_step = idx_step + 1
                If idx_step > AllBinCut(p_mode).Mode_Step Then
                    '''20210610: Modified to revise the message for the vbt function find_next_bin_eq_interpolation if it can't find any step to define the product voltage.
'                    TheExec.Datalog.WriteComment "The Performance Mode " & VddBinName(p_mode) & " EQ 1 Product Value is wrong. Error!!!"
                    Call Print_Error_Message(Error_Info, "LIB_VDD_BINNING", "", "site:" & site & "," & VddBinName(p_mode) & ",bin=" & CurrentPassBinCutNum(site) & ",but it can't find any step to define the product voltage. Error!!!")
'                    TheExec.Flow.TestLimit resultVal:=999, lowVal:=1, hiVal:=1, Tname:="find_next_bin_eq_interpolation Error"
                    '''20240303: Use modular function to bin out site
                    Call BinoutSite(site, VddbinFailStopSortNumber, VddbinFailStopBinNumber, strGlb_Flag_Vddbinning_Fail_Stop)
                    
                    exit_while_flag = True
                    Exit Function
                Else
                    
                    If CurrentPassBinCutNum < DYNAMIC_VBIN_IDS_ZONE(p_mode).passBinCut(idx_step) Then
                        CurrentPassBinCutNum(site) = DYNAMIC_VBIN_IDS_ZONE(p_mode).passBinCut(idx_step)
                    End If
                    voltage_in_dynamic_ids_zone = DYNAMIC_VBIN_IDS_ZONE(p_mode).Voltage(idx_step)
                    VBIN_RESULT(p_mode).GRADE = voltage_in_dynamic_ids_zone
                    VBIN_RESULT(p_mode).GRADEVDD = voltage_in_dynamic_ids_zone + BinCut(p_mode, VBIN_RESULT(p_mode).passBinCut).CP_GB(DYNAMIC_VBIN_IDS_ZONE(p_mode).EQ_Num(idx_step) - 1)
                    next_bin_flag = True
                    printflag = True
                End If
            Wend
        End If
        
        If next_bin_flag = True Then
            VBIN_RESULT(p_mode).step_in_BinCut = DYNAMIC_VBIN_IDS_ZONE(p_mode).EQ_Num(idx_step) - 1
            VBIN_RESULT(p_mode).step_in_IDS_Zone = idx_step
        End If
          
        If printflag = True Then
            TheExec.Datalog.WriteComment "site" & TheExec.sites.siteNumber & " Adjust EQ for " & VddBinName(p_mode) & ":" & AllBinCut(p_mode).powerPin & "=" & VBIN_RESULT(p_mode).GRADE & ", Binnum : " & CStr(DYNAMIC_VBIN_IDS_ZONE(p_mode).passBinCut(idx_step))
        End If
    End If
Exit Function
errHandler:
    Call Print_Error_Message(Error_Info, "LIB_VDD_BINNING", "find_next_bin_eq_interpolation_fullsteps") 'Add ErrHandler 2023/08/01
    If AbortTest Then Exit Function Else Resume Next
End Function

'20200118: Add for preAdjustVddbinning
Public Function check_pmode_for_pre_adjust_VddBinning(site As Variant, core_power_seq() As String, ids_PowerDomain As SiteDouble)
On Error GoTo errHandler 'Add ErrHandler 2023/08/01
    Dim p_mode As Integer
    Dim gradevdd_last As New SiteDouble
    Dim grade_last As New SiteDouble
    Dim i As Long
    '''//Only CP1 test has adjust_VddBinning.
    If UBound(core_power_seq) >= 1 Then
        gradevdd_last = VBIN_RESULT(VddBinStr2Enum(core_power_seq(0))).GRADEVDD
        grade_last = VBIN_RESULT(VddBinStr2Enum(core_power_seq(0))).GRADE
        
        '''//Adjust the vdd binning value to make voltage is always higher on higher performance mode for each p_mode of BinCut powerPin
        For i = 1 To UBound(core_power_seq)
            p_mode = VddBinStr2Enum(core_power_seq(i))
            gradevdd_last = VBIN_RESULT(AllBinCut(VddBinStr2Enum(core_power_seq(i))).PREVIOUS_Performance_Mode).GRADEVDD
            grade_last = VBIN_RESULT(AllBinCut(VddBinStr2Enum(core_power_seq(i))).PREVIOUS_Performance_Mode).GRADE
            
            If BinCut(p_mode, VBIN_RESULT(p_mode).passBinCut).ExcludedPmode = True Then
                '''Set Grade and GradeVdd to 0 if the Performance Mode is not enabled in the Test Result (Bin1 or BinX).
                VBIN_RESULT(p_mode).GRADE = 0
                VBIN_RESULT(p_mode).step_in_BinCut = -1
                VBIN_RESULT(p_mode).GRADEVDD = 0
            
            '''//For complete BinCut search steps, we replace "find_next_bin_eq" with "find_next_bin_eq_interpolation" according to BinCut monthly meeting Dec-2018.
            Else '''20190227: Modified to use DYNAMIC_VBIN_IDS_ZONE for complete BinCut search steps.
                find_next_bin_eq_interpolation_fullsteps p_mode, site, gradevdd_last, grade_last
            End If
        Next i
    End If
Exit Function
errHandler:
    Call Print_Error_Message(Error_Info, "LIB_VDD_BINNING", "check_pmode_for_pre_adjust_VddBinning") 'Add ErrHandler 2023/08/01
    If AbortTest Then Exit Function Else Resume Next
End Function

'''20220125,Add for check arg about harvPinGrp
Public Function check_arg_to_enable_HarvPinGrp(inst_info As Instance_Info)
On Error GoTo errHandler
    Dim i As Long, j As Long
    Dim strAry_PathSplit As Variant, strAry_PatNameSplit As Variant
    Dim patcnt As Long
    Dim patkeyword As String, patkeywordtmp As String
    Dim site As Variant 'Carter, 20240304
    'HarvPinGrp_Condition = P:ECPU_All(F_ECPU_All);P:ECPU_CORE0(F_ECPU_CORE0);P:ECPU_CORE1(F_ECPU_CORE1), HarvPinGrp_ComFailFlag = F_ECPU_All
    'If inst_info.enable_HarvPinGrp = True Then Stop
    If Flag_VDD_Binning_Offline = True Or TheExec.flow.enableWord("Vddbinning_OpenSocket") = True Then '''If the tester is offline or opensocket.
        inst_info.enable_HarvPinGrp = False
    ElseIf Flag_HarvPinFlag_Mapping_Table_Parsed = True Then        ''20220712: HarvPinGrp Table exist
        If inst_info.enable_HarvPinGrp = True Then
            'Only bin1 need to do harv update result check
            For Each site In TheExec.sites
                If CurrentPassBinCutNum(site) > 1 Then
                    inst_info.HarvStartWithBin1(site) = False
                Else
                    inst_info.HarvStartWithBin1(site) = True
                End If
            Next site
        End If
        
            '''Seperate HarvGrp conditions and flags
        If Not inst_info.count_FuncPat_decomposed = -1 Then
            patcnt = 0
            inst_info.patCnt_HarvPinGrp = 0
            patkeyword = vbNullString
            patkeywordtmp = vbNullString
            ReDim inst_info.HarvPinGrpCheckData(0)
            
            For i = 0 To UBound(inst_info.ary_FuncPat_decomposed)
                strAry_PathSplit = Split(LCase(inst_info.ary_FuncPat_decomposed(i)), "\")
                strAry_PatNameSplit = Split(LCase(strAry_PathSplit(UBound(strAry_PathSplit))), "_")
                'Only payload pat need to check harv pin
                If strAry_PatNameSplit(3) Like "*pl*" Then
                    For j = 0 To UBound(HarvPinFlagMapping)
                        If LCase(strAry_PathSplit(UBound(strAry_PathSplit))) Like LCase(HarvPinFlagMapping(j).Pattern) Then
                            patkeyword = HarvPinFlagMapping(j).Pattern
                            If patkeyword <> patkeywordtmp Then
                                ReDim Preserve inst_info.HarvPinGrpCheckData(patcnt)
                                
                                inst_info.HarvPinGrpCheckData(patcnt).keyword = HarvPinFlagMapping(j).Pattern
                                inst_info.HarvPinGrpCheckData(patcnt).MappingIndex = j
                                ReDim Preserve inst_info.HarvPinGrpCheckData(patcnt).HarvFailResult(UBound(HarvPinFlagMapping(j).HarvPinGrpConditionArr))
                                patcnt = patcnt + 1
                            End If
                            patkeywordtmp = patkeyword
                            Exit For
                        End If
                    Next j
                End If
            Next i
        End If
            
            If patcnt = 0 Then
                inst_info.enable_HarvPinGrp = False
                inst_info.patCnt_HarvPinGrp = 0
            Else
                inst_info.patCnt_HarvPinGrp = patcnt
            End If
        
        'Only CMEM_collection turn off need to do, or it done before
        If inst_info.enable_CMEM_Collection = False Then
            Call TheHdw.Digital.CMEM.SetCaptureConfig(0, CmemCaptNone) ' Resets CMEM
            TheHdw.Digital.CMEM.CentralFields = tlCMEMPatternName + _
                                        tlCMEMVMVectorOffset + _
                                        tlCMEMModCycle
        End If
    End If

Exit Function
errHandler:
    Call Print_Error_Message(Error_Info, "LIB_VDD_BINNING", "check_arg_to_enable_HarvPinGrp") 'Add ErrHandler 2023/08/01
    If AbortTest Then Exit Function Else Resume Next
End Function

'''20220125,Use CMEM to update HarvPinGrp result
Public Function update_HarvPinGrp_Result_UF(Harv_Condition As String, Harv_FailResult As SiteBoolean, HarvFailPinCnt As SiteLong, Optional ShowFailDetail As Boolean = False)
On Error GoTo errHandler 'Add ErrHandler 2023/08/01
    Dim i As Long, j As Long, k As Long
    Dim vsite As Variant
    Const Cond_Case = 0
    Const Cond_Pin = 0
'   Const ShowFailDetail_MaxNum = 1 '-1 for all
    Const ShowFailDetail_MaxNum = 511 ' Setting same as Ellis
    
    Dim str_InputCondArr() As String
    Dim str_CondArr() As String
    Dim str_CheckPin As String
    Dim str_CheckPinsArr() As String
    Dim lng_CheckPinsNum As Long
    Dim str_CheckPins_ReUnite As String
    
    Dim Sbln_PatternPass As New SiteBoolean
    Dim patternNames() As String
    Dim lastFailPerPattern() As Long
    Dim dbl_CmemCycleDataArr() As Double 'Data type is Double because cycle count is 40 bits.
    Dim dbl_CmemVectorDataArr() As Double
    Dim dbl_MinCmemIndex As Double
    Dim SiteIndexData  As New SiteVariant
    Dim SitePinData  As New SiteVariant
    Dim indexArr() As Long 'Used in site loop
    Dim pinDataArr() As Double 'Used in site loop
    Dim bigPinList As Integer 'Used when pinlist is greater than 32.
    Dim patIdx As Long
    Dim pinStr As String
    Dim PinDict As New Dictionary
    Dim lng_FailDetailStartIndex As Long
    
    Dim FailCount As New PinListData
                   
    'EX :Harv_Condition = P:JTAG_TDO:323,325,330;V:EUSB_P1:0.05;F:EUSB_P2:3000000;I:EUSB_P1:0~0.05
    
    If (Harv_Condition <> "") Then
        str_InputCondArr = Split(Harv_Condition, ";")
        
        For i = 0 To UBound(str_InputCondArr)
            str_CondArr = Split(str_InputCondArr(i), ":")
            'EX:Harv_Condition = P:ECPU_CORE0
            str_CheckPin = str_CondArr(Cond_Pin)
            Call TheExec.DataManager.DecomposePinList(str_CheckPin, str_CheckPinsArr, lng_CheckPinsNum)
            str_CheckPins_ReUnite = Join(str_CheckPinsArr, ",")
            
            Sbln_PatternPass = TheHdw.Digital.Patgen.PatternBurstPassedPerSite
            If Sbln_PatternPass.Any(False) Then
                'Read back the last index where data was captured for each pattern.
                Call TheHdw.Digital.CMEM.PatternName(lastFailPerPattern, patternNames)
                'Read back the central data.
                dbl_CmemCycleDataArr = TheHdw.Digital.CMEM.PatGenInfo(tlCMEMModCycle, -1)
                dbl_CmemVectorDataArr = TheHdw.Digital.CMEM.PatGenInfo(tlCMEMVMVectorOffset, -1)
                dbl_MinCmemIndex = Min(UBound(dbl_CmemCycleDataArr), UBound(dbl_CmemVectorDataArr))
                'Read failing pins for all sites, compressing the failures.
                '20220620: Size = dbl_MinCmemIndex -> -1 for UF and UFP
                Call TheHdw.Digital.pins(str_CheckPins_ReUnite).CMEM.StoredCycleData(SiteIndexData, SitePinData, -1, True)
                For Each vsite In TheExec.sites.Selected
                    If Sbln_PatternPass(vsite) = False Then
                        'Move siteVariant to normal array of Long/Double
                        indexArr = SiteIndexData
                        pinDataArr = SitePinData
                        'pat fail at pinA, pinB. but capture pinC => UBound(IndexArr) = -1
                        If UBound(indexArr) <> -1 Then
                            
                            Harv_FailResult(vsite) = True
                            
                            'being here means capture pins did fail
                            If CurrentPassBinCutNum(vsite) = 1 Then
                                TheExec.Datalog.WriteComment "Site" & vsite & " Pin : " & str_CheckPin & " total fail cycle = " & CStr(UBound(indexArr) + 1) & "."
                            End If
                                
                            If ShowFailDetail And CurrentPassBinCutNum(vsite) = 1 Then
                                'bigPinlist will equal 1 for <32 pins; 2 for 64 to 33 pins; and so on.
                                bigPinList = (UBound(pinDataArr) + 1) \ (UBound(indexArr) + 1)
                                patIdx = 0
                                If ShowFailDetail_MaxNum > -1 Then
                                    TheExec.Datalog.WriteComment "Print out all fail cycle."
                                    lng_FailDetailStartIndex = LBound(indexArr)
                                ElseIf UBound(indexArr) > ShowFailDetail_MaxNum Then
                                    TheExec.Datalog.WriteComment "Fail Cycle over " & ShowFailDetail_MaxNum & ", Only print the last " & ShowFailDetail_MaxNum & "."
                                    lng_FailDetailStartIndex = UBound(indexArr) - ShowFailDetail_MaxNum
                                Else
                                    lng_FailDetailStartIndex = LBound(indexArr)
                                End If
                                
                                For j = lng_FailDetailStartIndex To UBound(indexArr)
                                    pinStr = vbNullString
                                    PinDict.RemoveAll
                                    For k = 0 To bigPinList - 1
                                            pinStr = FailingPins(pinDataArr(j * bigPinList + k), str_CheckPinsArr, k, PinDict) & "," & pinStr
                                    Next k
                                    If patIdx > UBound(patternNames) Then
                                            'Raise error. This should never happen.
                                    End If
                                    If j > lastFailPerPattern(patIdx) Then
                                            'Go to the next pattern in the list.
                                            patIdx = patIdx + 1
                                    End If
                                    
                                    If j >= ShowFailDetail_MaxNum Then Exit For
                                    
                                    If j < ShowFailDetail_MaxNum Then
                                            TheExec.Datalog.WriteComment "vector : " & CStr(dbl_CmemVectorDataArr(indexArr(j))) + _
                                            " cycle : " & CStr(dbl_CmemCycleDataArr(indexArr(j))) + _
                                            " pin : " & pinStr
                                    End If
                                Next j
                            End If
                                    
                            FailCount = TheHdw.Digital.pins(str_CheckPins_ReUnite).FailCount
                            For j = 0 To FailCount.pins.Count - 1
                                If FailCount.pins.item(j).value <> 0 Then
                                    HarvFailPinCnt(vsite) = HarvFailPinCnt(vsite) + 1
                                End If
                            Next j
                        Else
                            'CMEM capture pins didnot fail
                            TheExec.Datalog.WriteComment "Site" & vsite & " Pat Fail, but Pin : " & str_CheckPin & " did not fail."
                        End If
                    End If
                Next vsite
            End If
        Next
    Else
        Call Print_Error_Message(Error_Info, "LIB_VDD_BINNING", "update_HarvPinGrp_Result_UF", "Input Harv_Condition it is blank. Please check Harv_Condition. Error!!!")
        'TheExec.ErrorLogMessage "Input Harv_Condition it is blank. Please check Harv_Condition. Error!!!"    ''20230807: Removed by sanity check
    End If
Exit Function 'Add ErrHandler 2023/08/01
errHandler: 'Add ErrHandler 2023/08/01
    Call Print_Error_Message(Error_Info, "LIB_VDD_BINNING", "update_HarvPinGrp_Result_UF") 'Add ErrHandler 2023/08/01
    If AbortTest Then Exit Function Else Resume Next 'Add ErrHandler 2023/08/01
End Function
'''20220928, Split into 2 function "update_HarvPinGrp_Result_UF" and "update_HarvPinGrp_Result_UFP" for IG-XL 10.30
'''20220125, Use CMEM to update HarvPinGrp result
Public Function update_HarvPinGrp_Result_UFP(Harv_Condition As String, Harv_FailResult As SiteBoolean, HarvFailPinCnt As SiteLong, Optional ShowFailDetail As Boolean = False)
On Error GoTo errHandler 'Add ErrHandler 2023/08/01
    Dim i As Long, j As Long, k As Long
    Dim vsite As Variant
    Const Cond_Case = 0
    Const Cond_Pin = 0
'   Const ShowFailDetail_MaxNum = 1 '-1 for all
    Const ShowFailDetail_MaxNum = 511 ' Setting same as Ellis
    
    Dim str_InputCondArr() As String
    Dim str_CondArr() As String
    Dim str_CheckPin As String
    Dim str_CheckPinsArr() As String
    Dim lng_CheckPinsNum As Long
    Dim str_CheckPins_ReUnite As String
    
    Dim Sbln_PatternPass As New SiteBoolean
    Dim patternNames() As String
    Dim lastFailPerPattern() As Long
    '' Mod 220610 use site variant instead of Double array
    Dim dbl_CmemCycleDataArr As New SiteVariant
    Dim dbl_CmemVectorDataArr As New SiteVariant
    Dim dbl_MinCmemIndex As Double
    Dim SiteIndexData  As New SiteVariant
    Dim SitePinData  As New SiteVariant
    Dim indexArr() As Long 'Used in site loop
    Dim pinDataArr() As Double 'Used in site loop
    Dim bigPinList As Integer 'Used when pinlist is greater than 32.
    Dim patIdx As Long
    Dim pinStr As String
    Dim PinDict As New Dictionary
    Dim lng_FailDetailStartIndex As Long
    
    Dim FailCount As New PinListData
                   
    'EX :Harv_Condition = P:JTAG_TDO:323,325,330;V:EUSB_P1:0.05;F:EUSB_P2:3000000;I:EUSB_P1:0~0.05
    
    If (Harv_Condition <> "") Then
        str_InputCondArr = Split(Harv_Condition, ";")
        
        For i = 0 To UBound(str_InputCondArr)
            str_CondArr = Split(str_InputCondArr(i), ":")
            'EX:Harv_Condition = P:ECPU_CORE0
            str_CheckPin = str_CondArr(Cond_Pin)
            Call TheExec.DataManager.DecomposePinList(str_CheckPin, str_CheckPinsArr, lng_CheckPinsNum)
            str_CheckPins_ReUnite = Join(str_CheckPinsArr, ",")
            
            Sbln_PatternPass = TheHdw.Digital.Patgen.PatternBurstPassedPerSite
            If Sbln_PatternPass.Any(False) Then
                For Each vsite In TheExec.sites.Selected
                    'Read back the last index where data was captured for each pattern.
                    Call TheHdw.Digital.CMEM.PatternName(lastFailPerPattern, patternNames)
                    'Read back the central data.
                    dbl_CmemCycleDataArr = TheHdw.Digital.CMEM.PatGenInfo(tlCMEMModCycle, -1)
                    dbl_CmemVectorDataArr = TheHdw.Digital.CMEM.PatGenInfo(tlCMEMVMVectorOffset, -1)
                    dbl_MinCmemIndex = Min(UBound(dbl_CmemCycleDataArr(vsite)), UBound(dbl_CmemVectorDataArr(vsite)))
                    'Read failing pins for all sites, compressing the failures.
                    '20220620: Size = dbl_MinCmemIndex -> -1 for UF and UFP
                    Call TheHdw.Digital.pins(str_CheckPins_ReUnite).CMEM.StoredCycleData(SiteIndexData, SitePinData, -1, True)
                    If Sbln_PatternPass(vsite) = False Then
                        'Move siteVariant to normal array of Long/Double
                        indexArr = SiteIndexData
                        pinDataArr = SitePinData
                        'pat fail at pinA, pinB. but capture pinC => UBound(IndexArr) = -1
                        If UBound(indexArr) <> -1 Then
                            
                            Harv_FailResult(vsite) = True
                            
                            'being here means capture pins did fail
                            If CurrentPassBinCutNum(vsite) = 1 Then
                                TheExec.Datalog.WriteComment "Site" & vsite & " Pin : " & str_CheckPin & " total fail cycle = " & CStr(UBound(indexArr) + 1) & "."
                            End If
                                
                            If ShowFailDetail And CurrentPassBinCutNum(vsite) = 1 Then
                                'bigPinlist will equal 1 for <32 pins; 2 for 64 to 33 pins; and so on.
                                bigPinList = (UBound(pinDataArr) + 1) \ (UBound(indexArr) + 1)
                                patIdx = 0
                                If ShowFailDetail_MaxNum > -1 Then
                                    TheExec.Datalog.WriteComment "Print out all fail cycle."
                                    lng_FailDetailStartIndex = LBound(indexArr)
                                ElseIf UBound(indexArr) > ShowFailDetail_MaxNum Then
                                    TheExec.Datalog.WriteComment "Fail Cycle over " & ShowFailDetail_MaxNum & ", Only print the last " & ShowFailDetail_MaxNum & "."
                                    lng_FailDetailStartIndex = UBound(indexArr) - ShowFailDetail_MaxNum
                                Else
                                    lng_FailDetailStartIndex = LBound(indexArr)
                                End If
                                
                                For j = lng_FailDetailStartIndex To UBound(indexArr)
                                    pinStr = vbNullString
                                    PinDict.RemoveAll
                                    For k = 0 To bigPinList - 1
                                            pinStr = FailingPins(pinDataArr(j * bigPinList + k), str_CheckPinsArr, k, PinDict) & "," & pinStr
                                    Next k
                                    If patIdx > UBound(patternNames) Then
                                            'Raise error. This should never happen.
                                    End If
                                    If j > lastFailPerPattern(patIdx) Then
                                            'Go to the next pattern in the list.
                                            patIdx = patIdx + 1
                                    End If
                                    
                                    If j >= ShowFailDetail_MaxNum Then Exit For
                                    
                                    If j < ShowFailDetail_MaxNum Then
                                            TheExec.Datalog.WriteComment "vector : " & CStr(dbl_CmemVectorDataArr(vsite)(indexArr(j))) + _
                                            " cycle : " & CStr(dbl_CmemCycleDataArr(vsite)(indexArr(j))) + _
                                            " pin : " & pinStr
                                    End If
                                Next j
                            End If
                                    
                            FailCount = TheHdw.Digital.pins(str_CheckPins_ReUnite).FailCount
                            For j = 0 To FailCount.pins.Count - 1
                                If FailCount.pins.item(j).value <> 0 Then
                                    HarvFailPinCnt(vsite) = HarvFailPinCnt(vsite) + 1
                                End If
                            Next j
                        Else
                            'CMEM capture pins didnot fail
                            TheExec.Datalog.WriteComment "Site" & vsite & " Pat Fail, but Pin : " & str_CheckPin & " did not fail."
                        End If
                    End If
                Next vsite
            End If
        Next
    Else
        Call Print_Error_Message(Error_Info, "LIB_VDD_BINNING", "update_HarvPinGrp_Result_UFP", "Input Harv_Condition it is blank. Please check Harv_Condition. Error!!!")
        'TheExec.ErrorLogMessage "Input Harv_Condition it is blank. Please check Harv_Condition. Error!!!"    ''20230807: Removed by sanity check
    End If
Exit Function 'Add ErrHandler 2023/08/01
errHandler: 'Add ErrHandler 2023/08/01
    Call Print_Error_Message(Error_Info, "LIB_VDD_BINNING", "update_HarvPinGrp_Result_UFP") 'Add ErrHandler 2023/08/01
    If AbortTest Then Exit Function Else Resume Next 'Add ErrHandler 2023/08/01
End Function

'''20220208,HarvPinGrp ControlFailFlag for binout
Public Function CtrlFailFlag_byHarvPinGrpResult(inst_info As Instance_Info, Failflag As String)
On Error GoTo errHandler 'Add ErrHandler 2023/08/01
    Dim site As Variant
    Dim HarvCnt As Long, HarvPatCnt As Long
    If inst_info.enable_HarvPinGrp = True Then
        For Each site In TheExec.sites
            '''20220125,judge bin1 eq1 harv result. if only harv fail use bin1eq1 result and turn on harv flag
            '''20230926: Added to save non-harvest pattern result in harvest instance
            If inst_info.HarvPinGrpCommnFailResult(site) = True Or inst_info.HarvPinGrp_NonHarvFail(site) = True Then
                '''20230821: Modified to support multi harvest flag, other pin group
                If inst_info.HarvPinGrpCommnFailResult(site) = True Then
                    Call Set_MultiHarvestFlag_ForBV(OtherPin, site, , inst_info.performance_mode)
                End If
                If inst_info.HarvPinGrp_NonHarvFail(site) = True Then
                    Call Set_MultiHarvestFlag_ForBV(NonHarvest, site, , inst_info.performance_mode)
                End If
                
                '''//Shut down the failed site. As per discussion with Chihome, he suggested us to ensure that Sort Number/Bin Number/fail-stop should be updated before .result = tlResultFail.
                'TheExec.sites.item(site).result = tlResultFail
                TheExec.sites.item(site).FlagState(Failflag) = logicTrue
                TheExec.sites.item(site).FlagState(strGlb_Flag_Vddbinning_Fail_Stop) = logicTrue
            Else
                For HarvPatCnt = 0 To UBound(inst_info.HarvPinGrpCheckData)
                    For HarvCnt = 0 To UBound(HarvPinFlagMapping(inst_info.HarvPinGrpCheckData(HarvPatCnt).MappingIndex).HarvPinGrpConditionArr)
'                        If HarvPinFlagMapping(inst_info.HarvPinGrpCheckData(HarvPatCnt).MappingIndex).HarvPinGrpConditionArr(HarvCnt) <> "" Then
                        If inst_info.HarvPinGrpCheckData(HarvPatCnt).HarvFailResult(HarvCnt)(site) = True Then
                            '''20230821: Modified to support multi harvest flag, other pin group
                            Call Set_MultiHarvestFlag_ForBV(PinHarvest, site, HarvPinFlagMapping(inst_info.HarvPinGrpCheckData(HarvPatCnt).MappingIndex).HarvPinGrpCondFailFlagArr(HarvCnt))
                            
                            TheExec.sites.item(site).FlagState(HarvPinFlagMapping(inst_info.HarvPinGrpCheckData(HarvPatCnt).MappingIndex).HarvPinGrpCondFailFlagArr(HarvCnt)) = logicTrue
                        End If
                    Next HarvCnt
                Next HarvPatCnt
            End If
        Next site
        '20220512, HVCC end cmem setting
        TheHdw.Digital.CMEM.SetCaptureConfig 0, CmemCaptNone '''CmemCaptNone: Capture no cycles.
        TheHdw.Digital.CMEM.CentralFields = tlCMEMNone  '20220512, add for turn off cmem capture
    End If
Exit Function
errHandler:
    Call Print_Error_Message(Error_Info, "LIB_VDD_BINNING", "CtrlFailFlag_byHarvPinGrpResult") 'Add ErrHandler 2023/08/01
    If AbortTest Then Exit Function Else Resume Next
End Function

'20220621: Parse All BinTables and save Sortbin & Hardbin
Public Function Parsing_BinTable_ForBinCut()
On Error GoTo errHandler 'Add ErrHandler 2023/08/01
    Dim wb As Workbook
    Dim ws_def As Worksheet
    Dim MaxRow As Long
    Dim maxcol As Long
    Dim Row As Long, Col As Long
    Dim sort_num As Long, bin_num As Long
    Dim spec_name As String, result_type As String
    Dim Row_of_Title As Integer
    Dim col_name As Integer, col_sort As Integer, col_bin As Integer, col_result As Integer
    Dim enableParsing As Boolean
    Dim sheet_idx As Integer
    Dim strAry_sheetnames() As String

    Sortbin2hardbinDict.RemoveAll
    Sortbin2PmodeDict.RemoveAll
    Sortbin2ResultDict.RemoveAll
    Flag_BinTable_Parsed_ForBinCut = False
    Set wb = Application.ActiveWorkbook
    '' Find all Bin Tables
    strAry_sheetnames = TheExec.job.GetSheetNamesOfType(DMGR_SHEET_TYPE_BINTABLESSHEET)
    
    For sheet_idx = 0 To UBound(strAry_sheetnames)
        Set ws_def = wb.Sheets(strAry_sheetnames(sheet_idx))
        'Init
        Row_of_Title = 0
        col_name = 0
        col_sort = 0
        col_bin = 0
        enableParsing = False
        
        '''//Check ranges of row and column
        MaxRow = ws_def.Cells.SpecialCells(xlCellTypeLastCell).Row
        maxcol = ws_def.Cells.SpecialCells(xlCellTypeLastCell).Column
        If MaxRow < 1 Or maxcol < 1 Then GoTo next_sheet
        
        '''//Check the header of the table.
        For Row = 1 To MaxRow
            For Col = 1 To maxcol
                If UCase(ws_def.Cells(Row, Col).value) = "NAME" Then
                '' Finds "name" & "sort" & "bin" on BinTable
                    Row_of_Title = Row
                    col_name = Col
                ElseIf UCase(ws_def.Cells(Row, Col).value) = "SORT" Then
                    col_sort = Col
                ElseIf UCase(ws_def.Cells(Row, Col).value) = "BIN" Then
                    col_bin = Col
                ElseIf UCase(ws_def.Cells(Row, Col).value) = "RESULT" Then
                    col_result = Col
                End If
                '' Finds all header, exit loop
                If Row_of_Title > 0 And col_name > 0 And col_sort > 0 And col_bin > 0 And col_result > 0 Then Exit For
            Next Col
        Next Row
        
        If Row_of_Title > 0 And col_name > 0 And col_sort > 0 And col_bin > 0 Then
            enableParsing = True
        Else
            TheExec.Datalog.WriteComment "Can't find correct header 'Name'/'Sort'/'Bin' in " & strAry_sheetnames(sheet_idx) & " !!!"
            enableParsing = False
            GoTo next_sheet
        End If
        
        If enableParsing = True Then
            '''//Start parsing the cells
            For Row = Row_of_Title + 1 To MaxRow

                sort_num = ws_def.Cells(Row, col_sort).value                '' Sortbin
                bin_num = ws_def.Cells(Row, col_bin).value                  '' Hardbin
                spec_name = ws_def.Cells(Row, col_name).value               '' Pmode
                result_type = UCase(ws_def.Cells(Row, col_result).value)    '' Result: "PASS"/"Fail"/"Fail-Stop"
                
                '' End of bintable
                If sort_num = 0 Or bin_num = 0 Or spec_name = vbNullString Or result_type = vbNullString Then Exit For
                
                '' Sortbin doesn't duplicates
                If Not (Sortbin2hardbinDict.Exists(sort_num)) And Not (Sortbin2PmodeDict.Exists(sort_num)) And Not Sortbin2ResultDict.Exists(sort_num) Then
                    '' SortBin -> HardBin
                    Sortbin2hardbinDict.Add sort_num, bin_num
                    
                    '' SortBin -> Pmode             '' Get second element: Pmode
                    If UCase(spec_name) Like "*_HBV" Then
                        spec_name = Split(spec_name, "_")(2)
                    ElseIf UCase(spec_name) Like "*_BV" Then
                        spec_name = Split(spec_name, "_")(1)
                    End If
                    Sortbin2PmodeDict.Add sort_num, spec_name
                    
                    '' SortBin -> Result
                    Sortbin2ResultDict.Add sort_num, result_type
                    
                    Flag_BinTable_Parsed_ForBinCut = True
                Else
                    '' Sortbin duplicates
                    If sort_num = 971 Or sort_num = 972 Or sort_num = 973 Or sort_num = 974 Or sort_num = 975 Or sort_num = 976 Or sort_num = 977 Or sort_num = 978 Then
                    Else
                        Call Print_Error_Message(Error_Info, "LIB_VDD_BINNING", "Parsing_BinTable_ForBinCut", "Error!! Sortbin: " & sort_num & ", duplicates in " & strAry_sheetnames(sheet_idx))
                        If isDebugMode = True Then
                            TheExec.AddOutput "Error!! Sortbin: " & sort_num & ", duplicates in " & strAry_sheetnames(sheet_idx)
                        End If
                    End If
                End If
            Next Row
        Else
            TheExec.Datalog.WriteComment "Can't find correct header 'Name'/'Sort'/'Bin'/'Result' in " & strAry_sheetnames(sheet_idx) & "!!"
            Exit Function
        End If
next_sheet:
    Next sheet_idx
Exit Function
errHandler:
    Call Print_Error_Message(Error_Info, "LIB_VDD_BINNING", "Parsing_BinTable_ForBinCut") 'Add ErrHandler 2023/08/01
    If AbortTest Then Exit Function Else Resume Next
End Function

'''20221125: Use pattern pins for CMEM capture, "TheHdw.Digital.Patterns.Files().DigitalPins.List" for UFP only.
Public Function GetPatternPinStr(str_pattern As String, str_patternPins As String)
On Error GoTo errHandler 'Add ErrHandler 2023/08/01
    Dim PatternPinArr() As String
    Dim str_pattern_nameonly As String
    Dim arr_Patname() As String
    
    If str_pattern = "" Then
        str_patternPins = vbNullString
        Exit Function
    End If
    
    str_pattern_nameonly = Split(str_pattern, ":")(0)           '' remove module name
    arr_Patname = Split(str_pattern_nameonly, "\")
    str_pattern_nameonly = arr_Patname(UBound(arr_Patname))     '' str_pattern_nameonly = "PP_XXX.patx"
    
    '''20230201: Added for ffc on UFP, save pattern name
    If UCase(str_pattern_nameonly) Like "*.PAT" Then
        str_pattern_nameonly = Replace(UCase(str_pattern_nameonly), "PAT", "PATX")
    End If
    
    ''Use this command to get pattern pins
    PatternPinArr = TheHdw.Digital.patterns.Files(str_pattern_nameonly).DigitalPins.list
    str_patternPins = Join(PatternPinArr, ",")     ''Save all pins a,b,c
    
Exit Function
errHandler:
    Call Print_Error_Message(Error_Info, "LIB_VDD_BINNING", "GetPatternPinStr") 'Add ErrHandler 2023/08/01
    If AbortTest Then Exit Function Else Resume Next
End Function

''20221223: For print voltage output(Vmain/Valt) in debug print mode
Public Function Print_BincutVoltageOutput_AfterPattern()
On Error GoTo errHandler 'Add ErrHandler 2023/08/01
    Dim VoltageOutput As Long
    Dim Temp_PinStr As String
    Dim Temp_PinAry() As String
    Dim i As Long
    Dim str_Vmain As String, str_Valt As String
    Dim cnt_Vmain As Long, cnt_Valt As Long

    VoltageOutput = 0
    Temp_PinStr = vbNullString
    cnt_Vmain = 0
    cnt_Valt = 0
    str_Vmain = vbNullString
    str_Valt = vbNullString
    
    ''Compose all bincut power pins
    For i = 0 To UBound(pinGroup_BinCut)
        If domain2pinDict.Exists(UCase(pinGroup_BinCut(i))) = True Then     ''Decompose PinGroup to Pins
            '' VddbinDomain2Pin("VDD_FIXED_ALL") = "VDD_FIXED_AMUX,VDD_FIXED_CPU,VDD_FIXED_GRP,VDD_FIXED_LPDP_RX....."
            '' VddbinDomain2Pin("VDD_Low_ALL") = "VDD_LOW,VDD_LOW_FLPPLL,VDD_LOW_ULPPLL,VDD_LOW_USB_DEBUG"
            '' VddbinDomain2Pin("VDD_SOC") = "VDD_SOC"
            If Temp_PinStr = "" Then
                Temp_PinStr = VddbinDomain2Pin(pinGroup_BinCut(i))
            Else
                Temp_PinStr = Temp_PinStr & "," & VddbinDomain2Pin(pinGroup_BinCut(i))
            End If
        Else
            TheExec.Datalog.WriteComment pinGroup_BinCut(i) & " is not a Bincut power domain!"
        End If
    Next i
    Temp_PinAry = Split(Temp_PinStr, ",")   ''Save all bincut power pins to PinAry
    
    ''Find voltage output and compose print string
    For i = 0 To UBound(Temp_PinAry)
        VoltageOutput = TheHdw.DCVS.pins(Temp_PinAry(i)).Voltage.Output
        
        ''Compose string
        If VoltageOutput = 1 Then       ''Vmain
            cnt_Vmain = cnt_Vmain + 1
            If cnt_Vmain > 7 Then
                str_Vmain = str_Vmain & vbCrLf
                cnt_Vmain = 0
            End If
            str_Vmain = str_Vmain & Temp_PinAry(i)
            If i <> UBound(Temp_PinAry) Then str_Vmain = str_Vmain & ", "
        ElseIf VoltageOutput = 2 Then   ''Valt
            cnt_Valt = cnt_Valt + 1
            If cnt_Valt > 7 Then
                str_Valt = str_Valt & vbCrLf
                cnt_Valt = 0
            End If
            str_Valt = str_Valt & Temp_PinAry(i)
            If i <> UBound(Temp_PinAry) Then str_Valt = str_Valt & ", "
        Else    ''Cannot find voltage output
            str_Vmain = str_Vmain & vbCrLf & Temp_PinAry(i) & " = Cannot Find Voltage Output!!" & vbCrLf
        End If
    Next i
    
    ''Print out voltage output result
    TheExec.Datalog.WriteComment "--------------------Bincut Voltage Output start--------------------"
    TheExec.Datalog.WriteComment "---Vmain Pin: " & str_Vmain
    TheExec.Datalog.WriteComment "---Valt Pin: " & str_Valt
    TheExec.Datalog.WriteComment "-------------------- Bincut Voltage Output end --------------------"
    
    ''Reset
    Temp_PinStr = vbNullString
    str_Vmain = vbNullString
    str_Valt = vbNullString
Exit Function
errHandler:
    Call Print_Error_Message(Error_Info, "LIB_VDD_BINNING", "Print_BincutVoltageOutput_AfterPattern") 'Add ErrHandler 2023/08/01
    If AbortTest Then Exit Function Else Resume Next 'Add ErrHandler 2023/08/01
End Function
Public Function CountMaxAdditionModeCnt(keyword_sheetName As String)
    Dim i As Long
    Dim count_WorkSheet As Integer
    Dim BincutSheetsArr() As String
    Dim idxSheet As Integer
    Dim maxSheetsNum As Integer
    Dim sheetName As String
    Dim TotalAdditionMode As New Dictionary
    Dim splittmp As Variant
On Error GoTo errHandler
'''//==================================================================================================================================================================================//'''
'''//Note:
'''Please check sheetName keyword "Non_Binning_Rail_Outside" of sheetName for the vbt functions "initVddBinCondition" and "parsing_OutsideBinCut_flow_table".
'''//==================================================================================================================================================================================//'''
    '''init
    idxSheet = -1
    TotalAdditionMode.RemoveAll
    '''//Check if the sheet exists.
    If keyword_sheetName = "Non_Binning_Rail" Then
        count_WorkSheet = Application.ActiveWorkbook.Worksheets.Count
        
        '''//Check if sheet name contains keyword_sheetName.
        For i = 1 To count_WorkSheet
            If ActiveWorkbook.Worksheets(i).name Like "*" & keyword_sheetName & "*" Then
                idxSheet = idxSheet + 1
                ReDim Preserve BincutSheetsArr(idxSheet)
                BincutSheetsArr(idxSheet) = ActiveWorkbook.Worksheets(i).name
            End If
        Next i
    ElseIf keyword_sheetName <> "" Then
        TheExec.Datalog.WriteComment keyword_sheetName & " is not the correct keyword to find Outside BinCut flow table for parsing_OutsideBinCut_flow_table. Error!!!"
        TheExec.ErrorLogMessage keyword_sheetName & " is not the correct keyword to find Outside BinCut flow table for parsing_OutsideBinCut_flow_table. Error!!!"
        Exit Function
    Else '''If keyword_sheetName is empty...
        Exit Function
    End If
    
    '''//Use sheet-loop to parse each Outside BinCut flow table.
    If idxSheet > -1 Then
        For i = 0 To idxSheet
            sheetName = BincutSheetsArr(i)
            
            '''//Parsing each of sheets with Keyword in sheetName.
            CntAdditionMode sheetName, TotalAdditionMode
        Next i
    End If

    MaxAdditionalModeCount = TotalAdditionMode.Count
    TotalAdditionMode.RemoveAll
Exit Function
errHandler:
    TheExec.Datalog.WriteComment "Error encountered in VBT Function of initVddBinCondition"
    TheExec.ErrorLogMessage "Error encountered in VBT Function of initVddBinCondition"
    If AbortTest Then Exit Function Else Resume Next
End Function
Public Function CountMaxEquationVoltagesMode(sheetName As String)
    Dim sheetName_EquationVoltages As String
    Dim sheetName_temp As String
    Dim Row As Long, Col As Long
    Dim MaxRow As Long
    Dim maxcol As Long
    Dim EquationVoltages_temp As String
    '''for testjob mapping
    Dim Row_of_Title As Long
    Dim row_of_testJob As Long
    '''
    Dim ws_def As Worksheet
    Dim isSheetFound As Boolean, isParsingOutsideBinCutFlow As Boolean
    'Dim cntEqVoltagePmode As Long ''20231102 Removed for correct dynamic stucture usage
On Error GoTo errHandler
'''//==================================================================================================================================================================================//'''
'''//Note:
'''1. Non_Binning_Rail should be parsed prior to sheet Non_Binning_Rail_Post_BinCut.
'''2. As per discussion with TSMC PCLINZG, he suggested us to use the same testCondition for outsideBinCutFlow BV and HBV.
'''3. Please check keyword "Non_Binning_Rail_Outside" of sheetName for the vbt functions "initVddBinCondition" and "parsing_OutsideBinCut_flow_table".
'''//==================================================================================================================================================================================//'''
    '''*****************************************************************'''
    '''//Check if the sheet exists
    'sheetName = "Non_Binning_Rail"
    sheetName_temp = sheetName
    sheetName = sheetName & "_" & TheExec.CurrentJob
    sheetName_EquationVoltages = "EquationVoltages" '''for EquationVoltages.
    dict_EquationVoltages_ALL.RemoveAll
    Set wb = Application.ActiveWorkbook
    Call check_Sheet_Range(sheetName, wb, ws_def, MaxRow, maxcol, isSheetFound, False)
    If isSheetFound = False Then                        ''' If "Non_Binning_Rail_JOBNAME" does not exist
        sheetName = sheetName_temp
        Call check_Sheet_Range(sheetName, wb, ws_def, MaxRow, maxcol, isSheetFound)
    End If
    '''*****************************************************************'''
    
    If isSheetFound = True Then
        '''//init
        '''Since all col_XXX and row_XXX related variables with default values=0, no need to initialize them as 0.
        Row_of_Title = 0
        row_of_testJob = 0
    Else
        Exit Function
    End If
    
    For Row = 1 To MaxRow
        For Col = 1 To maxcol
            If LCase(ws_def.Cells(Row, Col).value) Like LCase("DC Spec") Then
                If row_of_testJob = 0 Then
                    row_of_testJob = Row
                    Exit For
                End If
            End If
        Next Col
        
        If row_of_testJob > 0 Then
            Exit For
        End If
    Next Row

    Row_of_Title = row_of_testJob + 2
  
    If Row_of_Title > 0 Then '''It means that columns of BinCut powerDomains are found.
        While LCase(ws_def.Cells(Row, 1).value) <> "end"
            If Row > 2 Then
                EquationVoltages_temp = UCase(ws_def.Cells(Row, 1).value)
                If Not UCase(EquationVoltages_temp) Like "*EQN*" Then
                    TheExec.ErrorLogMessage "EquationVoltages don't have EQN Error!!!"
                    TheExec.Datalog.WriteComment "EquationVoltages don't have EQN Error!!! "
                Else
                    cntEqVoltagePmode = cntEqVoltagePmode + 1
                End If
            End If
            Row = Row + 1
        Wend
    End If '''If row_of_title > 0
Exit Function
errHandler:
    TheExec.Datalog.WriteComment "Error encountered in VBT Function of initVddBinCondition"
    TheExec.ErrorLogMessage "Error encountered in VBT Function of initVddBinCondition"
    If AbortTest Then Exit Function Else Resume Next
End Function
Public Function CountMaxIDS_DIS_ZoneCnt()
    Dim sheetName As String
    Dim MaxRow As Long
    Dim maxcol As Long
    Dim Row As Long, Col As Long
    Dim p_mode As Integer
    Dim idx_step As Long
    Dim Test_Type As Long
    Dim col_testTypeSelected As Long
    Dim zoneTemp As Long
    Dim Row_of_Title As Integer
    Dim Row_of_step0 As Integer
    Dim col_TestType(testType.Total_Count) As Integer 'Dim col_TestType(testType.TotalCnt) As Integer
    Dim strTemp As String
    Dim isSheetFound As Boolean
    Dim ws_def As Worksheet
    'Dim Max_IDS_Dis_Zone As Long
On Error GoTo errHandler
    '''*****************************************************************'''
    '''//Check if the sheet exists
    sheetName = "IDS_Distribution_" & TheExec.CurrentJob
    Set wb = Application.ActiveWorkbook
    Call check_Sheet_Range(sheetName, wb, ws_def, MaxRow, maxcol, isSheetFound, False)
    If isSheetFound = False Then
        sheetName = "IDS_Distribution"
        Call check_Sheet_Range(sheetName, wb, ws_def, MaxRow, maxcol, isSheetFound)
    End If
    Max_IDS_Dis_Zone = 0
    '''*****************************************************************'''
    If isSheetFound = True Then
        '''//init

        '''//Please check "Enum TestType" and "MaxTestType" in GlobalVariable.
        For Test_Type = 0 To testType.Total_Count - 1
            '''init the array to store the column number of each TestType.
            col_TestType(Test_Type) = 0
        Next Test_Type

        '''//Find the start point of the header.
        For Row = 1 To MaxRow
            For Col = 1 To maxcol
                If LCase(ws_def.Cells(Row, Col).value) Like "*rev*" And LCase(ws_def.Cells(Row + 1, Col).value) Like "td" Then
                    Row_of_Title = Row + 1
                End If

                If Row_of_Title > 0 And Row = Row_of_Title Then
                    If LCase(ws_def.Cells(Row, Col).value) <> "" Then
                        col_TestType(decide_test_type_for_string(ws_def.Cells(Row, Col).value)) = Col
                    End If
                End If
            Next Col
        Next Row

        If Row_of_Title > 0 And col_TestType(testType.TD) > 0 Then
            For Row = Row_of_Title + 1 To MaxRow
                If LCase(ws_def.Cells(Row, col_TestType(testType.TD)).value) = "ids range" And LCase(ws_def.Cells(Row, col_TestType(testType.TD) + 1).value) = "start bin" Then '''//Find the row with "IDS Range","Start Bin"

                    strTemp = ws_def.Cells(Row - 1, col_TestType(testType.TD)).value
                    Row_of_step0 = Row + 1
                    For Test_Type = 0 To testType.Total_Count - 1
                        col_testTypeSelected = col_TestType(Test_Type)

                        If col_testTypeSelected > 0 Then
                            If ws_def.Cells(Row_of_step0 - 2, col_testTypeSelected).value = strTemp Then
                                idx_step = 0
                                Row = Row_of_step0

                                If IsNumeric(ws_def.Cells(Row, col_testTypeSelected).value) And (IsEmpty(ws_def.Cells(Row, col_testTypeSelected).value) = False) Then
                                    While (LCase(ws_def.Cells(Row, col_testTypeSelected).value) <> "end" And (ws_def.Cells(Row, col_testTypeSelected).value) <> "")
                                        idx_step = idx_step + 1
                                        Row = Row + 1 '''Row Offset
                                    Wend
                                End If
                                If Max_IDS_Dis_Zone < idx_step Then Max_IDS_Dis_Zone = idx_step
                                Row = Row_of_step0
                            Else
                                TheExec.Datalog.WriteComment sheetName & " doesn't have the correct Performance_mode in row" & (Row_of_step0 - 2) & ", col" & col_testTypeSelected & " consistent with other TestType columns. Error!!!"
                                'TheExec.ErrorLogMessage SheetName & " doesn't have the correct Performance_mode in row" & Row & ", col" & col_TestType(TestType.TD) & " consistent with other TestType columns. Error!!!"
                            End If
                        End If
                    Next Test_Type

                End If
            Next Row
        Else
            TheExec.Datalog.WriteComment sheetName & " doesn't have correct format of the header. Error!!!"
            TheExec.ErrorLogMessage sheetName & " doesn't have correct format of the header. Error!!!"
        End If
    End If '''If isSheetFound = True
Exit Function
errHandler:
    TheExec.Datalog.WriteComment "Error encountered in VBT Function of initIDSTable"
    TheExec.ErrorLogMessage "Error encountered in VBT Function of initIDSTable"
    If AbortTest Then Exit Function Else Resume Next
End Function
Private Function CntAdditionMode(sheetName As String, ByRef additionMode As Dictionary)
    Dim sheetName_OutsideBinCut As String
    Dim sheetName_EquationVoltages As String
    Dim sheetName_temp As String
    Dim Row As Long, Col As Long
    Dim MaxRow As Long
    Dim maxcol As Long
    Dim split_content() As String
    Dim performance_mode As String
    Dim additional_mode As String
    Dim main_p_mode As Integer
    Dim p_mode As Integer
    Dim addi_mode As Integer '''For the additional mode
    Dim i As Long
    Dim other_voltage_start_point As Long
    Dim passBinCut As Long
    '''for testjob mapping
    Dim Row_of_Title As Long
    Dim row_of_testJob As Long
    Dim job_keyword As String
    Dim row_jobBlock As Long
    '''
    Dim cnt_testJob As Long
    Dim idx_testJob As Long
    Dim col_testJob() As Long
    Dim col_jobBlock_start As Long
    Dim col_jobBlock_end As Long
    '''for Performance Mode
    Dim col_mode As Long
    '''for powerDomain
    Dim got_correct_header As Boolean
    Dim got_CorrectDomain As Boolean
    '''for trackPower
    Dim trackpowerTemp As String
    Dim strAry_trackpower() As String
    '''variables
    Dim COREPOWER As Long
    Dim powerDomain As String
    Dim selected_powerDomain As String
    Dim str_mode_temp As String
    Dim isSheetFound As Boolean
    Dim str_mainColumn_content As String
    Dim strTemp As String
    Dim ws_def As Worksheet
    
On Error GoTo errHandler

    sheetName = sheetName
    Set wb = Application.ActiveWorkbook
    Call check_Sheet_Range(sheetName, wb, ws_def, MaxRow, maxcol, isSheetFound)

    '''*****************************************************************'''
    If isSheetFound = True Then
        '''//init
        other_voltage_start_point = 0
        trackpowerTemp = ""
        Row_of_Title = 0
        row_of_testJob = 0
        cnt_testJob = 0
        idx_testJob = -1
        got_correct_header = True
        got_CorrectDomain = False
        str_mainColumn_content = ""
        str_mode_temp = ""
        performance_mode = ""
        additional_mode = ""
        job_keyword = ""
        col_jobBlock_start = 0
        col_jobBlock_end = 0
        col_mode = 0
        addi_mode = 0
        '''//Get keyword for BinCut testJob mapping.
        '''20210819: Modified to assemble job_keyword by bincutJobName according to the vbt function Mapping_TestJobName_to_BincutJobName.
        job_keyword = LCase("*" & bincutJobName & "*") '''ex: "*cp1*", "*cp2*", "*ft_room*", "*ft_hot*", "*qa*"
    Else
        Exit Function
    End If
    
    '''//Find the keyword of BinCut testjob, and find column of the selected testJob.
    '''20210819: As per discussion with Jeff, he suggested us to use the keyword "Binning Domain" for each testJob.
    If job_keyword <> "" Then
        For Row = 1 To MaxRow
            For Col = 1 To maxcol
                If LCase(ws_def.Cells(Row, Col).value) Like LCase("Binning Domain") Then
                    ReDim Preserve col_testJob(cnt_testJob)
                    col_testJob(cnt_testJob) = Col
                    cnt_testJob = cnt_testJob + 1
                    
                    If row_of_testJob = 0 Then
                        row_of_testJob = Row
                    End If
                End If
            Next Col
            
            If row_of_testJob > 0 Then
                Exit For
            End If
        Next Row
    End If
    
    '''//Check if any matched block for IGXL Job.
    If cnt_testJob > 0 Then
        For i = 0 To cnt_testJob - 1
            If LCase(ws_def.Cells(row_of_testJob, col_testJob(i) + 2).value) Like job_keyword Then
                idx_testJob = i
                other_voltage_start_point = col_testJob(i) + 2
                
                '''//Get start/stop columns for block of the selected testJob.
                col_jobBlock_start = col_testJob(idx_testJob) '''column of Domain
                If idx_testJob = UBound(col_testJob) Then
                    col_jobBlock_end = maxcol
                Else
                    col_jobBlock_end = col_testJob(idx_testJob + 1) - 1
                End If
                
                '''//Get column of "Performance Mode"
                For Col = col_jobBlock_start To col_jobBlock_end
                    If LCase(Trim(ws_def.Cells(row_of_testJob, Col))) Like LCase("Performance Mode") Then
                        col_mode = Col
                        Exit For
                    End If
                Next Col
                Exit For
            End If
        Next i
    End If
    
    '''//Find the column of BinCut 1st powerDomain.
    If idx_testJob > -1 And col_mode > 0 And other_voltage_start_point > 0 Then
        '''Do nothing...
    Else
        other_voltage_start_point = 0
        TheExec.Datalog.WriteComment "sheet:" & sheetName & ", CntAdditionMode doesn't have the correct header for the current testJob:" & bincutJobName & ". Error!!!"
        TheExec.ErrorLogMessage "sheet:" & sheetName & ", CntAdditionMode doesn't have the correct header for the current testJob:" & bincutJobName & ". Error!!!"
        Exit Function
    End If
            
    '''====================================================-====================================================================
    '''[Step1] Find row of the Header with column of all BinCut powerDomains and pattern keywords.
    '''====================================================-====================================================================
    '''//If column of the selected testJob is found, start to parse the header VDD.
    For Row = row_of_testJob + 1 To MaxRow
        str_mainColumn_content = LCase(ws_def.Cells(Row, other_voltage_start_point).value)

        '''//Check if powerDomain is listed in VddbinPinDict (defined by sheet "Vdd_Binning_Def").
        For Col = col_jobBlock_start To col_jobBlock_end
            strTemp = UCase(Trim(ws_def.Cells(Row, Col).value))

            If strTemp <> "" Then
                If Col >= other_voltage_start_point And Col < other_voltage_start_point + cntVddbinPin Then '''powerDomain
                    '''//Check if the tracking power exists in the column of the header.
                    If strTemp Like "*,*" Then
                        strAry_trackpower = Split(strTemp, ",")
                        powerDomain = UCase(Trim(strAry_trackpower(0)))
                        trackpowerTemp = UCase(Trim(Replace(strTemp, (UCase(strAry_trackpower(0)) & ","), "")))
                    Else
                        powerDomain = strTemp
                        trackpowerTemp = ""
                    End If

                    '''//Check if powerDomain is CorePower or OtherRail shown in BinCut sheet "Vdd_Binning_Def".
                    '''//dict_IsCorePower is dictionary of BinCut CorePower/OtherRail.
                    If dict_IsCorePower.Exists(UCase(powerDomain)) = True Then
    
                    Else '''If dict_IsCorePower.Exists(UCase(powerDomain)) = False
                        got_correct_header = got_correct_header And False
                        TheExec.Datalog.WriteComment "sheet:" & sheetName & ", the header has the undefined powerDomain:" & strTemp & ", it doesn't show in sheet Vdd_Binning_Def sheet_appA_1. Error!!!"
                        TheExec.ErrorLogMessage "sheet:" & sheetName & ", the header has the undefined powerDomain:" & strTemp & ", it doesn't show in sheet Vdd_Binning_Def sheet_appA_1. Error!!!"
                    End If '''If dict_IsCorePower.Exists(UCase(powerDomain)) = True
                Else
                    powerDomain = ""
                    trackpowerTemp = ""
                End If
            End If '''If LCase(ws_def.Cells(row, col).Value) <> ""
        Next Col
        
        If got_correct_header = True Then
            Row_of_Title = Row
            Exit For
        End If
    Next Row
    
    '''====================================================-====================================================================
    '''[Step2] Parse each row to get testConditions of powerDomain.
    '''====================================================-====================================================================
    If Row_of_Title > 0 Then '''It means that columns of BinCut powerDomains are found.
        While LCase(ws_def.Cells(Row, 1).value) <> "end"
            '''//Check if any testcondition contains "#REF!" or "#NAME?".
            If IsError(ws_def.Cells(Row, other_voltage_start_point).value) Then
                TheExec.Datalog.WriteComment "sheet:" & sheetName & ", cell (row:" & Row & ",column:" & other_voltage_start_point & "), content:" & ws_def.Cells(Row, 1).value & ". The cell contains the incorrect content. Error!!!"
                TheExec.ErrorLogMessage "sheet:" & sheetName & ", cell (row:" & Row & ",column:" & other_voltage_start_point & "), content:" & ws_def.Cells(Row, 1).value & ". The cell contains the incorrect content. Error!!!"
            Else
                If LCase(ws_def.Cells(Row, col_mode).value) Like "m*" Then '''ex: "MS001", "MS001_GPU".
                    '''//Get performance_mode
                    str_mode_temp = UCase(ws_def.Cells(Row, col_mode).value)
                    split_content = Split(str_mode_temp, "_")
                    performance_mode = UCase(split_content(0))
                    
                    '''//If with all empty conditions, it means performance_mode without any additional_mode.
                    If UBound(split_content) > 0 Then
                        additional_mode = UCase(Replace(UCase(str_mode_temp), (performance_mode & "_"), ""))
                        
                        If Not additionMode.Exists(additional_mode) Then
                            additionMode.Add additional_mode, True
                        End If
                    End If
                End If
            End If '''If IsError(ws_def.Cells(row, other_voltage_start_point).Value) Then
            
            Row = Row + 1
        Wend
    End If '''If row_of_title > 0
Exit Function
errHandler:
    TheExec.Datalog.WriteComment "Error encountered in VBT Function of CntAdditionMode"
    TheExec.ErrorLogMessage "Error encountered in VBT Function of CntAdditionMode"
    If AbortTest Then Exit Function Else Resume Next
End Function
Function initDynamicArray()
    Dim idx_powerDomain As Long
    Dim addi_mode As Integer '''For the additional mode
    Dim passBinCut As Long
    Dim COREPOWER As Long
    Dim p_mode As Integer, i As Long
    On Error GoTo errHandler
    MaxJobCountInVbt = BinCutJobDefinition.Total_Count
    MaxPerformanceModeCount = cntVddbinPmode + 1 + cntEqVoltagePmode
    MaxEqnNum = MaxEqnNum + (Total_Bincut_Num - 1)  'Bin1 max step +Bin2/3 1 step
    Max_IDS_Zone = MaxEqnNum + 2 + Max_IDS_Dis_Zone   'Zone need to add 1 for start and end
    MaxBincutPowerdomainCount = cntVddbinPin
    ReDim AllBinCut(MaxPerformanceModeCount) As ALL_BINCUT_TYPE
    ReDim BinCut(MaxPerformanceModeCount, Total_Bincut_Num) As BINCUT_TYPE
    ReDim DC_Mapping(MaxPerformanceModeCount, Total_Bincut_Num) As DC_Mapping_Table
    ReDim IDS_for_BinCut(cntVddbinPin) As IDS_value
    ReDim ExcludedPmode(MaxPerformanceModeCount) As Boolean
    ReDim CPIDS_Spec(cntVddbinPin, Total_Bincut_Num) As Double
    ReDim FTIDS_Spec(cntVddbinPin, Total_Bincut_Num) As Double
    ReDim gb_IDS_hi_limit(cntVddbinPin, Total_Bincut_Num) As Double
    ReDim TestTypeName(testType.Total_Count) As String
    ReDim BinCut_Power_Seq(cntVddbinPin) As Power_Seq
    ReDim VBIN_IDS_ZONE(MaxPerformanceModeCount) As VBIN_IDS_ZONE
    ReDim VBIN_IDS_ZONE_Temp(MaxPerformanceModeCount) As VBIN_IDS_ZONE
    ReDim IDS_Distribution_Table(MaxPerformanceModeCount) As IDS_Distribution_TYPE
    ReDim gb_bincut_power_list(cntVddbinPin) As String
    'ReDim DC_Mapping(MaxPerformanceModeCount, MaxPassBinCut) As DC_Mapping_Table
    ReDim col_dynamic_offset(BinCutJobDefinition.Total_Count, testType.Total_Count) As Integer 'Dim col_dynamic_offset(MaxJobCountInVbt, MaxTestType) As Integer 20230526
    'ReDim Static_Mapping(MaxPerformanceModeCount, Total_Bincut_Num) As New BCType_Static_Mapping_Table
    ReDim FirstChangeBinInfo.str_Pmode_Test(MaxSiteCount - 1) As String  'str_Pmode_Test(MaxSiteCount - 1) As String
    ReDim VBIN_RESULT(MaxPerformanceModeCount) As VBIN_RESULT_TYPE
    ReDim Static_Mapping(MaxPerformanceModeCount, MaxPassBinCut) As Static_Mapping_Table
    'ReDim Static_Mapping.Addtional_Selsrm_Value(MaxAdditionalModeCount) As String
    'ReDim Static_Mapping.HVCC_Addtional_Selsrm_Value(MaxAdditionalModeCount) As String
    
    ReDim DYNAMIC_VBIN_IDS_ZONE(MaxPerformanceModeCount) As DYNAMIC_VBIN_IDS_ZONE
    For p_mode = 0 To MaxPerformanceModeCount - 1 'MaxPerformanceModeCount - 1
        For i = 0 To Total_Bincut_Num
            InitializeArrbincut BinCut(p_mode, i), cntVddbinPin, MaxEqnNum, MaxAdditionalModeCount
  
            InitializeArrDC DC_Mapping(p_mode, i), MaxAdditionalModeCount
            
            InitializeArr_Static_Mapping Static_Mapping(p_mode, i), MaxAdditionalModeCount

            'Static_Mapping(p_mode, i).InitializeArr MaxAdditionalModeCount
        Next

        'InitializeArr_Static_Mapping Static_Mapping, MaxAdditionalModeCount

        InitializeArr_VBIN_IDS_ZONE VBIN_IDS_ZONE(p_mode), Max_IDS_Zone, MaxEqnNum, testType.Total_Count '??

        InitializeArr_VBIN_IDS_ZONE_temp VBIN_IDS_ZONE_Temp(p_mode), Max_IDS_Zone, MaxEqnNum, testType.Total_Count '??

        InitializeArr_IDS_Distribution_Table IDS_Distribution_Table(p_mode), Max_IDS_Zone, testType.Total_Count

        InitializeArr_DYNAMIC_VBIN_IDS_ZONE DYNAMIC_VBIN_IDS_ZONE(p_mode), Total_Bincut_Num, MaxEqnNum
    Next p_mode
    
    For i = 0 To cntVddbinPin
        ReDim IDS_for_BinCut(i).ids_name(MaxSiteCount - 1) As String
    Next i
    
    ReDim BinCut_Init_Voltage(cntVddbinPin) As New SiteDouble
    ReDim BinCut_Payload_Voltage(cntVddbinPin) As New SiteDouble
    ReDim Previous_Payload_Voltage(cntVddbinPin) As New SiteDouble
     
    '''//Clear the array gb_bincut_power_list to reset the list of all performance_modes in each powerDomain.
    '''20210701: Modified to reset gb_bincut_power_list().
    For idx_powerDomain = 0 To UBound(gb_bincut_power_list) '20230629 Dynamic initial
        gb_bincut_power_list(idx_powerDomain) = vbNullString
    Next idx_powerDomain
    
    '''//Initialize the array of BV and HBV testConditions by empty string "".
    '''//The vbt function "initVddBinCondition" supported multiple "Non_Binning_Rail_Outside_BinCut" sheets.
    '''20210819: Modified to move the vbt code about resetting globalVariables of BinCut testCondition from the vbt function initVddBinCondition to the vbt function Reset_BinCut_GlobalVariable_for_initVddBinning.
    For p_mode = 0 To MaxPerformanceModeCount - 1 '0~60 '20230629 Dynamic initial
        For passBinCut = 0 To MaxPassBinCut         '0~3
            For COREPOWER = 0 To cntVddbinPin
                BinCut(p_mode, passBinCut).OTHER_VOLTAGE(COREPOWER) = vbNullString
                BinCut(p_mode, passBinCut).HVCC_OTHER_VOLTAGE(COREPOWER) = vbNullString
                '''for OutsideBinCut sheet
                BinCut(p_mode, passBinCut).OutsideBinCut_OTHER_VOLTAGE(COREPOWER) = vbNullString
                BinCut(p_mode, passBinCut).OutsideBinCut_HVCC_OTHER_VOLTAGE(COREPOWER) = vbNullString
                
                For addi_mode = 0 To MaxAdditionalModeCount
                    BinCut(p_mode, passBinCut).Addtional_OTHER_VOLTAGE(COREPOWER, addi_mode) = vbNullString
                    BinCut(p_mode, passBinCut).HVCC_Addtional_OTHER_VOLTAGE(COREPOWER, addi_mode) = vbNullString
                    '''for OutsideBinCut sheet
                    BinCut(p_mode, passBinCut).OutsideBinCut_Addtional_OTHER_VOLTAGE(COREPOWER, addi_mode) = vbNullString
                    BinCut(p_mode, passBinCut).OutsideBinCut_HVCC_Addtional_OTHER_VOLTAGE(COREPOWER, addi_mode) = vbNullString
                Next addi_mode
            Next COREPOWER
        Next passBinCut
    Next p_mode
    
Exit Function
errHandler:
    TheExec.Datalog.WriteComment "Error encountered in VBT Function of initDynamicArray"
    TheExec.ErrorLogMessage "Error encountered in VBT Function of initDynamicArray"
    If AbortTest Then Exit Function Else Resume Next
End Function
Public Sub InitializeArrbincut(BinCut As BINCUT_TYPE, MaxBincutPowerdomainCount As Integer, TotalStepPerMode As Long, MaxAdditionalModeCount As Long)
    ReDim BinCut.EQ_Num(TotalStepPerMode) As Long
    ReDim BinCut.c(TotalStepPerMode) As Double
    ReDim BinCut.m(TotalStepPerMode) As Double
    ReDim BinCut.CP_Vmax(TotalStepPerMode) As Double
    ReDim BinCut.CP_Vmin(TotalStepPerMode) As Double
    '''Montonicity_Offset
    ReDim BinCut.Monotonicity_Offset(TotalStepPerMode) As Double
    '''GuardBand
    ReDim BinCut.CP_GB(TotalStepPerMode) As Double
    ReDim BinCut.CP2_GB(TotalStepPerMode) As Double
    ReDim BinCut.FT1_GB(TotalStepPerMode) As Double
    ReDim BinCut.FT2_GB(TotalStepPerMode) As Double
    ReDim BinCut.SLT_GB(TotalStepPerMode) As Double
    ReDim BinCut.FTQA_GB(TotalStepPerMode) As Double
    ReDim BinCut.HTOL_RO_GB(TotalStepPerMode) As Double
    ReDim BinCut.HTOL_RO_GB_ROOM(TotalStepPerMode) As Double
    ReDim BinCut.HTOL_RO_GB_HOT(TotalStepPerMode) As Double
    ReDim BinCut.SLT_FTQA_GB(TotalStepPerMode) As Double
    '''IDS_Limit
    ReDim BinCut.IDS_CP_LIMIT(TotalStepPerMode) As Double
    ReDim BinCut.IDS_FT_LIMIT(TotalStepPerMode) As Double
    ReDim BinCut.IDS_QA_LIMIT(TotalStepPerMode) As Double
    ReDim BinCut.IDS_FT2_LIMIT(TotalStepPerMode) As Double
    ReDim BinCut.IDS_FT2_QA_LIMIT(TotalStepPerMode) As Double
    
    ReDim BinCut.HVCC_CP(TotalStepPerMode) As Double
    ReDim BinCut.HVCC_FT(TotalStepPerMode) As Double
    ReDim BinCut.HVCC_QA(TotalStepPerMode) As Double
    
    ReDim BinCut.SBIN_BINNING_FAIL(TotalStepPerMode, testType.Total_Count) As Long
    ReDim BinCut.SBIN_LVCC_FAIL(TotalStepPerMode, testType.Total_Count) As Long
    ReDim BinCut.HBIN_BINNING_FAIL(TotalStepPerMode, testType.Total_Count) As Long
    ReDim BinCut.HBIN_LVCC_FAIL(TotalStepPerMode, testType.Total_Count) As Long
    
    ReDim BinCut.OTHER_CP_Vmax(MaxBincutPowerdomainCount) As Double
    ReDim BinCut.OTHER_CP_Vmin(MaxBincutPowerdomainCount) As Double
    ReDim BinCut.OTHER_VOLTAGE(MaxBincutPowerdomainCount) As String      '''store testCondition of each powerDomain for performance mode in BV test instance.
    ReDim BinCut.HVCC_OTHER_VOLTAGE(MaxBincutPowerdomainCount) As String '''store testCondition of each powerDomain for performance mode in HBV test instance.
    '''GuardBand for nonbinning CorePower and OtherRail
    ReDim BinCut.OTHER_FT1_GB(MaxBincutPowerdomainCount) As Double
    ReDim BinCut.OTHER_FT2_GB(MaxBincutPowerdomainCount) As Double
    ReDim BinCut.OTHER_CP1_RAIL(MaxBincutPowerdomainCount) As Double     '''CP voltage of otherRail. Since M of otherRail is 0, it can directly take C as CP voltage for otherRail.
    ReDim BinCut.OTHER_CP1_GB(MaxBincutPowerdomainCount) As Double       '''CP1_GB of otherRail
    ReDim BinCut.OTHER_CP2_GB(MaxBincutPowerdomainCount) As Double
    ReDim BinCut.OTHER_PRODUCT_RAIL(MaxBincutPowerdomainCount) As Double
    ReDim BinCut.OTHER_SLT_GB(MaxBincutPowerdomainCount) As Double
    ReDim BinCut.OTHER_ATE_FQA_GB(MaxBincutPowerdomainCount) As Double
    ReDim BinCut.OTHER_HTOL_RO_GB(MaxBincutPowerdomainCount) As Double
    ReDim BinCut.OTHER_HTOL_RO_GB_ROOM(MaxBincutPowerdomainCount) As Double
    ReDim BinCut.OTHER_HTOL_RO_GB_HOT(MaxBincutPowerdomainCount) As Double
    ReDim BinCut.OTHER_SLT_FQA_GB(MaxBincutPowerdomainCount) As Double
    ReDim BinCut.HVCC_OTHER_CP_RAIL(MaxBincutPowerdomainCount) As Double
    ReDim BinCut.HVCC_OTHER_FT_RAIL(MaxBincutPowerdomainCount) As Double
    ReDim BinCut.HVCC_OTHER_QA_RAIL(MaxBincutPowerdomainCount) As Double
    ReDim BinCut.OTHER_CPIDS(MaxBincutPowerdomainCount) As Double
    ReDim BinCut.OTHER_FTIDS(MaxBincutPowerdomainCount) As Double
    
    ReDim BinCut.Allow_Equal(TotalStepPerMode) As Integer                                                    '''20161223: added to assign the inheriting rule of allow equal for the performance mode.
    ReDim BinCut.Addtional_OTHER_VOLTAGE(MaxBincutPowerdomainCount, MaxAdditionalModeCount) As String        '''store testCondition of each powerDomain for performance mode with additional mode in BV test instance.
    ReDim BinCut.HVCC_Addtional_OTHER_VOLTAGE(MaxBincutPowerdomainCount, MaxAdditionalModeCount) As String   '''store testCondition of each powerDomain for performance mode with additional mode in HBV test instance.
    ReDim BinCut.INTP_MODE_L(TotalStepPerMode) As Integer                                                    '''start p_mode of interpolation.
    ReDim BinCut.INTP_MODE_H(TotalStepPerMode) As Integer                                                    '''end p_mode of interpolation.
    ReDim BinCut.INTP_MFACTOR(TotalStepPerMode) As Double                                                    '''factor of interpolation.
    ReDim BinCut.INTP_OFFSET(TotalStepPerMode) As Double                                                     '''offset of interpolation.
    ReDim BinCut.INTP_SKIPTEST(TotalStepPerMode) As Boolean                                                  '''flag to skip interpolation tests of p_mode.
    
    
    ReDim BinCut.DYNAMIC_OFFSET(MaxJobCountInVbt, testType.Total_Count) As Double 'DYNAMIC_OFFSET(MaxJobCountInVbt, MaxTestType) As Double 20230526
    ReDim SRAM_VTH_SPEC(1) As Double '''SRAM_VTH_SPEC(0): for CP1 BV binSearch and postBinCut/OutsideBinCut, SRAM_VTH_SPEC(1): for CP1 HBV and non-CP1 BV/HBV.
    ReDim BinCut.OutsideBinCut_OTHER_VOLTAGE(MaxBincutPowerdomainCount) As String 'OutsideBinCut_OTHER_VOLTAGE(MaxBincutPowerdomainCount) As String
    ReDim BinCut.OutsideBinCut_HVCC_OTHER_VOLTAGE(MaxBincutPowerdomainCount) As String 'OutsideBinCut_HVCC_OTHER_VOLTAGE(MaxBincutPowerdomainCount) As String
    ReDim BinCut.OutsideBinCut_Addtional_OTHER_VOLTAGE(MaxBincutPowerdomainCount, MaxAdditionalModeCount) As String 'OutsideBinCut_Addtional_OTHER_VOLTAGE(MaxBincutPowerdomainCount, MaxAdditionalModeCount) As String
    ReDim BinCut.OutsideBinCut_HVCC_Addtional_OTHER_VOLTAGE(MaxBincutPowerdomainCount, MaxAdditionalModeCount) As String 'OutsideBinCut_HVCC_Addtional_OTHER_VOLTAGE(MaxBincutPowerdomainCount, MaxAdditionalModeCount) As String
    ReDim BinCut.OutsideBinCut_EquationVoltages(MaxBincutPowerdomainCount) As String 'OutsideBinCut_EquationVoltages(MaxBincutPowerdomainCount) As String
    'bincut.DC_Bincut_X_X_X() As String 'WWWWW20220631 DC_Bincut_X_X_X() As String
    
    ReDim BinCut.OutsideBinCut_OTHER_VOLTAGE(MaxBincutPowerdomainCount) As String
    ReDim BinCut.OutsideBinCut_HVCC_OTHER_VOLTAGE(MaxBincutPowerdomainCount) As String
    
    ReDim BinCut.OutsideBinCut_Addtional_OTHER_VOLTAGE(MaxBincutPowerdomainCount, MaxAdditionalModeCount) As String
    ReDim BinCut.OutsideBinCut_HVCC_Addtional_OTHER_VOLTAGE(MaxBincutPowerdomainCount, MaxAdditionalModeCount) As String

    ReDim BinCut.OutsideBinCut_EquationVoltages(MaxBincutPowerdomainCount) As String
    
    
    
    
End Sub
Public Sub InitializeArrDC(DC_Mapping As DC_Mapping_Table, MaxAdditionalModeCount As Long)
    ReDim DC_Mapping.Addtional_LV_DC_Value(MaxAdditionalModeCount) As String
    ReDim DC_Mapping.Addtional_HVCC_DC_Value(MaxAdditionalModeCount) As String
End Sub
Public Sub InitializeArr_DYNAMIC_VBIN_IDS_ZONE(DYNAMIC_VBIN_IDS_ZONE As DYNAMIC_VBIN_IDS_ZONE, MaxPassBinCut As Long, Max_IDS_Step As Long)
    ReDim DYNAMIC_VBIN_IDS_ZONE.c(Max_IDS_Step) As New SiteDouble
    ReDim DYNAMIC_VBIN_IDS_ZONE.m(Max_IDS_Step) As New SiteDouble
    ReDim DYNAMIC_VBIN_IDS_ZONE.passBinCut(Max_IDS_Step) As New SiteLong
    ReDim DYNAMIC_VBIN_IDS_ZONE.EQ_Num(Max_IDS_Step) As New SiteLong
    ReDim DYNAMIC_VBIN_IDS_ZONE.step_Mapping(MaxPassBinCut, Max_IDS_Step + 1) As New SiteLong
    ReDim DYNAMIC_VBIN_IDS_ZONE.Voltage(Max_IDS_Step) As New SiteDouble
    ReDim DYNAMIC_VBIN_IDS_ZONE.Product_Voltage(Max_IDS_Step) As New SiteDouble
    ReDim DYNAMIC_VBIN_IDS_ZONE.IDS_Start_EQ_Num(testType.Total_Count) As New SiteLong
    ReDim DYNAMIC_VBIN_IDS_ZONE.IDS_START_STEP(testType.Total_Count) As New SiteLong
End Sub
Public Sub InitializeArr_IDS_Distribution_Table(IDS_Distribution_Table As IDS_Distribution_TYPE, Max_IDS_Zone As Long, MaxTestType As Long)
    ReDim IDS_Distribution_Table.range(Max_IDS_Zone, MaxTestType) As Double
    ReDim IDS_Distribution_Table.Start_Bin(Max_IDS_Zone, MaxTestType) As Long
    ReDim IDS_Distribution_Table.START_STEP(Max_IDS_Zone, MaxTestType) As Long
End Sub
Public Sub InitializeArr_VBIN_IDS_ZONE(VBIN_IDS_ZONE As VBIN_IDS_ZONE, Max_IDS_Zone As Long, Max_IDS_Step As Long, MaxTestType As Long)
'    Dim i As Long, j As Long
'
    ReDim VBIN_IDS_ZONE.Ids_range(Max_IDS_Zone, MaxTestType) As Double
    ReDim VBIN_IDS_ZONE.IDS_Start_EQ_Num(Max_IDS_Zone, MaxTestType) As Long
    ReDim VBIN_IDS_ZONE.IDS_START_STEP(Max_IDS_Zone, MaxTestType) As Long
    ReDim VBIN_IDS_ZONE.IDS_RANGE_COUNT(MaxTestType) As Long
    ReDim VBIN_IDS_ZONE.c(Max_IDS_Zone, Max_IDS_Step) As Double
    ReDim VBIN_IDS_ZONE.m(Max_IDS_Zone, Max_IDS_Step) As Double
    ReDim VBIN_IDS_ZONE.passBinCut(Max_IDS_Zone, Max_IDS_Step) As Long
    ReDim VBIN_IDS_ZONE.EQ_Num(Max_IDS_Zone, Max_IDS_Step) As Long
    ReDim VBIN_IDS_ZONE.Voltage(Max_IDS_Zone, Max_IDS_Step) As New SiteDouble
    ReDim VBIN_IDS_ZONE.Product_Voltage(Max_IDS_Zone, Max_IDS_Step) As New SiteDouble    '''for GradeVDD
    ReDim VBIN_IDS_ZONE.Max_Step(Max_IDS_Zone) As Long
    'redim VBIN_IDS_ZONE.
     
End Sub

Public Sub InitializeArr_Static_Mapping(Static_Mapping As Static_Mapping_Table, MaxAdditionalModeCount As Long)
'    Dim i As Long, j As Long
'
    ReDim Static_Mapping.Addtional_Selsrm_Value(MaxAdditionalModeCount) As String
    ReDim Static_Mapping.HVCC_Addtional_Selsrm_Value(MaxAdditionalModeCount) As String
    
     
End Sub
Public Sub InitializeArr_VBIN_IDS_ZONE_temp(VBIN_IDS_ZONE_Temp As VBIN_IDS_ZONE, Max_IDS_Zone As Long, Max_IDS_Step As Long, MaxTestType As Long)

    ReDim VBIN_IDS_ZONE_Temp.Ids_range(Max_IDS_Zone, MaxTestType) As Double
    ReDim VBIN_IDS_ZONE_Temp.IDS_Start_EQ_Num(Max_IDS_Zone, MaxTestType) As Long
    ReDim VBIN_IDS_ZONE_Temp.IDS_START_STEP(Max_IDS_Zone, MaxTestType) As Long
    ReDim VBIN_IDS_ZONE_Temp.IDS_RANGE_COUNT(MaxTestType) As Long
    ReDim VBIN_IDS_ZONE_Temp.c(Max_IDS_Zone, Max_IDS_Step) As Double
    ReDim VBIN_IDS_ZONE_Temp.m(Max_IDS_Zone, Max_IDS_Step) As Double
    ReDim VBIN_IDS_ZONE_Temp.passBinCut(Max_IDS_Zone, Max_IDS_Step) As Long
    ReDim VBIN_IDS_ZONE_Temp.EQ_Num(Max_IDS_Zone, Max_IDS_Step) As Long
    ReDim VBIN_IDS_ZONE_Temp.Voltage(Max_IDS_Zone, Max_IDS_Step) As New SiteDouble
    ReDim VBIN_IDS_ZONE_Temp.Product_Voltage(Max_IDS_Zone, Max_IDS_Step) As New SiteDouble    '''for GradeVDD
    ReDim VBIN_IDS_ZONE_Temp.Max_Step(Max_IDS_Zone) As Long
    
     
End Sub
