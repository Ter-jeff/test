Attribute VB_Name = "VBT_LIB_VDD_Binning"
Option Explicit

'''//************************************************************************************************************************************************************************************************//'''
'''//Warning!!!!!! Read the following instructions before using GradeSearch_CallInstance_VT / GradeSearch_HVCC_CallInstance_VT / run_patt_only_CallInstance_VT.
'''1. For instance with "Call TheHdw.Patterns(ary_FuncPat_decomposed(indexPatt)).test(pfAlways, 0, result_mode)", pfAlways caused "TheExec.sites(site).LastTestResultRaw" and "TheExec.Flow.LastFlowStepResult" get the incorrect TestReseult.
'''20200819: Discussed "pfAlways" issue with Chihome, he found the same case in the offline simulation. He suggested us to ask TER factory for patch or .dll to fix this.
'''20200901: TER factory thought that pfAlways didn't cause "TheExec.sites(Site).LastTestResultRaw" issue...
'''//For instance with pfAlways, maybe it can use failFlag or BV_Pass to get testResult about Pass/Fail.
'''
'''2. For Multi-Instances with use-limit, we found that IGXL gave incorrect "testLimitIndex=0" for each instance with use-limit.
'''20200903: Discussed "call Multi-Instances with use-limit" with Chihome. He suggested us to create "Instance Group" for "Multi-Instances with use-limit" to avoid incorrect TestLimitIndex.
'''20200908: For better workaround, we suggested users to copy those called instances with use-limit from HardIP to the position right after BinCut call-instance in BinCut testFlow.
'''//************************************************************************************************************************************************************************************************//'''
'1.1: fixed pin level not restored after run_prepat
'1.2: add tracking power without tracking level and use the same level as main power
'1.3: add adjust VDD binning value for higher performance mode
'1.4: 140312 add column for FT IDS limit
'1.5: 140324 different binning for TD/SPI/MBIST
'1.6: 140404 allow assign default voltage for performance mode without VDD binning
'1.7: 140408 Do shmoo for all BV tests
'1.8: 140430 add binary search
'1.9: 140505 add find_start_voltage to start higher voltage at higher performance mode
'1.11: 140508 fix the search voltage not increasing to the top voltage for passing die
'1.12: 140509 use IDS to decide starting voltage
'1.13: 140514 fix repeat search to the top voltage and wrong binning assignment for IDS failure
'1.14: 140514 fix first executed instance does not consider previous executed instance at lower performance mode
'1.15: 140515 add test type for ldcsdd and ldcbfd
'1.15: 140515 add IDS for FT and FT_QA
'1.16: 140516 fix wrong reading sequence of bin number for Binning/LVCC failure in initVddBinTable
'1.16: 140516 only execute find_start_search_voltage at CP
'1.16: 140516 Vdd Binning column for FQA is modified to column 5
'1.17: 140521 fix bug of ids fail bin swapped to lvcc fail bin if one of the site is in IDS algorithm
'1.18: 140523 modify HV_step_x_percent for HBV tests to set correct SRAM voltage
'1.19: 140523 fix only count the num of execution before grade is found
'1.19: 140605 modify IDS_Check to check IDS with the limit relative to the step voltage
'1.20: 140528 add type VDD_BIN_TYPE and VBIN_RESULT_TYPE to replace orignal VDD_BIN_DEF_LVCC_VDD and EcidVddGrade array
'1.21: 140605 modify the vdd_bin_def and code to allow more than 2 pass bin
'1.21: 140611 fix the bug that show LVCC failure but actually is IDS failure (IndexLevelFailIdsPerSite(Site) = IndexLevelPerSite(Site))
'1.22: 140624 adjust test voltage to make fuse voltage of higher performance mode is alway higher or equal that of lower performance mode (find_start_voltage,find_next_bin)
'1.22: 140624 no need to test voltage for J42 lower than the highest Fiji voltage
'1.23: 141027 Consider CPU P8 and GPU P5 only exist for Tazanite and not for Dazzle devices
'1.24: 150701 modfify enum for new format bincut
'2.01: 151126 modify for EQ format bincut tables
'2.02: 151126 modify for other rail table format is changed
'2.03: 170629 add Interpolation function
'2.04: 170905 combine addtional vbt module

'20210910: Modified to revise the format of FirstChangeBinInfo, as requested by C651 Si and TSMC ZYLINI.
'20210906: Modified to adjust the priorty of Check_BinCut_flag_globalVariable in initVddBinning.
'20210701: As per discussion with TSMC SWLINZA, he told us that "BinCut flow table contains the super set of all BinCut performance modes."
'20210701: Modified to use update_bincut_pmode_list for BinCut_Power_Seq.
'20210701: Modified to adjust the priority of Parsing_IDSname_from_BDF_Table in the vbt function initVddBinning.
'20210617: Check_IDS and judge_IDS are dedicated to Efuse processed IDS for BinCut search, as requested by C651 Si.
'20210526: Modified to add "VBIN_Result(p_mode).is_Monotonicity_Offset_triggered" for Monotonicity_Offset check because C651 Si revised the check rules to ensure that: (GradeVDD(P_mode)-GradeVDD(previous perfromance_mode))> Monotonicity_Offset.
'20210525: Modified to reset siteMasks for MultiFSTP in CP1.
'20210510: Modified to parse HARV_Pmode_Table and HARVMappingTable for Harvest Core DSSC.
'20210420: Modified to remove the unused vbt code of GradeSearchMethod. Once if C651 provided any new definition of GradeSearchMethod, we will revise the vbt code.
'20210120: Modified to use VBIN_RESULT(p_mode).step_1stPass_in_IDS_Zone to store the first pass step in Dynamic IDS Zone and find the correspondent PassBinCut number.
'20210107: Modified to add for recording first changed binnum mode data, requested by C651 Si.
'20201222: Modified to use parsing_OutsideBinCut_flow_table for parsing multiple "Non_Binning_Rail_Outside_BinCut" sheets.
'20201210: Modified to use the flag "is_BinCutJob_for_StepSearch" for "check_bincutJob_for_StepSearch" to check if the test program is binSearch or functional test.
'20201130: Modified to clear the redundant dictionaries by setting then as Nothing.
'20201021: Modified to revise the vbt code for parsing "Non_Binning_Rail" and "Non_Binning_Rail_outside_BinCut".
'20200825: Modified to remove the conventional BinCut functions: GradeSearch_ELB_VT, GradeSearch_HVCC_ELB_VT, GradeSearch_TMPS_VT, GradeSearch_HVCC_TMPS_VT, run_patt_only_forLBK_VT, GradeSearch_HVCC_VT_RtosNewFeature.
'20200731: Modified to merge MappingBincutJobName and Mapping_TPJobName_to_BincutJobName into Mapping_TestJobName_to_BincutJobName.
'20200709: Modified to add "initGradeSearchMethodName".
'20200702: Modified to add "initVddBinCondition_Outside_BinCut".
'20200622: Modified to use "Reset_BinCut_GlobalVariable_for_initVddBinning" to reset BinCut globalVariable for initVddBinning.
'20200609: Modified to use Check_alarmFail_before_BinCut_Initial.
'20200526: Modified to remove the unused GlobalVariable "Public RtosSelSramAry() As String".
'20200326: Modified to use the funtion "Decide_PowerBinning_Type" to select "Parsing_Power_Bin_Table_Harvest" or "Parsing_Power_Bin_Table".
'20200325: Modified to decide the type of PowerBinning.
'20200211: Modified to replace "FlowTestCondName" with "AdditionalModeName".
'20200211: Modified to replace "cntFlowTestCond" with "cntAdditionalMode".
'20200106: Modified to check SRAMthresh for p_mode of selsram powerDomain.
'20200102: Modified to init the flag "Flag_BinCut_Config_Printed".
'20191230: Modified to use "initVbinTest" for each touchdown.
'20191219: Modified for Domain2Pin and Pin2Domain.
'20191210: Modified to check if selsramPin exists.
'20191204: Modified to the revised InitVddBinCondition.
'20191127: Modified for the revised InitVddBinTable.
'20190522: Modified to add "Parsing_IDSname_from_BDF_Table" to get IDS names for BinCut powerPin.
'20190321: Modified to add "Flag_SyncUp_DCVS_Output_enable".
'20190313: Modified to init FIRSTPASSBINCUT=999.
'20180626: Modified to add "Mapping_TPJobName_to_BincutJobName" for BinCut testjob mapping.
'20160614: Modified initIDSTable/Generate_IDS_ZONE_RANGE/Generate_IDS_ZONE_CONTENT by TSMC Jack.
Public Function initVddBinning()
On Error GoTo errHandler 'Add ErrHandler 2023/08/01
    Dim site As Variant
    Dim p_mode As Integer
    Dim inst_name As String
    Dim idx_powerDomain As Long
    'Dim IGXL_Version As String
'''//==================================================================================================================================================================================//'''
'''//Note:
'''1. For reparsing all BinCut related table, remember to set FlagInitVddBinningTable = False.
'''2. Remember to use "Cdec" for values comparison to avoid double format accuracy issues.
'''3. Remember to check hard-code Sort Number and Bin Number with Bin_Table for the vbt functions:
'''check_IDS, adjust_VddBinning, Adjust_Multi_PassBinCut_Per_Site, find_start_voltage,judge_PF_func, judge_PF, check_voltageInheritance_for_powerDomain.
'''4. Check_IDS and judge_IDS are dedicated to Efuse processed IDS for BinCut search, as requested by C651 Si,20210617.
'''5. As per discussion with TSMC SWLINZA, he told us that "BinCut flow table contains the super set of all BinCut performance modes.", 20210707.
'''//==================================================================================================================================================================================//'''
'''//==================================================================================================================================================================================//'''
'''//Tips:
'''Remember to check the following items!!!
'''We suggest BinCut owners to add "FlagInitVddBinningTable = False" into "Function OnProgramValidated" in "Exec_IP_Module.bas".
'''This can re-do initVddBinning when BinCut testjob or sheet is changed.
'''1. Do not add bin/sort and P/F in flow table.
'''2. Bin/sort is defined in Vdd_Binning_Def.
'''3. Put initVddBinning in the OnProgramStarted.
'''//==================================================================================================================================================================================//'''
    '''//Init variables
    'IGXL_Version = TheExec.SoftwareVersion
    inst_name = "initVddBinning"
    IsLevelLoadedForApplyLevelsTiming = False
    PreviousBinCutInstanceContext = vbNullString
    CurrentBinCutInstanceContext = vbNullString
    MaxSiteCount = TheExec.sites.Existing.Count '20230526 redim MaxSiteCount
    
    '''//Check if alarmFail was triggered before BinCut initial, and then reset the globalVariable of the alarm flag.
    Check_alarmFail_before_BinCut_Initial inst_name
    alarmFail = False
    
    ''20230608: Added for projects without bincut
    If Find_Sheet("Vdd_Binning_Def_appA_1_" & TheExec.CurrentJob) = False And Find_Sheet("Vdd_Binning_Def_appA_1") = False Then
        TheExec.Datalog.WriteComment "No Vdd_Binning_Def_appA_1* table exist. initVddbinning is skipped!!"
        If FlagInitVddBinningTable = False Then
            Dim jobName As String
            jobName = Mapping_TestJobName_to_BincutJobName(LCase(TheExec.CurrentJob))
            Parsing_SELSRM_Mapping_Table jobName
            FlagInitVddBinningTable = True
        End If
        Exit Function
    End If
    
    '''//C651 requested the feature to record first changed binnum mode data and print FirstChangeBinInfo in Adjust_VddBinning for each touchdown.
    FirstChangeBinInfo.FirstChangeBinMode = 999
    FirstChangeBinInfo.FirstChangeBinType = 999
    
    '''//**********************************************************************************************************//'''
    '''//For re-parsing all BinCut related table, remember to set FlagInitVddBinningTable = False.
    '''//Parsing BinCut related tables once for all touchdowns.
    '''//**********************************************************************************************************//'''
    If FlagInitVddBinningTable = False Then
        '''//Reset all BinCut pinGroups and globalVariables for initVddBinning.
        Reset_BinCut_GlobalVariable_for_initVddBinning
        
        '''//parse BinTable for checking BinCut LVCC bin number in VDD_def_Binning 20230526
        If Flag_Enable_Parsing_BinTable_ForBinCut = True Then
            Parsing_BinTable_ForBinCut
        End If

        CountMaxAdditionModeCnt "Non_Binning_Rail"
        
        'If TheExec.enableWord("Enable_ApplyVoltageFromBinCut") = True Then Parsing_EquationVoltages "EquationVoltages"

        ''20231102: Remove hardcode job definition on central
        CountMaxEquationVoltagesMode "EquationVoltages"
        
        If Flag_IDS_Distribution_enable = True Then
            CountMaxIDS_DIS_ZoneCnt
        End If
        
        '''//Decide the bincutJobName mapping.
        '''5 constant testJobs: "cp1, cp2, ft_room, ft_hot, qa" are enumerated in BinCut globalVariable "Enum BinCutJobDefinition".
        '''bincutJobName is BinCut globalVariable for testJob mapping.
        bincutJobName = Mapping_TestJobName_to_BincutJobName(LCase(TheExec.CurrentJob))
        
        '''//parse BinTable for checking BinCut LVCC bin number in VDD_def_Binning
        'If Flag_Enable_Parsing_BinTable_ForBinCut = True Then
            'Parsing_BinTable_ForBinCut
        'End If
        
        '''//Parse BinCut voltage tables (sheetNames with "Vdd_Binning_Def").
        '''Store IDS limit, BinCut EQN-based parameters (C, M, and Guardband) of CorePower and OtherRail for voltages calculation, and SortNumber into BinCut globalVariables.
        initVddBinTable
        
        '''//Parse BinCut flow table (sheet "Non_Binning_Rail") to get testConditions of CorePower and OtherRail for each performance mode according to bincutJobName.
        '''According to keyword "Evaluate Bin" in the testCondition of the table "Non_Binning_Rail", it can decide "is_BinCutJob_for_StepSearch = True".
        '''20210701: As per discussion with TSMC SWLINZA, he told us that "BinCut flow table contains the super set of all BinCut performance modes."
        initVddBinCondition "Non_Binning_Rail"
        
        Parsing_EquationVoltages "EquationVoltages"
        
        '''//Check powerDomain and powerPin.
        '''Check if powerDomain or powerPin is connected to DCVS.
        '''BinCut powerDdomains are the pinGroup from BinCut flow table (sheet "Non_Binning_Rail"), and each powerDomain might include pins.
        '''domain2pinDict, pin2domainDict are the dictionaries in GlobalVarible to store domains and pins.
        '''VddbinPinDcvstypeDict stores information about DCVS type for powerDomain.
        initDomain2Pin FullBinCutPowerinFlowSheet, domain2pinDict, pin2domainDict
        
        '''//Check Domain2DcSpecGrp
        '''check powerDomain--> powerPin --> DC Spec specName.
        initDomain2DcSpecGrp FullBinCutPowerinFlowSheet, dictDomain2DcSpecGrp
        
        '''//Paring EFuse_BitDef_Table to get IDS names for BinCut powerPin.
        '''//The vbt function check IDS name of each BinCut powerDomain and Efuse Product name of each BinCut performance mode from Efuse_BitDef_Table.
        '''//Update AllBinCut(p_mode).listed_in_Efuse_BDF in the vbt function Parsing_IDSname_from_BDF_Table.
        '''Caution!!! Remember check "Efuse_BitDef_Table" and testPlan for core number of the Harvest powerPin.
        '''20210617: Check_IDS and judge_IDS are dedicated to Efuse processed IDS for BinCut search, as requested by C651 Si.
        Parsing_IDSname_from_BDF_Table
        
        '''//Sort the Performance mode by MAX_ID for each BinCut powerDomain.
        '''After parsing Efuse_BitDef_Table, update "AllBinCut(p_mode).used" and sort_power_sequence for each BinCut powerDomain.
        '''BinCut_Power_Seq is the sequence of performance modes for each BinCut powerDomain.
        '''The sequence of performance modes is referred by voltage_inheritance.
        '''Update AllBinCut(p_mode).PREVIOUS_Performance_Mode from BinCut_Power_Seq of the BinCut powerDomain.
        '''<Note>: allbincut(p_mode).used is decided after parsing BinCut flow table and Efuse_BitDef_Table, and it can check if p_mode can be tested and fused for BinCut...
    
        
        update_bincut_pmode_list
        
        '''//Initialize IDS Zone table.
        init_IDS_ZONE
        
        '''//Parse the IDS distribution table to find start EQN for each p_mode.
        If Flag_IDS_Distribution_enable = True Then
           initIDSTable
        End If
        
        '''//Generate IDS ZONE RANGE from BinCut voltage tables.
        Generate_IDS_ZONE_RANGE
        
        '''//Generate the IDS ZONE with C, M, bin, and PassBinCut.
        Generate_IDS_ZONE_CONTENT
        
        '''//Initialize BincutVoltageTypeName for printing BV strings (enum2str).
        initBincutVoltageType
        
        '''//Initialize TestType names (enum2str).
        initTestTypeName
        
        '''**************Parsing Additional Tables Area**************'''
        '''***Please put the additional table parsing functions in this area!!!***
        '''//Define SyncUp pinGroup.
        '''20190626: As the discussion with TSMC PSYAO, we found vbump didn't exist in all BinCut powerDomain of BinCut patterns.
        '''So that we decided to detect the output status of Selsram Logic powerpins (refer to "SELSRM_Mapping_Table").
        '''If one of them is in Valt, the vbt code will switch other CorePower and OtherRail to Valt.
        'If Flag_SyncUp_DCVS_Output_enable Then
            SyncUp_PowerPin_Group = FullBinCutPowerinFlowSheet
        'End If
        
        '''//Parsing the SELSRAM bit order.
        Parsing_SELSRM_Mapping_Table bincutJobName, FullCorePowerinFlowSheet, FullOtherRailinFlowSheet
        
        '''20220127,HarvPinGrp Parsing
        '''20231120 move Parsing_HarvPinFlag_Mapping_Table to OnProgramStarted
        'Parsing_HarvPinFlag_Mapping_Table
        
        '''//Check SRAMthresh for p_mode of Selsrm powerpin.
        If Flag_SelsrmMappingTable_Parsed = True Then
            Precheck_SRAMthresh_for_Selsram_Power selsramLogicPin
        End If
        
        '''// Parsing the PowerBinning tables.
        '''//"PwrBinning_V*" is the keyword for "PowerBinning_Harvest", and "Pwrbin_Seq" is the keyword for conventional PowerBinning.
        Decide_PowerBinning_Type

        '''//Parse Outside BinCut flow table (sheet "Non_Binning_Rail_Outside_BinCut") to get outside BinCut testConditions of CorePower and OtherRail for each performance mode.
        '''================================================================================================================================================'''
        '''Note:
        '''1. Please contact C651 project DRI for the table "Outsite BinCut".
        '''2. Please check keyword "Non_Binning_Rail_Outside" of sheetName for the vbt functions "initVddBinCondition" and "parsing_OutsideBinCut_flow_table".
        '''================================================================================================================================================'''
        parsing_OutsideBinCut_flow_table "Non_Binning_Rail_Outside"
        
        '''//HARV_Pmode_Table and HARVMappingTable for Harvest Core DSSC.
        '''20210510: Modified to parse HARV_Pmode_Table and HARVMappingTable for Harvest Core DSSC.
        check_Harvest_Core_All_Table
        '''20211130 Rhodes Parsing special case VDD_GPU0~VDD_GPUX
        If Flag_cross_domain_check_enable = True Then
            create_cross_domain_voltageInheritance_for_VDD_GPU
        End If
        '''//**********************************************************************************************************//'''
        
        '''//After parsing tables, set FlagInitVddBinningTable as true.
        FlagInitVddBinningTable = True
        
        '''//Remove the redundant dictionaries to save system memory usage by setting them as Nothing.
        Set dict_BinCutFlow_Domain2Column = Nothing
        Set dict_BinCutFlow_Column2Domain = Nothing
        Set dict_Binned_Mode_Column2Ratio = Nothing
        Set dict_Binned_Mode_Ratio2Column = Nothing
        Set dict_Other_Mode_Column2Ratio = Nothing
        Set dict_Other_Mode_Ratio2Column = Nothing
        Set dict_OutsideBinCut_additionalMode = Nothing
        
        '''//Initialize the flag of printing BinCut configs.
        Flag_BinCut_Config_Printed = False
        
        '''20230323: Add enable word to print out pattern names under domain mode, requested by C651 Si
        If TheExec.enableWord("Vddbin_PatternBurst_PrintOut") = True Then
            Flag_Vddbin_PatternBurst_PrintOut = True
        Else
            Flag_Vddbin_PatternBurst_PrintOut = False
        End If
    
    Else '''If FlagInitVddBinningTable = True
        '''20230323: Add enable word to print out pattern names under domain mode, requested by C651 Si, force = false after 1st touchdown
        Flag_Vddbin_PatternBurst_PrintOut = False
    End If '''If FlagInitVddBinningTable = False
        '''//VBIN_RESULT(p_mode).step_1stPass_in_IDS_Zone is to store the first pass step in Dynamic IDS Zone of each p_mode for COF_StepInheritance.
    '''//VBIN_Result(p_mode).is_Monotonicity_Offset_triggered is for Monotonicity_Offset check of each p_mode.
    '''because C651 Si revised the check rules to ensure that: (GradeVDD(P_mode)-GradeVDD(previous perfromance_mode))> Monotonicity_Offset.
    For p_mode = 0 To MaxPerformanceModeCount - 1
        VBIN_RESULT(p_mode).step_1stPass_in_IDS_Zone = -1
        VBIN_RESULT(p_mode).is_Monotonicity_Offset_triggered = False
    Next p_mode
    
    '''20210910: Modified to revise the format of FirstChangeBinInfo, as requested by C651 Si and TSMC ZYLINI.
    For Each site In TheExec.sites
        FirstChangeBinInfo.str_Pmode_Test(site) = vbNullString
    Next site
    
    '''//Reset MultiFSTP related globalVariables for each touchdown.
    '''The globalVariable gb_siteMask_original is to store the original status of siteMask before MultiFSTP.
    '''The globalVariable gb_sitePassBin_original is to store the PassBin status of each site before MultiFSTP.
    gb_siteMask_original = True
    gb_siteMask_current = True
    gb_sitePassBin_original = -1
    
    
    '''//Initialize offline flag, and check if BinCut flags/globalVariables have any conflict.
    '''Check if tester is online or opensocket with conflicts.
    '''20210906: Modified to adjust the priorty of Check_BinCut_flag_globalVariable in initVddBinning.
    Check_BinCut_flag_globalVariable
    
    '''//Check if the step voltage is same as the one shown in the BinCut Table.
    TheExec.Datalog.WriteComment "Version of Vdd_Binning_Def = " & Version_Vdd_Binning_Def
    
    ReDim Bincut_Domain2BinX(cntVddbinPin)
    For idx_powerDomain = 0 To UBound(pinGroup_BinCut)
        Dim site_BINX As New SiteBoolean: site_BINX = False
        Bincut_Domain2BinX(VddBinStr2Enum(pinGroup_BinCut(idx_powerDomain))) = site_BINX
    Next idx_powerDomain
    '''//Check if StepVoltage from Vdd_Binning_Def matches globalVariable(default value of gC_StepVoltage is 3.125).
    If BV_StepVoltage = gC_StepVoltage Then
        TheExec.Datalog.WriteComment "Step Voltage in Vdd_Binning = " & BV_StepVoltage
    Else
        Call Print_Error_Message(Error_Info, "VBT_LIB_VDD_Binning", "initVddBinning", "The StepVoltage in BinCut = " & BV_StepVoltage & ", The Gc_Stepvoltage = " & gC_StepVoltage & ". Error!!!")
        'TheExec.ErrorLogMessage "The StepVoltage in BinCut = " & BV_StepVoltage & ", The Gc_Stepvoltage = " & gC_StepVoltage & ". Error!!!"    ''20230807: Removed by sanity check
    End If
    
    '''//init VBIN_RESULT and VBIN_RESULT(p_mode).tested for each touchdown.
    initVbinTest
    
    '''//initial Dynamic IDS zone.
    init_Dynamic_IDS_ZONE
Exit Function
errHandler: 'Add ErrHandler 2023/08/01
    Call Print_Error_Message(Error_Info, "VBT_LIB_VDD_Binning", "initVddBinning") 'Add ErrHandler 2023/08/01
    If AbortTest Then Exit Function Else Resume Next 'Add ErrHandler 2023/08/01
End Function

'20210813: Modified to use BVmin(CPVmin) and BVmax(CPVmax) as lo_limit and hi_limit of BinCut voltage for p_mode, as C651 Toby's new rules about Guardband and product voltage.
'20210730: C651 Toby revised the rule of adjust_VddBinning that "If CP1 has only PCPU bin cut search, then we only need PCPU for PrintOut_VddBinning and Adjust_VddBinning. The other domains can be skipped at CP1."
'20210726: Modified to add the argument "bincutNum As Long" to the vbt function Adjust_Multi_PassBinCut_Per_Site.
'20210623: Modified to get lo_limit and hi_limit of GradeVDD.
'20210621: Modified to revise the vbt function adjust_VddBinning for BinCut search in FT.
'20210420: Modified to remove the unused vbt code of GradeSearchMethod. Once if C651 provided any new definition of GradeSearchMethod, we will revise the vbt code.
'20201215: Modified to move "step_lowest = BinCut(p_mode, PassBinNum).Mode_Step" prior to CPVmin.
'20201210: Modified to use the flag "is_BinCutJob_for_StepSearch" for "check_bincutJob_for_StepSearch" to check if the test program is binSearch or functional test.
'20201126: Modified to remove the redundant vbt code of CMEM datalog setup.
'20200826: Modified to replace "If..Else" with "Select Case".
'20200826: Modified to merge the redundant branches.
'20200810: Modified to clear CMEM prior to BinCut HVCC tests.
'20200317: Modified for SearchByPmode.
'20191226: For opensocket CMEM overflow issues, we modified to clear and renew CMEM.
'20191127: Modified for the revised InitVddBinTable.
'20190722: Modified to printout the scale and the unit for BinCut voltages and IDS values.
'20190716: Modified to unify the unit for Voltage.
'20180723: Modified for BinCut testjob mapping.
Public Function PrintOut_VDD_BIN()
On Error GoTo errHandler 'Add ErrHandler 2023/08/01
    Dim site As Variant
    Dim p_mode As Integer
    Dim passBinCut As Variant
    Dim Active_site_count As Long
    Dim F_BlowConfig As New SiteLong
    Dim performance_mode As String
    Dim PassBinNum As New SiteLong
    Dim str_PmodeGradeTestJob As String
    '''variant
    Dim step_lowest As Long
    Dim dbl_CPVmin As Double
    Dim dbl_CPVmax As Double
    Dim dbl_GB_BinCutJob As Double
    Dim dbl_gradeVdd_lolimit As Double
    Dim dbl_gradeVdd_hilimit As Double
    Dim str_Efuse_read_pmode As String
    Dim str_Efuse_write_pmode As String
'''//==================================================================================================================================================================================//'''
'''//Note:
'''//This vbt function ensures that voltages and efuse product voltage will be matched to CurrentPassBinCutNum for BinCut search.
'''The flag "is_BinCutJob_for_StepSearch" = True is for BinCut search while testBinJob with testCondition from the table "Non_Binning_Rail" has the keyword "*Evaluate*Bin*".
'''1. The vbt function aligns Efuse product voltages of each performanc mode with the same PassBin for CP1 by adjusting step in Dynamic_IDS_zone.
'''2. The vbt function prints Grade and GradeVDD of each performance mode and check if it is in limits of PassBinCut ("PassBin" for CP1 and "Efuse Product_Identifier + 1" for non-CP1).
'''3. C651 Toby revised the rule of adjust_VddBinning that "If CP1 has only PCPU bin cut search, then we only need PCPU for PrintOut_VddBinning and Adjust_VddBinning. The other domains can be skipped at CP1.", 20210730.
'''4. C651 Toby updated the rules that Product voltage=BinCut voltage+binning_GuardBand, 20210812.
'20210813: Modified to use BVmin(CPVmin) and BVmax(CPVmax) as lo_limit and hi_limit of BinCut voltage for p_mode, as C651 Toby's new rules about Guardband and product voltage.
'''//Tips:
'''1. VBIN_RESULT(p_mode).PASSBINCUT is passbin number of p_mode, and CurrentPassBinCutNum defines current PassBin number for DUT.
'''//==================================================================================================================================================================================//'''
    If is_BinCutJob_for_StepSearch = True Then
        For p_mode = 0 To MaxPerformanceModeCount - 1
            '''//Set the excluded performance mode if the device is Bin2 die and the performance mode doesn't exist in Bin2 table.
            SkipTestBin2Site p_mode, Active_site_count
            
            '''//If PassBinCut of P_mode doesn't match CurrentPassBinCutNum, adjust step in Dynamic_IDS_Zone to match CurrentPassBinCutNum.
            '20211022: Modify to support SRAM Voltage calculate with equation
            For Each site In TheExec.sites
                If VBIN_RESULT(p_mode).passBinCut <> CurrentPassBinCutNum And _
                    (AllBinCut(p_mode).Used = True Or (dict_IsCorePower(AllBinCut(p_mode).powerPin) = True And UCase(AllBinCut(p_mode).powerPin) Like "*VDD*_SRAM*")) Then
                    Adjust_Multi_PassBinCut_Per_Site p_mode, site, CurrentPassBinCutNum(site)
                End If
            Next site
            
            RestoreSkipTestBin2Site p_mode
        Next p_mode
    End If
    
    '''********************************************************************************************************************'''
    '''//If the Bin result is Bin2, do not fuse and retest the chip again.
    '''********************************************************************************************************************'''
    For Each site In TheExec.sites
        '''init
        F_BlowConfig(site) = 0
        
        '''//solution1
        '''TheExec.Datalog.WriteComment "print: Dazzle retest solution1, use ECID blank to judge, only blow Fiji good at 1st time, no matter what bin it is, blow fuse at 2nd time"
        If IsEmpty(TheExec.sites(site).SiteVariableValue("ECIDBlankChk_Var")) = False Then
            If TheExec.sites(site).SiteVariableValue("ECIDBlankChk_Var") = 2 Then
                TheExec.sites.item(site).FlagState("F_BlowConfig") = logicTrue
                F_BlowConfig(site) = 1
            ElseIf TheExec.sites(site).SiteVariableValue("ECIDBlankChk_Var") = 1 And CurrentPassBinCutNum = 1 Then
                TheExec.sites.item(site).FlagState("F_BlowConfig") = logicTrue
                F_BlowConfig(site) = 1
            Else
                TheExec.sites.item(site).FlagState("F_BlowConfig") = logicFalse
                F_BlowConfig(site) = 0
            End If
            TheExec.Datalog.WriteComment "print: Site(" & site & "), ECIDBlankChk_Var = " & TheExec.sites(site).SiteVariableValue("ECIDBlankChk_Var") & ", CurrentPassBinCutNum = " & CurrentPassBinCutNum & ", Fuse CFG & UDR = " & F_BlowConfig
        End If ''' If IsEmpty(TheExec.sites(site).SiteVariableValue("ECIDBlankChk_Var")) = False Then
    Next site
    
    For Each site In TheExec.sites
        For p_mode = 0 To MaxPerformanceModeCount - 1
            '20211022: Modify to support SRAM Voltage calculate with equation
            If IsExcludedVddBin(p_mode) = False And _
                (AllBinCut(p_mode).Used = True Or (dict_IsCorePower(AllBinCut(p_mode).powerPin) = True And UCase(AllBinCut(p_mode).powerPin) Like "*VDD*_SRAM*")) Then
                '''//Get performance mode from p_mode
                performance_mode = VddBinName(p_mode)
                
                '''//Get Efuse category of Efuse product voltage(GradeVDD) for p_mode.
                str_Efuse_read_pmode = get_Efuse_category_by_BinCut_testJob("read", VddBinName(p_mode))
                str_Efuse_write_pmode = get_Efuse_category_by_BinCut_testJob("write", VddBinName(p_mode))
                
                '''//Check if p_mode for BinCut search has the dedicated Efuse category in the current testJob.
                '''20210730: C651 Toby revised the rule of adjust_VddBinning that "If CP1 has only PCPU bin cut search, then we only need PCPU for PrintOut_VddBinning and Adjust_VddBinning. The other domains can be skipped at CP1."
                If str_Efuse_read_pmode <> "" Or str_Efuse_write_pmode <> "" Then
                    '''//Get PassBinNumber / Mode_Step / CPVmax of p_mode by GradeSearchMethod.
                    '''ToDo: Check if we move the PassBinCut_ary loop to outside if..else.
                    PassBinNum = VBIN_RESULT(p_mode).passBinCut
                    '''//Need to reset the flags "F_PassBinCut_1", "F_PassBinCut_2" from "printout_vddbin" due to new binx judgement.
                    If Flag_BinX_Info_Parsed = True Then
                        If Flag_Fuse_BinX(site) = True Then
                            TheExec.sites(site).FlagState("F_PassBinCut_1") = logicFalse
                            TheExec.sites(site).FlagState("F_PassBinCut_2") = logicTrue
                        Else
                            TheExec.sites(site).FlagState("F_PassBinCut_1") = logicTrue
                            TheExec.sites(site).FlagState("F_PassBinCut_2") = logicFalse
                        End If
                    Else
                        '''//Remember to check if the flags "F_PassBinCut_1", "F_PassBinCut_2", and "F_PassBinCut_3" all exist in "Bin Table".
                        For Each passBinCut In PassBinCut_ary
                            If PassBinNum(site) = passBinCut Then
                                TheExec.sites(site).FlagState("F_PassBinCut_" & passBinCut) = logicTrue
                            Else
                                TheExec.sites(site).FlagState("F_PassBinCut_" & passBinCut) = logicFalse
                            End If
                        Next passBinCut
                    End If
                    
                    '''//Get CPVmin / CPVmax / Mode_Step of p_mode.
                    step_lowest = BinCut(p_mode, PassBinNum).Mode_Step
                    dbl_CPVmin = BinCut(p_mode, PassBinNum).CP_Vmin(step_lowest)
                    dbl_CPVmax = BinCut(p_mode, PassBinNum).CP_Vmax(0)
                    
                    '''//Get the matched Guardband(GB) according to the BinCut testjob.
                    '''20210812: C651 Toby updated the rules that Product voltage=BinCut voltage+binning_GuardBand.
                    Select Case LCase(bincutJobName)
                        Case "cp1":
                            dbl_GB_BinCutJob = BinCut(p_mode, PassBinNum).CP_GB(0)
                            str_PmodeGradeTestJob = VddBinName(p_mode) & " CP1"
                        Case "cp2":
                            dbl_GB_BinCutJob = BinCut(p_mode, CurrentPassBinCutNum).CP2_GB(0)
                            str_PmodeGradeTestJob = VddBinName(p_mode) & " CP2"
                        Case "ft_room":
                            dbl_GB_BinCutJob = BinCut(p_mode, CurrentPassBinCutNum).FT1_GB(0)
                            str_PmodeGradeTestJob = VddBinName(p_mode) & " FT1"
                        Case "ft_hot":
                            dbl_GB_BinCutJob = BinCut(p_mode, CurrentPassBinCutNum).FT2_GB(0)
                            str_PmodeGradeTestJob = VddBinName(p_mode) & " FT2"
                        Case "qa":
                            dbl_GB_BinCutJob = BinCut(p_mode, CurrentPassBinCutNum).FTQA_GB(0)
                            str_PmodeGradeTestJob = VddBinName(p_mode) & " QA"
                        Case Else:
                            dbl_GB_BinCutJob = 0
                            str_PmodeGradeTestJob = vbNullString
                            TheExec.sites.item(site).FlagState(strGlb_Flag_Vddbinning_Fail_Stop) = logicTrue
                            Call Print_Error_Message(Error_Info, "VBT_LIB_VDD_Binning", "PrintOut_VDD_BIN", "site:" & site & ", PrintOut_VDD_BIN has the incorrect BinCut TestJob selection. Error!!!")
                            ''TheExec.ErrorLogMessage "site:" & site & ", PrintOut_VDD_BIN has the incorrect BinCut TestJob selection. Error!!!"    ''20230807: Removed by sanity check
                    End Select
                    If TheExec.enableWord("T0TX_ROOM") = True Then dbl_GB_BinCutJob = BinCut(p_mode, CurrentPassBinCutNum).HTOL_RO_GB_ROOM(0) 'WWWWW20220631 T0TX modify by Wyatt 20230418
                    If TheExec.enableWord("T0TX_HOT") = True Then dbl_GB_BinCutJob = BinCut(p_mode, CurrentPassBinCutNum).HTOL_RO_GB_HOT(0) 'WWWWW20220631  T0TX modify by Wyatt 20230418

                    '''********************************************************************************************************************'''
                    '''//Print out all Grade and GradeVdd of performance mode.
                    '''//Remember to check the following items:
                    '''1. BinCut voltage calculation uses the scale and the unit in "mV"
                    '''2. TheExec.Flow.TestLimit should convert the voltage value into "V" with settings "unit:=unitVolt" and "scaleMilli".
                    '''********************************************************************************************************************'''
                    If str_PmodeGradeTestJob <> "" Then
                        '''//Get lo_limit and hi_limit for product voltage(GradeVDD) of p_mode.
                        dbl_gradeVdd_lolimit = dbl_CPVmin + BinCut(p_mode, PassBinNum).CP_GB(step_lowest)
                        dbl_gradeVdd_hilimit = dbl_CPVmax + BinCut(p_mode, PassBinNum).CP_GB(0)
                        
                        '''//Check if Grade and GradeVDD of p_mode are in limit.
                        '''//BinCut voltage(Grade)
                        '''20210812: C651 Toby updated the rules that Product voltage=BinCut voltage+binning_GuardBand.
                        '''20210813: Modified to use BVmin(CPVmin) and BVmax(CPVmax) as lo_limit and hi_limit of BinCut voltage for p_mode, as C651 Toby's new rules about Guardband and product voltage.
                        If str_Efuse_write_pmode <> "" And Not LCase(TheExec.CurrentJob) = "wlft1" Then
                            TheExec.Flow.TestLimit VBIN_RESULT(p_mode).GRADE / 1000, (dbl_CPVmin) / 1000, (dbl_CPVmax) / 1000, Tname:=str_PmodeGradeTestJob, scaletype:=scaleMilli, unit:=unitVolt, ForceUnit:=unitVolt
                        Else
                            If LCase(VddBinName(p_mode)) Like "vdd_cpu_sram_m[pe]s*" And LCase(TheExec.CurrentJob) = "wlft2" Then
                                TheExec.Flow.TestLimit VBIN_RESULT(p_mode).GRADE / 1000, (dbl_gradeVdd_lolimit - BinCut(p_mode, PassBinNum).CP_GB(step_lowest)) / 1000, (dbl_gradeVdd_hilimit - BinCut(p_mode, PassBinNum).CP_GB(0)) / 1000, Tname:=str_PmodeGradeTestJob, scaletype:=scaleMilli, unit:=unitVolt, ForceUnit:=unitVolt
                            Else
                                TheExec.Flow.TestLimit VBIN_RESULT(p_mode).GRADE / 1000, (dbl_gradeVdd_lolimit - dbl_GB_BinCutJob) / 1000, (dbl_gradeVdd_hilimit - dbl_GB_BinCutJob) / 1000, Tname:=str_PmodeGradeTestJob, scaletype:=scaleMilli, unit:=unitVolt, ForceUnit:=unitVolt
                            End If
                        End If
                        
                        '''//Efuse Product voltage(GradeVdd)
                        TheExec.Flow.TestLimit VBIN_RESULT(p_mode).GRADEVDD / 1000, dbl_gradeVdd_lolimit / 1000, dbl_gradeVdd_hilimit / 1000, Tname:=VddBinName(p_mode) & " VDD Define", scaletype:=scaleMilli, unit:=unitVolt, ForceUnit:=unitVolt
                    End If
                End If '''If str_Efuse_read_pmode <> "" Or str_Efuse_write_pmode <> ""
            End If '''If IsExcludedVddBin(p_mode) = False And AllBinCut(p_mode).Used = True
        Next p_mode
    Next site
    
    '''//Align the format of Judge_PF and Judge_PF_func in the datalog.
    TheExec.Datalog.WriteComment ""
    TheExec.Datalog.WriteComment ""
    TheExec.Datalog.WriteComment vbCrLf
Exit Function
errHandler: 'Add ErrHandler 2023/08/01
    Call Print_Error_Message(Error_Info, "VBT_LIB_VDD_Binning", "PrintOut_VDD_BIN") 'Add ErrHandler 2023/08/01
    If AbortTest Then Exit Function Else Resume Next 'Add ErrHandler 2023/08/01
End Function
'20230615: Modified udrP udrE allow Efuse
'20210928: Modified to check if p_mode is for BinCut search, as requested by TSMC ZYLINI.
'20210924: Modified to check product voltages of performance modes for BinCut powerDomain "VDD_PCPU" only while testJob="wlft2" in Crete A0.
'20210909: Modified to add the special case to get CP1-fused Prodcut_Identifier and Efuse product voltages while testJob="ft_hot" for T-Cre A0.
'20210812: C651 Toby updated the rules that Product voltage=BinCut voltage+binning_GuardBand.
'20210726: Modified to remove the vbt code of binning out the non-fused p_mode because some p_modes might not be fused in previous testJobs.
'20210723: Modified to move the vbt code of checking Efuse category "power_binning".
'20210719: Modified to revise the vbt code for BinCut search in FT.
'20210716: Modified to get Efuse category of Efuse product voltage for p_mode.
'20210710: As per discussion with TSMC ZYLINI, Crete had Efuse category "Product_Identifier" and "Product_Identifier_cp1" for testJob "cp1", but Efuse postCheck only supported "Product_Identifier". We had to use hard-code "Product_Identifier" here as Efuse workaround for Crete.
'20210709: Modified to move the vbt code for Product_Identifier offline simulation from the vbt function Read_DVFM_To_GradeVDD to generate_offline_IDS_IGSim_Parallel.
'20210709: Modified to move the vbt code for Efuse product voltage offline simulation from the vbt function Read_DVFM_To_GradeVDD to generate_offline_IDS_IGSim_Parallel.
'20210708: Modified to check if Efuse category "power_binning" for the current BinCut testJob has the matched programming stage in Efuse_BitDef_Table.
'20210708: Modified to check if column "power_binning" exists in PowerBinning flow table.
'20210707: Modified to check if it's OK to read Efuse category "power_binning".
'20210706: Modified to use the vbt function get_Efuse_category_by_BinCut_testJob to find the Efuse Category.
'20210705: Modified to set VBIN_RESULT(p_mode).tested = True because all p_mode will be based on CP1 results for search in FT.
'20210705: Modified to find Efuse Category for the BinCut performance mode by checking the current BinCut testjob.
'20210705: Modified to find Efuse Category for product_identifier by checking the current BinCut testjob.
'20210705: Modified to check if "power_binning" from the header BinCut powerbinning flow table exists in dict_EfuseCategory2BinCutTestJob.
'20210703: Modified to use dict_strPmode2EfuseCategory as the dictionary of p_mode and array of the related Efuse category.
'20210703: Modified to use dict_EfuseCategory2BinCutTestJob as the dictionary of Efuse category and the matched programming state in Efuse.
'20210703: Modified to Get all Efuse category names for each p_mode.
'20210630: Modified to merge site-loop of the vbt function Read_DVFM_To_GradeVDD.
'20210628: Modified to revise the vbt code to generate Grade and GradeVDD for opensocket in Read_DVFM_To_GradeVDD.
'20210623: Modified to get lo_limit and hi_limit of GradeVDD.
'20210623: Modified to check if "power_binning" exists in Efuse_BitDef_Table. If not, skip updating the flags.
'20210610: Modified to check if it got the incorrect value "power_binning" from Efuse. If that, bin out the failed DUT.
'20210510: Modified to check "power_binning" from Efuse and update the flags in Bin_Table for Read_DVFM_To_GradeVDD.
'20210415: Modified to set VBIN_RESULT(p_mode).tested = True if AllBinCut(p_mode).is_for_BinSearch = False.
'20210415: Modified to skip printing PassBin/Grade/GradeVDD if Flag_Remove_Printing_BV_voltages=True. It's the request as conclusion from BinCut central library meeting 20210413.
'20210217: Modified to print and check lo_limit and hi_limit of Grade and GradeVdd.
'20210207: Modified to generate offline simulation with passbin number for Read_DVFM_To_GradeVDD.
'20210207: Modified the branches of "If Flag_VDD_Binning_Offline = False Then" or "Vddbinning_OpenSocket".
'20210207: Modified to prevent Read_DVFM_To_GradeVDD from the incorrect Efuse product_identifier.
'20210129: Modified the branches of "If Flag_VDD_Binning_Offline = False Then".
'20210128: Modified to print info about Efuse product identifier and product voltages because Efuse won't print these anymore.
'20201228: Patty asked us to adapt Efuse object vbt code.
'20201118: Modified to updated VBIN_RESULT(nonbinning_pmode).tested = True for non-CP1, requested by PCLINZG.
'20200825: Modified to replace "If..Else" with "Select Case".
'20200825: Modified to bin out the DUT without any correct Efuse product voltage. Discussed this with projects BinCut owners, we decided to add "Vddbinning_Fail_Stop" to bin out the failed DUT.
'20200707: Modified to merge the branch of checking "VBIN_RESULT(p_mode).GRADEVDD=0".
'20200106: Modified to remove the ErrorLogMessage.
'20180724: Modified to prevent the misjudgement from GRADEVDD=0. We add the condition only for the pmode in use.
'20180705: Modified for BinCut testjob mapping.
Public Function Read_DVFM_To_GradeVDD() As Long
On Error GoTo errHandler 'Add ErrHandler 2023/08/01
    Dim site As Variant
    Dim p_mode As Long
    Dim performance_mode As String
    Dim efuse_gradevdd_val As Double
    '''variant
    Dim step_lowest As Long
    Dim dbl_CPVmin As Double
    Dim dbl_CPVmax As Double
    Dim dbl_GB_BinCutJob As Double
    Dim dbl_gradeVdd_lolimit As Double
    Dim dbl_gradeVdd_hilimit As Double
    Dim str_Efuse_read_pmode As String
    Dim str_Efuse_read_ProductIdentifier As String
    Dim str_Efuse_read_PowerBinning As String
    '''for Efuse obj vbt.
    Dim field As eFuseBdfField
    Dim siteLong_fusepowerBinning As New SiteLong
'''//==================================================================================================================================================================================//'''
'''//Note:
'''1. Read "Product_Identifier" from Efuse, then identify the chip is Bin1 or BinX or BinY.
'''2. CurrentPassBinCutNum = Product_Identifier+1, ex: Product_Identifier=0 means CurrentPassBinCutNum=1.
'''3. Update VBIN_RESULT(nonbinning_pmode).tested = True for non-CP1, as requested by TSMC PCLINZG.
'''4. Please check column "power_binning" in PowerBinning flow table, Efuse category "power_binning" in Efuse_BitDef_Table, and the correspondent flag in Bin_Table.
'''20210812: C651 Toby updated the rules that Product voltage=BinCut voltage+binning_GuardBand.
'''//==================================================================================================================================================================================//'''
    '''//Get Efuse catergory of "Product_Identifier" to "read" PassBin (Product_Identifier+1).
    str_Efuse_read_ProductIdentifier = get_Efuse_category_by_BinCut_testJob("read", "Product_Identifier")
    
    '''//Check if DUT is BinCut-searched and fused in the previous testJob.
    If str_Efuse_read_ProductIdentifier <> "" Then
'''********************************************************************************************************'''
'''//Warning!!! This block is the special case for Crete A0 only, and it will be removed while Crete B0!!!
'''20210909: Modified to add the special case to get CP1-fused Prodcut_Identifier and Efuse product voltages while testJob="ft_hot" for T-Cre A0.
'''20211118: Modified to read the cp1 product_idenfifier instead of shadow for wlft1, only for A0!
        If bincutJobName = "ft_hot" Or LCase(TheExec.CurrentJob) = "wlft1" Then
            str_Efuse_read_ProductIdentifier = "Product_Identifier"
        End If
'''********************************************************************************************************'''

        If Flag_Remove_Printing_BV_voltages = False Then
            TheExec.Datalog.WriteComment "Product_Identifier" & ",it can use Efuse category:" & str_Efuse_read_ProductIdentifier
        End If
    Else
        TheExec.Datalog.WriteComment "Efuse category:" & "Product_Identifier" & ", it wasn't fused before the current testJob, so that skip Read_DVFM_To_GradeVDD."
        Exit Function
    End If
    
    '''***********************************************************************************************************************************************'''
    '''[Step2] After getting Efuse categories of "Product_Identifier" and product voltages, check Grade and GradeVDD for each BinCut performance mode.
    '''***********************************************************************************************************************************************'''
    For Each site In TheExec.sites
        '''//Get PassBin from Efuse product identifier. PassBin=Product_Identifier+1.
        '''For project with Efuse obj vbt code.
        Set field = BdfDataBase.Bank_Cfg.Fields(str_Efuse_read_ProductIdentifier)
        CurrentPassBinCutNum(site) = field.DsscDecValue + 1
            
        '''//Print PassBin number for each site.
        If Flag_Remove_Printing_BV_voltages = False Then
            TheExec.Flow.TestLimit CurrentPassBinCutNum(site), 1, PassBinCut_ary(UBound(PassBinCut_ary)), , , scaleNoScaling, unitNone, formatStr:="%0.f", Tname:="PASSBIN NUMBER"
        End If
        
        If Flag_BinX_Info_Parsed = True Then
            If CurrentPassBinCutNum(site) = 2 Then
                Flag_Fuse_BinX(site) = True
                TheExec.Datalog.WriteComment "BinX is merged into bin1, force CurrentPassBinCutNum = 1!!"
                CurrentPassBinCutNum(site) = 1
            End If
        End If
        
        '''*****************************************************************************************************'''
        '''1. Read Efuse product value(GradeVdd) from efuse.
        '''2. Calculate GradeVdd - GB (by testJob) to get BinCut voltage (Grade). ex: CP2 = Product - CP2GB.
        '''*****************************************************************************************************'''
        For p_mode = 0 To MaxPerformanceModeCount - 1
            '20211022: Added to support SRAM Voltage calculate with equation
            If AllBinCut(p_mode).Used = True Or (dict_IsCorePower(AllBinCut(p_mode).powerPin) = True And UCase(AllBinCut(p_mode).powerPin) Like "*VDD*_SRAM*") Then
                '''//Get Efuse category of Efuse product voltage(GradeVDD) for p_mode.
                str_Efuse_read_pmode = get_Efuse_category_by_BinCut_testJob("read", VddBinName(p_mode))
                
'''********************************************************************************************************'''
'''//Warning!!! This block is the special case for Crete A0 only, and it will be removed while Crete B0!!!
'''20210924: Modified to check product voltages of performance modes for BinCut powerDomain "VDD_PCPU" only while testJob="wlft2" in Crete A0.
                If LCase(TheExec.CurrentJob) = "wlft2" Then
                    '''20210928: Modified to check if p_mode is for BinCut search, as requested by TSMC ZYLINI.
                    '''<org>
                        'If LCase(AllBinCut(p_mode).powerPin) <> "vdd_pcpu" Then
                    '''<new>
                    If AllBinCut(p_mode).is_for_BinSearch = True Then ''Or LCase(VddBinName(p_mode)) Like "vdd_cpu_sram_m[pe]s*" Then
                        str_Efuse_read_pmode = vbNullString
                    End If
                End If
'''********************************************************************************************************'''
            
                '''//If p_mode has Efuse category for the current BinCut testJob, get Efuse product voltage for P_mode.
                If str_Efuse_read_pmode <> "" Then
'''********************************************************************************************************'''
'''//Warning!!! This block is the special case for Crete A0 only, and it will be removed while Crete B0!!!
'''20210909: Modified to add the special case to get CP1-fused Prodcut_Identifier and Efuse product voltages while testJob="ft_hot" for T-Cre A0.
                    If bincutJobName = "ft_hot" Then
                        str_Efuse_read_pmode = VddBinName(p_mode)
                    End If
'''********************************************************************************************************'''
                
                    '''//Get name of the performance mode.
                    performance_mode = VddBinName(p_mode)
                    
                    '''//Get CPVmin / CPVmax / Mode_Step of p_mode.
                    step_lowest = BinCut(p_mode, CurrentPassBinCutNum).Mode_Step
                    dbl_CPVmin = BinCut(p_mode, CurrentPassBinCutNum).CP_Vmin(step_lowest)
                    dbl_CPVmax = BinCut(p_mode, CurrentPassBinCutNum).CP_Vmax(0)
                    
                    '''//Get the matched Guardband(GB) according to the BinCut testjob.
                    '''20210812: C651 Toby updated the rules that Product voltage=BinCut voltage+binning_GuardBand.
                    Select Case LCase(bincutJobName)
                        Case "cp1": dbl_GB_BinCutJob = BinCut(p_mode, CurrentPassBinCutNum).CP_GB(0)
                        Case "cp2": dbl_GB_BinCutJob = BinCut(p_mode, CurrentPassBinCutNum).CP2_GB(0)
                        Case "ft_room": dbl_GB_BinCutJob = BinCut(p_mode, CurrentPassBinCutNum).FT1_GB(0)
                        Case "ft_hot": dbl_GB_BinCutJob = BinCut(p_mode, CurrentPassBinCutNum).FT2_GB(0)
                        Case "qa": dbl_GB_BinCutJob = BinCut(p_mode, CurrentPassBinCutNum).FTQA_GB(0)
                        Case Else:
                                TheExec.sites.item(site).FlagState(strGlb_Flag_Vddbinning_Fail_Stop) = logicTrue
                                Call Print_Error_Message(Error_Info, "VBT_LIB_VDD_Binning", "Read_DVFM_To_GradeVDD", "site:" & site & ", Read_DVFM_To_GradeVDD has the incorrect BinCut TestJob selection. Error!!!")
                                ''TheExec.ErrorLogMessage "site:" & site & ", Read_DVFM_To_GradeVDD has the incorrect BinCut TestJob selection. Error!!!"    ''20230807: Removed by sanity check
                    End Select
                    
                    If TheExec.enableWord("T0TX_ROOM") = True Then dbl_GB_BinCutJob = BinCut(p_mode, CurrentPassBinCutNum).HTOL_RO_GB_ROOM(0) 'WWWWW20220631 T0TX
                    If TheExec.enableWord("T0TX_HOT") = True Then dbl_GB_BinCutJob = BinCut(p_mode, CurrentPassBinCutNum).HTOL_RO_GB_HOT(0) 'WWWWW20220631  T0TX
                    
                    
                    '''//Get Efuse product voltages (GradeVDD) from Efuse Category of each p_mode.
                    '''ToDo: Remember to check "EFUSE_BitDef_Table" for UDRP, UDRE, and CFG.
                    '''For project with Efuse obj vbt code.
'                    If LCase(AllBinCut(p_mode).powerPin) Like "vdd_pcpu0*" Then
'                        Set field = BdfDataBase.Bank_UdrP0.Fields(str_Efuse_read_pmode)
'                    ElseIf LCase(AllBinCut(p_mode).powerPin) Like "vdd_pcpu1*" Then
'                        Set field = BdfDataBase.Bank_UdrP1.Fields(str_Efuse_read_pmode)
'                    ElseIf LCase(AllBinCut(p_mode).powerPin) Like "vdd_ecpu" Then
'                        Set field = BdfDataBase.Bank_UdrE.Fields(str_Efuse_read_pmode)
'                    Else
'                        Set field = BdfDataBase.Bank_Cfg.Fields(str_Efuse_read_pmode)
'                    End If
                    ''20230704: Modified to check efuse bank for all pmode for T-Bra.
                    ''20230914: Added a boolean for getting correct field
                    Dim got_correct_field As Boolean
                    got_correct_field = False
                    If Not (BdfDataBase.Bank_UdrP Is Nothing) And got_correct_field = False Then
                        If BdfDataBase.Bank_UdrP.Fields.Exists(str_Efuse_read_pmode) = True Then
                            Set field = BdfDataBase.Bank_UdrP.Fields(str_Efuse_read_pmode)
                            got_correct_field = True
                        End If
                    End If
                    If Not (BdfDataBase.Bank_UdrP0 Is Nothing) And got_correct_field = False Then
                        If BdfDataBase.Bank_UdrP0.Fields.Exists(str_Efuse_read_pmode) = True Then
                            Set field = BdfDataBase.Bank_UdrP0.Fields(str_Efuse_read_pmode)
                            got_correct_field = True
                        End If
                    End If
                    If Not (BdfDataBase.Bank_UdrP1 Is Nothing) And got_correct_field = False Then
                        If BdfDataBase.Bank_UdrP1.Fields.Exists(str_Efuse_read_pmode) = True Then
                            Set field = BdfDataBase.Bank_UdrP1.Fields(str_Efuse_read_pmode)
                            got_correct_field = True
                        End If
                    End If
                    If Not (BdfDataBase.Bank_UdrE Is Nothing) And got_correct_field = False Then
                        If BdfDataBase.Bank_UdrE.Fields.Exists(str_Efuse_read_pmode) = True Then
                        Set field = BdfDataBase.Bank_UdrE.Fields(str_Efuse_read_pmode)
                            got_correct_field = True
                        End If
                    End If
                    If Not (BdfDataBase.Bank_Cfg Is Nothing) And got_correct_field = False Then
                        If BdfDataBase.Bank_Cfg.Fields.Exists(str_Efuse_read_pmode) = True Then
                        Set field = BdfDataBase.Bank_Cfg.Fields(str_Efuse_read_pmode)
                            got_correct_field = True
                        End If
                    End If
                    If got_correct_field = False Then
                        Call Print_Error_Message(Error_Info, "VBT_LIB_VDD_Binning", "Read_DVFM_To_GradeVDD", "Cannot find pmode=" & str_Efuse_read_pmode & " from eFuse bank!!! Error!!!")
                        '''20240303: Use modular function to bin out site
                        Call BinoutSite(site, VddbinFailStopSortNumber, VddbinFailStopBinNumber, strGlb_Flag_Vddbinning_Fail_Stop)

                        ''Bypass empty field VBT error, binout current site and go to next site
                        GoTo nextSite
                    End If
                    efuse_gradevdd_val = field.FuseMeasureValue
                    
                    '''//Calculate BinCut voltages (Grade) from Efuse product voltages (GradeVdd), ex: BinCut voltage CP2 = product - CP2GB.
                    VBIN_RESULT(p_mode).GRADEVDD = efuse_gradevdd_val
                    VBIN_RESULT(p_mode).GRADE = efuse_gradevdd_val - dbl_GB_BinCutJob
                    VBIN_RESULT(p_mode).passBinCut = CurrentPassBinCutNum
                    
                    '''//Update VBIN_RESULT(nonbinning_pmode).tested = True for non-CP1, as requested by TSMC PCLINZG.
                    '''If strAry_Efuse_read_pmode(p_mode) <> "", it means that p_mode is tested and fused in one of previous testJobs.
                    VBIN_RESULT(p_mode).tested = True
                    
                    '''//Print BinCut voltages (Grade) and Efuse product voltages (GradeVdd) for each site.
                    '''20210415: Modified to skip printing PassBin/Grade/GradeVDD if Flag_Remove_Printing_BV_voltages=True. It's the request as conclusion from BinCut central library meeting 20210413.
                    If Flag_Remove_Printing_BV_voltages = False Then
                        '''//Get lo_limit and hi_limit for product voltage(GradeVDD) of p_mode.
                        dbl_gradeVdd_lolimit = dbl_CPVmin + BinCut(p_mode, CurrentPassBinCutNum).CP_GB(step_lowest)
                        dbl_gradeVdd_hilimit = dbl_CPVmax + BinCut(p_mode, CurrentPassBinCutNum).CP_GB(0)
                        
                        '''//Check if BinCut voltage(Grade) and product voltage(GradeVDD) of p_mode are in limit.
                        '''//BinCut voltage(Grade).
                        TheExec.Flow.TestLimit VBIN_RESULT(p_mode).GRADE / 1000, (dbl_gradeVdd_lolimit - dbl_GB_BinCutJob) / 1000, (dbl_gradeVdd_hilimit - dbl_GB_BinCutJob) / 1000, Tname:=VddBinName(p_mode) & " VDD_Grade", scaletype:=scaleMilli, unit:=unitVolt, ForceUnit:=unitVolt
                        '''//Efuse product voltage(GradeVDD).
                        TheExec.Flow.TestLimit VBIN_RESULT(p_mode).GRADEVDD / 1000, dbl_gradeVdd_lolimit / 1000, dbl_gradeVdd_hilimit / 1000, Tname:=VddBinName(p_mode) & " VDD_Product", scaletype:=scaleMilli, unit:=unitVolt, ForceUnit:=unitVolt
                    End If
                    
                    '''//Check if Efuse product voltage(GradeVdd) is 0...
                    If VBIN_RESULT(p_mode).GRADEVDD = 0 Then
                        Call Print_Error_Message(Error_Info, "VBT_LIB_VDD_Binning", "Read_DVFM_To_GradeVDD", "site:" & site & "," & VddBinName(p_mode) & ". Efuse CFG provided the incorrect Efuse product voltage(GradeVdd) for Read_DVFM_To_GradeVDD. Error!!!")
                        ''TheExec.ErrorLogMessage "site:" & Site & "," & VddBinName(p_mode) & ". Efuse CFG provided the incorrect Efuse product voltage(GradeVdd) for Read_DVFM_To_GradeVDD. Error!!!"    ''20230807: Removed by sanity check
                        TheExec.sites.item(site).FlagState(strGlb_Flag_Vddbinning_Fail_Stop) = logicTrue
                    End If
                End If '''If strAry_Efuse_read_pmode(p_mode)<> ""
            Else '''If AllBinCut(p_mode).Used = False
                '''//For those peformance modes not used in BinCut flow or vdd_binning_def, add product voltages and bin numbers here.
                '''ex: MG008, MG009 are psuedo performance modes for BinCut flow sheet ver0.65 and earlier versions.
                '''So that we define Grade and GradeVDD(efuse product values) for MG008 and MG009.
            End If
        Next p_mode
nextSite:
    Next site
    
'''ToDo: Maybe we can make this block as the vbt function...
    '''***********************************************************************************************************************************************'''
    '''[Step3] Get the value of "power_binning" from Efuse, and update the related flags.
    '''Step3.1: gb_str_EfuseCategory_for_powerbinning is not empty if it is parsed from the header of column "PASS: power_binning" in power_binning flow table.
    '''Step3.2: If power_binning_flow table is correct, it can use gb_str_EfuseCategory_for_powerbinning = "power_binning" to find the matched Efuse category.
    '''Step3.3: If Efuse category "power_binning" exists in Efuse_BitDef_Table, read values from Efuse category "power_binning".
    '''***********************************************************************************************************************************************'''
    If gb_str_EfuseCategory_for_powerbinning <> "" Then
        '''//Get Efuse category "power_binning" for BinCut.
        str_Efuse_read_PowerBinning = get_Efuse_category_by_BinCut_testJob("read", gb_str_EfuseCategory_for_powerbinning)
    Else
        str_Efuse_read_PowerBinning = vbNullString
        Call Print_Error_Message(Warning_Info, "LIB_Vdd_Binning_customer", "Read_DVFM_To_GradeVDD", "No Efuse category for power_binning, it can't update the related failFlags in Read_DVFM_To_GradeVDD. Please check power_binning flow table and Efuse_BitDef_Table. Warning!!!")
    End If
    
    '''//Check values from Efuse category "power_binning" and update the flags in Bin_Table for Read_DVFM_To_GradeVDD.
    '''20210708: Modified to check if Efuse category "power_binning" for the current BinCut testJob has the matched programming stage in Efuse_BitDef_Table.
    If str_Efuse_read_PowerBinning <> "" Then
        '''//Please check if "power_binning" exists in Efuse_BitDef_Table and PowerBinning table "PwrBinning_V##".
        '''For project with Efuse obj vbt code.
        Set field = BdfDataBase.Bank_Cfg.Fields(str_Efuse_read_PowerBinning)
        siteLong_fusepowerBinning = field.DsscDecValue

        '''//According to the value of "power_binning" from Efuse, update the flag of "power_binning" in Bin_Table.
        For Each site In TheExec.sites
            '''***********************************************************************************************************************************************'''
            '''//Please check the following items:
            '''1. Column "PASS: power_binning" in PowerBinning flow table "PwrBinning_V***".
            '''2. Efuse category "power_binning" in Efuse_BitDef_Table.
            '''3. The correspondent flags and SortBin of power_binning in Bin_Table, ex: "F_PWRBIN_LOW", "F_PWRBIN_HIGH", "F_PWRBIN_HIGH", and "F_PWRBIN_LOWLOW".
            '''***********************************************************************************************************************************************'''
            '''For project with Efuse obj vbt code.
            If siteLong_fusepowerBinning(site) = 3 Then
                TheExec.sites.item(site).FlagState("F_PWRBIN_LOW") = logicTrue
            ElseIf siteLong_fusepowerBinning(site) = 2 Then
                TheExec.sites.item(site).FlagState("F_PWRBIN_HIGH") = logicTrue
            ElseIf siteLong_fusepowerBinning(site) = 1 Then
                TheExec.sites.item(site).FlagState("F_PWRBIN_LOWLOW") = logicTrue
            ElseIf siteLong_fusepowerBinning(site) = 0 Then
                '''do nothing, there is no power binning in binx/y(the value will be zero).
            Else
                '''20210610: Modified to check if it got the incorrect value "power_binning" from Efuse. If that, bin out the failed DUT.
                TheExec.sites.item(site).FlagState(strGlb_Flag_Vddbinning_Fail_Stop) = logicTrue
                '''For project with Efuse obj vbt code.
                Call Print_Error_Message(Warning_Info, "LIB_Vdd_Binning_customer", "Read_DVFM_To_GradeVDD", "Site:" & site & ",power binning value:" & siteLong_fusepowerBinning(site) & ", it is not defined in power binning table for Read_DVFM_To_GradeVDD,Error!!!")
                'TheExec.ErrorLogMessage "Site:" & site & ",power binning value:" & siteLong_fusepowerBinning(site) & ", it is not defined in power binning table for Read_DVFM_To_GradeVDD"    ''20230807: Removed by sanity check
            End If
        Next site
    End If '''If str_Efuse_read_PowerBinning <> ""
    
    '''//Insert 3 empty rows to separate blocks in the datalog.
    TheExec.Datalog.WriteComment ""
    TheExec.Datalog.WriteComment ""
    TheExec.Datalog.WriteComment vbCrLf
Exit Function
errHandler: 'Add ErrHandler 2023/08/01
    Call Print_Error_Message(Error_Info, "VBT_LIB_VDD_Binning", "Read_DVFM_To_GradeVDD") 'Add ErrHandler 2023/08/01
    If AbortTest Then Exit Function Else Resume Next 'Add ErrHandler 2023/08/01
End Function

'20210806: Modified to merge the branches for ReGenerate_IDS_ZONE_Voltage_Per_Site_ver2.
'20210729: Modified to check "AllBinCut(intAry_pmode(idx)).is_for_BinSearch=True" for Interpolation in the vbt function ReGenerate_IDS_ZONE_Voltage_Per_Site_ver2.
'20210405: Modified to check if powerDomain is one of BinCut CorePowers.
'20210405: Modified to check the input string is the correct powerDomain or p_mode for ReGenerate_IDS_ZONE_Voltage_Per_Site_ver2.
'20200730: Modified to print Dynamic_VBIN_IDS_ZONE for powerDomain.
'20200709: Modified to support interpolation of the single Pmode and powerDomain.
'20200709: Modified to skip interpolation if p_mode is used and tested.
'20200427: Modified to move "Flag_Interpolation_enable" from "ReGenerate_IDS_ZONE_Voltage_Per_Site_ver2" to "initVddBinning".
'20200106: Modified to remove the ErrorLogMessage.
'20191127: Modified for the revised InitVddBinTable.
'20190706: Modified to replace the hard-code "**_power_seq" with "BinCut_Power_Seq(VddBinStr2Enum(powerDomain)).Power_Seq(i)".
'20181026: Modified for interpolation, by Oscar.
Public Function ReGenerate_IDS_ZONE_Voltage_Per_Site_ver2(powerDomain As String)
On Error GoTo errHandler 'Add ErrHandler 2023/08/01
    Dim i As Long
    Dim p_mode As Integer
    Dim Flag_Print_Out_Dynamic_zone_voltage As Boolean: Flag_Print_Out_Dynamic_zone_voltage = False
    Dim intAry_pmode() As Integer
    Dim idx As Integer
    Dim isPmodeCorrect As Boolean
'''//==================================================================================================================================================================================//'''
'''//Note:
'''This is the EBB TTR function requested by C651 Shankar.
'''1. This function is to calculate the interpolated voltage for performance_mode by start_performance_mode(Int_Mode_L) and end_performancemode(Int_Mode_H).
'''2. It finds the nearest step (just great than or equal to the interpolated voltage) in Dynamic_IDS_Zone of p_mode.
'''//Method:
'''VddBinStr2Enum(powerDomain)).Power_Seq is array of all performance modes of the powerDomain.
'''We need this information to do interpolation for p_modes between start_performance_mode(Int_Mode_L) and end_performancemode(Int_Mode_H).
'''ReGenerate_DYNAMIC_IDS_ZONE_Voltage_Per_Site checked start_performance_mode(Int_Mode_L) and end_performancemode(Int_Mode_H) for p_mode.
'''//Important!!!
'''Remember to adjust the test flow to run both start_performance_mode and end_performancemode first.
'''20200709: C651 Si Li said that it can skip interpolation if p_mode is used and tested.
'''//==================================================================================================================================================================================//'''
    '''//Check if BinCut testJob is for BinCut search.
    '''//The flag "is_BinCutJob_for_StepSearch" is True if any testCondition from the table "Non_Binning_Rail" has the keyword "*Evaluate*Bin*".
    '''20210709: As the request from TER Verity, the branch of checking testJobs of BinCut search was removed.
    If is_BinCutJob_for_StepSearch = False Then
        '''//Align the format of Judge_PF and Judge_PF_func in the datalog.
        TheExec.Datalog.WriteComment ""
        Exit Function
    End If
    
    '''init
    isPmodeCorrect = False

    '''//Check if powerDomain is one of BinCut powerDomains or performance modes listed in sheet "Non_Binning_Rail".
    If VddbinPmodeDict.Exists(UCase(powerDomain)) = True Then
        p_mode = VddBinStr2Enum(powerDomain)
        
        '''*****************************************************************************************************'''
        '''//Tips:
        '''1. BinCut powerDomains (ex: VDD_PCPU) are enumerated from 1 to cntVddbinPin.
        '''2. BinCut performance_modes (ex: VDD_PCPU_MC601) are enumerated from (cntVddbinPin+1) to cntVddbinPmode.
        '''*****************************************************************************************************'''
        If p_mode > 0 And p_mode <= cntVddbinPin Then '''BinCut powerDomain.
            If dict_IsCorePowerInBinCutFlowSheet(UCase(powerDomain)) = True Then
                isPmodeCorrect = True
                
                '''//If it is one of BinCut CorePowers, resize the array to store all performance modes of the BinCut powerDomain, ex: VDD_PCPU includes MC601, MC602, MC603...
                ReDim intAry_pmode(UBound(BinCut_Power_Seq(VddBinStr2Enum(powerDomain)).Power_Seq))
                
                For i = 0 To UBound(BinCut_Power_Seq(VddBinStr2Enum(powerDomain)).Power_Seq)
                    intAry_pmode(i) = VddBinStr2Enum(BinCut_Power_Seq(VddBinStr2Enum(powerDomain)).Power_Seq(i))
                Next i
            Else
                isPmodeCorrect = False
                Call Print_Error_Message(Error_Info, "VBT_LIB_VDD_Binning", "ReGenerate_IDS_ZONE_Voltage_Per_Site_ver2", "powerDomain: " & powerDomain & " is not a powerDomain used for vddbinning. It can't be used for interpolation (ReGenerate_IDS_ZONE_Voltage_Per_Site_ver2). Error!!!")
                'TheExec.ErrorLogMessage "powerDomain: " & powerDomain & " is not a powerDomain used for vddbinning. It can't be used for interpolation (ReGenerate_IDS_ZONE_Voltage_Per_Site_ver2). Error!!!"    ''20230807: Removed by sanity check
            End If
        Else '''If p_mode > (cntVddbinPin) And p_mode <= cntVddbinPmode Then. It means that p_mode is performance_mode.
            isPmodeCorrect = True
            ReDim intAry_pmode(0)
            intAry_pmode(0) = p_mode
        End If
            
        If isPmodeCorrect = True Then
            TheExec.Datalog.WriteComment "Equation_N_Start"
            
            '''//If p_mode is used and tested, skip interpolation for p_mode. C651 Si Li said that only p_mode not tested can do interpolation.
'''ToDo: VBIN_RESULT(p_mode).tested is siteBoolean. Maybe we can use any method to check if VBIN_RESULT(p_mode).tested = True.
            For idx = 0 To UBound(intAry_pmode)
                '''//Check if p_mode is for BinCut search.
                '''20210806: Modified to merge the branches for ReGenerate_IDS_ZONE_Voltage_Per_Site_ver2.
                If AllBinCut(intAry_pmode(idx)).is_for_BinSearch = True And AllBinCut(intAry_pmode(idx)).Used = True And isPmodeTested(intAry_pmode(idx)) = False Then
                    ReGenerate_DYNAMIC_IDS_ZONE_Voltage_Per_Site intAry_pmode(idx), CurrentPassBinCutNum
                End If
            Next idx
            
            TheExec.Datalog.WriteComment "Equation_N_End"
            
            '''//Print Dynamic_VBIN_IDS_ZONE for powerDomain.
            If Flag_Print_Out_Dynamic_zone_voltage = True Then
                Print_DYNAMIC_VBIN_IDS_ZONE_to_sheet powerDomain
            End If
        End If
    Else '''If it isn't any BinCut powerDomain or performance mode.
         Call Print_Error_Message(Error_Info, "VBT_LIB_VDD_Binning", "ReGenerate_IDS_ZONE_Voltage_Per_Site_ver2", powerDomain & " isn't the correct BinCut powerDomain or p_mode for ReGenerate_IDS_ZONE_Voltage_Per_Site_ver2. Error!!!")
        ''TheExec.ErrorLogMessage powerDomain & " isn't the correct BinCut powerDomain or p_mode for ReGenerate_IDS_ZONE_Voltage_Per_Site_ver2. Error!!!"    ''20230807: Removed by sanity check
    End If
Exit Function
errHandler: 'Add ErrHandler 2023/08/01
    Call Print_Error_Message(Error_Info, "VBT_LIB_VDD_Binning", "ReGenerate_IDS_ZONE_Voltage_Per_Site_ver2") 'Add ErrHandler 2023/08/01
    If AbortTest Then Exit Function Else Resume Next 'Add ErrHandler 2023/08/01
End Function

'20210223: Modified to replace "Dim step As Long" with "Dim idx_step As Long".
'20200730: Modified to print Dynamic_VBIN_IDS_ZONE for powerDomain.
'20191204: Modified for the revised initVddBinTable.
Public Function Print_DYNAMIC_VBIN_IDS_ZONE_to_sheet(powerDomain As String)
On Error GoTo errHandler 'Add ErrHandler 2023/08/01
    Dim site As Variant
    Dim ws_def As Worksheet
    Dim wb As Workbook
    Dim Test_Type As testType
    Dim idx_step As Long
    Dim p_mode As Integer
    Dim i As Long
    Dim ids_range_step() As Long 'Dim ids_range_step(MaxPerformanceModeCount) As Long
    Dim IDS_current_Max() As Double 'Dim IDS_current_Max(MaxPerformanceModeCount) As Double
    Dim p_col As Integer, p_row As Integer
    Dim SheetCnt As Integer
    Dim SheetExist As Boolean
    Dim enable_Pmode As Boolean
    Dim sheetName As String
    '''init
    p_col = 1
    p_row = 1
    Test_Type = testType.TD
    For i = 0 To MaxPerformanceModeCount - 1
        ids_range_step(i) = 0
        IDS_current_Max(i) = 0
    Next i
    
    '''//Print Dynamic_VBIN_IDS_ZONE for powerDomain.
    sheetName = "DYNAMIC_VBIN_IDS_ZONE" + "_" + powerDomain

    SheetExist = False
    SheetCnt = ActiveWorkbook.Sheets.Count
    For i = 1 To SheetCnt
        If LCase(Sheets(i).name) Like LCase(sheetName) Then
            SheetExist = True
            Exit For
        End If
    Next i

    If SheetExist = False Then
        Worksheets.Add(after:=Worksheets(Worksheets.Count)).name = sheetName
    End If

    Set wb = Application.ActiveWorkbook
    Set ws_def = wb.Sheets(sheetName)
    ws_def.Select
    ws_def.Cells.Clear
    ws_def.Cells.Select
    Selection.ColumnWidth = 10

    '''Print out the header for ids zone voltage
    Cells(p_row, p_col).Select
    Selection.ColumnWidth = 18
    ws_def.Cells(p_row, p_col).value = "Site"
    p_col = p_col + 1
    Cells(p_row, p_col).Select
    Selection.ColumnWidth = 25
    ws_def.Cells(p_row, p_col).value = "Ids_zone Num belong to"
    p_col = p_col + 1

    For i = 0 To Max_V_Step_per_IDS_Zone - 1
        ws_def.Cells(p_row, p_col).value = "Vstep_" + CStr(i)
        p_col = p_col + 1
    Next i

'    For i = 0 To Max_V_Step_per_IDS_Zone - 1
'        ws_def.Cells(p_row, p_col).Value = "C"
'        p_col = p_col + 1
'    Next i
'
'    For i = 0 To Max_V_Step_per_IDS_Zone - 1
'        ws_def.Cells(p_row, p_col).Value = "M"
'        p_col = p_col + 1
'    Next i

    For i = 0 To Max_V_Step_per_IDS_Zone - 1
        ws_def.Cells(p_row, p_col).value = "EQ_Num"
        p_col = p_col + 1
    Next i

    For i = 0 To Max_V_Step_per_IDS_Zone - 1
        ws_def.Cells(p_row, p_col).value = "Bincut_Num"
        p_col = p_col + 1
    Next i

    p_row = p_row + 1

    For i = 0 To UBound(BinCut_Power_Seq(VddBinStr2Enum(powerDomain)).Power_Seq)
        p_mode = VddBinStr2Enum(BinCut_Power_Seq(VddBinStr2Enum(powerDomain)).Power_Seq(i))

        enable_Pmode = False

        For Each site In TheExec.sites
            If DYNAMIC_VBIN_IDS_ZONE(p_mode).Used = True Then
               enable_Pmode = True
               Exit For
            End If
        Next site

        If enable_Pmode = True Then
            p_col = 1
            ws_def.Cells(p_row, p_col).value = VddBinName(p_mode)
            p_row = p_row + 1
            For Each site In TheExec.sites
                p_col = 1
                ws_def.Cells(p_row, p_col).value = site
                p_col = p_col + 1
                ws_def.Cells(p_row, p_col).value = DYNAMIC_VBIN_IDS_ZONE(p_mode).IDS_ZONE_NUMBER
                p_col = p_col + 1
                p_col = 3

                '''//Print out calculated voltage
                For idx_step = 0 To DYNAMIC_VBIN_IDS_ZONE(p_mode).Max_Step - 1
                    ws_def.Cells(p_row, p_col).value = DYNAMIC_VBIN_IDS_ZONE(p_mode).Voltage(idx_step)
                    p_col = p_col + 1
                Next idx_step

                '''Print out "C"
'                p_col = 3 + Max_V_Step_per_IDS_Zone
'                For idx_step = 0 To DYNAMIC_VBIN_IDS_ZONE(p_mode).Max_Step - 1
'                    ws_def.Cells(p_row, p_col).Value = DYNAMIC_VBIN_IDS_ZONE(p_mode).C(idx_step)
'                    p_col = p_col + 1
'                Next idx_step

                '''//Print out "M"
'                p_col = 3 + Max_V_Step_per_IDS_Zone * 2
'                For idx_step = 0 To DYNAMIC_VBIN_IDS_ZONE(p_mode).Max_Step - 1
'                    ws_def.Cells(p_row, p_col).Value = DYNAMIC_VBIN_IDS_ZONE(p_mode).m(idx_step)
'                    p_col = p_col + 1
'                Next idx_step

                '''//Print out "EQ_Num"
                '''20200730: Modified to print Dynamic_VBIN_IDS_ZONE for powerDomain.
                '''<org>
                    'p_col = 3 + Max_V_Step_per_IDS_Zone * 3
                '''<new>
                p_col = 3 + Max_V_Step_per_IDS_Zone
                For idx_step = 0 To DYNAMIC_VBIN_IDS_ZONE(p_mode).Max_Step - 1
                    ws_def.Cells(p_row, p_col).value = DYNAMIC_VBIN_IDS_ZONE(p_mode).EQ_Num(idx_step)
                    p_col = p_col + 1
                Next idx_step

                '''//Print out "Bincut_Num"
                '''20200730: Modified to print Dynamic_VBIN_IDS_ZONE for powerDomain.
                '''<org>
                    'p_col = 3 + Max_V_Step_per_IDS_Zone * 4
                '''<new>
                p_col = 3 + Max_V_Step_per_IDS_Zone * 2
                For idx_step = 0 To DYNAMIC_VBIN_IDS_ZONE(p_mode).Max_Step - 1
                    ws_def.Cells(p_row, p_col).value = DYNAMIC_VBIN_IDS_ZONE(p_mode).passBinCut(idx_step)
                    p_col = p_col + 1
                Next idx_step

                p_row = p_row + 1
                p_row = p_row + 1
            Next site
        End If
    Next i
Exit Function
errHandler: 'Add ErrHandler 2023/08/01
    Call Print_Error_Message(Error_Info, "VBT_LIB_VDD_Binning", "Print_DYNAMIC_VBIN_IDS_ZONE_to_sheet") 'Add ErrHandler 2023/08/01
    If AbortTest Then Exit Function Else Resume Next 'Add ErrHandler 2023/08/01
End Function

'20201126: Modified to remove the redundant vbt code of CMEM datalog setup.
'20201125: As suggestion from Chihome, modified to set TheHdw.Digital.CMEM.CentralFields for initializing CMEM in GradeSearch_XXX_VT.
'20201103: Modified to check if the tester is not offline or not opensocket.
'20201030: Modified to clear CMEM.
'20190628: Modified the datalog initial setup for capture memeory of F.F.C (CMEM).
Public Function UpdateDLogColumns_Bincut(tsNameWidth As Long)
On Error GoTo errHandler 'Add ErrHandler 2023/08/01
    If (gB_newDlog_Flag) Then Exit Function

    tsNameWidth = CLng(tsNameWidth)
    TheExec.Datalog.Setup.Shared.ascii.Columns.EnableCustomWidths = True
    
    With TheExec.Datalog.Setup.Shared.ascii.Columns.Parametric
        .TestName.Enable = True
        .TestName.Width = tsNameWidth
        .pin.Enable = True
        .pin.Width = 40
    End With
    
    With TheExec.Datalog.Setup.Shared.ascii.Columns.Functional
        .Pattern.Enable = True
        .Pattern.Width = 128 '.Pattern.DefaultWidth
        .TestName.Enable = True
        .TestName.Width = tsNameWidth
    End With
    
    '''***************************************************************************'''
    '''//Initialize the datalog setup for capture memeory of F.F.C (CMEM).
    TheExec.Datalog.Setup.DatalogSetup.PartResult = True
    TheExec.Datalog.Setup.DatalogSetup.XYCoordinates = True
    TheExec.Datalog.Setup.DatalogSetup.DisableChannelNumberInPTR = True 'disable channel name to stdf, PE's datalog request -- 131225, chihome
    TheExec.Datalog.Setup.DatalogSetup.OutputWidth = 0
    
    If Flag_Enable_CMEM_Collection = True Then
        TheExec.Datalog.Setup.DatalogSetup.SetupStndInfo.FuncDispFormat() = 0 '''0: shortlog
    End If
    '''***************************************************************************'''
    '''//must need to apply after datalog setup
    TheExec.Datalog.ApplySetup
Exit Function
errHandler: 'Add ErrHandler 2023/08/01
    Call Print_Error_Message(Error_Info, "VBT_LIB_VDD_Binning", "UpdateDLogColumns_Bincut") 'Add ErrHandler 2023/08/01
    If AbortTest Then Exit Function Else Resume Next 'Add ErrHandler 2023/08/01
End Function

'20210709: Modified to remove the redundant branch of the vbt code because generate_offline_IDS_IGSim_Parallel was moved from check_IDS to Print_BinCut_config.
'20210708: Modified to move generate_offline_IDS_IGSim_Parallel from check_IDS to Print_BinCut_config.
'20210707: As per discussion with TSMC SWLINZA, BinCut retest and correlation should use Efuse IDS values only, not Efuse Product_Identifier!!!
'20210707: Modified to move init_IDS_ZONE_Voltage prior to site-loop because VBIN_IDS_ZONE(p_mode).Voltage were siteDouble.
'20210706: Modified to use the vbt function get_Efuse_category_by_BinCut_testJob to find the Efuse Category.
'20210705: Modified to find Efuse Category for product_identifier by checking the current BinCut testjob.
'20210621: Modified to move DisableCompare from the vbt function check_IDS to Print_BinCut_config.
'20210618: Modified to update SortNumber and binNumber if Vddbinning_IDS_fail is triggered in the vbt function check_IDS, as suggested by Chihome.
'20210617: Discussed this with TSMC T-Cre team and C651 Si. C651 Si said that Check_IDS and judge_IDS are dedicated to Efuse processed IDS for BinCut search.
'20210610: Modified to replace "SortNumber = 999 and binNumber = 10" with the triggered failflag strGlb_Flag_Vddbinning_IDS_fail for the vbt function check_IDS.
'20210507: Modified to remove the redundant testLimit for check_IDS.
'20210406: Modified to replace "bincutJobName = "cp1" with "is_BinCutJob_for_StepSearch = True".
'20210325: Modified to use Flag_Vddbin_DoAll_DebugCollection for TheExec.EnableWord("Vddbin_DoAll_DebugCollection").
'20201210: Modified to use the flag "is_BinCutJob_for_binSearch" for "check_bincutJob_for_binSearch" to check if the test program is binSearch or functional test.
'20201228: Patty asked us to adapt Efuse object vbt code.
'20200922: Modified to remove the redundant vbt code of "KeepAliveFlag".
'20200921: Discussed "RTOS_bootup_relay" / "KeepAliveFlag" / "spi_ttr_flag" with SWLINZA and PCLINZG. We decided to remove these SPI/RTOS branches because RTOS didn't use pattern test since Cebu/Sicily/Tonga/JC-Chop/Ellis/Bora.
'20200415: Modified to move "UpdateDLogColumns_Bincut" from check_IDS into Print_BinCut_config.
'20200317: Modified for SearchByPmode.
'20200210: Modified to use "Print_BinCut_config" to check GlobalVariable settings.
'20200129: Modified to check "DoAll" and "Override Fail-stop" in Run Options.
'20200120: Modified to check is powerDomain exists in domain2pinDict or pin2domainDict.
'20200102: Modified to use "Flag_BinCut_Config_Printed".
'20191219: Modified to enable word "Vddbin_DoAll_DebugCollection") for Bincut Do all debug.
'20191127: Modified for the revised InitVddBinTable.
'20190813: Modified to use different IDS lo_limit by BinCut testjobs.
'20190722: Modified to printout the scale and the unit for BinCut voltages and IDS values.
'20190716: Modified to unify the unit for IDS
'20190702: Modified to remove the hard-code "I_VDD_***" by the new data type "IDS_for_BinCut(VddBinStr2Enum(powerDomain)).Real".
'20190702: Modified to remove IDS_for_BinCut(VddBinStr2Enum("VDD_CPU_SRAM")).Real by pin-loop.
'20190626: Modified to unify the IDS scale in unit "A".
'20190606: Modified for CPIDS_Spec_OtherRail and FTIDS_Spec_OtherRail.
'20190523: Modified for the new data type "IDS_for_BinCut".
'20190507: Modified to add "Cdec" for IDS to avoid double format accuracy issues.
'20190104: Modified for GradeSearch on CP2.
'20180723: Modified for BinCut testjob mapping.
Public Function check_IDS(Optional RetestReset As Boolean = False)
On Error GoTo errHandler 'Add ErrHandler 2023/08/01
    Dim site As Variant
    Dim i As Integer, j As Integer
    Dim p_mode As Integer
    Dim powerDomain As String
    Dim powerPin As String
    Dim performance_mode As String
    Dim ids_lo_limit() As New SiteDouble 'Dim ids_lo_limit(1 To MaxBincutPowerdomainCount) As New SiteDouble
    Dim IsOtherRailInLimit As Boolean
    '''for Efuse obj vbt.
    Dim field As eFuseBdfField
    Dim opbank As eFuseBdfBank
    Dim Pass_Fail_Flag As New SiteBoolean
    '''20210707: Modified to remove the vbt code of updating Efuse Product_Identifier after generating Dynamic_IDS_Zone for each p_mode.
    Dim CurrentPassBinCutNumTemp As New SiteLong
    ReDim ids_lo_limit(1 To MaxPerformanceModeCount) As New SiteDouble

'''//==================================================================================================================================================================================//'''
'''//Note:
'''1. Check IDS values of BinCut CorePower and OtherRail with IDS limits from BinCut voltage tables (sheets "Vdd_Binning_Def").
'''2. For CP1, it decides IDS zone and IDS_Start_Step in Dynamic_IDS_Zone for each p_mode of CorePower.
'''3. For CP1, it decides start bin of OtherRail by IDS limits.
'''4. It updates flags "F_IDS_Binx" and "F_IDS_Biny" for Bin_Table.
'''5. Check_IDS and judge_IDS are dedicated to Efuse processed IDS for BinCut search, as requested by C651 Si, 20210617.
'''6. Warning!!!please check SortNumber and binNumber of Flag_Vddbinning_IDS_fail in Bin_Table before using this.
'''7. Since Efuse obj vbt code only provides one time permission to set the item, it can't update Efuse product_identifier in the vbt function check_IDS...
'''8. As per discussion with TSMC SWLINZA, BinCut retest and correlation should use Efuse IDS values only, but do not refer to Efuse Product_Identifier, 20210707.
'''9. As per discussion with TSMC SWLINZA, for powerPin group, it should use 1st powerPin to check IDS limit of powerPin group, 20210707.
'''ex: powerGroup: VDD_FIXED_GRP, and its 1st powerPin: VDD_FIXED, so that compare IDS value of VDD_FIXED with IDS_limit of VDD_FIXED_GRP. It must have Efuse category in Efuse_BitDef_Table to store IDS for VDD_FIXED.
'''10. C651 Toby provided the BinCut flow with testCondition "M*### E1 voltage" for non BinCut search in CP1, so that judge_stored_IDS(check_IDS) should be compatible with all BinCut testJobs, 20210727.
'''//==================================================================================================================================================================================//'''
    '''//Check if the Online tester has the incorrect Offline setting.
    If TheExec.TesterMode = testModeOnline And Flag_VDD_Binning_Offline = True Then
        Call Print_Error_Message(Error_Info, "VBT_LIB_VDD_Binning", "check_IDS", "Flag_VDD_Binning_Offline = True, but TheExec.TesterMode = testModeOnline. Test settings have conflicts for check_IDS. Error!!!")
        'TheExec.ErrorLogMessage "Flag_VDD_Binning_Offline = True, but TheExec.TesterMode = testModeOnline. Test settings have conflicts for check_IDS. Error!!!"    ''20230807: Removed by sanity check
        
        For Each site In TheExec.sites
            '''Warning!!!please check SortNumber and binNumber of Flag_Vddbinning_IDS_fail in Bin_Table before using this.
            '''20210618: Modified to update SortNumber and binNumber if Vddbinning_IDS_fail is triggered in the vbt function check_IDS, as suggested by Chihome.
            '''20240303: Use modular function to bin out site
            Call BinoutSite(site, VddbinIDSFailSortNumber, VddbinIDSFailBinNumber, strGlb_Flag_Vddbinning_IDS_fail)
        Next site
    End If
    
    If RetestReset = True Then
        CurrentPassBinCutNumTemp = CurrentPassBinCutNum
        TheExec.Datalog.WriteComment "inital Vbin Test Buffer"
        initVbinTest
        CurrentPassBinCutNum = CurrentPassBinCutNumTemp
        For Each site In TheExec.sites
            TheExec.Datalog.WriteComment CStr(site) & ",Update CurrentPassBinCutNum : " & CurrentPassBinCutNum(site)
        Next
    End If
    
    '''//Get IDS lo_limit for each BinCut OtherRail.
    For i = 0 To UBound(pinGroup_OtherRail)
        Call get_lo_limit_for_IDS(pinGroup_OtherRail(i), ids_lo_limit(VddBinStr2Enum(pinGroup_OtherRail(i)))) '''IDS lo_limit, unit: mA
    Next i
    
    '''//Initialize the voltage of each step in IDS_zone (set VBIN_IDS_ZONE(p_mode).Voltage = 0).
    init_IDS_ZONE_Voltage
    
    For Each site In TheExec.sites
        If RetestReset = False Then
                '''//Clear the flags.
                Binx_fail_power(site) = vbNullString
                Binx_fail_flag = False
                Biny_fail_power(site) = vbNullString
                Biny_fail_flag = False
        
                '''//Get IDS values for each BinCut powerDomain.
                '''********************************************************************************************************************'''
                '''1. Check_IDS and judge_IDS are dedicated to Efuse processed IDS for BinCut search, as requested by C651 Si.
                '''2. As per discussion with TSMC SWLINZA, BinCut retest and correlation should use Efuse IDS values only, but do not refer to Efuse Product_Identifier!!!
                '''3. Opensocket or offline simulation IDS values for are generated by generate_offline_IDS_IGSim_Parallel while Print_BinCut_config.
                '''********************************************************************************************************************'''
                For i = 0 To UBound(pinGroup_BinCut)
                    powerDomain = pinGroup_BinCut(i)
                    Call get_I_VDD_values(site, powerDomain, IDS_for_BinCut(VddBinStr2Enum(powerDomain)).Real)
                Next i
        End If
        
        '''//Generate IDS zone for each p_mode of CorePower.
        '''********************************************************************************************************************'''
        '''1. Judge the IDS Spec by real IDS current value.
        '''2. Base on IDS value to find out the IDS Zone Number.(for CP1 only)(for BinCut search)
        '''3. Base on IDS value to calculate the voltage for all IDS zone and steps.(for CP1 only)(for BinCut search)
        '''********************************************************************************************************************'''
        For p_mode = 0 To MaxPerformanceModeCount - 1
            If AllBinCut(p_mode).Used = True Then
                powerDomain = AllBinCut(p_mode).powerPin
                performance_mode = VddBinName(p_mode)
                
                '''//Check if the IDS value is in the range between IDS lo_limt and hi_limit for BinCut CorePower.
                If RetestReset = False Then judge_IDS IDS_for_BinCut(VddBinStr2Enum(powerDomain)).Real, performance_mode, site
                
                '''//Generate Dynamic_IDS_zone of p_mode and determine BinCut voltages for each step in Dynamic_IDS_zone.
                If IDS_for_BinCut(VddBinStr2Enum(powerDomain)).Real <> 0 Then
                    Find_IDS_ZONE_per_site IDS_for_BinCut(VddBinStr2Enum(powerDomain)).Real, p_mode
                    If Flag_IDS_Mapping_enable = True Then 'WWW20220630 for ids mapping
                        Generate_IDS_ZONE_Voltage_Per_Site IDS_for_BinCut(VddBinStr2Enum(AllBinCut(p_mode).IDS_MAPPING)).Real, p_mode
                    Else
                        Generate_IDS_ZONE_Voltage_Per_Site IDS_for_BinCut(VddBinStr2Enum(powerDomain)).Real, p_mode
                    End If
                    
                    Generate_DYNAMIC_IDS_ZONE_Voltage_Per_Site p_mode
                End If
            '20211022: Added to support SRAM Voltage calculate with equation
            ElseIf (AllBinCut(p_mode).Used = False And AllBinCut(p_mode).Mode_Step <> -1) _
                    And (dict_IsCorePower(AllBinCut(p_mode).powerPin) = True And UCase(AllBinCut(p_mode).powerPin) Like "*VDD*_SRAM*") Then
                powerDomain = AllBinCut(p_mode).powerPin
                performance_mode = VddBinName(p_mode)
                
                If RetestReset = False Then
                Call get_lo_limit_for_IDS(powerDomain, ids_lo_limit(VddBinStr2Enum(powerDomain)))
                judge_IDS IDS_for_BinCut(VddBinStr2Enum(powerDomain)).Real, performance_mode, site
                End If
                
                'If UCase(performance_mode) Like "*MPS*" Then
                If IDS_for_BinCut(VddBinStr2Enum(powerDomain)).Real <> 0 Then
                    Find_IDS_ZONE_per_site IDS_for_BinCut(VddBinStr2Enum(powerDomain)).Real, p_mode
                    '''20211202: IDS Mapping for SRAM domain
                    If Flag_IDS_Mapping_enable = True Then
                        Generate_IDS_ZONE_Voltage_Per_Site IDS_for_BinCut(VddBinStr2Enum(AllBinCut(p_mode).IDS_MAPPING)).Real, p_mode
                    Else
                        Generate_IDS_ZONE_Voltage_Per_Site IDS_for_BinCut(VddBinStr2Enum(powerDomain)).Real, p_mode
                    End If
                    Generate_DYNAMIC_IDS_ZONE_Voltage_Per_Site p_mode
                End If
'                ElseIf UCase(performance_mode) Like "*MES*" Then
'                    If IDS_for_BinCut(VddBinStr2Enum(powerDomain)).Real <> 0 Then
'                        Find_IDS_ZONE_per_site IDS_for_BinCut(VddBinStr2Enum(powerDomain)).Real, p_mode
'                        '''20211202: IDS Mapping for SRAM domain
'                        If Flag_IDS_Mapping_enable = True Then
'                            Generate_IDS_ZONE_Voltage_Per_Site IDS_for_BinCut(VddBinStr2Enum(AllBinCut(p_mode).IDS_MAPPING)).Real, p_mode
'                        Else
'                            Generate_IDS_ZONE_Voltage_Per_Site IDS_for_BinCut(VddBinStr2Enum("VDD_ECPU")).Real, p_mode
'                        End If
'                        Generate_DYNAMIC_IDS_ZONE_Voltage_Per_Site p_mode
'                    End If
'                End If
            End If
        Next p_mode
        
        '''//Check if the IDS value is in the range between IDS lo_limt and hi_limit for BinCut OtherRail.
        '''********************************************************************************************************************'''
        '''1. Loop the BinCut table, and use the different IDS limit to print out the datalog.
        '''2. Only judge the IDS by AllBinCut(P_mode).IDS_CP_LIMIT.
        '''********************************************************************************************************************'''
        '''20210617: Discussed this with TSMC T-Cre team and C651 Si. C651 Si said that Check_IDS and judge_IDS are dedicated to Efuse processed IDS for BinCut search.
        '''20210727: C651 Toby provided the BinCut flow with testCondition "M*### E1 voltage" for non BinCut search in CP1, so that judge_stored_IDS(check_IDS) should be compatible with all BinCut testJobs.
        If RetestReset = False Then
                For j = 0 To UBound(PassBinCut_ary)
                    '''Warning!!! Please check the unit of IDS applied to IDS values and limit.
                    '''//IDS calculation uses the scale and the unit in "mA", but TheExec.Flow.TestLimit should convert IDS value into "A" with settings "unit:=unitAmp" and "scaleMilli".
                    If j = UBound(PassBinCut_ary) Then
                        For i = 0 To UBound(pinGroup_OtherRail)
                            powerDomain = pinGroup_OtherRail(i)
                            '20240131 Bin1 BinX merge judge Bin1 IDS or BinX IDS
                            If Flag_BinX_Info_Parsed = True Then
                                If Pmode2BinX_IDS_Dict.Exists(VddBinStr2Enum(powerDomain)) Then
                                    If IDS_for_BinCut(VddBinStr2Enum(powerDomain)).Real(site) < Pmode2BinX_IDS_Dict(VddBinStr2Enum(powerDomain)) Then 'Bin1 IDS judge
                                        TheExec.Flow.TestLimit IDS_for_BinCut(VddBinStr2Enum(powerDomain)).Real(site) / 1000, ids_lo_limit(VddBinStr2Enum(powerDomain)) / 1000, Pmode2BinX_IDS_Dict(VddBinStr2Enum(powerDomain)) / 1000, , tlSignLess, scaleMilli, unit:=unitAmp, _
                                                    PinName:=powerDomain, Tname:=powerDomain & " BinCut1 IDS", ForceUnit:=unitAmp
                                    Else 'BinX IDS judge
                                        TheExec.Flow.TestLimit IDS_for_BinCut(VddBinStr2Enum(powerDomain)).Real(site) / 1000, _
                                                    ids_lo_limit(VddBinStr2Enum(powerDomain)) / 1000, AllBinCut(VddBinStr2Enum(powerDomain)).IDS_CP_LIMIT / 1000, , _
                                                    tlSignLess, scaleMilli, unit:=unitAmp, PinName:=powerDomain, Tname:=powerDomain & " BinCut1 IDS", ForceResults:=tlForcePass, ForceUnit:=unitAmp
                                                    
                                        TheExec.Flow.TestLimit IDS_for_BinCut(VddBinStr2Enum(powerDomain)).Real(site) / 1000, ids_lo_limit(VddBinStr2Enum(powerDomain)) / 1000, AllBinCut(VddBinStr2Enum(powerDomain)).IDS_CP_LIMIT / 1000, , tlSignLess, scaleMilli, unit:=unitAmp, _
                                                    PinName:=powerDomain, Tname:=powerDomain & " BinCut2 IDS", ForceUnit:=unitAmp
                                        Flag_Fuse_BinX(site) = True
                                    End If
                                Else
                                    Call Print_Error_Message(Error_Info, "VBT_LIB_VDD_Binning", "check_IDS", powerDomain & " doesn't have CP1 IDS limit, just use BinX CP1 IDS limit, please check definition sheet.")
                                    TheExec.Flow.TestLimit IDS_for_BinCut(VddBinStr2Enum(powerDomain)).Real(site) / 1000, ids_lo_limit(VddBinStr2Enum(powerDomain)) / 1000, AllBinCut(VddBinStr2Enum(powerDomain)).IDS_CP_LIMIT / 1000, , tlSignLess, scaleMilli, unit:=unitAmp, _
                                                    PinName:=powerDomain, Tname:=powerDomain & " BinCut2 IDS", ForceUnit:=unitAmp
                                End If
                            Else
                                TheExec.Flow.TestLimit IDS_for_BinCut(VddBinStr2Enum(powerDomain)).Real(site) / 1000, _
                                                    ids_lo_limit(VddBinStr2Enum(powerDomain)) / 1000, AllBinCut(VddBinStr2Enum(powerDomain)).IDS_CP_LIMIT / 1000, , _
                                                    tlSignLess, scaleMilli, unit:=unitAmp, PinName:=powerDomain, Tname:=powerDomain & " BinCut" & PassBinCut_ary(j) & " IDS", ForceUnit:=unitAmp
                            End If
                        Next i
                    Else
                        For i = 0 To UBound(pinGroup_OtherRail)
                            powerDomain = pinGroup_OtherRail(i)
                    
                            If CDec(IDS_for_BinCut(VddBinStr2Enum(powerDomain)).Real(site)) < CDec(gb_IDS_hi_limit(VddBinStr2Enum(powerDomain), PassBinCut_ary(j))) Then
                                TheExec.Flow.TestLimit IDS_for_BinCut(VddBinStr2Enum(powerDomain)).Real(site) / 1000, _
                                                        ids_lo_limit(VddBinStr2Enum(powerDomain)) / 1000, gb_IDS_hi_limit(VddBinStr2Enum(powerDomain), PassBinCut_ary(j)) / 1000, , _
                                                        tlSignLess, scaleMilli, unit:=unitAmp, PinName:=powerDomain, Tname:=powerDomain & " BinCut" & PassBinCut_ary(j) & " IDS", ForceUnit:=unitAmp
                            Else
                                TheExec.Flow.TestLimit IDS_for_BinCut(VddBinStr2Enum(powerDomain)).Real(site) / 1000, _
                                                        ids_lo_limit(VddBinStr2Enum(powerDomain)) / 1000, gb_IDS_hi_limit(VddBinStr2Enum(powerDomain), PassBinCut_ary(j)) / 1000, , _
                                                        tlSignLess, scaleMilli, unit:=unitAmp, PinName:=powerDomain, Tname:=powerDomain & " BinCut" & PassBinCut_ary(j) & " IDS", ForceResults:=tlForcePass, ForceUnit:=unitAmp
                            End If
                        Next i
                    End If
                Next j
        
                '''//Check CurrentPassbin and update the flags of "F_IDS_Binx" and "F_IDS_BinY" in Bin_Table.
                For j = 0 To UBound(PassBinCut_ary)
                    '''init
                    IsOtherRailInLimit = True
            
                    '''//Check if IDS values of OtherRail in limit.
                    For i = 0 To UBound(pinGroup_OtherRail)
                        powerDomain = pinGroup_OtherRail(i)
                        If Flag_BinX_Info_Parsed = True Then
                            If CDec(IDS_for_BinCut(VddBinStr2Enum(powerDomain)).Real(site)) <= CDec(Pmode2BinX_IDS_Dict(VddBinStr2Enum(powerDomain))) Then
                                'Bin1
                            ElseIf CDec(IDS_for_BinCut(VddBinStr2Enum(powerDomain)).Real(site)) <= AllBinCut(VddBinStr2Enum(powerDomain)).IDS_CP_LIMIT Then
                                TheExec.sites.item(site).FlagState("F_IDS_Binx") = logicTrue
                            Else
                                TheExec.sites.item(site).FlagState("F_IDS_BinY") = logicTrue
                            End If
                            IsOtherRailInLimit = False
                        '''20190716: Modified to unify the unit for IDS with mA.
                        ElseIf CDec(IDS_for_BinCut(VddBinStr2Enum(powerDomain)).Real(site)) < CDec(gb_IDS_hi_limit(VddBinStr2Enum(powerDomain), PassBinCut_ary(j))) Then
                            IsOtherRailInLimit = IsOtherRailInLimit And True
                        Else
                            IsOtherRailInLimit = IsOtherRailInLimit And False
                        End If
                    Next i
            
                    '''//OtherRail just updates CurrentPassBinCutNum as the worst case of passbin number.
                    If IsOtherRailInLimit = True Then
                        '''//Update BinCut PassBin.
                        If CurrentPassBinCutNum >= PassBinCut_ary(j) Then
                            '''Do nothing...
                        Else '''//If CurrentPassBinCutNum< PassBinCut_ary(i)
                            CurrentPassBinCutNum = PassBinCut_ary(j)
                        End If
                
                        '''//Update the flags of "F_IDS_Binx" and "F_IDS_BinY" in Bin_Table
                        If CurrentPassBinCutNum = 2 Then
                            TheExec.sites.item(site).FlagState("F_IDS_Binx") = logicTrue     'for Binx IDS binning
                        ElseIf CurrentPassBinCutNum = 3 Then
                            TheExec.sites.item(site).FlagState("F_IDS_BinY") = logicTrue     'for Biny IDS binning
                        End If
            
                        Exit For
                    End If
                Next j
        End If
    Next site
    
    '''//Print IDS_ZONE_voltage
    If Flag_Print_Out_tables_enable = True Then
        Print_IDS_ZONE_voltage_to_sheet
    End If
    
    '''//Align the format of Judge_PF and Judge_PF_func in the datalog.
    TheExec.Datalog.WriteComment ""
    TheExec.Datalog.WriteComment ""
    TheExec.Datalog.WriteComment vbCrLf
Exit Function
errHandler: 'Add ErrHandler 2023/08/01
    Call Print_Error_Message(Error_Info, "VBT_LIB_VDD_Binning", "check_IDS") 'Add ErrHandler 2023/08/01
    If AbortTest Then Exit Function Else Resume Next 'Add ErrHandler 2023/08/01
End Function

'20210910: Modified to revise the format of FirstChangeBinInfo, as requested by C651 Si and TSMC ZYLINI.
'20210823: Modified to replace "hiVal:=(UBound(PassBinCut_ary) + 1)" with "hiVal:=PassBinCut_ary(UBound(PassBinCut_ary))".
'20210819: Modified to add the special case to skip fusing product_identfier and product voltages for testJob "wlft2" of T-Cre.
'20210813: Modified to use BVmin(CPVmin) and BVmax(CPVmax) as lo_limit and hi_limit of BinCut voltage for p_mode, as C651 Toby's new rules about Guardband and product voltage.
'20210812: C651 Toby updated the rules that Product voltage=BinCut voltage+binning_GuardBand.
'20210730: C651 Toby revised the rule of adjust_VddBinning that "If CP1 has only PCPU bin cut search, then we only need PCPU for PrintOut_VddBinning and Adjust_VddBinning. The other domains can be skipped at CP1."
'20210726: Modified to get CPIDSMax by passBinCut and step_in_BinCut (EQN-1) for p_mode.
'20210726: Modified to update values (CurrentPassBinCutNum(site) - 1) to Efuse category "Product_Identifier" for adjust_VddBinning.
'20210720: Modified to check if Efuse category "Product_Identifier" exists in Efuse_BitDef_Table.
'20210710: As per discussion with TSMC ZYLINI, Crete had Efuse category "Product_Identifier" and "Product_Identifier_cp1" for testJob "cp1", but Efuse postCheck only supported "Product_Identifier". We had to use hard-code "Product_Identifier" here as Efuse workaround for Crete.
'20210709: TER Leon said that using forceWrite can write the same Efuse category at 2nd time for Efuse obj vbt.
'20210709: As the request from TER Verity, the branch of checking testJobs of BinCut search was removed.
'20210707: Modified to check if it's OK to write Efuse category "Product_Identifier".
'20210706: Modified to use the vbt function get_Efuse_category_by_BinCut_testJob to find the Efuse Category.
'20210705: Modified to find Efuse Category for product_identifier by checking the current BinCut testjob.
'20210629: Modified to check the EnableWord("Vddbin_PTE_Debug")=False to determine "Print Bincut Fail Info", as suggested by Chihome.
'20210629: Modified to update the fail-stop failFlag.
'20210623: Modified to get lo_limit and hi_limit of GradeVDD.
'20210621: Modified to remove testLimit and use the fail-stop flag here.
'20210621: Modified to revise the vbt function adjust_VddBinning for BinCut search in FT.
'20210621: Modified to use the revised vbt function check_voltageInheritance_for_powerDomain.
'20210610: Modified to add branches for the failed site.
'20210610: Modified to replace VBIN_RESULT(p_mode).passBinCut with CurrentPassBinCutNum for IDS.
'20210528: Modified to revise PTR format of "Monotonicity_Offset" with performance_mode, requested by C651 and TSMC ZQLin.
'20210526: Modified to add "VBIN_Result(p_mode).is_Monotonicity_Offset_triggered" for Monotonicity_Offset check because C651 Si revised the check rules to ensure that: (GradeVDD(P_mode)-GradeVDD(previous perfromance_mode))> Monotonicity_Offset.
'20210118: Modified to avoid Efuse incorrect flag.
'20210108: Modified the output format of Print Bincut Fail Info.
'20210107: Modified to add for recording first changed binnum mode data, requested by C651 Si.
'20201228: Patty asked us to adapt Efuse object vbt code.
'20201210: Modified to use the flag "is_BinCutJob_for_binSearch" for "check_bincutJob_for_binSearch" to check if the test program is binSearch or functional test.
'20200807: Modified to check BinCut testJob mapping.
'20200730: Modified to set "showPrint"=true for Efuse checkscript.
'20200424: Modified to use "Check_Adjust_Max_Min" to check Adjust_Max and Adjust_Min for adjust_VddBinning.
'20200225: Modified to change the data format for Efuse DSP func.
'20200106: Modified to remove the ErrorLogMessage.
'20191127: Modified to use the revised InitVddBinTable.
'20191023: Modified to check if "MaxPV(pmode0/pmode1)" is in the column "Comment" or not.
'20190722: Modified to print out the scale and the unit for BinCut voltages and IDS values.
'20190722: Modified to print out BinCut EQN and Bin values.
'20190717: Modified the printing format with "#.000" for EQN.
'20190716: Modified to unify the unit for IDS.
'20190706: Modified to replace the hard-code "**_power_seq" with "BinCut_Power_Seq(VddBinStr2Enum(powerDomain)).Power_Seq(i)".
'20190626: Modified to change IDS printing format.
'20190529: Modified to use the new data type "IDS_for_BinCut".
'20190507: Modified to use "Cdec" to avoid double format accuracy issues.
'20190108: Modified to correct Adjust_Min_Enable and Adjust_Max_Enable.
'20181226: Modified to add the optional parameters of ScaleType:=scaleNoScaling, FormatStr:="%.3f" for TheExec.Flow.TestLimit to avoid the truncation issue when voltage > 1v.
Public Function adjust_VddBinning(Adjust_Max_Enable As Boolean, Adjust_Min_Enable As Boolean, Optional Adjust_Power_Max_list As String, Optional Adjust_Power_Min_list As String)
On Error GoTo errHandler 'Add ErrHandler 2023/08/01
    Dim site As Variant
    Dim p_mode As Integer
    Dim dbl_ids_hi_limit As Double
    Dim i As Integer
    Dim j As Integer
    Dim binx_flag_name As String
    Dim biny_flag_name As String
    Dim nu As Long
    Dim counter As Long
    Dim strOutput() As String
    Dim site_num As Integer
    Dim powerDomain As String
    Dim EQN_lowest As Long
    '''variant
    Dim step_lowest As Long
    Dim dbl_CPVmin As Double
    Dim dbl_CPVmax As Double
    Dim str_PmodeGradeTestJob As String
    Dim dbl_GB_BinCutJob As Double
    Dim dbl_gradeVdd_lolimit As Double
    Dim dbl_gradeVdd_hilimit As Double
    Dim PassBinNum As Long
    Dim str_Efuse_write_ProductIdentifier As String
    Dim local_efuseval As New SiteLong '''Add for new efuse DSP func
    Dim str_Efuse_write_pmode As String
    '''for Efuse odj vbt.
    Dim Pass_Fail_Flag As New SiteBoolean: Pass_Fail_Flag = True
    Dim opbank As eFuseBdfBank
    Dim strBINXOfPowerDomain() As String ''20240221 Summary of BINX which power domain have adjust
    Dim count_strBINXOfPowerDomain As Long: count_strBINXOfPowerDomain = 0 ''20240221 count how much domain of every site have been adjust
'    Dim Flag_Fuse_BinX As New SiteBoolean  ''20240118: Move to global variable


'''//==================================================================================================================================================================================//'''
'''//Note:
'''1. For CP1, check monotonicity of Grade (CP) and GradeVdd (Efuse product voltages) between performance modes for the powerDomain.
'''2. It updates flags "F_IDS_Binx" and "F_IDS_Biny" for Bin_Table.
'''3. Print the status of VBIN_RESULT(p_mode).is_Monotonicity_Offset_triggered(site) with PTR format in Adjust_Binning for datalogs.
'''4. TER Leon said that using forceWrite can write the same Efuse category at 2nd time for Efuse obj vbt, 20210709.
'''5. C651 Toby revised the rule of adjust_VddBinning that "If CP1 has only PCPU bin cut search, then we only need PCPU for PrintOut_VddBinning and Adjust_VddBinning. The other domains can be skipped at CP1.", 20210730.
'''6. C651 Toby updated the rules that Product voltage=BinCut voltage+binning_GuardBand, 20210812.
'''Use BVmin(CPVmin) and BVmax(CPVmax) as lo_limit and hi_limit of BinCut voltage for p_mode, as C651 Toby's new rules about Guardband and product voltage.
'''7. If the testJob is for BinCut search, it should have the dedicated "Product_Identifier", as commented by C651 Si and Toby.
'''//==================================================================================================================================================================================//'''
    '''//Check if BinCut testJob is for BinCut search.
    '''//The flag "is_BinCutJob_for_StepSearch" is True if any testCondition from the table "Non_Binning_Rail" has the keyword "*Evaluate*Bin*".
    If is_BinCutJob_for_StepSearch = False Then
        '''//Align the format of Judge_PF and Judge_PF_func in the datalog.
        TheExec.Datalog.WriteComment ""
        Exit Function
    End If
    
    '''//Get Efuse category "Product_Identifier" for BinCut.
    str_Efuse_write_ProductIdentifier = get_Efuse_category_by_BinCut_testJob("write", "Product_Identifier")
    
    '''//If the testJob is for BinCut search, it should have the dedicated "Product_Identifier", as commented by C651 Si and Toby.
    If str_Efuse_write_ProductIdentifier <> "" Then
        If Flag_Remove_Printing_BV_voltages = False Then
            TheExec.Datalog.WriteComment "Product_Identifier" & ",it can use Efuse category:" & str_Efuse_write_ProductIdentifier
        End If
    Else
        Call Print_Error_Message(Error_Info, "VBT_LIB_VDD_Binning", "adjust_VddBinning", "No Efuse category for updating Product_Identifier for adjust_VddBinning, it can't be fused in the current testJob. Error!!!")
        'TheExec.ErrorLogMessage "No Efuse category for updating Product_Identifier for adjust_VddBinning, it can't be fused in the current testJob. Error!!!"    ''20230807: Removed by sanity check
    End If
    
    '''init
    site_num = TheExec.sites.Existing.Count
    ReDim strOutput(MaxPerformanceModeCount * site_num) As String
    ReDim strBINXOfPowerDomain(cntVddbinPin * site_num) As String
    
    '''//Check Adjust_Max and Adjust_Min from the arguments of the test instance "Adjust_VddBinning".
    Call Check_Adjust_Max_Min(Adjust_Max_Enable, Adjust_Min_Enable, Adjust_Power_Max_list, Adjust_Power_Min_list)
    
    For Each site In TheExec.sites
        For i = 0 To UBound(pinGroup_CorePower)
            '''20240220 adjust all pmode to BINX of specified powerdomain
            check_pmode_BINX_adjust_powerdomainEveryPmode pinGroup_CorePower(i), site, strBINXOfPowerDomain, count_strBINXOfPowerDomain
        Next i
    Next site
        
        '''20211130 dict_cross_domain_compare VDD_GPU0~VDD_GPUX
    If Flag_cross_domain_check_enable = True Then
        For i = 0 To UBound(pinGroup_CorePower)
            If Not UCase(pinGroup_CorePower(i)) Like "*VDD*_SRAM*" Then
                Call check_cross_voltageInheritance_for_powerDomain(pinGroup_CorePower(i))
            End If
        Next i
    End If

    '''//Check the voltage inheritance between p_mode and the previous performance_mode for each BinCut powerDomain.
    For i = 0 To UBound(pinGroup_CorePower)
        Call check_voltageInheritance_for_powerDomain(pinGroup_CorePower(i))
    Next i
        

    For Each site In TheExec.sites
        '''*****************************************************************************************'''
        '''ToDo: Maybe we can use this fail-stop flag to mask the failed DUT in Adjust_VddBinning...
        'If TheExec.sites.Item(site).FlagState(strGlb_Flag_Vddbinning_Fail_Stop) = logicFalse then
        '''*****************************************************************************************'''
        '''//After checking voltage inheritance of p_mode, print PassBin/EQN/voltage for each p_mode.
        For p_mode = 0 To MaxPerformanceModeCount - 1
            If IsExcludedVddBin(p_mode) Then
                '''skip it If the performance mode is not enabled.
            Else
                '20211022: Modify to support SRAM Voltage calculate with equation
                If AllBinCut(p_mode).Used = True Or _
                (AllBinCut(p_mode).Used = True Or (dict_IsCorePower(AllBinCut(p_mode).powerPin) = True And UCase(AllBinCut(p_mode).powerPin) Like "*VDD*_SRAM*")) Then
                    '''//Get powerDomain for the performance_mode.
                    powerDomain = AllBinCut(p_mode).powerPin
                    
                    '''//Check if p_mode for BinCut search has the dedicated Efuse category in the current testJob.
                    '''20210730: C651 Toby revised the rule of adjust_VddBinning that "If CP1 has only PCPU bin cut search, then we only need PCPU for PrintOut_VddBinning and Adjust_VddBinning. The other domains can be skipped at CP1."
                    '''//Get Efuse category of Efuse product voltage(GradeVDD) for p_mode.
                    str_Efuse_write_pmode = get_Efuse_category_by_BinCut_testJob("write", VddBinName(p_mode))
                    
                    '''//If p_mode has Efuse category for the current BinCut testJob, get Efuse product voltage for P_mode.
                    If str_Efuse_write_pmode <> "" Then
                        '''********************************************************************************************************************'''
                        '''[Step1] Check PassBin and EQN Num.
                        '''********************************************************************************************************************'''
                        '''//PassBin from BinCut search result of p_mode for each site.
                        PassBinNum = VBIN_RESULT(p_mode).passBinCut
                        
                        '''//Get CPVmin / CPVmax / Mode_Step of p_mode.
                        '''//step_lowest is the max Equation number stored of p_mode, ex: BinCut(P_mode,PassBinNum).EQ_Num(Max_EQ_Num_address)=Max EQ Number.
                        step_lowest = BinCut(p_mode, PassBinNum).Mode_Step
                        dbl_CPVmin = BinCut(p_mode, PassBinNum).CP_Vmin(step_lowest)
                        dbl_CPVmax = BinCut(p_mode, PassBinNum).CP_Vmax(0)
                        '''EQN_lowest = MODE_STEP + 1
                        EQN_lowest = BinCut(p_mode, PassBinNum).Mode_Step + 1
    
                        '''//Get the matched Guardband(GB) according to the BinCut testjob.
                        '''<org>
                            'str_PmodeGradeTestJob=VddBinName(p_mode) & " CP"
                        '''
                        Select Case LCase(bincutJobName)
                            Case "cp1":
                                dbl_GB_BinCutJob = BinCut(p_mode, PassBinNum).CP_GB(0)
                                str_PmodeGradeTestJob = VddBinName(p_mode) & " CP1"
                            Case "cp2":
                                dbl_GB_BinCutJob = BinCut(p_mode, PassBinNum).CP2_GB(0)
                                str_PmodeGradeTestJob = VddBinName(p_mode) & " CP2"
                            Case "ft_room":
                                dbl_GB_BinCutJob = BinCut(p_mode, PassBinNum).FT1_GB(0)
                                str_PmodeGradeTestJob = VddBinName(p_mode) & " FT1"
                            Case "ft_hot":
                                dbl_GB_BinCutJob = BinCut(p_mode, PassBinNum).FT2_GB(0)
                                str_PmodeGradeTestJob = VddBinName(p_mode) & " FT2"
                            Case "qa":
                                dbl_GB_BinCutJob = BinCut(p_mode, PassBinNum).FTQA_GB(0)
                                str_PmodeGradeTestJob = VddBinName(p_mode) & " QA"
                            Case Else:
                                dbl_GB_BinCutJob = 0
                                str_PmodeGradeTestJob = vbNullString
                                TheExec.sites.item(site).FlagState(strGlb_Flag_Vddbinning_Fail_Stop) = logicTrue
                                Call Print_Error_Message(Error_Info, "VBT_LIB_VDD_Binning", "adjust_VddBinning", "site:" & site & ", adjust_vddbinning has the incorrect BinCut TestJob selection. Error!!!")
                                ''TheExec.ErrorLogMessage "site:" & site & ", adjust_vddbinning has the incorrect BinCut TestJob selection. Error!!!"    ''20230807: Removed by sanity check
                        End Select
                        
                        '''//PassBinNumber.
                        '''20210823: Modified to replace "hiVal:=(UBound(PassBinCut_ary) + 1)" with "hiVal:=PassBinCut_ary(UBound(PassBinCut_ary))".
                        TheExec.Flow.TestLimit resultVal:=VBIN_RESULT(p_mode).passBinCut, lowVal:=1, hiVal:=PassBinCut_ary(UBound(PassBinCut_ary)), _
                                                scaletype:=scaleNoScaling, unit:=unitNone, formatStr:="%.0f", Tname:=VddBinName(p_mode) & " BinCut Num"
                        
                        '''//EQN
                        TheExec.Flow.TestLimit VBIN_RESULT(p_mode).step_in_BinCut + 1, 1, EQN_lowest, , , _
                                                scaleNoScaling, unitNone, formatStr:="%.0f", Tname:=VddBinName(p_mode) & " EQN"
                        
                        If TheExec.enableWord("T0TX_ROOM") = True And Not TheExec.enableWord("HardIP_RMA") = True Then 'WWWWW
                            dbl_GB_BinCutJob = BinCut(p_mode, PassBinNum).HTOL_RO_GB_ROOM(0)
                            str_PmodeGradeTestJob = VddBinName(p_mode) & bincutJobName
                        End If
                        If TheExec.enableWord("T0TX_HOT") = True And Not TheExec.enableWord("HardIP_RMA") = True Then 'WWWWW
                            dbl_GB_BinCutJob = BinCut(p_mode, PassBinNum).HTOL_RO_GB_HOT(0)
                            str_PmodeGradeTestJob = VddBinName(p_mode) & bincutJobName
                        End If
                        '''********************************************************************************************************************'''
                        '''[Step2] Check LVCC and Efuse Product voltage.
                        '''Efuse Product voltage (GradeVdd)  = LVCC voltage (Grade) + CPGB.
                        '''********************************************************************************************************************'''
                        '''//Get lo_limit and hi_limit for product voltage of p_mode.
                        dbl_gradeVdd_lolimit = dbl_CPVmin + BinCut(p_mode, PassBinNum).CP_GB(step_lowest)
                        dbl_gradeVdd_hilimit = dbl_CPVmax + BinCut(p_mode, PassBinNum).CP_GB(0)
                        
                        '''//Check if BinCut voltage(Grade) and Efuse Product voltage(GradeVdd) of p_mode are in limit.
                        '''//BinCut voltage(Grade)
                        '''20210812: C651 Toby updated the rules that Product voltage=BinCut voltage+binning_GuardBand.
                        '''20210813: Use BVmin(CPVmin) and BVmax(CPVmax) as lo_limit and hi_limit of BinCut voltage for p_mode, as C651 Toby's new rules about Guardband and product voltage.
                        TheExec.Flow.TestLimit VBIN_RESULT(p_mode).GRADE / 1000, (dbl_CPVmin) / 1000, (dbl_CPVmax) / 1000, _
                                                    Tname:=str_PmodeGradeTestJob, scaletype:=scaleMilli, unit:=unitVolt, formatStr:="%.3f", ForceUnit:=unitVolt
                        
                        '''//Efuse Product voltage(GradeVdd)
                        TheExec.Flow.TestLimit VBIN_RESULT(p_mode).GRADEVDD / 1000, dbl_gradeVdd_lolimit / 1000, dbl_gradeVdd_hilimit / 1000, _
                                                Tname:=VddBinName(p_mode) & " VDD Define", scaletype:=scaleMilli, unit:=unitVolt, formatStr:="%.3f", ForceUnit:=unitVolt
                        
                        '''//Check if Monotonicity_Offset is triggered.
                        '''//C651 Si revised the check rules of Montonicity_Offset to ensure that: (GradeVDD(P_mode)-GradeVDD(previous perfromance_mode))> Monotonicity_Offset.
                        '''20210528: Modified to revise PTR format of "Monotonicity_Offset" with performance_mode, requested by C651 and TSMC ZQLin.
                        If Flag_Get_column_Monotonicity_Offset = True Then
                            TheExec.Flow.TestLimit Abs(CLng(VBIN_RESULT(p_mode).is_Monotonicity_Offset_triggered(site))), 0, 1, _
                                                    , , scaleNoScaling, unitNone, "%.0f", Tname:=VddBinName(p_mode) & " Monotonicity_Offset"
                        End If
                        
                        '''********************************************************************************************************************'''
                        '''[Step3] Check IDS.
                        '''********************************************************************************************************************'''
                        '''//Get CPIDSMax by passBinCut and step_in_BinCut (EQN-1) for p_mode.
                        '''//IDS calculation uses the scale and the unit in "mA", but TheExec.Flow.TestLimit should convert IDS value into "A" with settings "unit:=unitAmp" and "scaleMilli".
                        '''Note: Check_IDS and judge_IDS are dedicated to Efuse processed IDS for BinCut search, as requested by C651 Si, 20210617.
                        If VBIN_RESULT(p_mode).step_in_BinCut < 0 Then Exit For     ''20221020: Added to avoid DoAll no shutdown scenario
                       
                        '20230216: Modified to use the IDS_Mapping domain's limit
                        'WWW20220630We modify IDS_Mapping to mapping the correct domain
                        If Flag_IDS_Mapping_enable = True And AllBinCut(p_mode).IDS_MAPPING <> AllBinCut(p_mode).powerPin Then
                            '20230216: Modified to use the IDS_Mapping domain's limit
                            dbl_ids_hi_limit = AllBinCut(p_mode).IDS_MAPPING_LIMIT '''unit: mA

                            TheExec.Flow.TestLimit IDS_for_BinCut(VddBinStr2Enum(AllBinCut(p_mode).IDS_MAPPING)).Real(site) / 1000, 0, dbl_ids_hi_limit / 1000, _
                                                , tlSignLess, scaleMilli, unitAmp, , Tname:=VddBinName(p_mode) & " IDS", ForceUnit:=unitAmp 'ids mapping
                        Else    ''Use its own IDS
                            dbl_ids_hi_limit = BinCut(p_mode, VBIN_RESULT(p_mode).passBinCut).IDS_CP_LIMIT(VBIN_RESULT(p_mode).step_in_BinCut) '''unit: mA
                        
                            TheExec.Flow.TestLimit IDS_for_BinCut(VddBinStr2Enum(powerDomain)).Real(site) / 1000, 0, dbl_ids_hi_limit / 1000, _
                                                , tlSignLess, scaleMilli, unitAmp, , Tname:=VddBinName(p_mode) & " IDS", ForceUnit:=unitAmp 'original
                        End If
                                                
                        '''//Bin out device if the adjust level is with higher IDS than that step.
                        '''Note: As suggested by CBCHENI, assumed default testType as "TD".
                        If CDec(IDS_for_BinCut(VddBinStr2Enum(powerDomain)).Real(site)) > CDec(dbl_ids_hi_limit) Then
                            '''Warning!!!please check SortNumber and binNumber of Flag_Vddbinning_Fail_Stop in Bin_Table before using this.
                            '''20240303: Use modular function to bin out site
                            Call BinoutSite(site, BinCut(p_mode, VBIN_RESULT(p_mode).passBinCut).SBIN_BINNING_FAIL(VBIN_RESULT(p_mode).step_in_BinCut, 0), _
                                                BinCut(p_mode, VBIN_RESULT(p_mode).passBinCut).HBIN_BINNING_FAIL(VBIN_RESULT(p_mode).step_in_BinCut, 0), strGlb_Flag_Vddbinning_Fail_Stop)
                        End If
                        
                        '''********************************************************************************************************************'''
                        '''[Step4] Set the flag to update the flag in Bin_Table for BinX and BinY.
                        '''********************************************************************************************************************'''
                        '''//Set flags of bintable according to different ids range.
                        If VBIN_RESULT(p_mode).passBinCut = 2 Then '''BinX
                            If CDec(IDS_for_BinCut(VddBinStr2Enum(powerDomain)).Real(site)) >= CDec(BinCut(p_mode, 1).IDS_CP_LIMIT(BinCut(p_mode, 1).Mode_Step)) Then '''unit: mA
                                TheExec.sites.item(site).FlagState("F_IDS_Binx") = logicTrue
                            End If
                            
                            binx_flag_name = "F_Binx_" & Binx_fail_power
                            TheExec.sites.item(site).FlagState(binx_flag_name) = logicTrue
                        ElseIf VBIN_RESULT(p_mode).passBinCut = 3 Then '''BinY
                            If CDec(IDS_for_BinCut(VddBinStr2Enum(powerDomain)).Real(site)) >= CDec(BinCut(p_mode, 2).IDS_CP_LIMIT(BinCut(p_mode, 2).Mode_Step)) Then
                                TheExec.sites.item(site).FlagState("F_IDS_BinY") = logicTrue
                            End If
                            
                            biny_flag_name = "F_Biny_" & Biny_fail_power
                            TheExec.sites.item(site).FlagState(biny_flag_name) = logicTrue
                        End If
                        
                        '''********************************************************************************************************************'''
                        '''[Step5] Store the info about adjust_vddBinning in DTR of STDF.
                        '''It will print the info at the end of adjust_VddBinning.
                        '''20190716: Modified to unify the unit for IDS with mA.
                        '''********************************************************************************************************************'''
                        strOutput(nu) = "VBIN," & "1," & site & "," & VddBinName(p_mode) & "," & IDS_for_BinCut(VddBinStr2Enum(powerDomain)).ids_name(site) & "," & _
                                        Int(IDS_for_BinCut(VddBinStr2Enum(powerDomain)).Real(site)) & "," & Int(VBIN_RESULT(p_mode).GRADEVDD) & "," & VBIN_RESULT(p_mode).step_in_BinCut + 1
                        nu = nu + 1
                    End If '''If str_Efuse_write_pmode <> ""
                End If '''If AllBinCut(p_mode).Used = True
            End If '''If IsExcludedVddBin(p_mode) Then
        Next p_mode

        '''********************************************************************************************************************'''
        '''Added for parsing sheet defining BinX EQNs and IDS value
        '''20230504: Decide correct bin number by EQN result
        '''20230706: Modified to use independent p_mode loop for binX ids & binX EQN, to get the correct binX trigger point, requested by Si
        '''20230706: Added to print PTR format of trigger point
        '''********************************************************************************************************************'''
        If Flag_BinX_Info_Parsed = True Then
            ''Step 1: To find if there's any Pmode contains binX IDS value
            For p_mode = 0 To MaxPerformanceModeCount - 1
                If Pmode2BinX_IDS_Dict.Exists(p_mode) = True And Flag_Fuse_BinX(site) = False Then
                    '''//Get powerDomain for the performance_mode.
                    powerDomain = AllBinCut(p_mode).powerPin
'                    If Flag_IDS_Mapping_enable = True Then
'                        If CDec(IDS_for_BinCut(VddBinStr2Enum(AllBinCut(p_mode).IDS_MAPPING)).Real) > CDec(Pmode2BinX_IDS_Dict(p_mode)) And CurrentPassBinCutNum(site) = 1 Then
'                            Flag_Fuse_BinX(site) = True
'                            ' TheExec.Datalog.WriteComment "site:" & site & ", IDS=" & CDec(IDS_for_BinCut(VddBinStr2Enum(powerDomain)).Real(site)) & " > binX limit=" & CDec(Pmode2BinX_IDS_Dict(p_mode)) _
'                            '                                 & ", Adjust fusing bin num to binX"
'                            TheExec.Datalog.WriteParametricResult site, TheExec.sites.item(site).TestNumber, logTestPass, parmTestLim, "", "", _
'                                                                0, 1, 1, unitNone, 0, unitNone, 0, , , powerDomain & " IDS BinX", scaleNoScaling, "%.0f"
'                        End If
'                    Else
                        If CDec(IDS_for_BinCut(VddBinStr2Enum(powerDomain)).Real(site)) > CDec(Pmode2BinX_IDS_Dict(p_mode)) And CurrentPassBinCutNum(site) = 1 Then
                            Flag_Fuse_BinX(site) = True
                            TheExec.Datalog.WriteParametricResult site, TheExec.sites.item(site).TestNumber, logTestPass, parmTestLim, "", "", _
                                                                0, 1, 1, unitNone, 0, unitNone, 0, , , powerDomain & " IDS BinX", scaleNoScaling, "%.0f"
                        End If
'                    End If
                End If
            Next p_mode
            ''Step 2: To find if there's any Pmode contains binX EQN
             For p_mode = 0 To MaxPerformanceModeCount - 1
                If Pmode2BinX_EQN_Dict.Exists(p_mode) And Check_EQN_UpToBinX(p_mode, site) = True And Flag_Fuse_BinX(site) = False Then
                    Flag_Fuse_BinX(site) = True
                    ' TheExec.Datalog.WriteComment "site:" & site & "," & VddBinName(p_mode) & ", VBinResult EQN=" & VBIN_RESULT(p_mode).step_in_BinCut + 1 _
                    '                             & ", matched the definition with 'EQN_Bin', Adjust fusing bin num to binX"
                    TheExec.Datalog.WriteParametricResult site, TheExec.sites.item(site).TestNumber, logTestPass, parmTestLim, "", "", _
                                                    0, 1, 1, unitNone, 0, unitNone, 0, , , VddBinName(p_mode) & " EQN BinX", scaleNoScaling, "%.0f"
                End If
            Next p_mode
            
            '''//Need to reset the flags "F_PassBinCut_1", "F_PassBinCut_2" from "printout_vddbin" due to new binx judgement.
            If Flag_Fuse_BinX(site) = True Then
                TheExec.sites(site).FlagState("F_PassBinCut_1") = logicFalse
                TheExec.sites(site).FlagState("F_PassBinCut_2") = logicTrue
            Else
                TheExec.sites(site).FlagState("F_PassBinCut_1") = logicTrue
                TheExec.sites(site).FlagState("F_PassBinCut_2") = logicFalse
            End If
        End If

        '''********************************************************************************************************************'''
        '''[Step6] If the fail stop flag is true, clear all efuse values for Vdd binning to avoid fusing wrong value.
        '''********************************************************************************************************************'''
        If TheExec.sites.item(site).FlagState(strGlb_Flag_Vddbinning_Fail_Stop) = logicTrue Then '''If the flag is true, we set the efuse flag to false and clear the VDD_RESULT.
            'Print_Error_Message (Error_Info, "VBT_LIB_VDD_Binning", "v","print: The Site " & site & " is failed, but it is not stop on fail. Please check!! Error!!!")
            
            If str_Efuse_write_ProductIdentifier <> "" Then
                '''For project with Efuse obj vbt code.
                Pass_Fail_Flag(site) = False
                local_efuseval(site) = 0 '''added for Efuse DSP func, 20200225.
            End If
   
            For p_mode = 0 To MaxPerformanceModeCount - 1
                VBIN_RESULT(p_mode).GRADE = 0
                VBIN_RESULT(p_mode).step_in_BinCut = -1
                VBIN_RESULT(p_mode).GRADEVDD = 0
            Next p_mode
        Else
            If Flag_VDD_Binning_Offline = False Then
                If CurrentPassBinCutNum(site) > 0 Then
                    If str_Efuse_write_ProductIdentifier <> "" Then
                        '''For project with Efuse obj vbt code.
                        Pass_Fail_Flag(site) = True
                        local_efuseval(site) = CurrentPassBinCutNum(site) - 1
                    End If
                Else
                    TheExec.Datalog.WriteComment "print: CurrentPassBinCutNum(" & site & ") = " & CurrentPassBinCutNum(site) & ", please check!!"
                    If str_Efuse_write_ProductIdentifier <> "" Then
                        '''For project with Efuse obj vbt code.
                        Pass_Fail_Flag(site) = False
                        local_efuseval(site) = 0
                    End If
                End If
            End If
        End If
        '''*****************************************************************************************'''
        '''ToDo: Maybe we can use this fail-stop flag to mask the failed DUT in Adjust_VddBinning...
        'End IF'''If TheExec.sites.Item(site).FlagState(strGlb_Flag_Vddbinning_Fail_Stop) = logicFalse
        '''*****************************************************************************************'''
    Next site
    
'''********************************************************************************************************'''
'''//Warning!!! This block is the special case for Crete A0 only, and it will be removed while Crete B0!!!
'''20210819: Modified to add the special case to skip fusing product_identfier and product voltages for testJob "wlft2" of T-Cre.
'''20210908: Modified to add the special case to fuse the product_identifier and product voltages for testJob "wlft2" of T-Cre.
    If LCase(TheExec.CurrentJob) = "wlft2" Then
        str_Efuse_write_ProductIdentifier = "Product_Identifier_Shadow"
    End If
'''********************************************************************************************************'''
    
    '''//Update PassBinNum of each site to Efuse "Product_Identifier" if the dedicated "Product_Identifier" for BinCut search exists.
    If str_Efuse_write_ProductIdentifier <> "" Then
        '''20230504: Decide correct bin number by EQN result
        For Each site In TheExec.sites
            If Flag_Fuse_BinX(site) = True Then
                local_efuseval(site) = 1
            End If
        Next site
        '''For project with Efuse obj vbt code.
        '''20210705: Modified to find Efuse Category for product_identifier by checking the current BinCut testjob.
        Set opbank = GetBdfBank("CFG")
        opbank.SetEfuse str_Efuse_write_ProductIdentifier, local_efuseval, Pass_Fail_Flag, , , , True
    End If
    
    '''//Print the info about adjust_vddBinning in DTR of STDF.
    For counter = 0 To nu
        If strOutput(counter) <> "" Then
            TheExec.Datalog.WriteComment strOutput(counter)
        End If
    Next counter
    
    '''20240220 summary which pmode cause power domain all pmode to BINX by site
    For counter = 0 To count_strBINXOfPowerDomain - 1
        TheExec.Datalog.WriteComment strBINXOfPowerDomain(counter)
    Next counter
    
    '''20240223 if more than 2 powerdomain up to BINX, binout by site
    For Each site In TheExec.sites
        Dim count_powerDomain_BINX As Long: count_powerDomain_BINX = 0
        For i = 0 To cntVddbinPin - 1
            If Bincut_Domain2BinX(i)(site) = True Then
                count_powerDomain_BINX = count_powerDomain_BINX + 1
            End If
        Next i
        If count_powerDomain_BINX >= 2 Then
            TheExec.sites.item(site).FlagState("F_two_BINX_PowerDomain") = logicTrue
        End If
    Next site
    
    
    
    '''********************************************************************************************************************'''
    '''[Optional] Print BinCut fail info to record the first changed binnum mode data, requested by C651 Si.
    '''********************************************************************************************************************'''
    '''20210629: Modified to check the EnableWord("Vddbin_PTE_Debug")=False to determine "Print Bincut Fail Info", as suggested by Chihome.
    If TheExec.enableWord("Vddbin_PTE_Debug") = False Then
        TheExec.Datalog.WriteComment "=============================================="
        TheExec.Datalog.WriteComment "======    " & "Print Bincut Fail Info" & "    ======"
        TheExec.Datalog.WriteComment "=============================================="
        For Each site In TheExec.sites
            If FirstChangeBinInfo.FirstChangeBinMode <> 999 Then
                '20240129 Bin1 BinX merge
                If Flag_BinX_Info_Parsed = True And CurrentPassBinCutNum(site) = 1 Then
                    TheExec.Datalog.WriteComment "XDMN," & site & "," & FirstChangeBinInfo.str_Pmode_Test(site)
                '''//The output format of FirstChangeBinInfo, ex: XDMN,<site>,<Pmode_Test>
                ElseIf CurrentPassBinCutNum(site) = 2 Then
                    '''20210910: Modified to revise the format of FirstChangeBinInfo, as requested by C651 Si and TSMC ZYLINI.
                    '''<org>
                        'TheExec.Datalog.WriteComment "site:" & site & ", Elevated_bincut_fail_info:" & "BinX_" & TestTypeName(FirstChangeBinInfo.FirstChangeBinType(site)) & "_" & VddBinName(FirstChangeBinInfo.FirstChangeBinMode(site))
                    '''<new>
                    TheExec.Datalog.WriteComment "XDMN," & site & "," & FirstChangeBinInfo.str_Pmode_Test(site)
                ElseIf CurrentPassBinCutNum(site) = 3 Then
                    '''20210910: Modified to revise the format of FirstChangeBinInfo, as requested by C651 Si and TSMC ZYLINI.
                    '''<org>
                        'TheExec.Datalog.WriteComment "site:" & site & ", Elevated_bincut_fail_info:" & "BinY_" & TestTypeName(FirstChangeBinInfo.FirstChangeBinType(site)) & "_" & VddBinName(FirstChangeBinInfo.FirstChangeBinMode(site))
                    '''<new>
                    TheExec.Datalog.WriteComment "XDMN," & site & "," & FirstChangeBinInfo.str_Pmode_Test(site)
                End If
            ElseIf TheExec.sites.item(site).FlagState("F_IDS_Binx") = logicTrue Then
                TheExec.Datalog.WriteComment "XDMN," & site & ", Elevated_bincut_fail_info:" & "BinX_IDS"
            ElseIf TheExec.sites.item(site).FlagState("F_IDS_BinY") = logicTrue Then
                TheExec.Datalog.WriteComment "XDMN," & site & ", Elevated_bincut_fail_info:" & "BinY_IDS"
            End If
        Next site
        TheExec.Datalog.WriteComment "=============================================="
    End If
    
    '''//Align the format of Judge_PF and Judge_PF_func in the datalog.
    TheExec.Datalog.WriteComment ""
    TheExec.Datalog.WriteComment ""
    TheExec.Datalog.WriteComment vbCrLf
Exit Function
errHandler: 'Add ErrHandler 2023/08/01
    Call Print_Error_Message(Error_Info, "VBT_LIB_VDD_Binning", "adjust_VddBinning") 'Add ErrHandler 2023/08/01
    If AbortTest Then Exit Function Else Resume Next 'Add ErrHandler 2023/08/01
End Function
'20230705: Check if doesn't have any site active then bin out. Request by T-Cai.
'20210908: Modified to merge vbt functions Non_Binning_Pwr_Setting_VT, HVCC_Set_VT, and PostBinCut_Voltage_Set_VT into the vbt function bincut_power_Setting_VT, as discussed with TSMC ZYLINI.
'20210908: Modified to add the argument "Optional enable_DynamicOffset As Boolean = False" to calculate BinCut payload voltage of the binning PowerDomain with DynamicOffset.
'20210903: Modified to merge properties of "Public Type Instance_Step_Control" into Public Type Instance_Info "Public Type Instance_Info".
'20210901: Modified to rename "stepcountMax As Long" as "maxStep As New SiteLong" for Public Type Instance_Info.
'20210901: Modified to rename "StepCount As Long" as "count_Step As New SiteLong" for Public Type Instance_Info.
'20210830: Modified to add the optional argument "Optional HarvestBinningFlag As String" for Harvest in BinCut, as requested by C651 Toby.
'20210824: Modified to rename the vbt function calculate_payload_voltage_for_BV as get_passBin_from_Step.
'20210824: Modified to move the vbt function Non_Binning_Pwr_Setting_VT from calculate_payload_voltage_for_BV to GradeSearch_VT.
'20210824: Modified to move the vbt function Calculate_Binning_CorePower_with_DynamicOffset from calculate_payload_voltage_for_BV to GradeSearch_VT.
'20210810: Modified to merge the vbt function Check_anySite_GradeFound into the vbt function Update_VBinResult_by_Step.
'20210603: Modified to move inst_info.Pattern_Pmode and inst_info.By_Mode from initialize_inst_info to GradeSearch_XXX_VT.
'20210530: Modified to update theExec.sites.Selected for MultiFSTP prior to the main flow of GradeSearch_VT.
'20210305: Modified to add the arguments "step_control As Instance_Step_Control" to the vbt function "StoreCaptureByStep".
'20210225: Modified to exit pattern-loop if all site fail and "inst_info.is_BV_Payload_Voltage_printed = True".
'20210225: Modified to move "Set_BinCut_Initial_by_ApplyLevelsTiming" prior to "decide_bincut_feature_for_stepsearch".
'20210223: Modified to move "Check_and_Decompose_PrePatt_FuncPat" prior to "decide_bincut_feature_for_stepsearch".
'20210207: Modified to move the vbt code of switching DCVS Valt from print_voltage_info_before_FuncPat to GradeSearch_XXX_VT.
'20210126: Modified to revise the vbt code for DevChar.
'20210125: Modified to move "select_DCVS_output_for_powerDomain" prior to "Set_PayloadVoltage_to_DCVS".
'20210122: Modified to check if FuncPat <> "" for print_voltage_info_before_FuncPat.
'20201217: Modified to use the vbt function decide_bincut_feature_for_stepsearch to decide if BinCut features are OK to be enabled for BinCut stepSearch.
'20201211: Modified to use the vbt function "initialize_control_flag_for_step_loop" to initialize control flags from "inst_info" and "step_control" at the beginning of each step in step-loop.
'20201211: Modified to use the vbt function "update_sort_result" to do Judge_PF for binSearch and Judge_PF_func for functional test.
'20201210: Modified to use run_patt_from_FuncPat_for_BinCut for running the pattern decomposed from FuncPat.
'20201210: Modified to rename the vbt function "calculate_payload_voltage_for_binning_CorePower" as "calculate_payload_voltage_for_BV".
'20201210: Modified to add the vbt functions "Get_PassBinNum_by_Step" and "Non_Binning_Pwr_Setting_VT" into the vbt function calculate_payload_voltage_for_binning_CorePower.
'20201210: Modified to use the arguments "inst_info As Instance_Info" and "step_control As Instance_Step_Control" for judge_PF, update_patt_result_for_COFInstance, Update_PassBinCut_for_GradeNotFound, Decide_NextStep_for_GradeSearch, update_control_flag_for_patt_loop, Check_anySite_GradeFound, and Update_VBinResult_by_Step.
'20201207: Modified to use "Dim step_control As Instance_Step_Control".
'20201204: Modified to add the argument "IndexLevelPerSite As SiteLong" for the vbt function Calculate_Binning_CorePower_with_DynamicOffset.
'20201204: Modified to initialize "inst_info.PrePattPass", "inst_info.FuncPatPass", and "inst_info.sitePatPass" in the vbt function initialize_inst_info.
'20201204: Modified to initialize "inst_info.IndexLevelPerSite = -1" in the vbt function initialize_inst_info.
'20201203: Modified to add the argument "enable_CMEM_collection As Boolean" for check_flag_to_enable_CMEM_collection.
'20201203: Modified to revise the vbt code for the undefined testJobs.
'20201202: Modified to add the argument "enable_CMEM_Collection as Boolean" for resize_CMEM_Data_by_pattern_number.
'20201201: Modified to use resize_CMEM_Data_by_pattern_number for CMEM.
'20201201: Modified to update CaptureSize, failpins, and PrintSize for CMEM.
'20201125: As suggestion from Chihome, modified to set TheHdw.Digital.CMEM.CentralFields for initializing CMEM in GradeSearch_XXX_VT.
'20201125: As suggestion from Chihome, modified to clear capture Memory (CMEM) after PostTestIPF.
'20201123: Modified to align the format of Judge_PF and Judge_PF_func in the datalog.
'20201118: Modified to use "TheHdw.Digital.Patgen.PatternBurstPassedPerSite" to get siteResult of pattern pass/fail.
'20201111: Modified to use "check_flag_to_enable_CMEM_collection".
'20201111: Modified to use "inst_info.voltage_SelsrmBitCalc".
'20201111: Modified to replace the vbt function name "print_bincut_power" with "print_bincut_voltage".
'20201110: Modified to check if FuncPat <> "".
'20201103: Modified to move "Dim ids_current As New SiteDouble", "IDS_current_fail As New SiteLong", and "Dim IDS_current_Min As Double" into "Public Type Instance_Info".
'20201103: Modified to move "Dim stepcount As Long" and "Dim stepcountMax As Long" into "Public Type Instance_Info".
'20201029: Modified to use inst_info.previousDcvsOutput and inst_info.currentDcvsOutput.
'20201029: Modified to use inst_info.is_BV_Safe_Voltage_printed and inst_info.is_BV_Payload_Voltage_printed.
'20201027: Modified to use "Public Type Instance_Info".
'20201026: Modified to switch DCVS to Vmain prior to PrePatt for TD pattern burst by C651 Toby.
'20201026: Modified to revise the vbt code for TD pattern burst proposed by C651 Toby.
'20201021: Modified to check DecomposedPat to decide Enable_CMEM_Collection on/off for PatternBurst.
'20201020: Modified to add the variables "COFInstance" and "PerEqnLog" for COFInstance.
'20201016: Modfied to save EQN-based BinCut Payload voltage of binning P_mode. Requested by C651 Si Li.
'20201016: Modified to use "decide_flag_for_COFInstance".
'20201015: Modified to use "update_patt_result_for_COFInstance" to record pattern pass/fail per site for "COFInstance".
'20201015: Modified to check "Enable_COFInstance".
'20201015: Modified to check the flag "Flag_Vddbin_COF_Instance".
'20201015: Modified to save result of pattern Pass/Fail in "sitePatPass".
'20201012: Modified to change the arguments of the vbt function "check_patt_Pass_Fail".
'20201008: Modified to replace "PrintedBVinDatalog" with "is_BV_Payload_Voltage_printed".
'20201006: Modified to add the condition "PrintedBVinDatalog = True" for pattern loop control, requested by CheckScript.
'20201006: Modified to merge the branches of cp1 and non-cp1.
'20200925: Modified to merge "run_patt_only_VT" of non-cp1 into "GradeSearch_VT" of cp1.
'20200924: Modified to merge the branches of "Calculate_Selsrm_DSSC_For_BinCut".
'20200924: Modified to move "select_DCVS_output_for_powerDomain" from GradeSearch_VT to "run_prepatt_decompose_VT".
'20200923: Modified to use "update_control_flag_for_patt_loop" to update pdate the status of "AllSiteFailPatt" and "All_Patt_Pass".
'20200923: Modified to remove "clear_after_patt".
'20200923: Modified to use "run_patt_offline_simulation" for offline simulation.
'20200923: Modified to call "check_patt_Pass_Fail".
'20200922: Modified the branch to simulate offline random Pass/Fail.
'20200922: Modified to update the status of "AllSiteFailPatt".
'20200922: Modified to align the branches of running FuncPat for online and offline tests.
'20200922: Modified to move the vbt block of "prepare_DCVS_Output_for_RailSwitch" from the branch "If Flag_VDD_Binning_Offline = False Then" to pattern-loop.
'20200922: Modified to remove the redundant vbt code of "KeepAliveFlag".
'20200921: Discussed "RTOS_bootup_relay" / "KeepAliveFlag" / "spi_ttr_flag" with SWLINZA and PCLINZG. We decided to remove these SPI/RTOS branches because RTOS didn't use pattern test since Cebu/Sicily/Tonga/JC-Chop/Ellis/Bora.
'20200918: Modified to add the argument "result_mode" for the vbt function "Check_and_Decompose_PrePatt_FuncPat".
'20200918: Modified to use "prepare_DCVS_Output_for_RailSwitch".
'20200918: Modified to use "print_voltage_info_before_FuncPat".
'20200909: Modified to replace "Flag_Enable_CMEM_Collection" with "Enable_CMEM_Collection".
'20200821: Modified to add "Dim str_Selsrm_DSSC_Bit(MaxSiteCount-1) As String".
'20200809: Modified to check DCVS output and Payload pattern.
'20200803: Modified to use "call Non_Binning_Pwr_Setting_VT".
'20200802: Modified to check patType init or payload, revised by Leon.
'20200730: Modified to add the EnableWord "VDDBinning_Offline_AllPattPass" for Offline simulation with all patterns pass.
'20200727: Modified to follow the new rule of SELSRM bit calculation proposed by C651 Toby.
'20200713: Modified to remove the argument "IndexLevelPerSite As SiteLong" from Non_Binning_Pwr_Setting_VT by using the function Get_PassBinNum_by_Step.
'20200711: Modified to use the siteDouble array "BinCut_Payload_Voltage" to store BinCut payload voltages.
'20200630: Modified to remove the unused "thehdw.Utility.Pins("k02,k04").State = tlUtilBitOff".
'20200615: Modified to get dynamic_offset type from the argument "offsetTestTypeIdx As Integer" for judge_PF.
'20200609: Modified to use Check_alarmFail_before_BinCut_Initial.
'20200525: Modified to use "Get_PassBinNum_by_Step".
'20200525: Modified to use "Get_Binning_CorePower_PayloadVoltage_by_Step" to get PayloadVoltage of the binning corePower by Step.
'20200520: Modified to use Check_Pattern_NoBurst_NoDecompose to show the errorLogMessage if "burst=no" and "Decompose_Pattern=false".
'20200520: Modified to use "Check_and_Decompose_PrePatt_FuncPat" to check and decompose patsets PrePatt and FuncPat, and find SELSRAM DSSC pattern for DSSC digSrc.
'20200511: Modified to use "Check_anySite_GradeFound".
'20200511: Modified to use "Update_PassBinCut_for_GradeNotFound".
'20200511: Modified to use Decide_NextStep_for_GradeSearch for deciding Next Search Step.
'20200508: Modified to use "Update_VBinResult_by_Step" to update VBIN_RESULT(p_mode) by BinCut Step.
'20200501: Modified to check if p_mode should be interpolated and test-skipped while testJob CP1 (Interpolation is only for CP1).
'20200424: Modified to use "Set_BinCut_Initial_by_ApplyLevelsTiming" to set BinCut initial voltage by ApplyLevelsTiming.
'20200423: Modified to replace "BinCut(p_mode, bincutNum(site)).tested = True" with "VBIN_RESULT(p_mode).tested=True".
'20200324: Modified to skip ApplyLevelsTiming when current instance has the same level/timing as previous instance for project with rail-switch.
'20200320: Modified to check instance contexts of current instance and previous instance.
'20200217: Modified to check if no vbump before running the payload pattern.
'20200214: Modified to print dynamic_offset by the vbt function "print_bincut_power".
'20200207: Modified to replace set_core_power_main and set_core_power_alt with select_DCVS_output_for_powerDomain.
'20200203: Modified to use the function "print_bincut_power".
'20200130: Modified to call Calculate_Selsrm_DSSC_For_BinCut for SELSRM DSSC bits calculation.
'20200130: Modified to call Get_Pmode_Addimode_Testtype_fromInstance to get pmode/addi_mode/testtype.
'20190106: Modified to add "TheHdw.Alarms.Check".
'20191226: Modified to clear and renew CMEM for opensocket CMEM overflow issues.
'20191224: Modified to calculate binning corePower with dynamic offset.
'20191127: Modified to use the revised InitVddBinTable.
'20191119: Modified to use the pattern keyword for SELSRM_Mapping_Table.
'20191113: If SelsrmMappingTable is parsed, check the pattern names with keyword of the matched SELSRAM blocktype.
'20191021: Modified to update array size for Step_CMEM_Data.
'20191009: Modified to print payload voltages for opensocket and offline simulation with rail-switch.
'20190627: Modified to use the global variable "pinGroup_BinCut" for BinCut powerDomains.
'20190618: Modified to deliver "Special_Voltage_setup" and "tcd_flow" to run_prepatt_decompose_VT.
'20190617: Modified to use siteDouble "'CorePowerStored()" to save/restore voltages for BinCut powerDomains.
'20190611: Modified for set voltages to "DC Specs" (by DcSpecsCategoryForInitPat) for initial and "init patterns" before ApplyLevelsTiming.
'20190606: Modified to add the argument "DcSpecsCategoryForInitPat as string" for Init patterns with the new test setting DC Specs.
'20190319: Modified to add "Flag_Enable_Rail_Switch" to turn VRS Rail Switch on/off.
'20190313: Modified to add FIRSTPASSBINCUT(p_mode) to store the first passbinnumber of P_mode GradeSearch (for Get_SRAM_Vth).
'20190311: Modified to add the flag "Flag_Enable_Rail_Switch" to switch the VRS rail switch utility on/off.
'20190304: Modified to call the new funciton Calculate_Extra_Voltage_for_PowerRail, especially for "MS001 Evaluate Bin + 15%".
'20181115: Modified for DSSC TD patt_group (init+payload1+init+payload).
'20181113: Modified for TD payload pattern (print_alt_power_payload).
'20180709: Modified for BinCut testjob mapping.
Public Function GradeSearch_VT(FuncPat As Pattern, performance_mode As String, result_mode As tlResultMode, DecomposePatt As String, _
                               FuncTestOnly As Boolean, PrePatt As Pattern, Optional SpiCounterValue As Integer, Optional Validating_ As Boolean, _
                               Optional DcSpecsCategoryForInitPat As String = vbNullString, Optional CaptureSize As Long, Optional failpins As String, Optional CollectOnEachStep As Boolean, _
                               Optional HarvestBinningFlag As String = vbNullString, _
                               Optional EnableCoreHarvest As Boolean = False, Optional EnableCoreMask As Boolean = true, _
                               Optional StartBin1E1 As Boolean = False, Optional SingleFSTP_Offset As Long = 0, Optional Harvest_With_BIN1E1 As Boolean = False, Optional Instance_Offset As Long = 0, Optional UserFunction As String = vbNullString)
On Error GoTo errHandler 'Add ErrHandler 2023/08/01
    Dim site As Variant
    Dim inst_info As Instance_Info
    Dim indexPatt As Long
    '''for binning p_mode
    Dim passBinFromStep As New SiteLong
    '''for testNumber alignment
    Dim Org_Test_Number As Long
    Dim instSSNinfo As Inst_SSN
    
    If gl_EnableCurrentProfile Or gl_EnableVoltageProfile Then
        ProfileCollect_Start glb_TestInstance
    End If
    
    glb_TestInstance = vbNullString
    glb_TestInstance = UCase(TheExec.DataManager.instanceName)
    
    '''//update theExec.sites.Selected at the begin of each MultiFSTP GradeSearch instance.
    '''20210530: Modified to update theExec.sites.Selected for MultiFSTP prior to the main flow of GradeSearch_VT.
    '''ToDo: Please check if EnableWord("Multifstp_Datacollection") exists in the flow table!!!
    If TheExec.Flow.enableWord("Multifstp_Datacollection") = True Then
        TheExec.sites.Selected = gb_siteMask_current
    End If

    If Validating_ Then
        If FuncPat.value <> "" Then Call PrLoadPattern(FuncPat.value)
        If PrePatt.value <> "" Then Call PrLoadPattern(PrePatt.value)
        Exit Function ''' Exit after validation
    End If
    
    If PrePatt <> "" Then
        Shmoo_Pattern = FuncPat.value & "," & PrePatt.value
    Else
        Shmoo_Pattern = FuncPat.value
    End If
    
    '''//Initialize inst_info.
    '''//Get p_mode, addi_mode, jobIdx, testtype, and offsettestype from test instance and performance_mode.
    '''//The flag "inst_info.is_BinSearch" is True if testCondition of the binning PowerDomain from the table "Non_Binning_Rail" has the keyword "*Evaluate*Bin*".
    Call initialize_inst_info(inst_info, performance_mode)
    inst_info.selsrm_DigSrc_Pin = "JTAG_TDI"
    inst_info.selsrm_DigSrc_SignalName = "DigSrcSignal"
    '''For Harvest MultiFSTP.
    inst_info.Harvest_Core_DigSrc_Pin = "JTAG_TDI"
    inst_info.Harvest_Core_DigSrc_SignalName = "Harvest_Core_DigSrcSignal"
    '20240124: New MFSTP defined by Alan
    'Dim instInfoUserFunction As UserDefinedFunctionType
    Call CheckInstForUserFunction(UserFunction, inst_info.digSrcLabel, inst_info.digSrcPatterns, CStr(FuncPat))
    
    If inst_info.digSrcLabel(0) <> "" Then
        inst_info.MultiFSTP_Enable = True
    Else
        inst_info.MultiFSTP_Enable = False
    End If
    '''//C651 Toby said that HarvestBinningFlag is for BinCur search only, 20210831.
    '''20210830: Modified to add the optional argument "Optional HarvestBinningFlag As String" for Harvest in BinCut, as requested by C651 Toby.
    inst_info.HarvestBinningFlag = HarvestBinningFlag
    
    '''20220125,HarvPinGrp variable set
    inst_info.enable_HarvPinGrp = EnableCoreHarvest
    '''20240604: Added independant boolean for saving pin/core mask
    inst_info.enable_CoreMask = EnableCoreMask
    
    inst_info.SingleFSTP_Offset = SingleFSTP_Offset '20240226 Temp SingleFSTP solution by add a argument of GradeSearch_VT
    inst_info.Harvest_With_BIN1E1 = Harvest_With_BIN1E1
    
    inst_info.Instance_Offset = Instance_Offset
    
    '''20220725,Turn off HarvPinGrp when mapping table does not exist
    '''20240726 add pin mask or core mask to check
    If (inst_info.enable_HarvPinGrp = True Or inst_info.enable_CoreMask = True) And Flag_HarvPinFlag_Mapping_Table_Parsed = False Then
        Call Print_Error_Message(Error_Info, "VBT_LIB_VDD_Binning", "GradeSearch_VT", "TestInst Argument 'HarvPinGrp_Enable' or 'enable_CoreMask' = True, but can't find HarvPinGrp Table! Error!!")
    End If
    '''20220125,SpecialSearch variable set
    If TheExec.Flow.enableWord("VddBinning_Pre_E1") = True Then
        inst_info.StartBin1E1_Flag = True
    Else
        inst_info.StartBin1E1_Flag = StartBin1E1
    End If
    
    ''20221223: For print voltage output(Vmain/Valt) in debug print mode
    Flag_isPrinted_BincutVoltageOutput = False
    
    '''//Check if DevChar Precondition is tested.
    If inst_info.is_DevChar_Running = True And inst_info.get_DevChar_Precondition = False Then
        Call Set_BinCut_Initial_by_ApplyLevelsTiming(inst_info)
        Exit Function
    End If
    
    '''//Set the excluded performance mode if the device is bin2 die and the performance mode do not exist in bin2 table.
    SkipTestBin2Site inst_info.p_mode, inst_info.Active_site_count
    
    If inst_info.Active_site_count = 0 Then
        RestoreSkipTestBin2Site inst_info.p_mode '''For the performance mode that does not exist in the bincut table
        Exit Function
    End If
    
    '''********************************************************************************************************************'''
    '''(1) For TestNumber align, the code is assembled with (2)
    '''********************************************************************************************************************'''
    For Each site In TheExec.sites.Active
        Org_Test_Number = TheExec.sites(site).TestNumber
        Exit For
    Next site
    
    '''//Check if alarmFail was triggered prior to BinCut initial(applyLevelsTiming).
    Check_alarmFail_before_BinCut_Initial inst_info.inst_name
    alarmFail = False
    
    '''//According to "inst_info.is_BinSearch", decide inst_info.maxStep for step-loop and find start_voltage(by start_Step in Dynamic_IDS_Zone of the binning p_mode).
    Call decide_binSearch_and_start_voltage(inst_info, FuncTestOnly)
    
    '''//If BinCut testJob is not defined for GradeSearch_XXX_VT, exit the vbt function.
    If inst_info.maxStep = -1 Then
        Call Print_Error_Message(Error_Info, "VBT_LIB_VDD_Binning", "GradeSearch_VT", "It can't get the correct maxStep for step-loop in GradeSearch_VT. Error!!!")
        'TheExec.ErrorLogMessage "It can't get the correct maxStep for step-loop in GradeSearch_VT. Error!!!"    ''20230807: Removed by sanity check
        Exit Function
    End If
    
        '' Check if doesn't have any site active then bin out. Request by T-Cai. 20230705
        If TheExec.sites.Active.Count = 0 Then
            Call Print_Error_Message(Error_Info, "VBT_LIB_VDD_Binning", "GradeSearch_VT", "Binout because we don't have any active site Warring!!!")
        Exit Function
    End If

    '''//Set initial voltages from category "Bincut_X_X_X" in DC_Specs sheet by ApplyLevelsTiming.
    '''Print the initial voltages, and applies them to DCVS Vmain and Valt by ApplyLevelsTiming (DCVS voltage source will be switched to Vmain).
    Call Set_BinCut_Initial_by_ApplyLevelsTiming(inst_info)

    '''//Decompose PrePatt and FuncPat to check if any DSSC digsrc pattern of SELSRM(defined in SELSRM_Mapping_Table) exists in the pattern sets.
    Call Check_and_Decompose_PrePatt_FuncPat(inst_info, result_mode, DecomposePatt, PrePatt.value, FuncPat.value)
    
    If Flag_HarvPinFlag_Mapping_Table_Parsed = True And ssnPatternsDict.Count > 0 Then
        Dim Decomposed_Pat_SSN() As String
        Dim patCnt As Long
        Decomposed_Pat_SSN = TheExec.DataManager.Raw.GetPatternsInSet(FuncPat.value, patCnt)
        Call CheckInstForSSN_BV(Decomposed_Pat_SSN, instSSNinfo, SSNMapping, inst_info.enable_HarvPinGrp, inst_info.enable_CoreMask)
        inst_info.BeSSNTest = instSSNinfo.bSSNTest '240703
    End If
    
    '''//Decide if BinCut features are OK to be enabled for BinCut stepSearch, ex: CMEM_collection, resize inst_info.BC_CMEM_StoreData, and COFInstance.
    Call decide_bincut_feature_for_stepsearch(inst_info, inst_info.count_FuncPat_decomposed, CaptureSize, failpins)
    
'**********************************************
'&& Search Grade Start
'**********************************************
    For inst_info.count_Step = 0 To inst_info.maxStep '''start Vdd binning search, use the full EQN count to loop.
        '''******************************************************************************************'''
        '''//Set DCVS voltage output to Vmain prior to PrePatt and FuncPat.
        '''******************************************************************************************'''
        select_DCVS_output_for_powerDomain tlDCVSVoltageMain
        inst_info.currentDcvsOutput = tlDCVSVoltageMain
        inst_info.previousDcvsOutput = tlDCVSVoltageMain        ''20230111: Added for SyncUp Vmain, to avoid SyncUp Vmain for every step IN01
        
        '''//Initialize control flags from "inst_info" and "step_control" at the beginning of each step in step-loop.
        '''Initialize flags of pattern pass/fail, BV Safe/Payload Voltage printed, and grade_found.
        Call initialize_control_flag_for_step_loop(inst_info)
        
        '''//Initialize array of BinCut_Init_Voltage and BinCut_Payload_Voltage before BinCut payload voltages calculation.
        Init_BinCut_Voltage_Array
        
        '''//Get passBin by the current step in Dynamic_IDS_Zone of the binning performance mode.
        '''//Update the current step in Dynamic_IDS_Zone if the perfromance mode is interpolated by Interpolation(ReGenerate_IDS_ZONE_Voltage_Per_Site_ver2).
        '''20210824: Modified to rename the vbt function calculate_payload_voltage_for_BV as get_passBin_from_Step.
        Call get_passBin_from_Step(inst_info, passBinFromStep)
        
        '''//Calculate BinCut payload voltages of BinCut CorePower and OtherRail.
        '''//It also calculates BinCut payload voltages with Dynamic_Offset for the binning PowerDomain.
        '''20210908: Modified to merge vbt functions Non_Binning_Pwr_Setting_VT, HVCC_Set_VT, and PostBinCut_Voltage_Set_VT into the vbt function bincut_power_Setting_VT, as discussed with TSMC ZYLINI.
        '''20210908: Modified to add the argument "Optional enable_DynamicOffset As Boolean = False" to calculate BinCut payload voltage of the binning PowerDomain with DynamicOffset.
        Call bincut_power_Setting_VT(inst_info, passBinFromStep, BinCut_Payload_Voltage, True)
        
        '''//If DSSC digsrc pattern of SELSRM exists in the pattern sets, calculate DSSC bits sequence by Selsrm Logic power/SRAMthresh, then prepare DSSC digsrc signal setups.
        Call Calculate_Selsrm_DSSC_For_BinCut(inst_info, passBinFromStep)
        
        '''//Set Payload voltages to DCVS. For projects with Rail-Switch, BinCut payload voltage values are applied to DCVS Valt.
        '''BinCut_Payload_Voltage is the siteDouble array for storing BinCut payload voltage values calculated from Non_Binning_Pwr_Setting_VT.
        Set_PayloadVoltage_to_DCVS Flag_Enable_Rail_Switch, pinGroup_BinCut, BinCut_Payload_Voltage
        TheHdw.Wait 0.001
        
        '''******************************************************************************************'''
        '''//Print BinCut init voltages for PrePatt, then run PrePatt(init pattern).
        '''******************************************************************************************'''
        Call run_prepatt_decompose_VT(inst_info, inst_info.PrePatt, inst_info.ary_PrePatt_decomposed, inst_info.count_PrePatt_decomposed, inst_info.PrePattPass, DcSpecsCategoryForInitPat)
        
        '''//Update siteResult of PrePatt.
        If inst_info.is_DevChar_Running = False Then '''for DevChar.
            Call update_Pattern_result_to_PattPass(inst_info.PrePattPass, inst_info.funcPatPass)
        End If
        
        '''//Clear capture Memory(CMEM) and resize array of CMEM_data if inst_info.enable_CMEM_Collection is enabled.
        Call resize_CMEM_Data_by_pattern_number(inst_info.enable_CMEM_Collection, inst_info.count_FuncPat_decomposed, inst_info.Step_CMEM_Data)

        '''//Check if FuncPat is empty...
        If FuncPat <> "" Then
            '''********************************************************************************************************************'''
            '''//For Mbist instances in project with rail-switch, set DCVS voltage output to Valt prior to FuncPat.
            '''********************************************************************************************************************'''
            '''20210207: Modified to move the vbt code of switching DCVS Valt from print_voltage_info_before_FuncPat to GradeSearch_XXX_VT.
            If inst_info.Test_Type = testType.Mbist Then '''ex: "*cpu*bist*", "*gfx*bist*", "*gpu*bist*", "*soc*bist*".
                '''====================================================================================================='''
                '''C651 didn't implement Vbump op-code in MBIST init pattern for project with rail-switch.
                '''So that we have to switch DCVS to Valt by VBT code here before running FuncPat for Mbist instances.
                '''====================================================================================================='''
                If Flag_Enable_Rail_Switch Then '''For projects with Rail Switch, BinCut payload voltages are applied to DCVS Valt.
                    select_DCVS_output_for_powerDomain tlDCVSVoltageAlt
                    inst_info.currentDcvsOutput = tlDCVSVoltageAlt
                Else '''For conventional projects without Rail Switch, BinCut payload voltages(BV) are applied to DCVS Vmain.
                    select_DCVS_output_for_powerDomain tlDCVSVoltageMain
                    inst_info.currentDcvsOutput = tlDCVSVoltageMain
                End If
            End If
        
            '''//Print BinCut voltage before running FuncPat.
            Call print_voltage_info_before_FuncPat(inst_info)
'**********************************************
'@@FuncPat pattern-loop Start
'**********************************************
            For indexPatt = 0 To inst_info.count_FuncPat_decomposed - 1
                '''====================================================================================================='''
                '''//Run pattern decomposed from FuncPatt patset, and get siteResult of pattern pass/fail.
                '''Sync up DCVS output and print BinCut payload voltage for projects with Rail Switch for TD instance.
                '''Run the pattern decomposed from FuncPat, and get siteResult of pattern pass/fail.
                '''====================================================================================================='''
                Call run_patt_from_FuncPat_for_BinCut(inst_info, indexPatt, inst_info.ary_FuncPat_decomposed(indexPatt), inst_info.funcPatPass, instSSNinfo, inst_info.idxBlock_Selsrm_FuncPat, CaptureSize, failpins)

                If inst_info.is_BinSearch = True Then
                    '''//Update the status of "AllSiteFailPatt" and "All_Patt_Pass".
                    Call update_control_flag_for_patt_loop(inst_info, inst_info.funcPatPass)
    
                    '''==============================================================================================================================
                    ''' Site 0 fail, Site 1 fail, site 2 fail =>  AllSiteFailPatt = 2^0 or 2^1 or 2^2 = 7.
                    ''' if there are 3 sites active, the All_Site_Mask =7, we do not need to run all patterns and exit the loop to save test time.
                    '''==============================================================================================================================
                    '''20210225: Modified to exit pattern-loop if all site fail and "inst_info.is_BV_Payload_Voltage_printed = True".
                    If inst_info.enable_COFInstance = False And inst_info.is_BV_Payload_Voltage_printed = True And inst_info.enable_HarvPinGrp = False Then
                        If inst_info.AllSiteFailPatt = inst_info.All_Site_Mask Then Exit For '''if all site fail exit for loop to save time
                    End If
                End If
            Next indexPatt
'**********************************************
'@@FuncPat pattern-loop End
'**********************************************
            DebugPrintFunc FuncPat.value
        End If '''If FuncPat <> ""

        '''//If CMEM_collection is enabled, collect the data of the failed pattern for current step for the failed pattern from "inst_info.Step_CMEM_Data" to "inst_info.BC_CMEM_StoreData".
        If inst_info.enable_CMEM_Collection = True And inst_info.AllSiteFailPatt > 0 Then
            Call StoreCaptureByStep(inst_info, inst_info.Step_CMEM_Data, inst_info.BC_CMEM_StoreData)
            If CollectOnEachStep = True Then Call PostTestIPF(inst_info.performance_mode, failpins, inst_info.PrintSize, inst_info.Step_CMEM_Data, inst_info.str_patternPins)
        End If

        If inst_info.is_BinSearch = True Then
            '''//According to the current BinCut step in Dynamic_IDS_zone of p_mode, update PassBin and BinCut voltage(Grade) to VBIN_Result of p_mode for the pass DUT.
            '''//Check if any site has found BinCut pass Grade (based on BinCut step).
            Call Update_VBinResult_by_Step(inst_info)

            '''================================================================================================================================================='''
            ''' If Grade_Found_Mask = All_Site_Mask (All sites had found BinCut Grade) or On_StopVoltage_Mask = Grade_Not_Found_Mask => exit the step-loop.
            ''' ex: The sites didn't find BinCut Grade, but the sites had reached the stopVoltage (step_Stop in Dynamic_IDS_zone), no chance to find the grade!!!
            '''================================================================================================================================================='''
            If inst_info.Grade_Found_Mask = inst_info.All_Site_Mask Or inst_info.On_StopVoltage_Mask = inst_info.Grade_Not_Found_Mask Then
                Exit For '''Exit for "Next StepCount"
            End If

            '''//Decide next step in DYNAMIC_IDS_ZONE of p_mode for GradeSearch.
            Call Decide_NextStep_for_GradeSearch(inst_info)
        End If
        
        If TheExec.enableWord("Vddbin_PTE_Debug") = True Then
            Exit For
        End If
    Next inst_info.count_Step
'**********************************************
'&& Search Grade End
'**********************************************
    '''//Check if running FuncPat with "burst=no" and "Decompose_Pattern=false".
    Call Check_Pattern_NoBurst_NoDecompose(inst_info.FuncPat, inst_info.count_FuncPat_decomposed, inst_info.enable_DecomposePatt)

    '''//Align testNumber and do judge_PF for binSearch; judge_PF_func for functional test.
    Call update_sort_result(inst_info, inst_info.funcPatPass, Org_Test_Number, failpins, CollectOnEachStep)
    
    '''//Align the format of Judge_PF and Judge_PF_func in the datalog.
    TheExec.Datalog.WriteComment ""
    TheExec.Datalog.WriteComment ""
    TheExec.Datalog.WriteComment vbCrLf
    If gl_EnableCurrentProfile Or gl_EnableVoltageProfile Then
        ProfileCollect_End glb_TestInstance
    End If
Exit Function 'Add ErrHandler 2023/08/01
errHandler: 'Add ErrHandler 2023/08/01
    Call Print_Error_Message(Error_Info, "VBT_LIB_VDD_Binning", "GradeSearch_VT") 'Add ErrHandler 2023/08/01
    If AbortTest Then Exit Function Else Resume Next 'Add ErrHandler 2023/08/01
End Function

'20210726: Modified to add the argument "bincutNum As Long" to the vbt function Adjust_Multi_PassBinCut_Per_Site.
'20210721: Modified to check if Efuse category "power_binning" is for the current BinCut testJob.
'20210713: Modified to check if "power_binning" is for the current BinCut testJob.
'20210708: Modified to check if Efuse category "power_binning" for the current BinCut testJob has the matched programming stage in Efuse_BitDef_Table.
'20210708: Modified to check if Efuse category "Product_Identifier" for the current BinCut testJob has the matched programming stage in Efuse_BitDef_Table.
'20210707: Modified to check if it's OK to write Efuse category "power_binning".
'20210705: Modified to check "power_binning" in dict_EfuseCategory2BinCutTestJob.
'20210628: Modified to use gb_str_EfuseCategory_for_powerbinning.
'20210610: Modified to check if it sent the incorrect value "power_binning" to Efuse. If that, bin out the failed DUT.
'20210517: Modified to remove If fusePwrbin(site) <> "", requedt by Efuse.'20210506: Modified to update Efuse product_identifier.
'20210420: Modified to remove the unused vbt code of GradeSearchMethod. Once if C651 provided any new definition of GradeSearchMethod, we will revise the vbt code.
'20210310: Modified to use PassBinCut_ary(LBound(PassBinCut_ary)) To PassBinCut_ary(UBound(PassBinCut_ary)) for PassBin loop.
'20210302: Modified to add the vbt code for F_IDS_Binx and F_IDS_Biny because PrintOut_VddBinning and Adjust_Binning will be adjusted prior to Power_Binning in Flow_VddBinning.
'20201216: Modified to print "C" value from powerbinning tables for each site, requested by PCLINZG and C651 Toby.
'20201211: Modified to use powerDomain from Binned_mode, ex: "LOW".
'20201211: Modified to use powerDomain from column "IDS" for Binned_Mode with SRAM 6-digit mode, ex: "MPS001".
'20201125: Modified to print Vbin for Other_Mode, requested by PCLINZG.
'20201111: Modified to merge branches for printing IDS and Power with TName by testLimit for Binned_Modeand Other_Mode.
'20201111: Modified to set Binned_Mode-loop and Other_Mode-loop start from 0.
'20201110: Modified to revised the vbt code for parsing PowerBinning sheets with the new format.
'20200708: Modified to use "CurrentPassBinCutNum(site) <= Total_Bincut_Num".
'20200707: Modified the branch to enable failFlag.
'20200629: Modified to use the flag "Flag_Vddbinning_Power_Binning_Fail_Stop".
'20200325: Modified to use the flag "Flag_PowerBinningTable_Parsed".
'20200324: Modified to merge "Power_Binning_Calculation_Harvest" into this function.
'20200317: Modified for SearchByPmode.
'20200305: Modified to reduce the flow complexity (requested by C651 Toby).
'20200226: Modified to check if next spec exists with same passbin and harvest_bin.
'20200226: Modified to check Harvest_bin for each site.
'20200221: Modified to use the revised data structures PWRBIN_BIN_Type, PWRBIN_CONTITION_Type, and PWRBIN_SPEC_Type.
'20200218: Modified for the revised PowerBinning for Harvest.
'20200106: Modified to remove the ErrorLogMessage.
'20191127: Modified for the revised InitVddBinTable.
'20190826: Modified to add RunPwrBinningFlag[0] for reset. RunPwrBinningFlag[1~3] for stored selected.
'20190813: Modified the parsing method for power binning tables.
'20190812: Modified to control currentPassNum.
'20190722: Modified to printout the scale and the unit for BinCut voltages and IDS values.
'20190711: Modified to replace site-loop with sheet loop
'20190615: Modified to align the unit of IDS with IDS datatype
'20190528: Modified to get IDS values from the new IDS datatype
'20190507: Modified to add "Cdec" to avoid double format accuracy issues.
'20190308: Modified for New Power Binning Sheet.
'20181004: Modified for SRAM in power binning table, sum of all SRAM IDS and calculate the total SRAM power at the end of calculation by Joseph.
'20180918: Modified for Multiple Ratio Name searching by Oscar
Public Function Power_Binning_Calculation(Optional remove_printing_power As Boolean = False)
On Error GoTo errHandler 'Add ErrHandler 2023/08/01
    Dim site As Variant
    Dim i As Integer
    Dim j As Integer
    Dim k As Integer
    Dim p_mode As Integer
    Dim idx_Condition As Integer
    Dim idx_Spec As Integer
    Dim idx_Sheet As Integer
    Dim idx_Binned_Mode As Integer
    Dim idx_Other_Mode As Integer
    Dim performance_mode As String
    Dim powerDomain As String
    Dim P_total As New SiteDouble
    Dim BinNumber As Long
    Dim sheetName As String
    Dim tName_Temp As String
    Dim voltage_Temp As New SiteDouble
    Dim ids_Temp As New SiteDouble
    Dim power_Temp As New SiteDouble
    Dim power_total_binned_mode As New SiteDouble
    Dim power_total_other_mode As New SiteDouble
    '''spec-loop control
    Dim RunPwrBinningFlag() As New SiteBoolean
    Dim PreRunPwrBinningFlag As New SiteBoolean
    ReDim RunPwrBinningFlag(Total_Bincut_Num)
    '''variables
    Dim harvestBit() As New SiteLong
    Dim flagPwrbinSheet() As New SiteBoolean
    Dim anySiteSelected As Boolean
    Dim foundSpec As New SiteBoolean
    Dim power_AllModesWithOffset() As New SiteDouble
    Dim isSheetCalculatedForSite() As New SiteBoolean
    Dim specName() As String 'Dim specName(MaxSiteCount - 1) As String 20230526
    Dim fusePwrbin() As String 'Dim fusePwrbin(MaxSiteCount - 1) As String
    Dim fuseValue() As String 'Dim fuseValue(MaxSiteCount - 1) As String
    Dim skipCalc As Boolean
    Dim recordflag As Boolean
    Dim F_BlowConfig As New SiteLong
    Dim binx_flag_name As String
    Dim biny_flag_name As String
    Dim pid_temp As New SiteLong
    Dim idx_CurrentBinCutJob As Long
    Dim str_Efuse_write_ProductIdentifier As String
    Dim str_Efuse_write_PowerBinning As String
    '''for Efuse obj vbt.
    Dim opbank As eFuseBdfBank ''202011xx for ap
    Dim field As eFuseBdfField
    Dim Pass_Fail_Flag As New SiteBoolean: Pass_Fail_Flag = True
    Dim value As New SiteVariant: value = "0"
    
    Dim idx_dictHarvCore As Integer
    Dim strCorePass As String
    Dim strHarvbin As String
    Dim str_Efuse_HarvestBin As String
    
'''//==================================================================================================================================================================================//'''
'''//Note:
'''We suggest BinCut owners to check the flag.
'''1. "LIB_Vdd_Binning_GlobalVariable": Public Const Flag_Vddbinning_Power_Binning_Fail_Stop = "F_Power_Binning_Fail".
'''2. "BinTable": Add the item with Name "Vddbinning_Power_Binning_Fail_Stop" and ItemList "F_Power_Binning_Fail" with Result "Fail-Stop".
'''3. It updates flags "F_IDS_Binx" and "F_IDS_Biny" for Bin_Table.
'''4. Please check if any column of "power_binning" in PowerBinning flow table, Efuse category "power_binning" in Efuse_BitDef_Table,and the correspondent flag in Bin_Table.
'''//==================================================================================================================================================================================//'''
    '// int
    ReDim specName(MaxSiteCount - 1) As String
    ReDim fusePwrbin(MaxSiteCount - 1) As String
    ReDim fuseValue(MaxSiteCount - 1) As String
    '''//If power_binning table doesn't is not in the workbook or not parsed, skip Power_Binning_Calculation.
    If Flag_PowerBinningTable_Parsed = False Then
        Exit Function
    End If
    
    '''//Get Efuse category "power_binning" for BinCut.
    '''//Please check column "power_binning" in PowerBinning flow table, Efuse category "power_binning" in Efuse_BitDef_Table, and the correspondent flag in Bin_Table.
    If gb_str_EfuseCategory_for_powerbinning <> "" Then
        '''//Check if powerBinning is for the current BinCut testJob.
        If dict_strPmode2EfuseCategory.Exists(UCase(gb_str_EfuseCategory_for_powerbinning)) = True Then
            str_Efuse_write_PowerBinning = get_Efuse_category_by_BinCut_testJob("write", gb_str_EfuseCategory_for_powerbinning)
            
            If str_Efuse_write_PowerBinning <> "" Then
                If Flag_Remove_Printing_BV_voltages = False Then
                    TheExec.Datalog.WriteComment str_Efuse_write_PowerBinning & ",it can use Efuse category:" & str_Efuse_write_PowerBinning
                End If
            Else
                Call Print_Error_Message(Warning_Info, "LIB_Vdd_Binning_customer", "Power_Binning_Calculation", gb_str_EfuseCategory_for_powerbinning & ", no matched Efuse category for the current BinCut testJob. Please check Programming Stage in Efuse_BitDef_Table and Job for Power_Binning_Calculation in Flow_VddBinning. Warning!!!")
                'TheExec.Datalog.WriteComment gb_str_EfuseCategory_for_powerbinning & ", no matched Efuse category for the current BinCut testJob. Please check Programming Stage in Efuse_BitDef_Table and Job for Power_Binning_Calculation in Flow_VddBinning. Warning!!!"
                'TheExec.Datalog.WriteComment gb_str_EfuseCategory_for_powerbinning & ", no matched Efuse category for the current BinCut testJob. Please check Programming Stage in Efuse_BitDef_Table and Job for Power_Binning_Calculation in Flow_VddBinning. Error!!!"
                ''TheExec.ErrorLogMessage gb_str_EfuseCategory_for_powerbinning & ",it can't use Efuse category for Power_Binning_Calculation due to no matched Efuse category for the current BinCut testJob. Please check Programming Stage in Efuse_BitDef_Table and Job for Power_Binning_Calculation in Flow_VddBinning. Error!!!"    ''20230807: Removed by sanity check
                Exit Function
            End If
        Else
            str_Efuse_write_PowerBinning = vbNullString
            Call Print_Error_Message(Warning_Info, "LIB_Vdd_Binning_customer", "Power_Binning_Calculation", "No Efuse category for power_binning, it can't update the related failFlags about power_binning for other testJobs. Please check power_binning flow table and Efuse_BitDef_Table. Warning!!!")
        End If
    Else
        str_Efuse_write_PowerBinning = vbNullString
    End If '''If gb_str_EfuseCategory_for_powerbinning<>""

    '''//Updated width of TName in the datalog.
    Call UpdateDLogColumns_Bincut(110)
    
    '''//init
    RunPwrBinningFlag(0) = True
    ReDim flagPwrbinSheet(PwrBin_SheetCnt - 1)
    ReDim isSheetCalculatedForSite(PwrBin_SheetCnt - 1)
    ReDim power_AllModesWithOffset(PwrBin_SheetCnt - 1)
    anySiteSelected = False
    foundSpec = False
    skipCalc = False
    '''vvv 210730 initial harvestBit array
    If dict_PwrBinning_Col_HarvBin.Count <> 0 Then
        ReDim harvestBit(dict_PwrBinning_Col_HarvBin.Count - 1)
        For idx_dictHarvCore = 0 To UBound(harvestBit)
            harvestBit(idx_dictHarvCore) = 0
        Next idx_dictHarvCore
    Else
        ReDim harvestBit(0)
        harvestBit(0) = 0
    End If
    
    For idx_Sheet = 0 To PwrBin_SheetCnt - 1
        power_AllModesWithOffset(idx_Sheet) = -1
        isSheetCalculatedForSite(idx_Sheet) = False
    Next idx_Sheet
    
    For Each site In TheExec.sites
        specName(site) = vbNullString
        fusePwrbin(site) = vbNullString
        fuseValue(site) = vbNullString
    Next site
        
    '''//Check Harvest_bin for each site.
    '''Print the header with the PowerBinning sheet name.
    If Flag_Enable_PowerBinning_Harvest = True Then
        TheExec.Datalog.WriteComment "=============================================="
        TheExec.Datalog.WriteComment "======    " & "PwrBin Harvest_Bin" & "    ======"
        TheExec.Datalog.WriteComment "=============================================="
        For Each site In TheExec.sites
            '''vvv 20210729 don't judge job to get core pass flag. It should judge by All_Core_Pass flag always         '''vvv NEED MODIFY  vvvv
            strHarvbin = vbNullString
            If dict_PwrBinning_Col_HarvBin.Count <> 0 Then
                For idx_dictHarvCore = 0 To (dict_PwrBinning_Col_HarvBin.Count - 1)
                    str_Efuse_HarvestBin = LCase(dict_PwrBinning_Col_HarvBin(idx_dictHarvCore))
                    str_Efuse_HarvestBin = Replace(str_Efuse_HarvestBin, "input", "")
                    str_Efuse_HarvestBin = Trim(Replace(str_Efuse_HarvestBin, ":", ""))

                    Set field = BdfDataBase.Bank_Cfg.Fields(str_Efuse_HarvestBin)
                    If field.TrimFlagBySite = True Then
                        '''use write buffer
                        If field.TrimAteDecValue = 0 Then
                            harvestBit(idx_dictHarvCore)(site) = 0
                        Else
                            harvestBit(idx_dictHarvCore)(site) = 1
                        End If
                    ElseIf field.Trimmed = True Then
                        ''' use read buffer
                        If field.DsscDecValue = 0 Then
                            harvestBit(idx_dictHarvCore)(site) = 0
                        Else
                            harvestBit(idx_dictHarvCore)(site) = 1
                        End If
                    Else
                        '''binout, setup one impossible bin number
                        harvestBit(idx_dictHarvCore)(site) = 999
                    End If
                    
                    If TheExec.TesterMode = testModeOffline Then harvestBit(idx_dictHarvCore)(site) = 0
                    
                    If strHarvbin <> "" Then
                        strHarvbin = strHarvbin + " , " + str_Efuse_HarvestBin + " = " + CStr(harvestBit(idx_dictHarvCore)(site))
                    Else
                        strHarvbin = str_Efuse_HarvestBin + " = " + CStr(harvestBit(idx_dictHarvCore)(site))
                    End If
                Next idx_dictHarvCore
                TheExec.Datalog.WriteComment "site:" & site & ", Product_identifier = " & CStr(CurrentPassBinCutNum(site) - 1) & " , " & strHarvbin
            Else
                '''************************************************************************************'''
                '''Harvest_bin:
                '''bin1: all cores pass. --> "0".
                '''bin2: one core fails. --> "1".
                '''<Note>: Please check powerBinning tables and discuss Harvest_Bin with project Integrators!!!
                '''************************************************************************************'''
                If bincutJobName = "cp1" Then
                        harvestBit(0) = 0
                Else
                    If TheExec.sites.item(site).FlagState("Harvesting_Bin_Fused") = logicTrue Then
                        harvestBit(0) = 1 '''Discussed this with Minder and Alfred, we thought it should be "1" if fused, 20200428.
                    Else
                        harvestBit(0) = 0
                    End If
                End If
                TheExec.Datalog.WriteComment "site:" & site & ", Product_identifier = " & CStr(CurrentPassBinCutNum(site) - 1) & ", Harvesting_bin = " & CStr(harvestBit(0))
            End If
            
'''            If bincutJobName = "cp1" Then
'''                    If dict_PwrBinning_Col_HarvBin.Count <> 0 Then
'''''                    If dict_HarvCoreMapTable.Count <> 0 And dict_PwrBinning_Col_HarvBin.Count <> 0 Then
''''                    If dict_HarvCoreMapTable.Count Then
'''                        For idx_dictHarvCoreMap = 0 To (dict_PwrBinning_Col_HarvBin.Count - 1)
'''                            strCorePass = dict_HarvCoreMapTable(dict_PwrBinning_Col_HarvBin(idx_dictHarvCoreMap))
'''                            If strCorePass <> "" Then
'''                                If TheExec.sites.item(site).FlagState(strCorePass) = logicTrue Then
'''                                    harvestBit(site)(idx_dictHarvCoreMap) = 0
'''                                Else
'''                                    harvestBit(site)(idx_dictHarvCoreMap) = 1
'''                                End If
'''
'''                                If strHarvbin <> "" Then
'''                                    strHarvbin = strHarvbin + " ," + CStr(harvestBit(site)(idx_dictHarvCoreMap))
'''                                Else
'''                                    strHarvbin = CStr(harvestBit(site)(idx_dictHarvCoreMap))
'''                                End If
'''                            End If
'''                        Next idx_dictHarvCoreMap
'''                        TheExec.Datalog.WriteComment "site:" & site & ", Product_identifier = " & CStr(CurrentPassBinCutNum(site) - 1) & ", Harvesting_bin = " & strHarvbin
'''                    Else
'''                        '''vvv 20210729 This is old. For without Harv_core_map_table
'''                        If TheExec.sites.item(site).FlagState("Gfx_All_Core_Pass") = logicTrue Then
'''                            harvestBit(0) = 0
'''                        Else
'''                            harvestBit(0) = 1
'''                        End If
'''                        TheExec.Datalog.WriteComment "site:" & site & ", Product_identifier = " & CStr(CurrentPassBinCutNum(site) - 1) & ", Harvesting_bin = " & CStr(harvestBit(0))
'''                    End If
''''                Else
''''                    If TheExec.sites.item(site).FlagState("Harvesting_Bin_Fused") = logicTrue Then
''''                        harvestBit(site) = 1 '''Discussed this with Minder and Alfred, we thought it should be "1" if fused, 20200428.
''''                    Else
''''                        harvestBit(site) = 0
''''                    End If
'''            End If
        Next site
    End If
        
    '''//PassBin-loop.
    '''//LBound(PassBinCut_ary) and UBound(PassBinCut_ary) define range of PassBin in the header "Bin Cut List =" of sheets Vdd_Binning_Def.
    For BinNumber = PassBinCut_ary(LBound(PassBinCut_ary)) To PassBinCut_ary(UBound(PassBinCut_ary))
        If BinNumber > 1 Then
            RunPwrBinningFlag(BinNumber) = False '''Add for initial states, 20190826.
            
            For Each site In TheExec.sites
                RunPwrBinningFlag(CurrentPassBinCutNum(site))(site) = True
                
                If CurrentPassBinCutNum(site) = BinNumber Then
                    For p_mode = 0 To MaxPerformanceModeCount - 1
                        If AllBinCut(p_mode).Used = True Then
                            '''****************************************************************************************************************************************************************************'''
                            '''//F_IDS_Binx and F_IDS_Biny are generated in Bin_Table by AutoGen.
                            '''20210302: Modified to add the vbt code for F_IDS_Binx and F_IDS_Biny because PrintOut_VddBinning and Adjust_Binning will be adjusted prior to Power_Binning in Flow_VddBinning.
                            '''****************************************************************************************************************************************************************************'''
                            If VBIN_RESULT(p_mode).passBinCut = 2 Then '''BinX
                                If CDec(IDS_for_BinCut(VddBinStr2Enum(powerDomain)).Real(site)) >= CDec(BinCut(p_mode, 1).IDS_CP_LIMIT(BinCut(p_mode, 1).Mode_Step)) Then '''unit: mA
                                    TheExec.sites.item(site).FlagState("F_IDS_Binx") = logicTrue
                                End If
                                
                                binx_flag_name = "F_Binx_" & Binx_fail_power
                                TheExec.sites.item(site).FlagState(binx_flag_name) = logicTrue
                            ElseIf VBIN_RESULT(p_mode).passBinCut = 3 Then '''BinY
                                If CDec(IDS_for_BinCut(VddBinStr2Enum(powerDomain)).Real(site)) >= CDec(BinCut(p_mode, 2).IDS_CP_LIMIT(BinCut(p_mode, 2).Mode_Step)) Then
                                    TheExec.sites.item(site).FlagState("F_IDS_BinY") = logicTrue
                                End If
                                
                                biny_flag_name = "F_Biny_" & Biny_fail_power
                                TheExec.sites.item(site).FlagState(biny_flag_name) = logicTrue
                            End If
                            
                            '''****************************************************************************************************************************************************************************'''
                            '''//For SearchByPmode, C651 Toby and Chris defined that it only align p_mode product voltage by change CP_GB with CurrentPassBinCutNum.
                            '''BinX and BinY only have Eqn1, so that use CP_GB(0).
                            '''ToDo: Discuss this with C651 Toby to see if VBIN_RESULT(p_mode).GRADEVDD should be smaller than (VBIN_RESULT(p_mode).GRADE+BinCut(p_mode, CurrentPassBinCutNum).CP_GB(0)).
                            '''****************************************************************************************************************************************************************************'''
                            '''20210726: Modified to add the argument "bincutNum As Long" to the vbt function Adjust_Multi_PassBinCut_Per_Site.
                            Adjust_Multi_PassBinCut_Per_Site p_mode, site, CurrentPassBinCutNum(site)
                        End If
                    Next p_mode
                End If
            Next site

            For idx_Sheet = 0 To PwrBin_SheetCnt - 1
                isSheetCalculatedForSite(idx_Sheet) = False
                power_AllModesWithOffset(idx_Sheet) = -1
                isSheetCalculatedForSite(idx_Sheet) = False
            Next idx_Sheet
        ElseIf BinNumber <= Total_Bincut_Num Then
            '''//Initialize the flag to control the site-loop
            RunPwrBinningFlag(BinNumber) = False
            
            For Each site In TheExec.sites
                If CurrentPassBinCutNum(site) = BinNumber And Not Flag_Fuse_BinX(site) = True And TheExec.sites.item(site).FlagState("BIN1") = logicTrue Then
                    RunPwrBinningFlag(CurrentPassBinCutNum(site))(site) = True
                End If
            Next site
        End If '''If binNumber > 1
        
        skipCalc = True
        
        '''//condition-loop.
        For idx_Condition = 0 To UBound(AllPwrBin)
            If AllPwrBin(idx_Condition).passBinCut = BinNumber Then
                '''//init
                anySiteSelected = False
                TheExec.sites.Selected = RunPwrBinningFlag(0)
                
                '''//Check passbin number for each site.
                For Each site In TheExec.sites
                    If CurrentPassBinCutNum(site) = BinNumber And Not Flag_Fuse_BinX(site) = True And TheExec.sites.item(site).FlagState("BIN1") = logicTrue Then
                        RunPwrBinningFlag(CurrentPassBinCutNum(site))(site) = True
                    End If
                Next site
                
                TheExec.sites.Selected = RunPwrBinningFlag(BinNumber)
                
                '''//Check the harvest bin.
                For Each site In TheExec.sites.Selected
                    If AllPwrBin(idx_Condition).harvestUsed = True Then
                        ''''vvv 210730 If one harvestbit isn't match with AllPwrBin().harvestBin then this site can't be use for this power binning spec
                        For idx_dictHarvCore = 0 To UBound(AllPwrBin(idx_Condition).harvestBin)
                            If Not (harvestBit(idx_dictHarvCore) = AllPwrBin(idx_Condition).harvestBin(idx_dictHarvCore)) Then
                                RunPwrBinningFlag(BinNumber)(site) = False
                            End If
                        Next idx_dictHarvCore
                    End If
                Next site
     
                For idx_Spec = 0 To UBound(AllPwrBin(idx_Condition).TestSpec)
                    anySiteSelected = False
                    
                    TheExec.sites.Selected = RunPwrBinningFlag(BinNumber)
                    
                    For Each site In TheExec.sites.Selected
                        anySiteSelected = True
                        Exit For
                    Next site
                    
                    '''//Print the header with "Test Name", ex: bin1_low_power, bin1_high_power.
                    If AllPwrBin(idx_Condition).TestSpec(idx_Spec).TestName <> "" And anySiteSelected = True Then
                        TheExec.Datalog.WriteComment "=============================================="
                        TheExec.Datalog.WriteComment "======    " & "PwrBin Test Name : " & CStr(AllPwrBin(idx_Condition).TestSpec(idx_Spec).TestName) & "    ======"
                        TheExec.Datalog.WriteComment "=============================================="
                    End If
                    
                    '''//Check if site finds powerBinning spec.
                    For Each site In TheExec.sites.Selected
                        If foundSpec(site) = False Then
                            RunPwrBinningFlag(BinNumber)(site) = True
                            foundSpec(site) = True
                        Else
                            RunPwrBinningFlag(BinNumber)(site) = False
                        End If
                    Next site
                    
                    recordflag = False
                    TheExec.sites.Selected = RunPwrBinningFlag(0)
                    PreRunPwrBinningFlag = False
                    TheExec.sites.Selected = RunPwrBinningFlag(BinNumber)
                    recordflag = False
                    
                    '''//Run the specSheet-loop to do PowerBinning.
                    For idx_Sheet = 0 To UBound(AllPwrBin(idx_Condition).TestSpec(idx_Spec).specCustomized)
                        anySiteSelected = False
                        For Each site In TheExec.sites.Selected
                            anySiteSelected = True
                            Exit For
                        Next site
                        
                        If recordflag = False Then
                            PreRunPwrBinningFlag = RunPwrBinningFlag(BinNumber)
                            recordflag = True
                        End If
                        
                        '''//Set the siteMask for powerbinning sheets.
                        If AllPwrBin(idx_Condition).TestSpec(idx_Spec).specUsed(idx_Sheet) = True Then
                            '''//Check if any site exists.
                            For Each site In TheExec.sites.Selected
                                power_total_other_mode(site) = 0
                                power_total_binned_mode(site) = 0
                                
                                If isSheetCalculatedForSite(idx_Sheet) = False Then
                                    anySiteSelected = True
                                End If
                            Next site
                            
                            sheetName = PwrBin_Sheet(idx_Sheet).sheetName
                        
                            If sheetName <> "" And anySiteSelected = True Then
                                '''Print the header with sheetName of the PowerBinning table.
                                TheExec.Datalog.WriteComment "=============================================="
                                TheExec.Datalog.WriteComment "======   " & "PwrBin Sheet : " & CStr(sheetName) & "    ======"
                                TheExec.Datalog.WriteComment "=============================================="
                            End If
                            
                            '''********************************************************************************************************************'''
                            '''//Summation of power consumption for Binned_Mode.
                            '''********************************************************************************************************************'''
                            If dict_Binned_Mode_Ratio2Idx.Count > 0 And PwrBin_Sheet(idx_Sheet).cnt_Binned_Mode > 0 Then
                                For idx_Binned_Mode = 0 To PwrBin_Sheet(idx_Sheet).cnt_Binned_Mode - 1
                                    '''//Check if Binned_Mode exists in the powerbinning spec sheet.
                                    performance_mode = PwrBin_Sheet(idx_Sheet).Binned_Mode(idx_Binned_Mode).pmode
            
                                    '''//Get powerDomain from the column of "Bin Voltage" or "IDS"
                                    '''If Both of these columns ("Bin Voltage" and "IDS") have "*vdd*vdd*", Parse Binned_Mode to get powerDomain, ex: "MAX(VDD_PCPU, VDD_CPU_SRAM)", "VDD_PCPU+VDD_CPU_SRAM"
                                    If LCase(performance_mode) Like "sram_*" Then
                                        If LCase(PwrBin_Sheet(idx_Sheet).Binned_Mode(idx_Binned_Mode).ratio(PwB_Binned_Mode_Ratio2Idx("IDS"))) Like "vdd*+*vdd*" Then
                                            powerDomain = Replace(PwrBin_Sheet(idx_Sheet).Binned_Mode(idx_Binned_Mode).ratio(PwB_Binned_Mode_Ratio2Idx("IDS")), "+", "_")
                                        Else
                                            powerDomain = PwrBin_Sheet(idx_Sheet).Binned_Mode(idx_Binned_Mode).ratio(PwB_Binned_Mode_Ratio2Idx("IDS"))
                                        End If
                                    ''' add to support MPS0*/MES0* for VDD_CPU_SRAM in power binning
                                    ElseIf LCase(performance_mode) Like "m*##*" Or LCase(performance_mode) Like "m*#*" Then
                                        '''//Get powerDomain powerDomain from column "IDS" for Binned_Mode with SRAM 6-digit mode, ex: "MPS001".
                                        If Len(Trim(performance_mode)) = 6 Then
                                            powerDomain = PwrBin_Sheet(idx_Sheet).Binned_Mode(idx_Binned_Mode).ratio(PwB_Binned_Mode_Ratio2Idx("IDS"))
                                        Else
                                            powerDomain = AllBinCut(VddBinStr2Enum(performance_mode)).powerPin
                                        End If
                                    Else
                                        '''//Get powerDomain from colunm "Binned_mode", ex: "VDD_LOW", "LOW".
                                        If dict_IsCorePowerInBinCutFlowSheet.Exists(UCase(performance_mode)) Then
                                            powerDomain = UCase(performance_mode)
                                        ElseIf dict_IsCorePowerInBinCutFlowSheet.Exists(UCase("VDD_" & performance_mode)) Then
                                            powerDomain = UCase("VDD_" & performance_mode)
                                        Else
                                            Call Print_Error_Message(Error_Info, "VBT_LIB_VDD_Binning", "Power_Binning_Calculation", "Power Binning can't get powerDomain name from the p_mode: " & performance_mode & ". Error!!!")
                                            ''TheExec.ErrorLogMessage "Power Binning can't get powerDomain name from the p_mode: " & Performance_mode & ". Error!!!"    ''20230807: Removed by sanity check
                                        End If
                                    End If
                                    
                                    '''//Check if Binned_Mode from PowerBinning table contains BinCut powerDomain or performance mode.
                                    ''' add to support MPS0*/MES0* for VDD_CPU_SRAM in power binning
                                    If LCase(performance_mode) Like "m*##*" Or LCase(performance_mode) Like "m*#*" Then '''ex: "MS001", "MG001"
                                        tName_Temp = powerDomain & "_" & performance_mode
                                    Else '''ex: "SRAM_MG001".
                                        tName_Temp = "VDD_" & performance_mode
                                    End If
                                    
                                    For Each site In TheExec.sites.Selected
                                        If isSheetCalculatedForSite(idx_Sheet) = False Then
                                            '''//Get powerDomain from "Bin Voltage" of Binned_Mode, then get Efuse product voltages by powerDomain.
                                            voltage_Temp = get_Voltage_for_PowerBinning(performance_mode, site, BinNumber, PwrBin_Sheet(idx_Sheet).Binned_Mode(idx_Binned_Mode).ratio(PwB_Binned_Mode_Ratio2Idx("Bin Voltage")))
                
                                            '''//Get IDS values from eFuse if they had been recorded in efuse.
                                            ids_Temp = get_IDS_for_PowerBinning(site, PwrBin_Sheet(idx_Sheet).Binned_Mode(idx_Binned_Mode).ratio(PwB_Binned_Mode_Ratio2Idx("IDS"))) '''unit:mA
                                            
                                            '''//Calculate power for Binned_Mode by the formula provided in PowerBinning tables or documents.
                                            power_Temp(site) = calculate_power_for_binned_mode(idx_Sheet, idx_Binned_Mode, voltage_Temp(site), ids_Temp(site))
                
                                            '''//Sum the power values of Binned_Mode.
                                            power_total_binned_mode = power_total_binned_mode + power_Temp
                
                                            If (remove_printing_power) Then
                                                '''Do nothing
                                            Else
                                                TheExec.Flow.TestLimit ids_Temp / 1000, , , , , scaleMilli, unitAmp, "%.4f", _
                                                                        sheetName & "_" & tName_Temp & "_Ids", , , , unitAmp, , , ForceResults:=tlForceNone
            
                                                TheExec.Flow.TestLimit voltage_Temp / 1000, , , , , scaleMilli, unitVolt, "%.4f", _
                                                                        sheetName & "_" & tName_Temp & "_Vbin", , , , unitVolt, , , ForceResults:=tlForceNone
                                                
                                                '''//Print "C" value from powerbinning tables for each site, requested by PCLINZG and C651 Toby.
                                                If dict_Binned_Mode_Ratio2Idx.Exists(UCase("Vdd0")) = True And dict_Binned_Mode_Ratio2Idx.Exists(UCase("Vdd1")) = True Then
                                                    TheExec.Flow.TestLimit PwrBin_Sheet(idx_Sheet).Binned_Mode(idx_Binned_Mode).ratio(PwB_Binned_Mode_Ratio2Idx("C")) / 1000, , , , , scaleMilli, unitVolt, "%.4f", _
                                                                            sheetName & "_" & tName_Temp & "_C", , , , unitVolt, , , ForceResults:=tlForceNone
                                                End If
            
                                                TheExec.Flow.TestLimit power_Temp, , , , , scaleNoScaling, unitNone, "%.4f", _
                                                                        sheetName & "_" & tName_Temp & "_P_Binned", , , , , , , ForceResults:=tlForceNone
                                            End If
                                        End If
                                    Next site
                                Next idx_Binned_Mode
                            End If
                                
                            '''********************************************************************************************************************'''
                            '''//Summation of power consumption for Other_Mode.
                            '''********************************************************************************************************************'''
                            If dict_Other_Mode_Ratio2Idx.Count > 0 And PwrBin_Sheet(idx_Sheet).cnt_Other_Mode > 0 Then
                                For idx_Other_Mode = 0 To PwrBin_Sheet(idx_Sheet).cnt_Other_Mode - 1
                                    '''//Check if Other_Mode exists in the powerbinning spec sheet.
                                    performance_mode = PwrBin_Sheet(idx_Sheet).Other_Mode(idx_Other_Mode).pmode
                                    
                                    '''//Check if Other_Mode from PowerBinning table contains BinCut powerDomain or performance mode.
                                    If UCase("*," & FullOtherRailinFlowSheet & ",*") Like UCase("*," & "VDD_" & performance_mode & ",*") Then '''ex: "VDD_FIXED", "VDD_LOW, "VDD_SRAM_GPU"
                                        '''//Get powerDomain from the column of "Other_Mode"
                                        If LCase(performance_mode) Like "vdd*" Then
                                            powerDomain = AllBinCut(VddBinStr2Enum(performance_mode)).powerPin
                                        Else
                                            powerDomain = AllBinCut(VddBinStr2Enum("VDD_" & UCase(performance_mode))).powerPin
                                        End If
                                        tName_Temp = powerDomain
                                    Else
                                        '''//Get powerDomain from the column of "Bin Voltage" or "IDS".
                                        ''' add to support MPS0*/MES0* for VDD_CPU_SRAM in power binning
                                        If LCase(performance_mode) Like "m*##*" Or LCase(performance_mode) Like "m*#*" Then
                                            powerDomain = UCase(Trim(PwrBin_Sheet(idx_Sheet).Other_Mode(idx_Other_Mode).ratio(PwB_Other_Mode_Ratio2Idx("Bin Voltage"))))
                                            tName_Temp = powerDomain & "_" & performance_mode
                                        Else
                                            tName_Temp = vbNullString
                                            Call Print_Error_Message(Error_Info, "VBT_LIB_VDD_Binning", "Power_Binning_Calculation", "Power Binning can't get powerDomain name from the p_mode: " & performance_mode & ". Error!!!")
                                            ''TheExec.ErrorLogMessage "Power Binning can't get powerDomain name from the p_mode: " & Performance_mode & ". Error!!!"    ''20230807: Removed by sanity check
                                        End If
                                    End If
                                    
                                    For Each site In TheExec.sites.Selected
                                        If isSheetCalculatedForSite(idx_Sheet) = False Then
                                            '''//Get powerDomain from "Bin Voltage" of Binned_Mode, then get Efuse product voltages by powerDomain.
                                            voltage_Temp = get_Voltage_for_PowerBinning(performance_mode, site, BinNumber, PwrBin_Sheet(idx_Sheet).Other_Mode(idx_Other_Mode).ratio(PwB_Other_Mode_Ratio2Idx("Bin Voltage")))
                
                                            '''//Get IDS values from eFuse if they had been recorded in efuse.
                                            ids_Temp = get_IDS_for_PowerBinning(site, PwrBin_Sheet(idx_Sheet).Other_Mode(idx_Other_Mode).ratio(PwB_Other_Mode_Ratio2Idx("IDS"))) '''unit:mA
                                            
                                            '''//Calculate power for Other_Mode by the formula provided in PowerBinning tables or documents.
                                            power_Temp(site) = calculate_power_for_other_mode(idx_Sheet, idx_Other_Mode, voltage_Temp(site), ids_Temp(site))
                                            
                                            '''//Sum the power values of Other_Mode.
                                            power_total_other_mode = power_total_other_mode + power_Temp
                
                                            If (remove_printing_power) Then
                                                '''nothing
                                            Else
                                                '''//IDS of Other_Mode.
                                                TheExec.Flow.TestLimit ids_Temp / 1000, , , , , scaleMilli, unitAmp, "%.4f", _
                                                                        sheetName & "_" & tName_Temp & "_Ids", , , , unitAmp, , , ForceResults:=tlForceNone
                                                '''//Vbin of Other_Mode.
                                                TheExec.Flow.TestLimit voltage_Temp / 1000, , , , , scaleMilli, unitVolt, "%.4f", _
                                                                        sheetName & "_" & tName_Temp & "_Vbin", , , , unitVolt, , , ForceResults:=tlForceNone
                                                '''//Power of Other_Mode.
                                                TheExec.Flow.TestLimit power_Temp, , , , , scaleNoScaling, , "%.4f", _
                                                                        sheetName & "_" & tName_Temp & "_P_Other", , , , , , , ForceResults:=tlForceNone
                                            End If
                                        End If
                                    Next site
                                Next idx_Other_Mode
                            End If
                                
                            '''********************************************************************'''
                            '''//Calculate the total power: P_total=Pbinned_total+Pother_total+Offset
                            '''********************************************************************'''
                            For Each site In TheExec.sites.Selected
                                If isSheetCalculatedForSite(idx_Sheet) = False Then
                                    power_AllModesWithOffset(idx_Sheet) = power_total_other_mode + power_total_binned_mode + PwrBin_Sheet(idx_Sheet).offset
                                    isSheetCalculatedForSite(idx_Sheet) = True
                                End If
                                
                                TheExec.Flow.TestLimit resultVal:=power_AllModesWithOffset(idx_Sheet), lowVal:=0, hiVal:=AllPwrBin(idx_Condition).TestSpec(idx_Spec).specCustomized(idx_Sheet), formatStr:="%.4f", _
                                                        ForceResults:=tlForceNone, scaletype:=scaleNoScaling, Tname:=sheetName & "_Power_Binning_P_total"
            
                                If power_AllModesWithOffset(idx_Sheet) > AllPwrBin(idx_Condition).TestSpec(idx_Spec).specCustomized(idx_Sheet) Then
                                    '''================================================================='''
                                    '''//Check if "CurrentPassBinCutNum(site) <= Total_Bincut_Num".
                                    '''================================================================='''
                                    If CurrentPassBinCutNum(site) <= Total_Bincut_Num Then
                                        '''//Check if next spec exists with same passbin and harvest_bin.
                                        If AllPwrBin(idx_Condition).TestSpec(idx_Spec).haveNextSpec = False Then
                                            '''//Since only bin1 dice can enter this func, here just adjust CurrentPassBinCutNum to next BinCutNum, ex: Bin1 to BinX
                                            CurrentPassBinCutNum(site) = BinNumber + 1
        
                                            If CurrentPassBinCutNum(site) <= Total_Bincut_Num Then
                                                TheExec.Datalog.WriteComment "site:" & site & ", Power_Binning adjusts from Bin" & BinNumber & " to Bin" & CurrentPassBinCutNum(site)
                                            End If
                                            
                                            '''//For startBin sorting of PowerBinning, user can use this part and add flags to Bin_Table.
                                            '    Select Case CurrentPassBinCutNum(site)
                                            '        Case 1: TheExec.sites.item(site).FlagState("F_IsBin1_PowerBinning") = logicTrue
                                            '        Case 2: TheExec.sites.item(site).FlagState("F_IsBinx_PowerBinning") = logicTrue
                                            '        Case 3: TheExec.sites.item(site).FlagState("F_IsBiny_PowerBinning") = logicTrue
                                            '        Case Else
                                            '    End Select
                                            
                                            '''//If the site fails on the spec, mask the site not to run the next spec.
                                            RunPwrBinningFlag(BinNumber)(site) = False
                                        Else
                                            RunPwrBinningFlag(BinNumber)(site) = False 'True
                                        End If
                                        
                                        foundSpec(site) = foundSpec(site) And False
                                    Else
                                        TheExec.sites.item(site).FlagState(strGlb_Flag_Vddbinning_Power_Binning_Fail_Stop) = logicTrue
                                        foundSpec(site) = False
                                    End If
                                Else
                                    foundSpec(site) = foundSpec(site) And True
                                End If
                            Next site
                        End If
                        
                        TheExec.sites.Selected = RunPwrBinningFlag(BinNumber)
                    Next idx_Sheet
                    
                    TheExec.sites.Selected = PreRunPwrBinningFlag

                    For Each site In TheExec.sites.Selected
                        If foundSpec(site) = True Then
                            specName(site) = AllPwrBin(idx_Condition).TestSpec(idx_Spec).TestName
                            fusePwrbin(site) = AllPwrBin(idx_Condition).TestSpec(idx_Spec).fusePwrbin
                            fuseValue(site) = AllPwrBin(idx_Condition).TestSpec(idx_Spec).fuseValue
                            RunPwrBinningFlag(BinNumber)(site) = False
                        Else
                            RunPwrBinningFlag(BinNumber)(site) = True
                        End If
                    Next site
                    
                    TheExec.sites.Selected = RunPwrBinningFlag(BinNumber)
                Next idx_Spec
            Else
                '''Do nothing
            End If
        Next idx_Condition
            
        '''//Reset Selected flag for next test instance
        TheExec.sites.Selected = RunPwrBinningFlag(0)
        
        '''//Update failFlag of BinTable for PowerBinning.
        For Each site In TheExec.sites
            '''Check if "CurrentPassBinCutNum(site) <= Total_Bincut_Num".
            If CurrentPassBinCutNum(site) <= Total_Bincut_Num Then
                TheExec.sites.item(site).FlagState(strGlb_Flag_Vddbinning_Power_Binning_Fail_Stop) = logicFalse
            Else
                TheExec.sites.item(site).FlagState(strGlb_Flag_Vddbinning_Power_Binning_Fail_Stop) = logicTrue
            End If
        Next site
    Next BinNumber
        
    '''//If "PowerBinning" has the column "fuse_value" or "PASS: power_binning", select the fuse_value for SPEC.
    TheExec.Datalog.WriteComment "======================================="
    TheExec.Datalog.WriteComment "======   Power Binning Summary   ======"
    TheExec.Datalog.WriteComment "======================================="

    For Each site In TheExec.sites
        If foundSpec(site) = True Then
            '''**********************************************************************************************************************************************'''
            '''//For projects with Harvest cores, it needs to update flags for Efuse. Please check Efuse_BitDef_Table with project BinCut and Efuse owners.
            '''**********************************************************************************************************************************************'''
            If PwrBin_SheetnameDict.Exists(fuseValue(site)) Then '''sheetName exists, fuse total power value
                TheExec.Datalog.WriteComment "site:" & site & ", currentpassbinnum=" & CurrentPassBinCutNum(site) & ", power_binning = " & fusePwrbin(site) & _
                                                ", fuse_name2 = " & Format(power_AllModesWithOffset(PwrBin_SheetnameDict.item(fuseValue(site))), "0.0000") & ", spec_name = " & specName(site)   '''siteDouble
            Else '''fuse number of the specName
                TheExec.Datalog.WriteComment "site:" & site & ", currentpassbinnum=" & CurrentPassBinCutNum(site) & ", power_binning = " & fusePwrbin(site) & _
                                                ", fuse_name2 = " & fuseValue(site) & ", spec_name = " & specName(site)
            End If
            
            '''//Check if "PASS: power_binning" exists in the header of PowerBinning table.
            If gb_str_EfuseCategory_for_powerbinning <> "" Then
                If fusePwrbin(site) <> "" Then
                    '''For project with Efuse obj vbt code.
                    value = fusePwrbin(site)
                
                    '''//According to the value of "power_binning" from Efuse, update the flag of "power_binning" in Bin_Table.
                    '''*****************************************************************************************************************************************''''''''''''''''''''''''''''''''''''''''''
                    '''//Please check the column of "PASS: power_binning" in powerBinning table "PwrBinning_V***" and see if any failflag for this in Bin_Table.
                    '''Flags "F_PWRBIN_LOW", "F_PWRBIN_HIGH", "F_PWRBIN_HIGH", and "F_PWRBIN_LOWLOW" are defined in the column "Comment" and related to items SortBin in Bin_Table.
                    '''Please check these flags in Bin_Table to see if any matched SortBin is available.
                    '''*****************************************************************************************************************************************''''''''''''''''''''''''''''''''''''''''''
                    If (fusePwrbin(site) = "3") Then
                        TheExec.sites.item(site).FlagState("F_PWRBIN_LOW") = logicTrue
                    ElseIf (fusePwrbin(site) = "2") Then
                        TheExec.sites.item(site).FlagState("F_PWRBIN_HIGH") = logicTrue
                    ElseIf (fusePwrbin(site) = "1") Then
                        TheExec.sites.item(site).FlagState("F_PWRBIN_LOWLOW") = logicTrue
                    ElseIf (fusePwrbin(site) = "0") Then
                        '''do nothing, there is no power binning in binx/y(the value will be zero).
                    Else
                        '''20210610: Modified to check if it sent the incorrect value "power_binning" to Efuse. If that, bin out the failed DUT.
                        TheExec.sites.item(site).FlagState(strGlb_Flag_Vddbinning_Power_Binning_Fail_Stop) = logicTrue
                        Call Print_Error_Message(Error_Info, "VBT_LIB_VDD_Binning", "Power_Binning_Calculation", "Site:" & site & ",power binning value:" & fusePwrbin(site) & ", it is not defined in power binning table for Power_Binning_Calculation,Error!!!")
                                                'TheExec.Datalog.WriteComment "Site:" & site & ",power binning value:" & fusePwrbin(site) & ", it is not defined in power binning table for Power_Binning_Calculation"
                        'TheExec.ErrorLogMessage "Site:" & site & ",power binning value:" & fusePwrbin(site) & ", it is not defined in power binning table for Power_Binning_Calculation"    ''20230807: Removed by sanity check
                    End If
                End If '''If fusePwrbin(site) <> ""
            End If '''If gb_str_EfuseCategory_for_powerbinning <> ""
        Else
            TheExec.Datalog.WriteComment "site:" & site & ", currentpassbinnum=" & CurrentPassBinCutNum(site)
        End If
    Next site
    
    '''//Update the value of "power_binning" to Efuse.
    '''//Please check if "power_binning" exists in Efuse_BitDef_Table and PowerBinning table "PwrBinning_V##".
    '''20210708: Modified to check if Efuse category "power_binning" for the current BinCut testJob has the matched programming stage in Efuse_BitDef_Table.
    If str_Efuse_write_PowerBinning <> "" Then
        '''For project with Efuse obj vbt code.
        '''20201228: Patty asked us to adapt Efuse object vbt code.
        ''====20201230 add for efuse new code====
        Set opbank = GetBdfBank("CFG")
        Set field = opbank.Fields(str_Efuse_write_PowerBinning)
        opbank.SetEfuse field.name, value, Pass_Fail_Flag, , , , True
    End If
    
    '''*************************************************************************************************'''
'    '''//This is optional for project, and please check Efuse_BitDef_Table.
'    '''//Efuse product_identifer = CurrentPassBinCutNum - 1
'    pid_temp = CurrentPassBinCutNum.Subtract(1)
'
'    '''//Update PassBinNum of each site to Efuse "Product_Identifier".
'    '''//Get Efuse category "Product_Identifier" for BinCut.
'    '''<Note>: Since Efuse obj vbt code only provides one time permission to set the item, it can't update Efuse product_identifier in the vbt function Power_Binning_Calculation...
'    '''20210706: Modified to use the vbt function get_Efuse_category_by_BinCut_testJob to find the Efuse Category.
'    str_Efuse_write_ProductIdentifier = get_Efuse_category_by_BinCut_testJob("write", "Product_identifier")
'
'    If str_Efuse_write_ProductIdentifier <> "" Then
'        '''For project with Efuse obj vbt code.
'        Set field = opbank.Fields(str_Efuse_write_ProductIdentifier)
'        opbank.SetEfuse field.name, pid_temp, Pass_Fail_Flag, , , , True
'    Else
'        Print_Error_Message (Error_Info, "VBT_LIB_VDD_Binning", "","Product_identifier" & ", it can't write Efuse Product_identifier for adjust_VddBinning due to no matched Efuse category for the current BinCut testJob. Please check Programming Stage in Efuse_BitDef_Table. Error!!!")
'        'TheExec.ErrorLogMessage "Product_identifier" & ", it can't write product voltage for adjust_VddBinning due to no matched Efuse category for the current BinCut testJob. Please check Programming Stage in Efuse_BitDef_Table. Error!!!"    ''20230807: Removed by sanity check
'    End If
    '''*************************************************************************************'''
    
    '''//Align the format of Judge_PF and Judge_PF_func in the datalog.
    TheExec.Datalog.WriteComment ""
    TheExec.Datalog.WriteComment ""
    TheExec.Datalog.WriteComment vbCrLf
Exit Function
errHandler: 'Add ErrHandler 2023/08/01
    Call Print_Error_Message(Error_Info, "VBT_LIB_VDD_Binning", "Power_Binning_Calculation") 'Add ErrHandler 2023/08/01
    If AbortTest Then Exit Function Else Resume Next 'Add ErrHandler 2023/08/01
End Function

'20210908: Modified to merge vbt functions Non_Binning_Pwr_Setting_VT, HVCC_Set_VT, and PostBinCut_Voltage_Set_VT into the vbt function bincut_power_Setting_VT, as discussed with TSMC ZYLINI.
'20210603: Modified to move inst_info.Pattern_Pmode and inst_info.By_Mode from initialize_inst_info to GradeSearch_XXX_VT.
'20210530: Modified to replace typo "Multisftp_Binout" with "MultiFstp_NoBinout".
'20210517: Modified to overwrite fail-stop for MultiFSTP if TheExec.sites.item(site).FlagState("Multisftp_Binout") = logicTrue.
'20210223: Modified to move "Check_and_Decompose_PrePatt_FuncPat" prior to "decide_bincut_feature_for_stepsearch".
'20210207: Modified to move the vbt code of switching DCVS Valt from print_voltage_info_before_FuncPat to GradeSearch_XXX_VT.
'20210126: Modified to revise the vbt code for DevChar.
'20210125: Modified to move "select_DCVS_output_for_powerDomain" prior to "Set_PayloadVoltage_to_DCVS".
'20210122: Modified to check if FuncPat <> "" for print_voltage_info_before_FuncPat.
'20201210: Modified to use run_patt_from_FuncPat_for_BinCut for running the pattern decomposed from FuncPat.
'20201204: Modified to initialize "inst_info.PrePattPass", "inst_info.FuncPatPass", and "inst_info.sitePatPass" in the vbt function initialize_inst_info.
'20201204: Modified to initialize "inst_info.IndexLevelPerSite = -1" in the vbt function initialize_inst_info.
'20201123: Modified to align the format of Judge_PF and Judge_PF_func in the datalog.
'20201118: Modified to use "TheHdw.Digital.Patgen.PatternBurstPassedPerSite" to get siteResult of pattern pass/fail.
'20201111: Modified to use "inst_info.voltage_SelsrmBitCalc".
'20201110: Modified to check if FuncPat <> "".
'20201029: Modified to use inst_info.previousDcvsOutput and inst_info.currentDcvsOutput.
'20201029: Modified to use inst_info.is_BV_Safe_Voltage_printed and inst_info.is_BV_Payload_Voltage_printed.
'20201027: Modified to use "Public Type Instance_Info".
'20201026: Modified to switch DCVS to Vmain prior to PrePatt for TD pattern burst by C651 Toby.
'20201015: Modified to save result of pattern Pass/Fail in "sitePatPass".
'20201012: Modified to change the arguments of the vbt function "check_patt_Pass_Fail".
'20201008: Modified to replace "PrintedBVinDatalog" with "is_BV_Payload_Voltage_printed".
'20200924: Modified to merge the branches of "Calculate_Selsrm_DSSC_For_BinCut".
'20200923: Modified to use "run_patt_offline_simulation" for offline simulation.
'20200923: Modified to call "check_patt_Pass_Fail".
'20200922: Modified to bin out the failed DUT if "PattPass = False".
'20200922: Modified the branch to simulate offline random Pass/Fail.
'20200922: Modified to use "ary_FuncPat_decomposed(indexPatt)" for RunFailCycle.
'20200922: Modified to align the branches of running FuncPat for online and offline tests.
'20200921: Modified to replace the PrePatt vbt block with calling "run_prepatt_decompose_VT".
'20200921: Discussed "RTOS_bootup_relay" / "KeepAliveFlag" / "spi_ttr_flag" with SWLINZA and PCLINZG. We decided to remove these SPI/RTOS branches because RTOS didn't use pattern test since Cebu/Sicily/Tonga/JC-Chop/Ellis/Bora.
'20200918: Modified to add the argument "result_mode" for the vbt function "Check_and_Decompose_PrePatt_FuncPat".
'20200918: Modified to use "prepare_DCVS_Output_for_RailSwitch".
'20200918: Modified to use "print_voltage_info_before_FuncPat".
'20200908: Modified to remove the redundant site-loop.
'20200821: Modified to add "Dim str_Selsrm_DSSC_Bit(MaxSiteCount-1) As String".
'20200802: Modified to check patType init or payload, revised by Leon.
'20200730: Modified to add the EnableWord "VDDBinning_Offline_AllPattPass" for Offline simulation with all patterns pass.
'20200727: Modified to follow the new rule of SELSRM bit calculation proposed by C651 Toby.
'20200711: Modified to use the siteDouble array "BinCut_Payload_Voltage" to store BinCut payload voltages.
'20200630: Modified to remove the unused "thehdw.Utility.Pins("k02,k04").State = tlUtilBitOff".
'20200609: Modified to use Check_alarmFail_before_BinCut_Initial.
'20200520: Modified to use Check_Pattern_NoBurst_NoDecompose to show the errorLogMessage if "burst=no" and "Decompose_Pattern=false".
'20200520: Modified to use "Check_and_Decompose_PrePatt_FuncPat" to check and decompose patsets PrePatt and FuncPat, and find SELSRAM DSSC pattern for DSSC digSrc.
'20200424: Modified to use "Set_BinCut_Initial_by_ApplyLevelsTiming" to set BinCut initial voltage by ApplyLevelsTiming.
'20200324: Modified to skip ApplyLevelsTiming when current instance has the same level/timing as previous instance for project with rail-switch.
'20200320: Modified to check instance contexts of current instance and previous instance.
'20200319: Modified to switch off save_core_power_vddbinning and restore_core_power_vddbinning if Flag_Enable_Rail_Switch = True.
'20200217: Modified to check if no vbump before running the payload pattern.
'20200207: Modified to replace set_core_power_main and set_core_power_alt with select_DCVS_output_for_powerDomain.
'20200203: Modified to use the function "print_bincut_power".
'20200130: Modified to call Calculate_Selsrm_DSSC_For_BinCut for SELSRM DSSC bits calculation.
'20200130: Modified to init BinCut_Init_Voltage and BinCut_Payload_Voltage.
'20200130: Modified to call Get_Pmode_Addimode_Testtype_fromInstance to get pmode/addi_mode/testtype.
'20200121: Modified for pattern decomposed.
'20200113: Modified for pattern bursted without decomposing pattern.
'20190106: Modified to add "TheHdw.Alarms.Check".
'20191204: Modified to check if init patts pass or fail...
'20191202: Modified for the revised initVddBinCondition.
'20191127: Modified for the revised InitVddBinTable.
'20191113: Modified to decide BlockType for SELSRM_Mapping_Table.
'20190627: Modified to use the global variable "pinGroup_BinCut" for BinCut powerDomains.
'20190617: Modified to use siteDouble "CorePowerStored()" to save/restore voltages for BinCut powerDomains.
'20190615: Modified for offline tests.
'20190611: Modified for set voltages to "DC Specs" (by DcSpecsCategoryForInitPat) for initial and "init patterns" before ApplyLevelsTiming.
'20190606: Modified to add the argument "DcSpecsCategoryForInitPat as string" for Init patterns with the new test setting DC Specs.
'20190508: Modified for C651 new performance mode naming rule, ex: "VDD_PCPU_MC60A".
'20190319: Modified to add "Flag_Enable_Rail_Switch" to turn VRS Rail Switch on/off.
'20190215: Modified for DCVS shadow voltages.
'20181224: Modified to add Flag_SyncUp_DCVS_Output_enable to control SyncUp on/off.
'20181115: Modified for DSSC TD patt_group (init+payload1+init+payload).
'20181113: Modified for TD payload pattern (print_alt_power_payload).
'20181004: Warning!!! Do not add "Optional No_Bin_Out As Boolean = False" into arguments the function.
Public Function GradeSearch_HVCC_VT(FuncPat As Pattern, performance_mode As String, result_mode As tlResultMode, DecomposePatt As String, _
                                    FuncTestOnly As Boolean, IDSCurrentLimitList As String, PrePatt As Pattern, _
                                    Optional SpiCounterValue As Integer, Optional RunFailCycle As Boolean, _
                                    Optional Validating_ As Boolean, Optional DcSpecsCategoryForInitPat As String = vbNullString, _
                                    Optional HarvestBinningFlag As String = vbNullString, _
                                    Optional EnableCoreHarvest As Boolean = False, Optional EnableCoreMask As Boolean = true, _
                                    Optional Harv_CommonFailFlag As String = vbNullString, Optional UserFunction As String = vbNullString, _
                                    Optional failpins As String = vbNullString, Optional CaptureSize As Long = 0)
On Error GoTo errHandler 'Add ErrHandler 2023/08/01
    Dim site As Variant
    Dim inst_info As Instance_Info
    Dim indexPatt As Long
    Dim instSSNinfo As Inst_SSN
    
    glb_TestInstance = vbNullString
    glb_TestInstance = UCase(TheExec.DataManager.instanceName)
    
    If Validating_ Then
        If FuncPat.value <> "" Then Call PrLoadPattern(FuncPat.value)
        If PrePatt.value <> "" Then Call PrLoadPattern(PrePatt.value)
        Exit Function ''' Exit after validation
    End If
    
    If PrePatt <> "" Then
        Shmoo_Pattern = FuncPat.value & "," & PrePatt.value
    Else
        Shmoo_Pattern = FuncPat.value
    End If

    '''//Initialize inst_info.
    '''//Get p_mode, addi_mode, jobIdx, testtype, and offsettestype from test instance and performance_mode.
    Call initialize_inst_info(inst_info, performance_mode)
    inst_info.selsrm_DigSrc_Pin = "JTAG_TDI"
    inst_info.selsrm_DigSrc_SignalName = "DigSrcSignal"
    '''For Harvest MultiFSTP.
    inst_info.Harvest_Core_DigSrc_Pin = "JTAG_TDI"
    inst_info.Harvest_Core_DigSrc_SignalName = "Harvest_Core_DigSrcSignal"
    '20240124: New MFSTP defined by Alan
    'Dim instInfoUserFunction As UserDefinedFunctionType
    Call CheckInstForUserFunction(UserFunction, inst_info.digSrcLabel, inst_info.digSrcPatterns, CStr(FuncPat))
    
    If inst_info.digSrcLabel(0) <> "" Then
        inst_info.MultiFSTP_Enable = True
    Else
        inst_info.MultiFSTP_Enable = False
    End If
    '''20220208,HarvPinGrp variable set
    inst_info.enable_HarvPinGrp = EnableCoreHarvest
    '''20240604: Added independant boolean for saving pin/core mask
    inst_info.enable_CoreMask = EnableCoreMask
    inst_info.HarvestBinningFlag = HarvestBinningFlag
    inst_info.Harv_CommonFailFlag = Harv_CommonFailFlag
    '''20240726 add pin mask or core mask to check
    If (inst_info.enable_HarvPinGrp = True Or inst_info.enable_CoreMask = True) And Flag_HarvPinFlag_Mapping_Table_Parsed = False Then
        Call Print_Error_Message(Error_Info, "VBT_LIB_VDD_Binning", "GradeSearch_VT", "TestInst Argument 'HarvPinGrp_Enable' or 'enable_CoreMask' = True, but can't find HarvPinGrp Table! Error!!")
    End If
    
    '''//Check if DevChar Precondition is tested.
    If inst_info.is_DevChar_Running = True And inst_info.get_DevChar_Precondition = False Then
        Call Set_BinCut_Initial_by_ApplyLevelsTiming(inst_info)
        Exit Function
    End If

    '''//Check if alarmFail was triggered prior to BinCut initial(applyLevelsTiming).
    Check_alarmFail_before_BinCut_Initial inst_info.inst_name
    alarmFail = False
    
    '''//Set initial voltages from category "Bincut_X_X_X" in DC_Specs sheet by ApplyLevelsTiming.
    '''Print the initial voltages, and applies them to DCVS Vmain and Valt by ApplyLevelsTiming (DCVS voltage source will be switched to Vmain).
    Call Set_BinCut_Initial_by_ApplyLevelsTiming(inst_info)

    '''//Decompose PrePatt and FuncPat to check if any DSSC digsrc pattern of SELSRM(defined in SELSRM_Mapping_Table) exists in the pattern sets.
    Call Check_and_Decompose_PrePatt_FuncPat(inst_info, result_mode, DecomposePatt, PrePatt.value, FuncPat.value)

    '''*********************************************************************************'''
    '''  Added for SSN Harvest (Core Diagnosis) & SSN Check (Core Mask) 20240702        '''
    '''*********************************************************************************'''
    If Flag_HarvPinFlag_Mapping_Table_Parsed = True And ssnPatternsDict.Count > 0 Then
        Dim Decomposed_Pat_SSN() As String
        Dim patCnt As Long
        Decomposed_Pat_SSN = TheExec.DataManager.Raw.GetPatternsInSet(FuncPat.value, patCnt)
        Call CheckInstForSSN_BV(Decomposed_Pat_SSN, instSSNinfo, SSNMapping, inst_info.enable_HarvPinGrp, inst_info.enable_CoreMask)
        inst_info.BeSSNTest = instSSNinfo.bSSNTest '240703
    End If
    
    '''20230912: Added FFC feature for HVCC
    Call check_flag_to_enable_CMEM_collection(inst_info, Flag_Enable_CMEM_Collection, inst_info.count_FuncPat_decomposed, CaptureSize, failpins, True)
    
    '''20230912: Added FFC feature for HVCC
    '''//Clear capture Memory(CMEM) and resize array of CMEM_data if inst_info.enable_CMEM_Collection is enabled.
    Call resize_CMEM_Data_by_pattern_number(inst_info.enable_CMEM_Collection, inst_info.count_FuncPat_decomposed, inst_info.Step_CMEM_Data)
    
    '''20220208,Add for check arg about harvPinGrp
    Call check_arg_to_enable_HarvPinGrp(inst_info)
    
    '''//Set the excluded performance mode if the device is bin2 die and the performance mode doesn't exist in bin2 table.
    SkipTestBin2Site inst_info.p_mode, inst_info.Active_site_count
    
    If inst_info.Active_site_count = 0 Then
        RestoreSkipTestBin2Site inst_info.p_mode
        Exit Function
    End If
    
    '''******************************************************************************************'''
    '''//Set DCVS voltage output to Vmain prior to PrePatt and FuncPat.
    '''******************************************************************************************'''
    select_DCVS_output_for_powerDomain tlDCVSVoltageMain
    inst_info.currentDcvsOutput = tlDCVSVoltageMain
    
    '''//Initialize array of BinCut_Init_Voltage and BinCut_Payload_Voltage before BinCut payload voltages calculation.
    Init_BinCut_Voltage_Array
    
    '''//Calculate BinCut payload voltages of BinCut CorePower and OtherRail.
    '''20210908: Modified to merge vbt functions Non_Binning_Pwr_Setting_VT, HVCC_Set_VT, and PostBinCut_Voltage_Set_VT into the vbt function bincut_power_Setting_VT, as discussed with TSMC ZYLINI.
    Call bincut_power_Setting_VT(inst_info, CurrentPassBinCutNum, BinCut_Payload_Voltage)
    
    '''//If DSSC digsrc pattern of SELSRM exists in the pattern sets, calculate DSSC bits sequence by Selsrm Logic power/SRAMthresh, then prepare DSSC digsrc signal setups.
    Call Calculate_Selsrm_DSSC_For_BinCut(inst_info, VBIN_RESULT(inst_info.p_mode).passBinCut)

    '''//Set Payload voltages to DCVS. For projects with Rail-Switch, BinCut payload voltage values are applied to DCVS Valt.
    '''BinCut_Payload_Voltage is the siteDouble array for storing BinCut Payload voltage values calculated from HVCC_Set_VT.
    Set_PayloadVoltage_to_DCVS Flag_Enable_Rail_Switch, pinGroup_BinCut, BinCut_Payload_Voltage
    TheHdw.Wait 0.001
    
    '===================================================================
    ' Print BinCut init voltages for PrePatt, then run PrePatt(init pattern).
    '===================================================================
    Call run_prepatt_decompose_VT(inst_info, inst_info.PrePatt, inst_info.ary_PrePatt_decomposed, inst_info.count_PrePatt_decomposed, inst_info.PrePattPass, DcSpecsCategoryForInitPat)
    
    '''//Update siteResult of PrePatt.
    If inst_info.is_DevChar_Running = False Then '''for DevChar.
        Call update_Pattern_result_to_PattPass(inst_info.PrePattPass, inst_info.funcPatPass)
    End If

    '''//Check if FuncPat is empty...
    If FuncPat <> "" Then
        '''******************************************************************************************'''
        '''//For Mbist instances in project with rail-switch, set DCVS voltage output to Valt prior to FuncPat.
        '''******************************************************************************************'''
        '''20210207: Modified to move the vbt code of switching DCVS Valt from print_voltage_info_before_FuncPat to GradeSearch_XXX_VT.
        If inst_info.Test_Type = testType.Mbist Then '''ex: "*cpu*bist*", "*gfx*bist*", "*gpu*bist*", "*soc*bist*".
            '''====================================================================================================='''
            '''C651 didn't implement Vbump op-code in MBIST init pattern for project with rail-switch.
            '''So that we have to switch DCVS to Valt by VBT code here before running FuncPat for Mbist instances.
            '''====================================================================================================='''
            If Flag_Enable_Rail_Switch Then '''For projects with Rail Switch, BinCut payload voltages are applied to DCVS Valt.
                select_DCVS_output_for_powerDomain tlDCVSVoltageAlt
                inst_info.currentDcvsOutput = tlDCVSVoltageAlt
            Else '''For conventional projects without Rail Switch, BinCut payload voltages(BV) are applied to DCVS Vmain.
                select_DCVS_output_for_powerDomain tlDCVSVoltageMain
                inst_info.currentDcvsOutput = tlDCVSVoltageMain
            End If
        End If
    
        '''//Print BinCut voltage before running FuncPat.
        Call print_voltage_info_before_FuncPat(inst_info)
'**********************************************
'@@FuncPat pattern-loop Start
'**********************************************
        For indexPatt = 0 To inst_info.count_FuncPat_decomposed - 1
            If RunFailCycle Then
                If TheExec.enableWord("Mbist_FingerPrint") = True And TheExec.enableWord("TTR_Enable") <> True Then
                    Call Finger_print(inst_info.ary_FuncPat_decomposed(indexPatt), RunFailCycle)      'Mbist finger print VBT '20190629 Oscar Compile Check
                Else
                    Call TheHdw.Patterns(inst_info.ary_FuncPat_decomposed(indexPatt)).test(pfAlways, 0, inst_info.result_mode)
                End If
            Else
                '''====================================================================================================='''
                '''//Run pattern decomposed from FuncPatt patset, and get siteResult of pattern pass/fail.
                '''Sync up DCVS output and print BinCut payload voltage for projects with Rail Switch for TD instance.
                '''Run the pattern decomposed from FuncPat, and get siteResult of pattern pass/fail.
                '''====================================================================================================='''
                
                Call run_patt_from_FuncPat_for_BinCut(inst_info, indexPatt, inst_info.ary_FuncPat_decomposed(indexPatt), inst_info.funcPatPass, instSSNinfo, inst_info.idxBlock_Selsrm_FuncPat, CaptureSize, failpins)
                
                '''//Bin out the failed DUT if "PattPass = False"...
                If inst_info.is_DevChar_Running = False Then '''for DevChar.
                    For Each site In TheExec.sites
                        If inst_info.funcPatPass(site) = False And inst_info.enable_HarvPinGrp = False And inst_info.HarvestBinningFlag = "" Then '20240719 normal BV check without any harvest
                            '''//Overwrite fail-stop for MultiFSTP if TheExec.sites.item(site).FlagState("MultiFstp_NoBinout") = logicTrue.
                            '''ToDo: Please check if the failFlag ""MultiFstp_NoBinout"" exists in the flow table!!!
                            If TheExec.sites.item(site).FlagState("MultiFstp_NoBinout") = logicTrue Then 'MultiFstp without Binout
                                inst_info.inst_name = TheExec.DataManager.instanceName
                                TheExec.Datalog.WriteComment "Site:" & site & "," & inst_info.inst_name & ", test failed, but MultiFSTP bypassed BinOut!"
                            ElseIf TheExec.Flow.enableWord("Vddbin_Skip_BinCheck_Enable") = True Then
                                'TheExec.Datalog.WriteComment "Site:" & site & ",Pattern Fails but the result is skipped due to Vddbin_Skip_BinCheck_Enable=True!!"
                            Else
                                TheExec.sites.item(site).testResult = siteFail
                                TheExec.sites.item(site).FlagState(strGlb_Flag_Vddbinning_Fail_Stop) = logicTrue
                            End If
                        End If
                    Next site
                End If
            End If '''If RunFailCycle Then
        Next indexPatt
'**********************************************
'@@FuncPat pattern-loop End
'**********************************************
        DebugPrintFunc FuncPat.value
    End If '''If FuncPat <> ""

    '''//Check if running FuncPat with "burst=no" and "Decompose_Pattern=false".
    Call Check_Pattern_NoBurst_NoDecompose(inst_info.FuncPat, inst_info.count_FuncPat_decomposed, inst_info.enable_DecomposePatt)
    
    '''20230912: Added FFC feature for HVCC
    If inst_info.enable_CMEM_Collection = True And inst_info.funcPatPass.Any(False) Then
        Call PostTestIPF(inst_info.performance_mode, failpins, inst_info.PrintSize, inst_info.Step_CMEM_Data, inst_info.str_patternPins, inst_info.PrintSize)
        TheHdw.Digital.CMEM.SetCaptureConfig 0, CmemCaptNone '''CmemCaptNone: Capture no cycles.
        TheHdw.Digital.CMEM.CentralFields = tlCMEMNone  '20220512, add for turn off cmem capture
    End If
    
    '''20220208,HarvPinGrp ControlFailFlag for binout
    Call CtrlFailFlag_byHarvPinGrpResult(inst_info)
    Call CtrlFailFlag_ByPatternHarvResult(inst_info)
    
    '==================================================
    'Restore the site which is disabled for bin2 chip
    '==================================================
    RestoreSkipTestBin2Site inst_info.p_mode
    
    '''//Align the format of Judge_PF and Judge_PF_func in the datalog.
    TheExec.Datalog.WriteComment ""
    TheExec.Datalog.WriteComment ""
    TheExec.Datalog.WriteComment vbCrLf
Exit Function
errHandler: 'Add ErrHandler 2023/08/01
    Call Print_Error_Message(Error_Info, "VBT_LIB_VDD_Binning", "GradeSearch_HVCC_VT") 'Add ErrHandler 2023/08/01
    If AbortTest Then Exit Function Else Resume Next 'Add ErrHandler 2023/08/01
End Function

'20210908: Modified to merge vbt functions Non_Binning_Pwr_Setting_VT, HVCC_Set_VT, and PostBinCut_Voltage_Set_VT into the vbt function bincut_power_Setting_VT, as discussed with TSMC ZYLINI.
'20210603: Modified to move inst_info.Pattern_Pmode and inst_info.By_Mode from initialize_inst_info to GradeSearch_XXX_VT.
'20210223: Modified to move "Check_and_Decompose_PrePatt_FuncPat" prior to "decide_bincut_feature_for_stepsearch".
'20210207: Modified to move the vbt code of switching DCVS Valt from print_voltage_info_before_FuncPat to GradeSearch_XXX_VT.
'20210126: Modified to revise the vbt code for DevChar.
'20210125: Modified to move "select_DCVS_output_for_powerDomain" prior to "Set_PayloadVoltage_to_DCVS".
'20210122: Modified to check if FuncPat <> "" for print_voltage_info_before_FuncPat.
'20201210: Modified to use run_patt_from_FuncPat_for_BinCut for running the pattern decomposed from FuncPat.
'20201204: Modified to initialize "inst_info.PrePattPass", "inst_info.FuncPatPass", and "inst_info.sitePatPass" in the vbt function initialize_inst_info.
'20201204: Modified to initialize "inst_info.IndexLevelPerSite = -1" in the vbt function initialize_inst_info.
'20201123: Modified to align the format of Judge_PF and Judge_PF_func in the datalog.
'20201118: Modified to use "TheHdw.Digital.Patgen.PatternBurstPassedPerSite" to get siteResult of pattern pass/fail.
'20201111: Modified to use "inst_info.voltage_SelsrmBitCalc".
'20201110: Modified to check if FuncPat <> "".
'20201029: Modified to use inst_info.previousDcvsOutput and inst_info.currentDcvsOutput.
'20201029: Modified to use inst_info.is_BV_Safe_Voltage_printed and inst_info.is_BV_Payload_Voltage_printed.
'20201027: Modified to use "Public Type Instance_Info".
'20201026: Modified to switch DCVS to Vmain prior to PrePatt for TD pattern burst by C651 Toby.
'20201015: Modified to save result of pattern Pass/Fail in "sitePatPass".
'20201012: Modified to change the arguments of the vbt function "check_patt_Pass_Fail".
'20201008: Modified to replace "PrintedBVinDatalog" with "is_BV_Payload_Voltage_printed".
'20200924: Modified to merge the branches of "Calculate_Selsrm_DSSC_For_BinCut".
'20200923: Modified to use "run_patt_offline_simulation" for offline simulation.
'20200923: Modified to call "check_patt_Pass_Fail".
'20200922: Modified the branch to simulate offline random Pass/Fail.
'20200922: Modified to align the branches of running FuncPat for online and offline tests.
'20200921: Modified to replace the PrePatt vbt block with calling "run_prepatt_decompose_VT".
'20200918: Modified to add the argument "result_mode" for the vbt function "Check_and_Decompose_PrePatt_FuncPat".
'20200918: Modified to use "prepare_DCVS_Output_for_RailSwitch".
'20200918: Modified to use "print_voltage_info_before_FuncPat".
'20200908: Modified to remove the redundant site-loop.
'20200903: Modified to align TestNumber from TestFlow table.
'20200821: Modified to add "Dim str_Selsrm_DSSC_Bit(MaxSiteCount-1) As String".
'20200809: Modified to check DCVS output and Payload pattern.
'20200730: Modified to add the EnableWord "VDDBinning_Offline_AllPattPass" for Offline simulation with all patterns pass.
'20200727: Modified to follow the new rule of SELSRM bit calculation proposed by C651 Toby.
'20200724: Modified to check PrePattPass.
'20200724: Modified to use judge_PF_func to update sort Number. and bin out the failed DUT.
'20200707: Modified to print the comment while the argument "EnableBinOut" is not enabled.
'20200630: Modified to remove the unused "thehdw.Utility.Pins("k02,k04").State = tlUtilBitOff".
'20200609: Modified to use Check_alarmFail_before_BinCut_Initial.
'20200526: Modified to simplfy the branch.
'20200520: Modified to use Check_Pattern_NoBurst_NoDecompose to show the errorLogMessage if "burst=no" and "Decompose_Pattern=false".
'20200520: Modified to use "Check_and_Decompose_PrePatt_FuncPat" to check and decompose patsets PrePatt and FuncPat, and find SELSRAM DSSC pattern for DSSC digSrc.
'20200424: Modified to use "Set_BinCut_Initial_by_ApplyLevelsTiming" to set BinCut initial voltage by ApplyLevelsTiming.
'20200324: Modified to skip ApplyLevelsTiming when current instance has the same level/timing as previous instance for project with rail-switch.
'20200320: Modified to check instance contexts of current instance and previous instance.
'20200319: Modified to switch off save_core_power_vddbinning and restore_core_power_vddbinning if Flag_Enable_Rail_Switch = True.
'20200217: Modified to check if no vbump before running the payload pattern.
'20200210: Modified to merge the branches to use PostBinCut_Voltage_Set_VT.
'20200207: Modified to replace set_core_power_main and set_core_power_alt with select_DCVS_output_for_powerDomain.
'20200203: Modified to use the function "print_bincut_power".
'20200130: Modified to call Calculate_Selsrm_DSSC_For_BinCut for SELSRM DSSC bits calculation.
'20200130: Modified to init BinCut_Init_Voltage and BinCut_Payload_Voltage.
'20200130: Modified to call Get_Pmode_Addimode_Testtype_fromInstance to get pmode/addi_mode/testtype.
'20200121: Modified for pattern decomposed.
'20200115: Modified to check if the project with rail-switch.
'20190106: Modified to add "TheHdw.Alarms.Check".
'20191205: Modified to call Find_DsscPatt_fromPattSet to find DSSC pattern.
'20191204: Modified to check if init patts pass or fail...
'20191202: Modified for the revised initVddBinCondition.
'20191127: Modified for the revised InitVddBinTable.
'20191113: Modified to decide BlockType for SELSRM_Mapping_Table.
'20191106: TSMC SWLINZA suggested us to add fail-stop flag (Flag_Vddbinning_Fail_Stop) to avoid no bin-out or no fail-stop in BinTable.
'20191030: Modified for BinOut control.
'20191009: Modified print payload voltages for opensocket and offline simulation with rail-switch
'20191009: Modified print payload voltages for offline simulation with rail-switch
'20190627: Modified to use the global variable "pinGroup_BinCut" for BinCut powerDomains.
'20190618: Modified to add the additional p_mode for print_main_power
'20190617: Modified to use siteDouble "CorePowerStored()" to save/restore voltages for BinCut powerDomains.
'20190615: Modified for offline tests.
'20190611: Modified for set voltages to "DC Specs" (by DcSpecsCategoryForInitPat) for initial and "init patterns" before ApplyLevelsTiming.
'20190606: Modified to add the argument "DcSpecsCategoryForInitPat as string" for Init patterns with the new test setting DC Specs.
'20190319: Modified to add "Flag_Enable_Rail_Switch" to turn VRS Rail Switch on/off.
'20181224: Modified to add Flag_SyncUp_DCVS_Output_enable to control SyncUp on/off.
'20181115: Modified for DSSC TD patt_group (init+payload1+init+payload).
'20181113: Modified for TD payload pattern (print_alt_power_payload).
'20181004: As the request from KTCHAN, we modify the function with the argument to control BinOut(judge_PF_func).
'20180926: Created for retention tests and postBinCut data collection.
Public Function GradeSearch_postBinCut_VT(FuncPat As Pattern, performance_mode As String, result_mode As tlResultMode, DecomposePatt As String, _
                                        FuncTestOnly As Boolean, IDSCurrentLimitList As String, PrePatt As Pattern, _
                                        Optional SpiCounterValue As Integer, Optional RunFailCycle As Boolean, _
                                        Optional Validating_ As Boolean, Optional EnableBinOut As Boolean, _
                                        Optional DcSpecsCategoryForInitPat As String = vbNullString, _
                                        Optional Flag_RetWait As Boolean = False, Optional Dbl_RetWaitTime_mS As Double = 100, _
                                        Optional failpins As String = vbNullString, Optional CaptureSize As Long = 0, _
                                        Optional HarvestBinningFlag As String = vbNullString, _
                                        Optional EnableCoreHarvest As Boolean = False, Optional EnableCoreMask As Boolean = true, _
                                        Optional Harv_CommonFailFlag As String = vbNullString, Optional UserFunction As String = vbNullString)
                                        'Optional Str_RetWait_bfrPat As String = vbNullString)
On Error GoTo errHandler 'Add ErrHandler 2023/08/01
    Dim site As Variant
    Dim inst_info As Instance_Info
    Dim indexPatt As Long
    '''variables specified for project.
    '''======================== Special case of Retention for Sicily ========================'''
    '''20190920: Modifed to addd for Sicily by SWLINZA
    Dim Current_PatName_PostBV As String '''20190920: added Special case of Retention for Sicily by SWLINZA.
    Dim Previous_PatName_PostBV As String
    Dim Flag_DetectBtwWR As Boolean
    Dim instSSNinfo As Inst_SSN
    'Dim RetWait_PatAry() As String
    'Dim RetWait_PatAry_Dic As New dictionary
    '''======================================================================================'''
    glb_TestInstance = vbNullString
    glb_TestInstance = UCase(TheExec.DataManager.instanceName)
'''//==================================================================================================================================================================================//'''
'''//Note:
'''1. GradeSearch_postBinCut_VT can use the argument "EnableBinOut" or failFlag in flow table to control BinOut.
'''
'''<Keyword replacement of BinCut test condition>
'''20180926: Currently C651 Toby didn't define "bin result" or "product-*gb" for postBinCut and retention tests.
'''So that we define the voltage as "VBIN_RESULT(P_mode).Grade". If we get the definition from Toby, we will update this.
'''ToDo: Warning!!! Please discuss this with C651 project DRIs to see if we can use the keyword in the instance names to decide the keyword replacement of BinCut test condition.
'''//==================================================================================================================================================================================//'''
    If Validating_ Then
        If FuncPat.value <> "" Then Call PrLoadPattern(FuncPat.value)
        If PrePatt.value <> "" Then Call PrLoadPattern(PrePatt.value)
        Exit Function    ' Exit after validation
    End If
    
    If PrePatt <> "" Then
        Shmoo_Pattern = FuncPat.value & "," & PrePatt.value
    Else
        Shmoo_Pattern = FuncPat.value
    End If
    
    '''//Initialize inst_info.
    '''//Get p_mode, addi_mode, jobIdx, testtype, and offsettestype from test instance and performance_mode.
    Call initialize_inst_info(inst_info, performance_mode)
    inst_info.selsrm_DigSrc_Pin = "JTAG_TDI"
    inst_info.selsrm_DigSrc_SignalName = "DigSrcSignal"
    '''For Harvest MultiFSTP.
    inst_info.Harvest_Core_DigSrc_Pin = "JTAG_TDI"
    inst_info.Harvest_Core_DigSrc_SignalName = "Harvest_Core_DigSrcSignal"
    
    '20240806 postbv also need below function, or it will be error, because inst_info.digSrcPatterns has no initial value
    Call CheckInstForUserFunction(UserFunction, inst_info.digSrcLabel, inst_info.digSrcPatterns, CStr(FuncPat))
    
    If inst_info.digSrcLabel(0) <> "" Then
        inst_info.MultiFSTP_Enable = True
    Else
        inst_info.MultiFSTP_Enable = False
    End If
    
    '''//Check if DevChar Precondition is tested.
    If inst_info.is_DevChar_Running = True And inst_info.get_DevChar_Precondition = False Then
        Call Set_BinCut_Initial_by_ApplyLevelsTiming(inst_info)
        Exit Function
    End If
    
    '''//Check if alarmFail was triggered prior to BinCut initial(applyLevelsTiming).
    Check_alarmFail_before_BinCut_Initial inst_info.inst_name
    alarmFail = False
    
    '240726 add for pattern harvest and PinGrp harvest
    inst_info.HarvestBinningFlag = HarvestBinningFlag
    inst_info.enable_HarvPinGrp = EnableCoreHarvest
    inst_info.enable_CoreMask = EnableCoreMask
    inst_info.Harv_CommonFailFlag = Harv_CommonFailFlag
    '''20240726 add pin mask or core mask to check
    If (inst_info.enable_HarvPinGrp = True Or inst_info.enable_CoreMask = True) And Flag_HarvPinFlag_Mapping_Table_Parsed = False Then
        Call Print_Error_Message(Error_Info, "VBT_LIB_VDD_Binning", "GradeSearch_VT", "TestInst Argument 'HarvPinGrp_Enable' or 'enable_CoreMask' = True, but can't find HarvPinGrp Table! Error!!")
    End If
    
    '''======================== Special case of Retention for Sicily ========================'''
    '''20190920: Modifed to addd by SWLINZA
    Current_PatName_PostBV = vbNullString
    Previous_PatName_PostBV = vbNullString
    '    RetWait_PatAry() = Split(Str_RetWait_bfrPat, ",")
    '    RetWait_PatAry_Dic.RemoveAll
    '    RetWait_PatAry_Dic.compareMode = 1
    '    For j = 0 To UBound(RetWait_PatAry())
    '        If Not RetWait_PatAry_Dic(RetWait_PatAry(j)).Exists Then
    '            RetWait_PatAry_Dic.Add RetWait_PatAry(j), 1
    '        End If
    '    Next j
    '''======================================================================================'''

    '''//Set initial voltages from category "Bincut_X_X_X" in DC_Specs sheet by ApplyLevelsTiming.
    '''Print the initial voltages, and applies them to DCVS Vmain and Valt by ApplyLevelsTiming (DCVS voltage source will be switched to Vmain).
    Call Set_BinCut_Initial_by_ApplyLevelsTiming(inst_info)

    '''//Decompose PrePatt and FuncPat to check if any DSSC digsrc pattern of SELSRM(defined in SELSRM_Mapping_Table) exists in the pattern sets.
    Call Check_and_Decompose_PrePatt_FuncPat(inst_info, result_mode, DecomposePatt, PrePatt.value, FuncPat.value)
    
    '240726 add
    '''*********************************************************************************'''
    '''  Added for SSN Harvest (Core Diagnosis) & SSN Check (Core Mask) 20240702        '''
    '''*********************************************************************************'''
    If Flag_HarvPinFlag_Mapping_Table_Parsed = True And ssnPatternsDict.Count > 0 Then
        Dim Decomposed_Pat_SSN() As String
        Dim patCnt As Long
        Decomposed_Pat_SSN = TheExec.DataManager.Raw.GetPatternsInSet(FuncPat.value, patCnt)
        Call CheckInstForSSN_BV(Decomposed_Pat_SSN, instSSNinfo, SSNMapping, inst_info.enable_HarvPinGrp, inst_info.enable_CoreMask)
        inst_info.BeSSNTest = instSSNinfo.bSSNTest '240703
    End If
    
    '''//Set the excluded performance mode if the device is bin2 die and the performance mode doesn't exist in bin2 table.
    SkipTestBin2Site inst_info.p_mode, inst_info.Active_site_count
    
    If inst_info.Active_site_count = 0 Then
        RestoreSkipTestBin2Site inst_info.p_mode
        Exit Function
    End If

    '''******************************************************************************************'''
    '''//Set DCVS voltage output to Vmain prior to PrePatt and FuncPat.
    '''******************************************************************************************'''
    select_DCVS_output_for_powerDomain tlDCVSVoltageMain
    inst_info.currentDcvsOutput = tlDCVSVoltageMain
    
    '''//Initialize array of BinCut_Init_Voltage and BinCut_Payload_Voltage before BinCut payload voltages calculation.
    Init_BinCut_Voltage_Array
    
    '''//Calculate BinCut payload voltages of BinCut CorePower and OtherRail.
    '''20210908: Modified to merge vbt functions Non_Binning_Pwr_Setting_VT, HVCC_Set_VT, and PostBinCut_Voltage_Set_VT into the vbt function bincut_power_Setting_VT, as discussed with TSMC ZYLINI.
    Call bincut_power_Setting_VT(inst_info, CurrentPassBinCutNum, BinCut_Payload_Voltage)
    
    '''//If DSSC digsrc pattern of SELSRM exists in the pattern sets, calculate DSSC bits sequence by Selsrm Logic power/SRAMthresh, then prepare DSSC digsrc signal setups.
    Call Calculate_Selsrm_DSSC_For_BinCut(inst_info, VBIN_RESULT(inst_info.p_mode).passBinCut)
    
    '''//Set Payload voltages to DCVS. For projects with Rail-Switch, BinCut payload voltage values are applied to DCVS Valt.
    '''BinCut_Payload_Voltage is the siteDouble array for storing BinCut payload voltage values calculated from PostBinCut_Voltage_Set_VT.
    Set_PayloadVoltage_to_DCVS Flag_Enable_Rail_Switch, pinGroup_BinCut, BinCut_Payload_Voltage
    TheHdw.Wait 0.001
    
    '===================================================================
    ' Print BinCut init voltages for PrePatt, then run PrePatt(init pattern).
    '===================================================================
    Call run_prepatt_decompose_VT(inst_info, inst_info.PrePatt, inst_info.ary_PrePatt_decomposed, inst_info.count_PrePatt_decomposed, inst_info.PrePattPass, DcSpecsCategoryForInitPat)
    
    '''20231207: Added FFC feature for postBincut
    Call check_flag_to_enable_CMEM_collection(inst_info, Flag_Enable_CMEM_Collection, inst_info.count_FuncPat_decomposed, CaptureSize, failpins, True)
    
    '''20231207: Added FFC feature for postBincut
    '''//Clear capture Memory(CMEM) and resize array of CMEM_data if inst_info.enable_CMEM_Collection is enabled.
    Call resize_CMEM_Data_by_pattern_number(inst_info.enable_CMEM_Collection, inst_info.count_FuncPat_decomposed, inst_info.Step_CMEM_Data)
    
    '''20240726 add
    Call check_arg_to_enable_HarvPinGrp(inst_info)
    
    '''//Update siteResult of PrePatt.
    '''20210129: Modified to revise the vbt code for DevChar.
    If inst_info.is_DevChar_Running = False Then
        Call update_Pattern_result_to_PattPass(inst_info.PrePattPass, inst_info.funcPatPass)
    End If

    '''//Check if FuncPat is empty...
    If FuncPat <> "" Then
        '''******************************************************************************************'''
        '''//For Mbist instances in project with rail-switch, set DCVS voltage output to Valt prior to FuncPat.
        '''******************************************************************************************'''
        '''20210207: Modified to move the vbt code of switching DCVS Valt from print_voltage_info_before_FuncPat to GradeSearch_XXX_VT.
        If inst_info.Test_Type = testType.Mbist Then '''ex: "*cpu*bist*", "*gfx*bist*", "*gpu*bist*", "*soc*bist*".
            '''====================================================================================================='''
            '''C651 didn't implement Vbump op-code in MBIST init pattern for project with rail-switch.
            '''So that we have to switch DCVS to Valt by VBT code here before running FuncPat for Mbist instances.
            '''====================================================================================================='''
            If Flag_Enable_Rail_Switch Then '''For projects with Rail Switch, BinCut payload voltages are applied to DCVS Valt.
                select_DCVS_output_for_powerDomain tlDCVSVoltageAlt
                inst_info.currentDcvsOutput = tlDCVSVoltageAlt
            Else '''For conventional projects without Rail Switch, BinCut payload voltages(BV) are applied to DCVS Vmain.
                select_DCVS_output_for_powerDomain tlDCVSVoltageMain
                inst_info.currentDcvsOutput = tlDCVSVoltageMain
            End If
        End If
    
        '''//Print BinCut voltage before running FuncPat.
        '''20210122: Modified to check if FuncPat <> "" for print_voltage_info_before_FuncPat.
        Call print_voltage_info_before_FuncPat(inst_info)
'**********************************************
'@@FuncPat pattern-loop Start
'**********************************************
        For indexPatt = 0 To inst_info.count_FuncPat_decomposed - 1
            '''======================== Special case of Retention for Sicily ========================'''
            '''20190920: Modifed to addd for Sicily by SWLINZA
            Current_PatName_PostBV = inst_info.ary_FuncPat_decomposed(indexPatt)
            
            '''Warning!!! Please check the keywords of patterns for your project before using this.
            If (UCase(Previous_PatName_PostBV) Like "*_1W*" And UCase(Current_PatName_PostBV) Like "*_1R*") _
            Or (UCase(Previous_PatName_PostBV) Like "*GRI1_SRT_JTG_ETR*" And UCase(Current_PatName_PostBV) Like "*GRI1_SRT_JTG_EXT*") Then Flag_DetectBtwWR = True
        
            If Flag_RetWait = True And Flag_DetectBtwWR = True Then
                ''If RetWait_PatAry_Dic(inst_info.ary_FuncPat_decomposed(indexPatt)).Exists Then
                    '----- Retention default Wait time 100 ms ------
                    TheHdw.Wait Dbl_RetWaitTime_mS * 0.001
                    TheExec.Flow.TestLimit Dbl_RetWaitTime_mS, PinName:="Wait_Time", unit:=unitCustom, customUnit:="mSec"
                    TheExec.Datalog.WriteComment "*************************************************"
                    TheExec.Datalog.WriteComment "*print: MbistRetention wait " & Dbl_RetWaitTime_mS & " ms*"
                    TheExec.Datalog.WriteComment "*************************************************"
                ''End If
                Flag_DetectBtwWR = False
            End If
        
            Previous_PatName_PostBV = Current_PatName_PostBV
            '''======================================================================================'''
            
            '''====================================================================================================='''
            '''//Run pattern decomposed from FuncPatt patset, and get siteResult of pattern pass/fail.
            '''Sync up DCVS output and print BinCut payload voltage for projects with Rail Switch for TD instance.
            '''Run the pattern decomposed from FuncPat, and get siteResult of pattern pass/fail.
            '''====================================================================================================='''
            Call run_patt_from_FuncPat_for_BinCut(inst_info, indexPatt, inst_info.ary_FuncPat_decomposed(indexPatt), inst_info.funcPatPass, instSSNinfo, inst_info.idxBlock_Selsrm_FuncPat, CaptureSize, failpins)
        Next indexPatt
'**********************************************
'@@FuncPat pattern-loop End
'**********************************************
        DebugPrintFunc FuncPat.value
    End If '''If FuncPat <> ""

    '''//Check if running FuncPat with "burst=no" and "Decompose_Pattern=false".
    Call Check_Pattern_NoBurst_NoDecompose(inst_info.FuncPat, inst_info.count_FuncPat_decomposed, inst_info.enable_DecomposePatt)
    
    '''========================== For BinOut control by argument "EnableBinOut" ==========================
    '''//GradeSearch_postBinCut_VT can use the argument "EnableBinOut" or failFlag in flow table to control BinOut.
    '''===================================================================================================
    '''20191106: TSMC SWLINZA suggested to add fail-stop to avoid no bin-out or no fail-stop in BinTable.
    '''20210126: Modified to revise the vbt code for DevChar.
    If inst_info.is_DevChar_Running = False Then
        If EnableBinOut Then
            If LCase(inst_info.inst_name) Like "*_bv" Then '''BV instance
                If LCase(inst_info.inst_name) Like "*outsidebincut_bv" Or LCase(inst_info.inst_name) Like "*postbincut_bv" Then
                    judge_PF_func inst_info, inst_info.Test_Type, inst_info.funcPatPass, True
                Else
                    judge_PF_func inst_info, inst_info.Test_Type, inst_info.funcPatPass, False
                End If
            Else '''HBV instance
                For Each site In TheExec.sites
                    If inst_info.funcPatPass(site) = False Then
                        If TheExec.Flow.enableWord("Vddbin_Skip_BinCheck_Enable") = True Then
                            'TheExec.Datalog.WriteComment "Site:" & site & ",Pattern Fails but the result is skipped due to Vddbin_Skip_BinCheck_Enable=True!!"
                        Else
                            TheExec.sites.item(site).testResult = siteFail
                            TheExec.sites.item(site).FlagState(strGlb_Flag_Vddbinning_Fail_Stop) = logicTrue
                        End If
                    End If
                Next site
            End If
        Else
            '''Do nothing, only for data collection
            TheExec.Datalog.WriteComment "Test Instance :" & inst_info.inst_name & ", judge_PF_func is not used. Please check if any device condition for BinOut control exists in the instance flow table!!!"
            TheExec.sites.item(site).TestNumber = TheExec.sites.item(site).TestNumber + 1 'modified for testnumber, 20180927
        End If
    End If
    
    '''20231207: Added FFC feature for postBincut
    If inst_info.enable_CMEM_Collection = True And inst_info.funcPatPass.Any(False) Then
        Call PostTestIPF(inst_info.performance_mode, failpins, inst_info.PrintSize, inst_info.Step_CMEM_Data, inst_info.str_patternPins, inst_info.PrintSize)
        TheHdw.Digital.CMEM.SetCaptureConfig 0, CmemCaptNone '''CmemCaptNone: Capture no cycles.
        TheHdw.Digital.CMEM.CentralFields = tlCMEMNone  '20220512, add for turn off cmem capture
    End If
    
    '==================================================
    'Restore the site which is disabled for bin2 chip
    '==================================================
    RestoreSkipTestBin2Site inst_info.p_mode
    
    '''//Align the format of Judge_PF and Judge_PF_func in the datalog.
    TheExec.Datalog.WriteComment ""
    TheExec.Datalog.WriteComment ""
    TheExec.Datalog.WriteComment vbCrLf
Exit Function
errHandler: 'Add ErrHandler 2023/08/01
    Call Print_Error_Message(Error_Info, "VBT_LIB_VDD_Binning", "GradeSearch_postBinCut_VT") 'Add ErrHandler 2023/08/01
    If AbortTest Then Exit Function Else Resume Next 'Add ErrHandler 2023/08/01
End Function

'''20201126: Modified to remove the redundant vbt code of CMEM datalog setup.
Public Function BV_Init_Datalog_Setup(Optional CaptureSize As Long = 512)
On Error GoTo errHandler 'Add ErrHandler 2023/08/01
    TheExec.Datalog.Setup.DatalogSetup.PartResult = True
    TheExec.Datalog.Setup.DatalogSetup.XYCoordinates = True
    TheExec.Datalog.Setup.DatalogSetup.DisableChannelNumberInPTR = True 'disable channel name to stdf, PE's datalog request -- 20131225, chihome
    TheExec.Datalog.Setup.DatalogSetup.OutputWidth = 0

    TheExec.Datalog.Setup.Shared.ascii.Columns.EnableCustomWidths = True
    TheExec.Datalog.Setup.Shared.ascii.Columns.Parametric.TestName.Width = 60
    TheExec.Datalog.Setup.Shared.ascii.Columns.Functional.TestName.Width = 60
    TheExec.Datalog.Setup.Shared.ascii.Columns.Functional.Pattern.Width = 128
    
    '''//If CMEM is enabled for BinCut, update the setting of Datalog display format.
    If Flag_Enable_CMEM_Collection = True Then
        TheExec.Datalog.Setup.DatalogSetup.SetupStndInfo.FuncDispFormat() = 0 '''0: shortlog
    End If
    
    '''//must need to apply after datalog setup
    TheExec.Datalog.ApplySetup
  
    If TheExec.Flow.enableWord("Vddbinning_OpenSocket") = True Then TheHdw.Digital.pins("All_Digital").DisableCompare = True
Exit Function 'Add ErrHandler 2023/08/01
errHandler: 'Add ErrHandler 2023/08/01
    Call Print_Error_Message(Error_Info, "VBT_LIB_VDD_Binning", "BV_Init_Datalog_Setup") 'Add ErrHandler 2023/08/01
    If AbortTest Then Exit Function Else Resume Next 'Add ErrHandler 2023/08/01
End Function

'''20201126: Modified to remove the redundant vbt code of CMEM datalog setup.
Public Function Restore_BV_DataLog_SetUp()
On Error GoTo errHandler 'Add ErrHandler 2023/08/01
    If Flag_Enable_CMEM_Collection = True Then
        TheExec.Datalog.Setup.DatalogSetup.SetupStndInfo.FuncDispFormat() = 0 '''0: shortlog
    End If
    
    '''must need to apply after datalog setup.
    TheExec.Datalog.ApplySetup
    
    If TheExec.Flow.enableWord("Vddbinning_OpenSocket") = True Then TheHdw.Digital.pins("All_Digital").DisableCompare = False
Exit Function
errHandler: 'Add ErrHandler 2023/08/01
    Call Print_Error_Message(Error_Info, "VBT_LIB_VDD_Binning", "Restore_BV_DataLog_SetUp") 'Add ErrHandler 2023/08/01
    If AbortTest Then Exit Function Else Resume Next 'Add ErrHandler 2023/08/01
End Function

'20210831: Modified to add the TheHdw.Alarms.Check.
'20210721: Modified to run generate_offline_IDS_IGSim_Parallel for offline and opensocket, as requested by TSMC ZYLINI.
'20210708: Modified to check is_BinCutJob_for_StepSearch.
'20210708: Modified to move generate_offline_IDS_IGSim_Parallel from check_IDS to Print_BinCut_config.
'20210629: Modified to print Version_Vdd_Binning_Def in the vbt function Print_BinCut_config.
'20210622: Modified to check if alarmFail is triggered, as requested by TSMC ZQLIN.
'20210621: Modified to move DisableCompare from the vbt function check_IDS to Print_BinCut_config.
'20210528: Modified to check "Flag_Get_column_Monotonicity_Offset" for CheckScript.
'20210420: Modified to remove the unused vbt code of GradeSearchMethod. Once if C651 provided any new definition of GradeSearchMethod, we will revise the vbt code.
'20210409: Modified to check the flag strGlb_Flag_Vddbinning_Interpolation_fail for Interpolation bin out control.
'20210325: Modified to use Flag_Vddbin_DoAll_DebugCollection for TheExec.EnableWord("Vddbin_DoAll_DebugCollection").
'20210322: Modified to decide Flag_Vddbin_COF_StepInheritance by checking TheExec.Flow.EnableWord("Vddbin_COF_StepInheritance").
'20210315: Modified to check the EnableWord "Vddbin_COF_StepInheritance".
'20210308: Modified to check each flag state of the flag Group for Harvest.
'20210219: Modified to check the flags "Flag_Skip_Printing_Safe_Voltage" and "Flag_Skip_Printing_SelSrm_DSSC_Info".
'20210129: Modified to print "bincutJobName" for Mapping_TestJobName_to_BincutJobName, requested by CheckScript.
'20210121: Modified to adjust the printing sequence, request by TSMC ZQLIN.
'20201125: As suggestion from Chihome, modified to clear capture Memory (CMEM) after PostTestIPF of GradeSearch_XXX_VT.
'20201120: Modified to print status of the flag "Flag_use_new_Interpolation_Monotonicity", requested by Autogen team.
'20201103: Modified to check if the tester is not offline or not opensocket.
'20201030: Modified to check the flag "Flag_Only_Check_PV_for_VoltageHeritage".
'20201015: Modified to check the flag "Flag_Vddbin_COF_Instance".
'20200815: Modified to prevent the error of getting current timing mode and min period failed.
'20200810: Modified to clear CMEM.
'20200730: Modified to check the EnableWord "VDDBinning_Offline_AllPattPass" and "Golden_Default".
'20200728: Modified to check the flag "Flag_Using_Payload_Voltage_for_Selsrm_Calc".
'20200717: Modified to check the flag "Flag_Vddbinning_IDS_fail".
'20200629: Modified to check the flag "Flag_Vddbinning_Power_Binning_Fail_Stop".
'20200505: Modified to add "Flag_IDS_Distribution_enable".
'20200415: Modified to move "UpdateDLogColumns_Bincut" from check_IDS into Print_BinCut_config.
'20200227: Modified to print header and footer.
'20200210: Created to print BinCut config from GlobalVariables.
'20200129: Modified to check "DoAll" and "Override Fail-stop" in Run Options.
Public Function Print_BinCut_config(Optional str_flag_Group As String = vbNullString)
On Error GoTo errHandler 'Add ErrHandler 2023/08/01
    Dim site As Variant
    Dim strAry_flag_Group() As String
    Dim idx_flag As Long
    Dim str_site_flagstate As String
'''//==================================================================================================================================================================================//'''
'''//Note:
'''Print the status of BinCut settings(defined in "LIB_Vdd_Binning_GlobalVariable.bas") and flag names.
'''//==================================================================================================================================================================================//'''
    '''//Disable compare pinLevel for OpenSocket since Print_BinCut_config is 1st instance of BinCut.
    '''20210621: Modified to move DisableCompare from the vbt function check_IDS to Print_BinCut_config.
    If TheExec.Flow.enableWord("Vddbinning_OpenSocket") = True Then TheHdw.Digital.pins("All_Digital").DisableCompare = True
    
    If Flag_BinCut_Config_Printed = False Then
        '''//If alarmFail is triggered, print the info in the datalog.
        ''20230823: Remove checking IGXL version condition, requested by Verity
        ''20230824: Added UF for using OnAlarmOccurred alarmfail array
        If glb_TesterType = "UltraFLEXplus" Then
            '''20220913: Added for new IG-XL syntax
            Dim alarmOccurred As New SiteBoolean
            alarmOccurred = False
            alarmOccurred = TheHdw.Alarms.GetAlarmingSites(True)        ''GetAlarmingSites(clearAlarm = True)
            For Each site In TheExec.sites
                If alarmOccurred(site) = True Then
                    TheExec.Datalog.WriteComment "site:" & site & ", alarm was triggered before BinCut test instances!!!"
                End If
            Next site
        Else
            For Each site In TheExec.sites
                If alarmFail(site) = True Then
                    TheExec.Datalog.WriteComment "site:" & site & ", alarmFail was triggered before BinCut test instances!!!"
                End If
            Next site
        End If
        '''//Check if alarmFail(site) is triggered or not before vddbinning.
        '''==============================================================================================='''
        '''This method forces an alarm check. It determines whether alarms are present and reports on them.
        '''This method clears alarms. For this reason, do not use it for monitoring alarms during debugging.
        '''20200106: As per discussion with SWLINZA, he suggested us to add this to check any alarm.
        '''==============================================================================================='''

        ''20231102: Remove conditional compiling
        If glb_TesterType = "Jaguar" Then
            TheHdw.Alarms.Check
        End If
        
        TheExec.Datalog.WriteComment "********************************"
        TheExec.Datalog.WriteComment "*print: " & "BinCut Config" & " start*"
        TheExec.Datalog.WriteComment "******************************"
        '''//TestJob Mapping
        TheExec.Datalog.WriteComment "Mapping_TestJobName_to_BincutJobName" & "=" & UCase(bincutJobName)
        TheExec.Datalog.WriteComment "is_BinCutJob_for_StepSearch" & "=" & CStr(is_BinCutJob_for_StepSearch)
        
        '''//Version of Vdd_Binning_Def (BinCut file)
        TheExec.Datalog.WriteComment "Version_Vdd_Binning_Def" & "=" & Version_Vdd_Binning_Def
        
        '' 20211224: Added to Print Bincut Sheet Name in Print_Bincut_Config (Originally from JadeC/D, 20210202 Jack)
        TheExec.Datalog.WriteComment "Vdd_Binning_Def_appA_1=" & gS_Vdd_Binning_Def_appA_1
        TheExec.Datalog.WriteComment "Vdd_Binning_Def_appA_2=" & gS_Vdd_Binning_Def_appA_2
        TheExec.Datalog.WriteComment "Vdd_Binning_Def_appA_3=" & gS_Vdd_Binning_Def_appA_3
        TheExec.Datalog.WriteComment "Non_Binning_Rail=" & gS_Non_Binning_Rail_sheet
        TheExec.Datalog.WriteComment "EquationVoltages=" & gS_EquationVoltages_sheet
        TheExec.Datalog.WriteComment "Power_Binning_Seq_Sheet=" & gS_PwrSeq_ShtName
        
        '''//Run Options
        TheExec.Datalog.WriteComment "IGXL_RunOptions_DoAll" & "=" & CStr(TheExec.RunOptions.DoAll)
        TheExec.Datalog.WriteComment "IGXL_RunOptions_OverrideFailstop" & "=" & CStr(TheExec.RunOptions.OverrideFailStop)
        
        '''//Run Options/Enable Words
        TheExec.Datalog.WriteComment "Vddbinning_OpenSocket" & "=" & CStr(TheExec.Flow.enableWord("Vddbinning_OpenSocket"))
        TheExec.Datalog.WriteComment "VDDBinning_Offline_AllPattPass" & "=" & CStr(TheExec.Flow.enableWord("VDDBinning_Offline_AllPattPass"))
        TheExec.Datalog.WriteComment "Golden_Default" & "=" & CStr(TheExec.Flow.enableWord("Golden_Default"))
        TheExec.Datalog.WriteComment "Vddbin_DoAll_DebugCollection" & "=" & CStr(Flag_Vddbin_DoAll_DebugCollection)
        TheExec.Datalog.WriteComment "Vddbin_PTE_Debug" & "=" & CStr(TheExec.enableWord("Vddbin_PTE_Debug"))
        TheExec.Datalog.WriteComment "Vddbin_Skip_BinCheck_Enable" & "=" & CStr(TheExec.Flow.enableWord("Vddbin_Skip_BinCheck_Enable"))
                
        '''//Flags in Bin Table
        TheExec.Datalog.WriteComment "Flag_Vddbinning_Fail_Stop" & "=" & strGlb_Flag_Vddbinning_Fail_Stop
        TheExec.Datalog.WriteComment "Flag_Vddbinning_IDS_fail" & "=" & strGlb_Flag_Vddbinning_IDS_fail
        TheExec.Datalog.WriteComment "Flag_Vddbinning_Power_Binning_Fail_Stop" & "=" & strGlb_Flag_Vddbinning_Power_Binning_Fail_Stop
        TheExec.Datalog.WriteComment "Flag_Enable_PowerBinning_Harvest" & "=" & CStr(Flag_Enable_PowerBinning_Harvest)
        TheExec.Datalog.WriteComment "Flag_Vddbinning_Interpolation_fail" & "=" & strGlb_Flag_Vddbinning_Interpolation_fail
        
        '''//Variables from BinCut Tables
        TheExec.Datalog.WriteComment "VddbinningBaseVoltage" & "=" & CStr(VddbinningBaseVoltage)
        TheExec.Datalog.WriteComment "Version_IDS_Distribution" & "=" & Version_IDS_Distribution
        
        '''//Flags of PTE optimization
        TheExec.Datalog.WriteComment "Flag_PrintDcvsShadowVoltage" & "=" & CStr(Flag_PrintDcvsShadowVoltage)
        TheExec.Datalog.WriteComment "Flag_noRestoreVoltageForPrepatt" & "=" & CStr(Flag_noRestoreVoltageForPrepatt)
        TheExec.Datalog.WriteComment "Flag_Skip_ReApplyPayloadVoltageToDCVS" & "=" & CStr(Flag_Skip_ReApplyPayloadVoltageToDCVS)
        TheExec.Datalog.WriteComment "Flag_Skip_Printing_Safe_Voltage" & "=" & CStr(Flag_Skip_Printing_Safe_Voltage)
        TheExec.Datalog.WriteComment "Flag_Skip_Printing_SelSrm_DSSC_Info" & "=" & CStr(Flag_Skip_Printing_SelSrm_DSSC_Info)
        
        '''//Flags for BinCut utilities.
        TheExec.Datalog.WriteComment "Flag_VDD_Binning_Offline" & "=" & CStr(Flag_VDD_Binning_Offline)
        TheExec.Datalog.WriteComment "Flag_Tester_Offline" & "=" & CStr(Flag_Tester_Offline)
        TheExec.Datalog.WriteComment "Flag_Interpolation_enable" & "=" & CStr(Flag_Interpolation_enable)
        TheExec.Datalog.WriteComment "Flag_SelsrmMappingTable_Parsed" & "=" & CStr(Flag_SelsrmMappingTable_Parsed)
        TheExec.Datalog.WriteComment "Flag_Enable_Rail_Switch" & "=" & CStr(Flag_Enable_Rail_Switch)
        TheExec.Datalog.WriteComment "Flag_SyncUp_DCVS_Output_enable" & "=" & CStr(Flag_SyncUp_DCVS_Output_enable)
        TheExec.Datalog.WriteComment "Flag_SyncUp_DCVS_Output_enable_Printing" & "=" & CStr(Flag_SyncUp_DCVS_Output_enable)
        TheExec.Datalog.WriteComment "Flag_Enable_CMEM_Collection" & "=" & CStr(Flag_Enable_CMEM_Collection)
        TheExec.Datalog.WriteComment "Flag_IDS_Distribution_enable" & "=" & CStr(Flag_IDS_Distribution_enable)
        TheExec.Datalog.WriteComment "Flag_IDS_Mapping_enable" & "=" & CStr(Flag_IDS_Mapping_enable)
        '''20211130 Rhodes special case GPU0~GPUX
        TheExec.Datalog.WriteComment "Flag_cross_domain_check_enable" & "=" & CStr(Flag_cross_domain_check_enable)
        TheExec.Datalog.WriteComment "Flag_Remove_Printing_BV_voltages" & "=" & CStr(Flag_Remove_Printing_BV_voltages)
        '''ToDo: Maybe we can remove this later. Discuss this with CheckScript owner...
        TheExec.Datalog.WriteComment "Flag_Using_Payload_Voltage_for_Selsrm_Calc" & "=" & CStr(False) '''real BinCut Payoad voltage(True) or EQN-based voltage without dynamic_offset(False).
        TheExec.Datalog.WriteComment "Flag_use_COFInstance" & "=" & CStr(Flag_Vddbin_COF_Instance)
        TheExec.Datalog.WriteComment "Flag_use_PerEqnLog" & "=" & CStr(Flag_Vddbin_COF_Instance_with_PerEqnLog)
        TheExec.Datalog.WriteComment "Vddbin_COF_StepInheritance" & "=" & CStr(Flag_Vddbin_COF_StepInheritance)
        TheExec.Datalog.WriteComment "Vddbin_COF_StepInheritance_New_Logic" & "=" & CStr(Flag_Vddbin_COF_StepInheritance)
        TheExec.Datalog.WriteComment "Flag_Only_Check_PV_for_VoltageHeritage" & "=" & CStr(Flag_Only_Check_PV_for_VoltageHeritage)
        TheExec.Datalog.WriteComment "Flag_use_new_Interpolation_Monotonicity" & "=" & CStr(True)
        TheExec.Datalog.WriteComment "Flag_New_Selsrm_Vth_enable" & "=" & CStr(True)
        '''20210528: Modified to check "Flag_Get_column_Monotonicity_Offset".
        TheExec.Datalog.WriteComment "Flag_Get_column_Monotonicity_Offset" & "=" & CStr(Flag_Get_column_Monotonicity_Offset)
        TheExec.Datalog.WriteComment "Flag_PowerBinning_Harvest_Bin_FieldName" & "=" & CStr(Flag_PowerBinning_Harvest_Bin_FieldName)
        TheExec.Datalog.WriteComment "Flag_T0TX_ROOM_format" & "=" & CStr(Flag_T0TX_ROOM_format)
        TheExec.Datalog.WriteComment "Flag_T0TX_HOT_format" & "=" & CStr(Flag_T0TX_HOT_format)
        
        '''//Check each flag state of the flag Group for Harvest.
        If str_flag_Group <> "" Then
            strAry_flag_Group = Split(str_flag_Group, ",")
            
            For Each site In TheExec.sites
                str_site_flagstate = "site:" & site
                
                For idx_flag = 0 To UBound(strAry_flag_Group)
                    If TheExec.sites.item(site).FlagState(strAry_flag_Group(idx_flag)) = logicTrue Then
                        str_site_flagstate = str_site_flagstate & "," & strAry_flag_Group(idx_flag) & "=" & "T"
                    Else
                        str_site_flagstate = str_site_flagstate & "," & strAry_flag_Group(idx_flag) & "=" & "F"
                    End If
                Next idx_flag
                
                TheExec.Datalog.WriteComment str_site_flagstate
            Next site
        End If
        
        TheExec.Datalog.WriteComment "******************************"
        TheExec.Datalog.WriteComment "*print: " & "BinCut Config" & " end*"
        TheExec.Datalog.WriteComment "******************************"
        
        '''//Use the flag to control printing the config once for all touchdowns.
        'Flag_BinCut_Config_Printed = True
    End If
    
    '''//Decide the print width of TName
    '''Since Prin_BinCut_config is 1st test instance in Flow_VddBinning, it can update column width of TName in the datalog for the following BinCut instances.
    Call UpdateDLogColumns_Bincut(110)
    
    '''*********************************************************************'''
    '''//Offline simulation.
    '''*********************************************************************'''
    '''//For offline simulation, DC_TEST_IDS might not be tested, so that we need to use the simulated IDS values for BinCut.
    '''20210721: Modified to run generate_offline_IDS_IGSim_Parallel for offline and opensocket, as requested by TSMC ZYLINI.
    If Flag_VDD_Binning_Offline = True Or TheExec.Flow.enableWord("Vddbinning_OpenSocket") = True Then
        generate_offline_IDS_IGSim_Parallel '''For opensocket or offline.
    End If
    
    '''//Align the format of Judge_PF and Judge_PF_func in the datalog.
    TheExec.Datalog.WriteComment ""
    TheExec.Datalog.WriteComment ""
    TheExec.Datalog.WriteComment vbCrLf
Exit Function
errHandler: 'Add ErrHandler 2023/08/01
    Call Print_Error_Message(Error_Info, "VBT_LIB_VDD_Binning", "Print_BinCut_config") 'Add ErrHandler 2023/08/01
    If AbortTest Then Exit Function Else Resume Next 'Add ErrHandler 2023/08/01
End Function

'20210813: Modified to use Set_VBinResult_by_Step for updating PassBin, Pass step, and voltage to VBIN_Result.
'20210812: Modified to remove checking the globalVariable "Flag_Only_Check_PV_for_VoltageHeritage" for the vbt function Set_VBinResult_without_Test.
'20210812: Modified to add the argument "Optional Force_Updating_VbinResult_without_Test As Boolean = False" to the vbt function Set_VBinResult_without_Test, as requested by TSMC ZYLINI.
'20210812: Modified to check if p_mode is tested or not.
'20210730: C651 Toby revised the rule of adjust_VddBinning that "If CP1 has only PCPU bin cut search, then we only need PCPU for PrintOut_VddBinning and Adjust_VddBinning. The other domains can be skipped at CP1."
'20210729: Modified to check "AllBinCut(p_mode).is_for_BinSearch=True" for Set_VBinResult_without_Test.
'20210722: Modified to use VBIN_IDS_ZONE(p_mode).Product_Voltage(Max_IDS_Zone, Max_IDS_Step) and DYNAMIC_VBIN_IDS_ZONE(p_mode).Product_Voltage(Max_IDS_Step) for GradeVDD.
'20210713: Modified to check if the current BinCut is for BinCut search.
'20210304: Modified to set VBIN_RESULT(p_mode).GRADEVDD = VBIN_RESULT(p_mode).GRADE + BinCut(p_mode, VBIN_RESULT(p_mode).passBinCut).CP_GB(VBIN_RESULT(p_mode).Step_in_BinCut).
'20210120: Modified to use VBIN_RESULT(p_mode).step_1stPass_in_IDS_Zone to store the first pass step in Dynamic IDS Zone and find the correspondent PassBinCut number.
'20201005: Modified to add "FIRSTPASSBINCUT(p_mode) = CurrentPassBinCutNum".
'20200425: Modified to change the output format of the interpolated voltage string.
'20200423: Modified to set "VBIN_RESULT(p_mode).tested=True".
'20200422: Created to set "VBIN_RESULT(p_mode)" without test.
Public Function Set_VBinResult_without_Test(Optional Force_Updating_VbinResult_without_Test As Boolean = False)
On Error GoTo errHandler 'Add ErrHandler 2023/08/01
    Dim site As Variant
    Dim p_mode As Integer
    Dim selected_pmode As Integer
    Dim bcType As BINCUT_TYPE
'''//==================================================================================================================================================================================//'''
'''//Note:
'''C651 Toby revised the rule of adjust_VddBinning that "If CP1 has only PCPU bin cut search, then we only need PCPU for PrintOut_VddBinning and Adjust_VddBinning. The other domains can be skipped at CP1.", 20210730.
'''//==================================================================================================================================================================================//'''
    '''//Check if BinCut testJob is for BinCut search.
    '''//The flag "is_BinCutJob_for_StepSearch" is True if any testCondition from the table "Non_Binning_Rail" has the keyword "*Evaluate*Bin*".
    If is_BinCutJob_for_StepSearch = False Then
        '''//Align the format of Judge_PF and Judge_PF_func in the datalog.
        TheExec.Datalog.WriteComment ""
        Exit Function
    End If

    TheExec.Datalog.WriteComment "=============================================="
    TheExec.Datalog.WriteComment "======    " & "Set_VBinResult_without_Test" & "    ======"
    TheExec.Datalog.WriteComment "=============================================="
    TheExec.Datalog.WriteComment "***** Start of Set_VBinResult_without_Test *****"
    
    If Force_Updating_VbinResult_without_Test = True Then
        Call Print_Error_Message(Warning_Info, "LIB_Vdd_Binning_customer", "Set_VBinResult_without_Test", "Warning!!!Force_Updating_VbinResult_without_Test=" & CStr(Force_Updating_VbinResult_without_Test))
    End If
    
    For p_mode = 0 To MaxPerformanceModeCount - 1
        '''//Check if p_mode is for BinCut search.
        If AllBinCut(p_mode).Used = True And AllBinCut(p_mode).is_for_BinSearch = True Then
            For Each site In TheExec.sites
                '''init
                selected_pmode = 0
                
                '''//Check if p_mode is tested or not.
                If VBIN_RESULT(p_mode).tested = False Then
                    '''//Check if p_mode has the allowEqual mode to skip test.
                    '''ex: Site:1,VDD_PCPU_MP009 doesn't need to be tested(no keyword in Non_Binning_Rail), VDD_PCPU_MP009 follows voltage from its Allow_Equal mode: VDD_PCPU_MP008
                    If AllBinCut(p_mode).Allow_Equal <> 0 And AllBinCut(p_mode).Allow_Equal <= cntVddbinPmode Then
                        If VBIN_RESULT(AllBinCut(p_mode).Allow_Equal).tested = True Then
                            selected_pmode = AllBinCut(p_mode).Allow_Equal
                            TheExec.Datalog.WriteComment "site:" & site & "," & VddBinName(p_mode) & " doesn't need to be tested(no keyword in Non_Binning_Rail), " & VddBinName(p_mode) & " follows voltage from its Allow_Equal mode: " & VddBinName(AllBinCut(p_mode).Allow_Equal) ' & "=" & VBIN_RESULT(AllBinCut(p_mode).Allow_Equal).PassBinCut
                        Else
                            Call Print_Error_Message(Warning_Info, "LIB_Vdd_Binning_customer", "Set_VBinResult_without_Test", VddBinName(p_mode) & " can't get VBIN_RESULT from the untested Allow_Equal:" & VddBinName(AllBinCut(p_mode).Allow_Equal) & ". Error!!!")
                            'TheExec.Datalog.WriteComment VddBinName(p_mode) & " can't get VBIN_RESULT from the untested Allow_Equal:" & VddBinName(AllBinCut(p_mode).Allow_Equal) & ". Error!!!"
                            'TheExec.ErrorLogMessage VddBinName(p_mode) & " can't get VBIN_RESULT from the untested Allow_Equal:" & VddBinName(AllBinCut(p_mode).Allow_Equal) & ". Error!!!"    ''20230807: Removed by sanity check
                        End If
                    Else
                        '''20210812: Modified to add the argument "Optional Force_Updating_VbinResult_without_Test As Boolean = False" to the vbt function Set_VBinResult_without_Test, as requested by TSMC ZYLINI.
                        If Force_Updating_VbinResult_without_Test = True Then
                            selected_pmode = p_mode
                            Call Print_Error_Message(Warning_Info, "LIB_Vdd_Binning_customer", "Set_VBinResult_without_Test", "site:" & site & "," & VddBinName(p_mode) & ", p_mode is not tested, but Force_Updating_VbinResult_without_Test overwrites testResult of p_mode for Set_VBinResult_without_Test. Warning!!!")
                        Else
                            Call Print_Error_Message(Error_Info, "VBT_LIB_VDD_Binning", "Set_VBinResult_without_Test", "site:" & site & "," & VddBinName(p_mode) & ", p_mode is not tested. Please check test flow table and instances. Error!!!")
                            'TheExec.ErrorLogMessage "site:" & site & "," & VddBinName(p_mode) & ", p_mode is not tested. Please check test flow table and instances. Error!!!"    ''20230807: Removed by sanity check
                        End If
                    End If
                End If '''If VBIN_RESULT(p_mode).tested = False
                
                If selected_pmode > 0 Then
                    '''//Update PassBin, Pass step, flag"VBIN_Result(p_mode).tested", and voltage to VBIN_Result by the step in Dynamic_IDS_Zone.
                    '''20210813: Modified to use Set_VBinResult_by_Step for updating PassBin, Pass step, and voltage to VBIN_Result.
                    Call Set_VBinResult_by_Step(site, p_mode, VBIN_RESULT(selected_pmode).step_in_IDS_Zone(site))
                    
                    VBIN_RESULT(p_mode).FLAGFAIL = False
                End If '''If selected_pmode > 0
            Next site
        '20211022: Added to support SRAM Voltage calculate with equation
        ElseIf dict_IsCorePower(AllBinCut(p_mode).powerPin) = True And UCase(AllBinCut(p_mode).powerPin) Like "*VDD*_SRAM*" Then
            For Each site In TheExec.sites
                                If VBIN_RESULT(p_mode).tested = False Then
                    TheExec.Datalog.WriteComment "site:" & site & "," & VddBinName(p_mode) & " doesn't need to be tested(no keyword in Non_Binning_Rail), " & VddBinName(p_mode) & " used EQ1 voltage."
                    Call Set_VBinResult_by_Step(site, p_mode, VBIN_RESULT(p_mode).step_in_IDS_Zone(site))
                
                    VBIN_RESULT(p_mode).FLAGFAIL = False
                                End If
             Next site
        End If '''If AllBinCut(p_mode).Used = True And AllBinCut(p_mode).is_for_BinSearch = True
    Next p_mode
    
    TheExec.Datalog.WriteComment "***** End of Set_VBinResult_without_Test *****"
    
    '''//Align the format of Judge_PF and Judge_PF_func in the datalog.
    TheExec.Datalog.WriteComment ""
    TheExec.Datalog.WriteComment ""
    TheExec.Datalog.WriteComment vbCrLf
Exit Function
errHandler: 'Add ErrHandler 2023/08/01
    Call Print_Error_Message(Error_Info, "VBT_LIB_VDD_Binning", "Set_VBinResult_without_Test") 'Add ErrHandler 2023/08/01
    If AbortTest Then Exit Function Else Resume Next 'Add ErrHandler 2023/08/01
End Function
'20230705: Check if doesn't have any site active then bin out. Request by T-Cai.
'20210908: Modified to merge vbt functions Non_Binning_Pwr_Setting_VT, HVCC_Set_VT, and PostBinCut_Voltage_Set_VT into the vbt function bincut_power_Setting_VT, as discussed with TSMC ZYLINI.
'20210903: Modified to merge properties of "Public Type Instance_Step_Control" into Public Type Instance_Info "Public Type Instance_Info".
'20210901: Modified to rename "stepcountMax As Long" as "maxStep As New SiteLong" for Public Type Instance_Info.
'20210901: Modified to rename "StepCount As Long" as "count_Step As New SiteLong" for Public Type Instance_Info.
'20210830: Modified to add the optional argument "Optional HarvestBinningFlag As String" for Harvest in BinCut, as requested by C651 Toby.
'20210824: Modified to rename the vbt function calculate_payload_voltage_for_BV as get_passBin_from_Step.
'20210810: Modified to merge the vbt function Check_anySite_GradeFound into the vbt function Update_VBinResult_by_Step.
'20210706: Modified to use AllBinCut(p_mode).is_for_BinSearch to decide if GradeSearch_XXX_VT is a BinCut search or a functional test.
'20210625: Modified to update HardIP globalVariable of Selsrm DSSC DigSrc for Crete, revised by Carter.
'20210603: Modified to move inst_info.Pattern_Pmode and inst_info.By_Mode from initialize_inst_info to GradeSearch_XXX_VT.
'20210602: Modified the printing sequence for CheckScript. Print alg first, then print the called instance.
'20210305: Modified to add the arguments "step_control As Instance_Step_Control" to the vbt function "StoreCaptureByStep".
'20210305: Modified to add the argument "siteResult" to the vbt function "StoreCapFailcycle".
'20201217: Modified to use the vbt function decide_bincut_feature_for_stepsearch to decide if BinCut features are OK to be enabled for BinCut stepSearch.
'20201211: Modified to use the vbt function "initialize_control_flag_for_step_loop" to initialize control flags from "inst_info" and "step_control" at the beginning of each step in step-loop.
'20201211: Modified to use the vbt function "update_sort_result" to do Judge_PF for binSearch and Judge_PF_func for functional test.
'20201210: Modified to rename the vbt function "calculate_payload_voltage_for_binning_CorePower" as "calculate_payload_voltage_for_BV".
'20201210: Modified to add the vbt functions "Get_PassBinNum_by_Step" and "Non_Binning_Pwr_Setting_VT" into the vbt function calculate_payload_voltage_for_binning_CorePower.
'20201210: Modified to use the arguments "inst_info As Instance_Info" and "step_control As Instance_Step_Control" for judge_PF, update_patt_result_for_COFInstance, Update_PassBinCut_for_GradeNotFound, Decide_NextStep_for_GradeSearch, update_control_flag_for_patt_loop, Check_anySite_GradeFound,
'20201209: Modified to remove the argument "ByRef voltage_SelsrmBitCalc() As SiteDouble" and use "inst_info.voltage_SelsrmBitCalc" for Non_Binning_Pwr_Setting_VT.
'20201207: Modified to use "Dim step_control As Instance_Step_Control".
'20201204: Modified to add the argument "IndexLevelPerSite As SiteLong" for the vbt function Calculate_Binning_CorePower_with_DynamicOffset.
'20201204: Modified to initialize "inst_info.PrePattPass", "inst_info.FuncPatPass", and "inst_info.sitePatPass" in the vbt function initialize_inst_info.
'20201204: Modified to initialize "inst_info.IndexLevelPerSite = -1" in the vbt function initialize_inst_info.
'20201203: Modified to add the argument "enable_CMEM_collection As Boolean" for check_flag_to_enable_CMEM_collection.
'20201203: Modified to revise the vbt code for the undefined testJobs.
'20201202: Modified to add the argument "enable_CMEM_Collection as Boolean" for resize_CMEM_Data_by_pattern_number.
'20201201: Modified to use resize_CMEM_Data_by_pattern_number for CMEM.
'20201201: Modified to update CaptureSize, failpins, and PrintSize for CMEM.
'20201126: As suggestion from Chihome, modified 2-dimensions array "Step_CMEM_Data()" and "BC_CMEM_StoreData()" into 1-dimension array to save memory.
'20201125: As suggestion from Chihome, modified to set TheHdw.Digital.CMEM.CentralFields for initializing CMEM in GradeSearch_XXX_VT.
'20201125: As suggestion from Chihome, modified to clear capture Memory (CMEM) after PostTestIPF.
'20201123: Modified to align the format of Judge_PF and Judge_PF_func in the datalog.
'20201111: Modified to use "check_flag_to_enable_CMEM_collection".
'20201111: Modified to use "inst_info.voltage_SelsrmBitCalc".
'20201111: Modified to replace the vbt function name "print_bincut_power" with "print_bincut_voltage".
'20201103: Modified to move "Dim ids_current As New SiteDouble", "IDS_current_fail As New SiteLong", and "Dim IDS_current_Min As Double" into "Public Type Instance_Info".
'20201103: Modified to move "Dim stepcount As Long" and "Dim stepcountMax As Long" into "Public Type Instance_Info".
'20201029: Modified to use inst_info.is_BV_Safe_Voltage_printed and inst_info.is_BV_Payload_Voltage_printed.
'20201027: Modified to use "Public Type Instance_Info".
'20201015: Modified to check the flag "Flag_Vddbin_COF_Instance".
'20201012: Modified to use "update_Pattern_result_to_PattPass" to update the result of multi-instances.
'20201006: Modified to merge the branches of cp1 and non-cp1.
'20200925: Modified to merge "run_patt_only_CallInstance_VT" of non-cp1 into "GradeSearch_CallInstance_VT" of cp1.
'20200923: Modified to remove "clear_after_patt".
'20200923: Modified to use "update_control_flag_for_patt_loop" to update pdate the status of "AllSiteFailPatt" and "All_Patt_Pass".
'20200923: Modified to remove the unused condition from the branch of "AllSiteFailPatt" and "All_Patt_Pass".
'20200923: Modified to check if alarmFail(site) is triggered or not.
'20200923: Modified to remove "clear_after_patt".
'20200922: Modified to remove the redundant vbt code of "KeepAliveFlag".
'20200921: Discussed "RTOS_bootup_relay" / "KeepAliveFlag" / "spi_ttr_flag" with SWLINZA and PCLINZG. We decided to remove these SPI/RTOS branches because RTOS didn't use pattern test since Cebu/Sicily/Tonga/JC-Chop/Ellis/Bora.
'20200908: For better workaround, we suggested users to copy those called instances with use-limit from HardIP to the position right after BinCut call-instance in BinCut testFlow.
'20200903: Discussed "call Multi-Instances with use-limit" with Chihome. He suggested us to create "Instance Group" for "Multi-Instances with use-limit" to avoid incorrect TestLimitIndex.
'20200903: Modified to align TestNumber from TestFlow table.
'20200901: Modified to remove the unused function "Set_BinCut_Initial_by_ApplyLevelsTiming".
'20200828: Modified to support calling multi-instances.
'20200819: Discussed "pfAlways" issue with Chihome, he found the same case in the offline simulation. He suggested us to ask TER factory for patch or .dll to fix this. But TER factory thought that pfAlways didn't cause this issue..
'20200803: Modified to use "call Non_Binning_Pwr_Setting_VT".
'20200727: Modified to follow the new rule of SELSRM bit calculation proposed by C651 Toby.
'20200713: Modified to remove the argument "IndexLevelPerSite As SiteLong" from Non_Binning_Pwr_Setting_VT by using the function Get_PassBinNum_by_Step.
'20200711: Modified to use the siteDouble array "BinCut_Payload_Voltage" to store BinCut payload voltages.
'20200622: Modified to use Decide_PattPass_by_failFlag.
'20200617: Modified to remove BinCut ApplyLevelsTiming.
'20200617: Modified to check if Overlay is applied.
'20200615: Modified to get dynamic_offset type from the argument "offsetTestTypeIdx As Integer" for judge_PF.
'20200612: Created for "Call Instance".
'20191224: Modified to use ResetPmodePowerforBincut for init BinCut pmode power.
Public Function GradeSearch_CallInstance_VT(performance_mode As String, result_mode As tlResultMode, DecomposePatt As String, FuncTestOnly As Boolean, inst_CallInstance As String, _
                                            Optional Validating_ As Boolean, Optional DcSpecsCategoryForInitPat As String = vbNullString, _
                                            Optional CaptureSize As Long, Optional failpins As String, Optional CollectOnEachStep As Boolean, _
                                            Optional HarvestBinningFlag As String = vbNullString)
On Error GoTo errHandler 'Add ErrHandler 2023/08/01
    Dim site As Variant
    Dim inst_info As Instance_Info
    Dim indexPatt As Long
    '''for binning p_mode
    Dim passBinFromStep As New SiteLong
    '''for testNumber alignment
    Dim Org_Test_Number As Long
    '''Calculate test number for each step
    Dim Test_Number_Per_Step As Long: Test_Number_Per_Step = 0
    '''for control of call Instance
    Dim idx_instance As Integer
    Dim flagName As String
    Dim str_Overlay_for_Bincut As String
    Dim strAry_inst_CallInstance() As String
    Dim str_inst_CallInstance As String
    
    glb_TestInstance = vbNullString
    glb_TestInstance = UCase(TheExec.DataManager.instanceName)
    
'''//==================================================================================================================================================================================//'''
'''//Warning!!!!!!
'''//Read the following instructions before using the function:
'''1. For instance with "Call TheHdw.Patterns(ary_FuncPat_decomposed(indexPatt)).test(pfAlways, 0, result_mode)", pfAlways caused "TheExec.sites(site).LastTestResultRaw" and "TheExec.Flow.LastFlowStepResult" get the incorrect TestReseult.
'''20200819: Discussed "pfAlways" issue with Chihome, he found the same case in the offline simulation. He suggested us to ask TER factory for patch or .dll to fix this.
'''20200901: TER factory thought that pfAlways didn't cause "TheExec.sites(Site).LastTestResultRaw" issue...
'''Workaround: For instance with pfAlways, maybe it can use failFlag or BV_Pass to get testResult about Pass/Fail.
'''2. For Multi-Instances with use-limit, we found that IGXL gave incorrect "testLimitIndex=0" for each instance with use-limit.
'''20200903: Discussed "call Multi-Instances with use-limit" with Chihome. He suggested us to create "Instance Group" for "Multi-Instances with use-limit" to avoid incorrect TestLimitIndex.
'''20200908: For better workaround, we suggested users to copy those called instances with use-limit from HardIP to the position right after BinCut call-instance in BinCut testFlow.
'''//Note:
'''//Call HardIP/RTOS Instance (inst_CallInstance)
'''1. Currently we can't directly access Lo/Hi limits of HardIP/RTOS Instance, so that we have to copy these HardIP/RTOS Instances with use-limit to position right after this BinCut call instance in test flow table.
'''2. Create "Overlay_BV_xxx" for the original HardIP/RTOS Instance.
'''3. Make sure failFlag (ex: F_BV_CALLINST) of the test instance and use-limit exist in the column "Fail" of the test flow.
'''4. Remember to add flag-clear for F_BV_CALLINST into Flow_Table_Main_Init_Flags.
'''5. Remember to check if BV_Pass is used in LIB_HardIP\HardIP_WriteFuncResult.
'''//==================================================================================================================================================================================//'''
    If Validating_ Then
        '    If DqsSwpPat.Value <> "" Then Call PrLoadPattern(DqsSwpPat.Value)
        '    If DqSwpPat.Value <> "" Then Call PrLoadPattern(DqSwpPat.Value)
        Exit Function    ' Exit after validation
    End If
   
    '''init
    strAry_inst_CallInstance = Split(inst_CallInstance, ",")
    
    '''//Initialize inst_info.
    '''//Get p_mode, addi_mode, jobIdx, testtype, and offsettestype from test instance and performance_mode.
    '''//The flag "inst_info.is_BinSearch" is True if testCondition from the table "Non_Binning_Rail" has the keyword "*Evaluate*Bin*".
    Call initialize_inst_info(inst_info, performance_mode)
    inst_info.selsrm_DigSrc_Pin = "JTAG_TDI"
    inst_info.selsrm_DigSrc_SignalName = "DigSrcSignal"
    '''For Harvest MultiFSTP.
    inst_info.Harvest_Core_DigSrc_Pin = "JTAG_TDI"
    inst_info.Harvest_Core_DigSrc_SignalName = "Harvest_Core_DigSrcSignal"
    '''20210830: Modified to add the optional argument "Optional HarvestBinningFlag As String" for Harvest in BinCut, as requested by C651 Toby.
    inst_info.HarvestBinningFlag = HarvestBinningFlag

    '''//Check if DevChar Precondition is tested.
    If inst_info.is_DevChar_Running = True And inst_info.get_DevChar_Precondition = False Then
        Call Print_Error_Message(Error_Info, "VBT_LIB_VDD_Binning", "GradeSearch_CallInstance_VT", inst_info.inst_name & " is used for Characterization, but it doesn't run DevChar Precondition. Error!!!")
                'TheExec.Datalog.WriteComment inst_info.inst_name & " is used for Characterization, but it doesn't run DevChar Precondition. Error!!!"
        'TheExec.ErrorLogMessage inst_info.inst_name & " is used for Characterization, but it doesn't run DevChar Precondition. Error!!!"    ''20230807: Removed by sanity check
        Exit Function
    End If
    
    '''//Set the excluded performance mode if the device is Bin2 die and the performance mode doesn't exist in Bin2 table.
    SkipTestBin2Site inst_info.p_mode, inst_info.Active_site_count
    
    If inst_info.Active_site_count = 0 Then
        RestoreSkipTestBin2Site inst_info.p_mode '''For the performance mode that does not exist in the BinCut table
        Exit Function
    End If
    
    '''//Check if alarmFail was triggered prior to BinCut initial(applyLevelsTiming).
    Check_alarmFail_before_BinCut_Initial inst_info.inst_name
    alarmFail = False
    
    '''********************************************************************************************************************'''
    '''(1)For TestNumber align, the code is assembled with (2)
    '''********************************************************************************************************************'''
    For Each site In TheExec.sites.Active
        Org_Test_Number = TheExec.sites(site).TestNumber
    Next site
    
    '''//According to "inst_info.is_BinSearch", decide inst_info.maxStep for step-loop and find start_voltage(by start_Step in Dynamic_IDS_Zone of the binning p_mode).
    Call decide_binSearch_and_start_voltage(inst_info, FuncTestOnly)
    
    '''//If BinCut testJob is not defined for GradeSearch_XXX_VT, exit the vbt function.
    If inst_info.maxStep = -1 Then
        Call Print_Error_Message(Error_Info, "VBT_LIB_VDD_Binning", "GradeSearch_CallInstance_VT", "It can't get the correct maxStep for step-loop in GradeSearch_CallInstance_VT. Error!!!")
        'TheExec.ErrorLogMessage "It can't get the correct maxStep for step-loop in GradeSearch_CallInstance_VT. Error!!!"    ''20230807: Removed by sanity check
        Exit Function
    End If

    '' Check if doesn't have any site active then bin out. Request by T-Cai. 20230705
        If TheExec.sites.Active.Count = 0 Then
            Call Print_Error_Message(Error_Info, "VBT_LIB_VDD_Binning", "GradeSearch_CallInstance_VT", "Binout because we don't have any active site Warring!!!")
        Exit Function
    End If

    '''//Print info about the called instance in the datalog.
    TheExec.Datalog.WriteComment inst_info.inst_name & ". It uses GradeSearch_CallInstance_VT to call instance: " & inst_CallInstance
    


    '''//Get failFlag from HardIP instcance name
    Call Get_flagName_from_instanceName(inst_info.inst_name, inst_info.p_mode, flagName)
    
    '''//Decide if BinCut features are OK to be enabled for BinCut stepSearch, ex: CMEM_collection, resize inst_info.BC_CMEM_StoreData, and COFInstance.
    Call decide_bincut_feature_for_stepsearch(inst_info, inst_info.count_FuncPat_decomposed, CaptureSize, failpins, inst_CallInstance)
    
'**********************************************
'&& Search Grade Start
'**********************************************
    For inst_info.count_Step = 0 To inst_info.maxStep '''start Vdd binning search, use the full EQN count to loop.
        '''//Initialize control flags from "inst_info" and "step_control" at the beginning of each step in step-loop.
        '''Initialize flags of pattern pass/fail, BV Safe/Payload Voltage printed, and grade_found.
        Call initialize_control_flag_for_step_loop(inst_info)
        
        '''//Initialize array of BinCut_Init_Voltage and BinCut_Payload_Voltage before BinCut payload voltages calculation.
        Init_BinCut_Voltage_Array

        '''//Get passBin by the current step in Dynamic_IDS_Zone of the binning performance mode.
        '''//Update the current step in Dynamic_IDS_Zone if the perfromance mode is interpolated by Interpolation(ReGenerate_IDS_ZONE_Voltage_Per_Site_ver2).
        '''20210824: Modified to rename the vbt function calculate_payload_voltage_for_BV as get_passBin_from_Step.
        Call get_passBin_from_Step(inst_info, passBinFromStep)
        
        '''//Calculate BinCut payload voltages of BinCut CorePower and OtherRail.
        '''//It also calculates BinCut payload voltages with Dynamic_Offset for the binning PowerDomain.
        '''20210908: Modified to merge vbt functions Non_Binning_Pwr_Setting_VT, HVCC_Set_VT, and PostBinCut_Voltage_Set_VT into the vbt function bincut_power_Setting_VT, as discussed with TSMC ZYLINI.
        '''20210908: Modified to add the argument "Optional enable_DynamicOffset As Boolean = False" to calculate BinCut payload voltage of the binning PowerDomain with DynamicOffset.
        Call bincut_power_Setting_VT(inst_info, passBinFromStep, BinCut_Payload_Voltage)
        
        '========================
        '''//Print BinCut voltages
        '========================
        '''******************************************************************************************'''
        '''//For projects with Rail Switch, BinCut voltage(BV) for binning power is applied to DCVS Valt.
        '''//For conventional projects without Rail Switch, BinCut voltage(BV) for binning power is applied to DCVS Vmain.
        '''******************************************************************************************'''
        If Flag_Enable_Rail_Switch Then '''For projects with Rail Switch
            '''Since pattern set without decomposing, it can directly print shadow voltages for payload.
            print_bincut_voltage inst_info, CurrentPassBinCutNum, False, True, BincutVoltageType.PayloadVoltage
        Else '''For conventional projects without Rail Switch
            print_bincut_voltage inst_info, CurrentPassBinCutNum, False, Flag_PrintDcvsShadowVoltage, BincutVoltageType.PayloadVoltage
        End If
        
        '''//Clear capture Memory(CMEM) and resize array of CMEM_data if inst_info.enable_CMEM_Collection is enabled.
        Call resize_CMEM_Data_by_pattern_number(inst_info.enable_CMEM_Collection, inst_info.count_FuncPat_decomposed, inst_info.Step_CMEM_Data)
        
'***********************
'@@Instance-loop Start
'***********************
        For idx_instance = 0 To UBound(strAry_inst_CallInstance)
            str_inst_CallInstance = strAry_inst_CallInstance(idx_instance)
            str_Overlay_for_Bincut = "Overlay_BV_" & str_inst_CallInstance

            '''//Set Payload voltages to Overlay "Overlay_BV_XXX" of DC Specs in inst_CallInstance.
            '''BinCut_Payload_Voltage is the siteDouble array for storing BinCut Payload voltage values calculated from Non_Binning_Pwr_Setting_VT.
            '''Note: Use Overlay to avoid HardIP/RTOS instances doing applyLevelsTiming and ForceCondition to overwrite BinCut payload voltages...
            Set_PayloadVoltage_to_Overlay Flag_Enable_Rail_Switch, pinGroup_BinCut, BinCut_Payload_Voltage, str_Overlay_for_Bincut
            
            '''//Set Flag status initialization for failFlag.
            BV_Pass = True
            
            For Each site In TheExec.sites
                TheExec.sites.item(site).FlagState(flagName) = logicFalse
            Next site

            '''//Only CP1 uses CMEM.
            If inst_info.enable_CMEM_Collection = True Then TheHdw.Digital.CMEM.SetCaptureConfig CaptureSize, CmemCaptFail, tlCMEMCaptureSourcePassFailData

            '''//Call HardIP/RTOS Instance (inst_CallInstance)
            '''Note: Currently we can't directly access Lo/Hi limits of HardIP/RTOS Instance, so that we have to copy these items as use-limit of this BinCut instance into test flow table.
'''//************************************************************************************************************************************************************************************************//'''
'''//Warning!!!!!!
'''1. For instance with "Call TheHdw.Patterns(ary_FuncPat_decomposed(indexPatt)).test(pfAlways, 0, result_mode)", pfAlways caused "TheExec.sites(site).LastTestResultRaw" and "TheExec.Flow.LastFlowStepResult" get the incorrect TestReseult.
'''20200819: Discussed "pfAlways" issue with Chihome, he found the same case in the offline simulation. He suggested us to ask TER factory for patch or .dll to fix this.
'''20200901: TER factory thought that pfAlways didn't cause "TheExec.sites(Site).LastTestResultRaw" issue...
'''//For instance with pfAlways, maybe it can use failFlag or BV_Pass to get testResult about Pass/Fail.
'''
'''2. For Multi-Instances with use-limit, we found that IGXL gave incorrect "testLimitIndex=0" for each instance with use-limit.
'''20200903: Discussed "call Multi-Instances with use-limit" with Chihome. He suggested us to create "Instance Group" for "Multi-Instances with use-limit" to avoid incorrect TestLimitIndex.
'''20200908: For better workaround, we suggested users to copy those called instances with use-limit from HardIP to the position right after BinCut call-instance in BinCut testFlow.
'''//************************************************************************************************************************************************************************************************//'''
            '''20211223: Modified AMPLP5 -> AMP
            '''20210625: Modified to update HardIP globalVariable of Selsrm DSSC DigSrc for Crete, revised by Carter.
            '''//------------------------HardIP AMPLP5 ELB/ILB do SELSRAM DSSC------------------------
            If (LCase(str_inst_CallInstance) Like "*amp*" And LCase(str_inst_CallInstance) Like "*elb*") _
                Or (LCase(str_inst_CallInstance) Like "*amp*" And LCase(str_inst_CallInstance) Like "*ilb*") Then
                glb_BCCallHIPinst_instinfo = inst_info
                glb_BCCallHIPinst_OverlayName = str_Overlay_for_Bincut
                glb_BCCallHIPinst_passBinFromStep = passBinFromStep
            End If
            '''//------------------------HardIP AMPLP5 ELB/ILB do SELSRAM DSSC------------------------
            
            TheExec.Flow.instance(str_inst_CallInstance).Execute
            
            '''//Check if Overlay of BinCut Payload voltages is applied.
            If inst_info.count_Step = 0 Then
                '''The specified overlay "Overlay_BV_XXX" should be applied while calling HardIP/RTOS Instance.
                '''ToDo: Check if failFlag of the instance and use-limit exist...
                If TheExec.Overlays.item(str_Overlay_for_Bincut).IsApplied = True Then
                    '''Note: After calling HardIP/RTOS Instance, remove BinCut Overlay.
                Else
                    Call Print_Error_Message(Error_Info, "VBT_LIB_VDD_Binning", "GradeSearch_CallInstance_VT", "Overlay: " & str_Overlay_for_Bincut & " of DC Specs from " & str_inst_CallInstance & " isn't applied for BinCut instance " & inst_info.inst_name & ". Error!!!")
                    'TheExec.ErrorLogMessage "Overlay: " & str_Overlay_for_Bincut & " of DC Specs from " & str_inst_CallInstance & " isn't applied for BinCut instance " & inst_info.inst_name & ". Error!!!"    ''20230807: Removed by sanity check
                End If
            End If
        
            '''//Decide results of Pattern pass/fail and use-limit by failFlag of the instance and use-limit.
            '''Warning!!! Remember to check if BV_Pass is used in LIB_HardIP\HardIP_WriteFuncResult.
'''//****************************************************************************************************************************//'''
'''//Note:
'''Check if "TheExec.sites(site).LastTestResultRaw=tlResultFail" for HardIP ELB vbt function "Meas_FreqVoltCurr_Universal_func" with "Call TheHdw.Patterns(Pat).Test(pfNever, 0)".
'''Check if flagState("F_BV_CALLINST") for HardIP ELB vbt function with "Call TheHdw.Patterns(Pat).Test(pfAlways, 0)".
'''//Warning!!!
'''"TheExec.Flow.LastFlowStepResult" has issues with "TheHdw.Patterns(Pat).test(pfAlways, 0)".
'''Please contact Teradyne factory/software team for this issue.
'''//****************************************************************************************************************************//'''
            Call Decide_PattPass_by_failFlag(flagName, inst_info.sitePatPass)
            
            '''//Only test instances for BinCut search can use CMEM.
'            If inst_info.enable_CMEM_Collection = True Then
            Dim siteResult As New SiteBoolean
            siteResult = TheHdw.Digital.Patgen.PatternBurstPassedPerSite
            If inst_info.sitePatPass.Any(False) And inst_info.enable_CMEM_Collection = True Then
                If glb_TesterType = "UltraFLEXplus" Then        ''UFP use pattern pin
                    ''Get pattern pins, save in inst_info
                    'Call GetPatternPinStr(str_pattern, inst_info.str_patternPins(indexPatt))
                    ''Use pattern pins for CMEM capture
                    Call StoreCapFailcycle_UFP(inst_info.sitePatPass, failpins, indexPatt, CaptureSize, inst_info.Step_CMEM_Data)
                Else    ''UltraFlex uses failpins = "All_Digital" for CMEM capture
                    Call StoreCapFailcycle_UF(inst_info.sitePatPass, failpins, indexPatt, CaptureSize, inst_info.Step_CMEM_Data)
                End If
            End If
 
            '''//Check if all called instances pass or fail, then update the result of multi-instances to "FuncPatPass".
            Call update_Pattern_result_to_PattPass(inst_info.sitePatPass, inst_info.funcPatPass)
            
            '''//After calling HardIP/RTOS Instance, remove BinCut Overlay.
            Call Remove_PayloadVoltage_from_Overlay(str_Overlay_for_Bincut)
            
            '''20211223: Modified AMPLP5 -> AMP
            '''20210625: Modified to update HardIP globalVariable of Selsrm DSSC DigSrc for Crete, revised by Carter.
            '''//------------------------HardIP AMPLP5 ELB/ILB do SELSRAM DSSC------------------------
            If (LCase(str_inst_CallInstance) Like "*amp*" And LCase(str_inst_CallInstance) Like "*elb*") _
                Or (LCase(str_inst_CallInstance) Like "*amp*" And LCase(str_inst_CallInstance) Like "*ilb*") Then
                glb_BCCallHIPinst_OverlayName = vbNullString
                glb_BCCallHIPinst_passBinFromStep = 0
            End If
            '''//------------------------HardIP AMPLP5 ELB/ILB do SELSRAM DSSC------------------------
        Next idx_instance
        
        ''Calculate test number per EQN, only calculate one time
        If Test_Number_Per_Step = 0 Then
            For Each site In TheExec.sites.Active '20240126 avoid any site binout before
                Test_Number_Per_Step = TheExec.sites(site).LastTestNumber - Org_Test_Number
                Exit For
            Next site
        End If
'***********************
'@@Instance-loop End
'***********************
        '''//If CMEM_collection is enabled, collect the data of the failed pattern for current step from "inst_info.Step_CMEM_Data" to "inst_info.BC_CMEM_StoreData".
        If inst_info.enable_CMEM_Collection = True And inst_info.AllSiteFailPatt > 0 Then
            Call StoreCaptureByStep(inst_info, inst_info.Step_CMEM_Data, inst_info.BC_CMEM_StoreData)
            If CollectOnEachStep = True Then Call PostTestIPF(inst_info.performance_mode, failpins, inst_info.PrintSize, inst_info.Step_CMEM_Data, inst_info.str_patternPins)
        End If
        
        If inst_info.is_BinSearch = True Then
            '''//Update the status of "AllSiteFailPatt" and "All_Patt_Pass".
            Call update_control_flag_for_patt_loop(inst_info, inst_info.funcPatPass)
        
            '''//According to the current BinCut step in Dynamic_IDS_zone of p_mode, update PassBin and BinCut voltage(Grade) to VBIN_Result of p_mode for the pass DUT.
            '''//Check if any site has found BinCut pass Grade (based on BinCut step).
            Call Update_VBinResult_by_Step(inst_info)
            
            '''================================================================================================================================================='''
            '''If Grade_Found_Mask = All_Site_Mask (All sites had found BinCut Grade) or On_StopVoltage_Mask = Grade_Not_Found_Mask => exit the step-loop.
            '''ex: The sites didn't find BinCut Grade, but the sites had reached the stopVoltage (step_Stop in Dynamic_IDS_zone), no chance to find the grade!!!
            '''================================================================================================================================================='''
            If inst_info.Grade_Found_Mask = inst_info.All_Site_Mask Or inst_info.On_StopVoltage_Mask = inst_info.Grade_Not_Found_Mask Then
                Exit For '''Exit for "Next StepCount"
            End If
            
            '''//Decide next step in DYNAMIC_IDS_ZONE of p_mode for GradeSearch.
            Call Decide_NextStep_for_GradeSearch(inst_info)
        End If
    Next inst_info.count_Step
'**********************************************
'&& Search Grade End
'**********************************************
    '''//Align testNumber and do judge_PF for binSearch; judge_PF_func for functional test.
    Call update_sort_result(inst_info, inst_info.funcPatPass, Org_Test_Number, failpins, CollectOnEachStep, Test_Number_Per_Step)
    
    '''//Clear the failFlag prior to the next test instance
    For Each site In TheExec.sites
        TheExec.sites.item(site).FlagState(flagName) = logicFalse
    Next site
    
    '''//Align the format of Judge_PF and Judge_PF_func in the datalog.
    TheExec.Datalog.WriteComment ""
    TheExec.Datalog.WriteComment ""
    TheExec.Datalog.WriteComment vbCrLf
Exit Function
errHandler: 'Add ErrHandler 2023/08/01
    Call Print_Error_Message(Error_Info, "VBT_LIB_VDD_Binning", "GradeSearch_CallInstance_VT") 'Add ErrHandler 2023/08/01
    If AbortTest Then Exit Function Else Resume Next 'Add ErrHandler 2023/08/01
End Function

'20210908: Modified to merge vbt functions Non_Binning_Pwr_Setting_VT, HVCC_Set_VT, and PostBinCut_Voltage_Set_VT into the vbt function bincut_power_Setting_VT, as discussed with TSMC ZYLINI.
'20210625: Modified to update HardIP globalVariable of Selsrm DSSC DigSrc for Crete, revised by Carter.
'20210603: Modified to move inst_info.Pattern_Pmode and inst_info.By_Mode from initialize_inst_info to GradeSearch_XXX_VT.
'20210602: Modified the printing sequence for CheckScript. Print alg first, then print the called instance.
'20210126: Modified to revise the vbt code for DevChar.
'20201204: Modified to initialize "inst_info.IndexLevelPerSite = -1" in the vbt function initialize_inst_info.
'20201123: Modified to align the format of Judge_PF and Judge_PF_func in the datalog.
'20201111: Modified to use "inst_info.voltage_SelsrmBitCalc".
'20201111: Modified to replace the vbt function name "print_bincut_power" with "print_bincut_voltage".
'20201029: Modified to use inst_info.is_BV_Safe_Voltage_printed and inst_info.is_BV_Payload_Voltage_printed.
'20201027: Modified to use "Public Type Instance_Info".
'20201012: Modified to use "update_Pattern_result_to_PattPass" to update the result of multi-instances.
'20200923: Modified to check if alarmFail(site) is triggered or not.
'20200921: Discussed "RTOS_bootup_relay" / "KeepAliveFlag" / "spi_ttr_flag" with SWLINZA and PCLINZG. We decided to remove these SPI/RTOS branches because RTOS didn't use pattern test since Cebu/Sicily/Tonga/JC-Chop/Ellis/Bora.
'20200908: For better workaround, we suggested users to copy those called instances with use-limit from HardIP to the position right after BinCut call-instance in BinCut testFlow.
'20200903: Discussed "call Multi-Instances with use-limit" with Chihome. He suggested us to create "Instance Group" for "Multi-Instances with use-limit" to avoid incorrect TestLimitIndex.
'20200903: Modified to align TestNumber from TestFlow table.
'20200901: Modified to remove the unused function "Set_BinCut_Initial_by_ApplyLevelsTiming".
'20200828: Modified to support calling multi-instances.
'20200819: Discussed "pfAlways" issue with Chihome, he found the same case in the offline simulation. He suggested us to ask TER factory for patch or .dll to fix this. But TER factory thought that pfAlways didn't cause this issue..
'20200727: Modified to follow the new rule of SELSRM bit calculation proposed by C651 Toby.
'20200711: Modified to use the siteDouble array to store BinCut payload voltages
'20200622: Modified to use Decide_PattPass_by_failFlag.
'20200622: Modified for the flag naming rule for the failFlag of Call Instance.
'20200617: Modified to remove BinCut ApplyLevelsTiming.
'20200617: Modified to check if Overlay is applied.
'20200615: Created for "Call Instance".
Public Function GradeSearch_HVCC_CallInstance_VT(performance_mode As String, result_mode As tlResultMode, DecomposePatt As String, FuncTestOnly As Boolean, inst_CallInstance As String, _
                                                Optional Validating_ As Boolean, Optional DcSpecsCategoryForInitPat As String = vbNullString)
On Error GoTo errHandler 'Add ErrHandler 2023/08/01
    Dim site As Variant
    Dim inst_info As Instance_Info
    '''for control of call Instance
    Dim idx_instance As Integer
    Dim flagName As String
    Dim str_Overlay_for_Bincut As String
    Dim strAry_inst_CallInstance() As String
    Dim str_inst_CallInstance As String
    
    glb_TestInstance = vbNullString
    glb_TestInstance = UCase(TheExec.DataManager.instanceName)
    
'''//==================================================================================================================================================================================//'''
'''//Warning!!!!!!
'''//Read the following instructions before using the function:
'''1. For instance with "Call TheHdw.Patterns(ary_FuncPat_decomposed(indexPatt)).test(pfAlways, 0, result_mode)", pfAlways caused "TheExec.sites(site).LastTestResultRaw" and "TheExec.Flow.LastFlowStepResult" get the incorrect TestReseult.
'''20200819: Discussed "pfAlways" issue with Chihome, he found the same case in the offline simulation. He suggested us to ask TER factory for patch or .dll to fix this.
'''20200901: TER factory thought that pfAlways didn't cause "TheExec.sites(Site).LastTestResultRaw" issue...
'''Workaround: For instance with pfAlways, maybe it can use failFlag or BV_Pass to get testResult about Pass/Fail.
'''2. For Multi-Instances with use-limit, we found that IGXL gave incorrect "testLimitIndex=0" for each instance with use-limit.
'''20200903: Discussed "call Multi-Instances with use-limit" with Chihome. He suggested us to create "Instance Group" for "Multi-Instances with use-limit" to avoid incorrect TestLimitIndex.
'''20200908: For better workaround, we suggested users to copy those called instances with use-limit from HardIP to the position right after BinCut call-instance in BinCut testFlow.
'''//Note:
'''//Call HardIP/RTOS Instance (inst_CallInstance)
'''1. Currently we can't directly access Lo/Hi limits of HardIP/RTOS Instance, so that we have to copy these HardIP/RTOS Instances with use-limit to position right after this BinCut call instance in test flow table.
'''2. Create "Overlay_BV_xxx" for the original HardIP/RTOS Instance.
'''3. Make sure failFlag (ex: F_BV_CALLINST) of the test instance and use-limit exist in the column "Fail" of the test flow.
'''4. Remember to add flag-clear for F_BV_CALLINST into Flow_Table_Main_Init_Flags.
'''5. Remember to check if BV_Pass is used in LIB_HardIP\HardIP_WriteFuncResult.
'''//==================================================================================================================================================================================//'''
    If Validating_ Then
        '    If DqsSwpPat.Value <> "" Then Call PrLoadPattern(DqsSwpPat.Value)
        '    If DqSwpPat.Value <> "" Then Call PrLoadPattern(DqSwpPat.Value)
        Exit Function    ' Exit after validation
    End If
    
    '''20200903: Modified to align TestNumber from TestFlow table.
    For Each site In TheExec.sites.Active
        TheExec.sites(site).TestNumber = TheExec.sites(site).TestNumber
    Next site
    
    '''init
    strAry_inst_CallInstance = Split(inst_CallInstance, ",")
    
    '''//Initialize inst_info.
    '''//Get p_mode, addi_mode, jobIdx, testtype, and offsettestype from test instance and performance_mode.
    Call initialize_inst_info(inst_info, performance_mode)
    inst_info.selsrm_DigSrc_Pin = "JTAG_TDI"
    inst_info.selsrm_DigSrc_SignalName = "DigSrcSignal"
    '''For Harvest MultiFSTP.
    inst_info.Harvest_Core_DigSrc_Pin = "JTAG_TDI"
    inst_info.Harvest_Core_DigSrc_SignalName = "Harvest_Core_DigSrcSignal"
    
    '''//Check if DevChar Precondition is tested.
    If inst_info.is_DevChar_Running = True And inst_info.get_DevChar_Precondition = False Then
        Call Print_Error_Message(Error_Info, "VBT_LIB_VDD_Binning", "GradeSearch_HVCC_CallInstance_VT", inst_info.inst_name & " is used for Characterization, but it doesn't run DevChar Precondition. Error!!!")
        'TheExec.Datalog.WriteComment inst_info.inst_name & " is used for Characterization, but it doesn't run DevChar Precondition. Error!!!"
        'TheExec.ErrorLogMessage inst_info.inst_name & " is used for Characterization, but it doesn't run DevChar Precondition. Error!!!"    ''20230807: Removed by sanity check
        Exit Function
    End If
    
    '''//Get failFlag from HardIP instcance name
    Call Get_flagName_from_instanceName(inst_info.inst_name, inst_info.p_mode, flagName)
    
    '''//Set the excluded performance mode if the device is bin2 die and the performance mode doesn't exist in bin2 table.
    SkipTestBin2Site inst_info.p_mode, inst_info.Active_site_count
    
    If inst_info.Active_site_count = 0 Then
        RestoreSkipTestBin2Site inst_info.p_mode
        Exit Function
    End If
    
    '''//Check if alarmFail was triggered prior to BinCut initial(applyLevelsTiming).
    Check_alarmFail_before_BinCut_Initial inst_info.inst_name
    alarmFail = False
    
    '''//Print info about the called instance in the datalog.
    TheExec.Datalog.WriteComment inst_info.inst_name & ". It uses GradeSearch_HVCC_CallInstance_VT to call instance: " & inst_CallInstance

    '''//Initialize array of BinCut_Init_Voltage and BinCut_Payload_Voltage before BinCut payload voltages calculation.
    Init_BinCut_Voltage_Array
    
    '''//Calculate BinCut payload voltages of BinCut CorePower and OtherRail.
    '''20210908: Modified to merge vbt functions Non_Binning_Pwr_Setting_VT, HVCC_Set_VT, and PostBinCut_Voltage_Set_VT into the vbt function bincut_power_Setting_VT, as discussed with TSMC ZYLINI.
    Call bincut_power_Setting_VT(inst_info, CurrentPassBinCutNum, BinCut_Payload_Voltage)
    
    '========================
    '''//Print BinCut voltages
    '========================
    '''******************************************************************************************'''
    '''//For projects with Rail Switch, BinCut voltage(BV) for binning power is applied to DCVS Valt.
    '''//For conventional projects without Rail Switch, BinCut voltage(BV) for binning power is applied to DCVS Vmain.
    '''******************************************************************************************'''
    If Flag_Enable_Rail_Switch Then '''For projects with Rail Switch
        '''Since pattern set without decomposing, it can directly print shadow voltages for payload.
        print_bincut_voltage inst_info, , Flag_Remove_Printing_BV_voltages, True, BincutVoltageType.PayloadVoltage
    Else '''For conventional projects without Rail Switch
        print_bincut_voltage inst_info, , Flag_Remove_Printing_BV_voltages, Flag_PrintDcvsShadowVoltage, BincutVoltageType.PayloadVoltage
    End If
        
'***********************
'@@Instance-loop Start
'***********************
    For idx_instance = 0 To UBound(strAry_inst_CallInstance)
        str_inst_CallInstance = strAry_inst_CallInstance(idx_instance)
        str_Overlay_for_Bincut = "Overlay_BV_" & str_inst_CallInstance
            
        '''//Set Payload voltages to Overlay "Overlay_BV_XXX" of DC Specs in inst_CallInstance.
        '''BinCut_Payload_Voltage is the siteDouble array for storing BinCut Payload voltage values calculated from HVCC_Set_VT.
        '''Note: Use Overlay to avoid HardIP applyLevelsTiming and ForceCondition to overwrite BinCut payload voltages...
        Set_PayloadVoltage_to_Overlay Flag_Enable_Rail_Switch, pinGroup_BinCut, BinCut_Payload_Voltage, str_Overlay_for_Bincut
        
        '''//Set Flag status initialization for failFlag.
        BV_Pass = True
        
        For Each site In TheExec.sites
            TheExec.sites.item(site).FlagState(flagName) = logicFalse
        Next site
    
        '''//Call HardIP/RTOS Instance (inst_CallInstance)
        '''Note: Currently we can't directly access Lo/Hi limits of HardIP/RTOS Instance, so that we have to copy these items as use-limit of this BinCut instance into test flow table.
'''//************************************************************************************************************************************************************************************************//'''
'''//Warning!!!!!!
'''1. For instance with "Call TheHdw.Patterns(ary_FuncPat_decomposed(indexPatt)).test(pfAlways, 0, result_mode)", pfAlways caused "TheExec.sites(site).LastTestResultRaw" and "TheExec.Flow.LastFlowStepResult" get the incorrect TestReseult.
'''20200819: Discussed "pfAlways" issue with Chihome, he found the same case in the offline simulation. He suggested us to ask TER factory for patch or .dll to fix this.
'''20200901: TER factory thought that pfAlways didn't cause "TheExec.sites(Site).LastTestResultRaw" issue...
'''//For instance with pfAlways, maybe it can use failFlag or BV_Pass to get testResult about Pass/Fail.
'''
'''2. For Multi-Instances with use-limit, we found that IGXL gave incorrect "testLimitIndex=0" for each instance with use-limit.
'''20200903: Discussed "call Multi-Instances with use-limit" with Chihome. He suggested us to create "Instance Group" for "Multi-Instances with use-limit" to avoid incorrect TestLimitIndex.
'''20200908: For better workaround, we suggested users to copy those called instances with use-limit from HardIP to the position right after BinCut call-instance in BinCut testFlow.
'''//************************************************************************************************************************************************************************************************//'''
        '''20211223: Modified AMPLP5 -> AMP
        '''20210625: Modified to update HardIP globalVariable of Selsrm DSSC DigSrc for Crete, revised by Carter.
        '''//------------------------HardIP AMPLP5 ELB/ILB do SELSRAM DSSC------------------------
        If (LCase(str_inst_CallInstance) Like "*amp*" And LCase(str_inst_CallInstance) Like "*elb*") _
            Or (LCase(str_inst_CallInstance) Like "*amp*" And LCase(str_inst_CallInstance) Like "*ilb*") Then
            glb_BCCallHIPinst_instinfo = inst_info
            glb_BCCallHIPinst_OverlayName = str_Overlay_for_Bincut
            glb_BCCallHIPinst_passBinFromStep = CurrentPassBinCutNum
        End If
        '''//------------------------HardIP AMPLP5 ELB/ILB do SELSRAM DSSC------------------------

        TheExec.Flow.instance(str_inst_CallInstance).Execute
        
        '''//Check if Overlay of BinCut Payload voltages is applied.
        '''The specified overlay "Overlay_BV_XXX" should be applied while calling HardIP/RTOS Instance.
        '''ToDo: Check if failFlag of the instance and use-limit exist...
        If TheExec.Overlays.item(str_Overlay_for_Bincut).IsApplied = True Then
            '''Note: After calling HardIP/RTOS Instance, remove BinCut Overlay.
        Else
            Call Print_Error_Message(Error_Info, "VBT_LIB_VDD_Binning", "GradeSearch_HVCC_CallInstance_VT", "Overlay: " & str_Overlay_for_Bincut & " of DC Specs from " & str_inst_CallInstance & " isn't applied for BinCut instance " & inst_info.inst_name & ". Error!!!")
            'TheExec.ErrorLogMessage "Overlay: " & str_Overlay_for_Bincut & " of DC Specs from " & str_inst_CallInstance & " isn't applied for BinCut instance " & inst_info.inst_name & ". Error!!!"    ''20230807: Removed by sanity check
        End If
        
        '''//Decide results of Pattern pass/fail and use-limit by failFlag of the instance and use-limit.
'''//****************************************************************************************************************************//'''
'''//Note:
'''Check if "TheExec.sites(site).LastTestResultRaw=tlResultFail" for HardIP ELB vbt function "Meas_FreqVoltCurr_Universal_func" with "Call TheHdw.Patterns(Pat).Test(pfNever, 0)".
'''Check if flagState("F_BV_CALLINST") for HardIP ELB vbt function with "Call TheHdw.Patterns(Pat).Test(pfAlways, 0)".
'''//Warning!!!
'''"TheExec.Flow.LastFlowStepResult" has issues with "TheHdw.Patterns(Pat).test(pfAlways, 0)".
'''Please contact Teradyne factory/software team for this issue.
'''//****************************************************************************************************************************//'''
        Call Decide_PattPass_by_failFlag(flagName, inst_info.sitePatPass)
        
        '''//Check if all called instances pass or fail, then update the result of multi-instances to "FuncPatPass".
        Call update_Pattern_result_to_PattPass(inst_info.sitePatPass, inst_info.funcPatPass)
        
        '''//After calling HardIP/RTOS Instance, remove BinCut Overlay.
        Call Remove_PayloadVoltage_from_Overlay(str_Overlay_for_Bincut)

        '''20211223: Modified AMPLP5 -> AMP
        '''20210625: Modified to update HardIP globalVariable of Selsrm DSSC DigSrc for Crete, revised by Carter.
        '''//------------------------HardIP AMPLP5 ELB/ILB do SELSRAM DSSC------------------------
        If (LCase(str_inst_CallInstance) Like "*amp*" And LCase(str_inst_CallInstance) Like "*elb*") _
            Or (LCase(str_inst_CallInstance) Like "*amp*" And LCase(str_inst_CallInstance) Like "*ilb*") Then
            glb_BCCallHIPinst_OverlayName = vbNullString
            glb_BCCallHIPinst_passBinFromStep = 0
        End If
        '''//------------------------HardIP AMPLP5 ELB/ILB do SELSRAM DSSC------------------------
    Next idx_instance
'***********************
'@@Instance-loop End
'***********************

    '''//Bin out the failed DUT if "PattPass = False"...
    '''20191106: TSMC SWLINZA suggested to add fail-stop to avoid no bin-out or no fail-stop in BinTable.
    If inst_info.is_DevChar_Running = False Then '''for DevChar.
        For Each site In TheExec.sites
            If inst_info.funcPatPass(site) = False Then
                If TheExec.Flow.enableWord("Vddbin_Skip_BinCheck_Enable") = True Then
                    'TheExec.Datalog.WriteComment "Site:" & site & ",Pattern Fails but the result is skipped due to Vddbin_Skip_BinCheck_Enable=True!!"
                Else
                    TheExec.sites.item(site).testResult = siteFail
                    TheExec.sites.item(site).FlagState(strGlb_Flag_Vddbinning_Fail_Stop) = logicTrue
                End If
            End If
        Next site
    End If
    
    '''//Clear the failFlag prior to the next test instance
    For Each site In TheExec.sites
        TheExec.sites.item(site).FlagState(flagName) = logicFalse
    Next site
    
    '==================================================
    'Restore the site which is disabled for bin2 chip
    '==================================================
    RestoreSkipTestBin2Site inst_info.p_mode
    
    '''//Align the format of Judge_PF and Judge_PF_func in the datalog.
    TheExec.Datalog.WriteComment ""
    TheExec.Datalog.WriteComment ""
    TheExec.Datalog.WriteComment vbCrLf
Exit Function
errHandler: 'Add ErrHandler 2023/08/01
    Call Print_Error_Message(Error_Info, "VBT_LIB_VDD_Binning", "GradeSearch_HVCC_CallInstance_VT") 'Add ErrHandler 2023/08/01
    If AbortTest Then Exit Function Else Resume Next 'Add ErrHandler 2023/08/01
End Function

'20210726: Modified to add the argument "bincutNum As Long" to the vbt function Adjust_Multi_PassBinCut_Per_Site.
'20201027: Created to overwrite CurrentPassBinCutNum for Harvest Core.
Public Function Overwrite_PassBinNum_by_ForcedBin(Optional enableForecedBin As Boolean = False, Optional BinNumber As Long = 0)
On Error GoTo errHandler 'Add ErrHandler 2023/08/01
    Dim site As Variant
    Dim p_mode As Integer
    If enableForecedBin = True Then
        TheExec.Datalog.WriteComment "=============================================="
        TheExec.Datalog.WriteComment "======    " & "Overwrite_PassBinNum_by_ForcedBin" & "    ======"
        TheExec.Datalog.WriteComment "=============================================="
        For Each site In TheExec.sites
            If CurrentPassBinCutNum(site) < BinNumber Then
                CurrentPassBinCutNum(site) = BinNumber
            
                For p_mode = 0 To MaxPerformanceModeCount - 1
                    If AllBinCut(p_mode).Used = True Then
                        '''20210726: Modified to add the argument "bincutNum As Long" to the vbt function Adjust_Multi_PassBinCut_Per_Site.
                        Adjust_Multi_PassBinCut_Per_Site p_mode, site, CurrentPassBinCutNum(site)
                    End If
                Next p_mode
            End If
        Next site
    End If
Exit Function
errHandler: 'Add ErrHandler 2023/08/01
    Call Print_Error_Message(Error_Info, "VBT_LIB_VDD_Binning", "Overwrite_PassBinNum_by_ForcedBin") 'Add ErrHandler 2023/08/01
    If AbortTest Then Exit Function Else Resume Next 'Add ErrHandler 2023/08/01
End Function

'20210531: Modified to print the header for MultiFSTP by save_siteMask_for_MultiFSTP.
'20210530: Modified to use gb_sitePassBin_original to save CurrentPassBinCutNum before MultiFSTP instances.
'20210528: Modified to update TheExec.sites.Selected.
'20210525: Created to save current siteMask for MultiFSTP in CP1.
Public Function save_siteMask_for_MultiFSTP(Optional str_printMsg As String = vbNullString)
On Error GoTo errHandler 'Add ErrHandler 2023/08/01
    '''//Save CurrentPassBinCutNum into gb_sitePassBin_original before MultiFSTP instances.
    gb_sitePassBin_original = CurrentPassBinCutNum
    
    '''//Save current sites selected into globalVariable before MultiFSTP.
    gb_siteMask_original = TheExec.sites.Selected
    gb_siteMask_current = TheExec.sites.Selected
    TheExec.sites.Selected = gb_siteMask_current
    
    '''//Print the header for MultiFSTP.
    If str_printMsg <> "" Then
        TheExec.Datalog.WriteComment "=============================================="
        TheExec.Datalog.WriteComment "======    " & str_printMsg & "    ======"
        TheExec.Datalog.WriteComment "=============================================="
    End If
Exit Function
errHandler: 'Add ErrHandler 2023/08/01
    Call Print_Error_Message(Error_Info, "VBT_LIB_VDD_Binning", "save_siteMask_for_MultiFSTP") 'Add ErrHandler 2023/08/01
    If AbortTest Then Exit Function Else Resume Next 'Add ErrHandler 2023/08/01
End Function

'20210720: Modified to check AllBinCut(p_mode).is_for_BinSearch to decide if it has to reset VBIN_Result for p_mode after MultiFSTP.
'20210713: Modified to check if the current testJob is for BinCut search.
'20210611: Modified to revise the message "", requested by TSMC PE.
'20210530: Modified to use gb_sitePassBin_original to restore CurrentPassBinCutNum after MultiFSTP instances.
'20210525: Created to restore current siteMask and reset VBinResult for MultiFSTP in CP1.
Public Function restore_siteMask_for_MultiFSTP(powerDomain As String, Optional str_printMsg As String = vbNullString)
On Error GoTo errHandler 'Add ErrHandler 2023/08/01
    Dim strTemp As String
    Dim strAry_pmode() As String
    Dim idx_pmode As Long
    '''//Reset siteMask after MultiFSTP.
    gb_siteMask_current = gb_siteMask_original
    TheExec.sites.Selected = gb_siteMask_original
    
    '''//Restore CurrentPassBinCutNum from gb_sitePassBin_original after MultiFSTP instances.
    CurrentPassBinCutNum = gb_sitePassBin_original
    
    If str_printMsg <> "" Then
        TheExec.Datalog.WriteComment "=============================================="
        TheExec.Datalog.WriteComment "======    " & str_printMsg & "    ======"
        TheExec.Datalog.WriteComment "=============================================="
    End If
    
    '''//Check if powerDomain is BinCut powerPin or performance_mode.
    strTemp = UCase(powerDomain)
    If VddbinPmodeDict.Exists(strTemp) Then
        '''//If it is a BinCut powerDomain, get all performance_modes of powerDomain.
        If dict_IsCorePowerInBinCutFlowSheet(strTemp) = True Then
            strAry_pmode = BinCut_Power_Seq(VddBinStr2Enum(powerDomain)).Power_Seq
        Else '''the input string is a BinCut performance_mode.
            ReDim strAry_pmode(0) As String
            strAry_pmode(0) = strTemp
        End If
        
        '''//Use pmode-loop to check each p_mode.
        For idx_pmode = 0 To UBound(strAry_pmode)
            '''//If testJob is for BinCut search, reset VBIN_Result for p_mode after MultiFSTP.
            '''20210720: Modified to check AllBinCut(p_mode).is_for_BinSearch to decide if it has to reset VBIN_Result for p_mode after MultiFSTP.
            If AllBinCut(VddBinStr2Enum(strAry_pmode(idx_pmode))).is_for_BinSearch = True Then
                Call Reset_VBinResult(VddBinStr2Enum(strAry_pmode(idx_pmode)))
            
                '''//Print info about resetting VBin_Result in the datalog for CheckScript.
                TheExec.Datalog.WriteComment strAry_pmode(idx_pmode) & ", it resets BinCut search result(VBin_Result) after data collection."
            End If
        Next idx_pmode
    Else
        Call Print_Error_Message(Error_Info, "VBT_LIB_VDD_Binning", "restore_siteMask_for_MultiFSTP", "powerDomain:" & powerDomain & " isn't the correct BinCut powerDomain or performance_mode for restore_siteMask. Error!!!")
        'TheExec.ErrorLogMessage "powerDomain:" & powerDomain & " isn't the correct BinCut powerDomain or performance_mode for restore_siteMask. Error!!!"    ''20230807: Removed by sanity check
    End If
Exit Function
errHandler: 'Add ErrHandler 2023/08/01
    Call Print_Error_Message(Error_Info, "VBT_LIB_VDD_Binning", "restore_siteMask_for_MultiFSTP") 'Add ErrHandler 2023/08/01
    If AbortTest Then Exit Function Else Resume Next 'Add ErrHandler 2023/08/01
End Function

'20210608: Modified to print header/footer for Check_flagstate_for_failflag, requested by CheckScript.
'20210604: Created to print flagstate of failFlag with PTR format in the datalog.
Public Function Check_flagstate_for_failflag(str_flag_Group As String)
On Error GoTo errHandler 'Add ErrHandler 2023/08/01
    Dim site As Variant
    Dim strAry_flag_Group() As String
    Dim idx_flag As Long
'''//==================================================================================================================================================================================//'''
'''//Note:
'''TSMC PE requested to print flagstate of failFlag with PTR format in the datalog.
'''//==================================================================================================================================================================================//'''
    If str_flag_Group <> "" Then
        strAry_flag_Group = Split(str_flag_Group, ",")
                
        TheExec.Datalog.WriteComment "*******************************************"
        TheExec.Datalog.WriteComment "*print: Check_flagstate_for_failflag start*"
        TheExec.Datalog.WriteComment "*******************************************"
        
        For Each site In TheExec.sites
            For idx_flag = 0 To UBound(strAry_flag_Group)
                '''//****************************************************************************************************************************//'''
                '''//IGXL flagstate => logicTrue=1; logicFalse=0; logicClear=-1.
                '''We defined that logicTrue is 1, and all other states as 0 for PTR format in the datalog.
                '''//****************************************************************************************************************************//'''
                If TheExec.sites.item(site).FlagState(strAry_flag_Group(idx_flag)) = logicTrue Then
                    TheExec.Datalog.WriteParametricResult site, TheExec.sites.item(site).TestNumber, logTestFail, parmTestLim, "", "", _
                                                            0, 1, 0, unitNone, 0, unitNone, 0, , , strAry_flag_Group(idx_flag), scaleNoScaling, "%.0f"

                Else
                    TheExec.Datalog.WriteParametricResult site, TheExec.sites.item(site).TestNumber, logTestPass, parmTestLim, "", "", _
                                                            0, 0, 0, unitNone, 0, unitNone, 0, , , strAry_flag_Group(idx_flag), scaleNoScaling, "%.0f"
                End If
                
                '''//Remember to check if any conflict on testNumer in the datalog.
                TheExec.sites.item(site).IncrementTestNumber
            Next idx_flag
        Next site
        
        TheExec.Datalog.WriteComment "*******************************************"
        TheExec.Datalog.WriteComment "*print: Check_flagstate_for_failflag end*"
        TheExec.Datalog.WriteComment "*******************************************"
    Else
        Call Print_Error_Message(Error_Info, "VBT_LIB_VDD_Binning", "Check_flagstate_for_failflag", "The argument str_flag_Group of Check_flagstate_for_failflag shouldn't empty. Error!!!")
        'TheExec.ErrorLogMessage "The argument str_flag_Group of Check_flagstate_for_failflag shouldn't empty. Error!!!"    ''20230807: Removed by sanity check
    End If
Exit Function
errHandler: 'Add ErrHandler 2023/08/01
    Call Print_Error_Message(Error_Info, "VBT_LIB_VDD_Binning", "Check_flagstate_for_failflag") 'Add ErrHandler 2023/08/01
    If AbortTest Then Exit Function Else Resume Next 'Add ErrHandler 2023/08/01
End Function

'20210831: Modified to remove the redundant branch of the vbt function align_startStep_to_GradeVDD.
'20210813: Modified to use Set_VBinResult_by_Step for updating PassBin, Pass step, and voltage to VBIN_Result.
'20210813: Modified to revise the format of the info about p_mode, as suggested by TSMC ZYLINI and TER Jeff.
'20210810: Modified to skip printing the info about the step-adjusted voltage for p_mode, requested by C651 Si and TSMC ZYLINI.
'20210727: Modified to revise the vbt code for testCondition "M*### E1 Voltage" in non-CP1.
'20210727: Modified to check if p_mode is already searched, as requested by C651 Si.
'20210726: Modified to add the argument "bincutNum As Long" to the vbt function Adjust_Multi_PassBinCut_Per_Site.
'20210726: Modified to check if DYNAMIC_VBIN_IDS_ZONE(p_mode).Used = True after judge_stored_IDS.
'20210721: Modified to check if Flag_Remove_Printing_BV_voltages = False for PTE or TTR.
'20210720: Modified to check if Efuse category "Product_Identifier" exists in Efuse_BitDef_Table.
'20210719: Modified to revise the vbt code for BinCut search in FT.
'20210713: Modified to add header/footer for align_startStep_to_GradeVDD.
'20210629: Created to align FT1 startStep (Bin and EQN) with CP1 results.
Public Function align_startStep_to_GradeVDD()
On Error GoTo errHandler 'Add ErrHandler 2023/08/01
    Dim site As Variant
    Dim p_mode As Integer
    Dim idx_step As Long
    Dim find_out_flag As Boolean
    Dim str_Efuse_read_ProductIdentifier As String
    Dim str_Efuse_write_pmode As String
'''//==================================================================================================================================================================================//'''
'''//Note:
'''1. For search in non-CP1, align_startStep_to_GradeVDD can update start Bin and EQN by the fused Product_Identifier and product voltages from Read_DVFM_To_GradeVDD.
'''2. As per discussion with C651 Si and Toby on 20210727, if any p_mode is searched and fused, product_identifier should be fused, too.
'''3. If the testJob is for BinCut search, it should have the dedicated "Product_Identifier", as commented by C651 Si and Toby.
'''//==================================================================================================================================================================================//'''
    '''//Get Efuse catergory of "Product_Identifier" to "read" PassBin (Product_Identifier+1).
    '''20210831: Modified to remove the redundant branch of the vbt function align_startStep_to_GradeVDD.
    str_Efuse_read_ProductIdentifier = get_Efuse_category_by_BinCut_testJob("read", "Product_Identifier")
    
    '''********************************************************************************************************'''
    '''//Warning!!! This block is the special case for Crete A0 only, and it will be removed while Crete B0!!!
    '''20210909: Modified to add the special case to get CP1-fused Prodcut_Identifier and Efuse product voltages while testJob="ft_hot" for T-Cre A0.
    '''20211118: Modified to read the cp1 product_idenfifier instead of shadow for wlft1, only for A0!
    If bincutJobName = "ft_hot" Or LCase(TheExec.CurrentJob) = "wlft1" Then
        str_Efuse_read_ProductIdentifier = "Product_Identifier"
    End If
    '''********************************************************************************************************'''
    
    '''//Check if DUT is BinCut-searched and fused in the previous testJob.
    '''//If the testJob is for BinCut search, it should have the dedicated "Product_Identifier", as commented by C651 Si and Toby.
    '''Note: As per discussion with C651 Si and Toby on 20210727, if any p_mode is searched and fused, product_identifier should be fused, too.
    If str_Efuse_read_ProductIdentifier <> "" Then
        TheExec.Datalog.WriteComment "Product_Identifier" & ",it can use Efuse category:" & str_Efuse_read_ProductIdentifier
    Else
        TheExec.Datalog.WriteComment "Efuse category:" & "Product_Identifier" & ", it wasn't fused before the current testJob, so that skip align_startStep_to_GradeVDD."
        TheExec.Datalog.WriteComment ""
        TheExec.Datalog.WriteComment ""
        TheExec.Datalog.WriteComment vbCrLf
        Exit Function
    End If
    
    TheExec.Datalog.WriteComment "=============================================="
    TheExec.Datalog.WriteComment "======    " & "align_startStep_to_GradeVDD" & "    ======"
    TheExec.Datalog.WriteComment "=============================================="
    TheExec.Datalog.WriteComment "***** Start of align_startStep_to_GradeVDD *****"
    
    For Each site In TheExec.sites
        For p_mode = 0 To MaxPerformanceModeCount - 1
            '''//Check if DYNAMIC_VBIN_IDS_ZONE(p_mode).Used = True after judge_stored_IDS.
            If DYNAMIC_VBIN_IDS_ZONE(p_mode).Used(site) = True Then
                '''//VBIN_RESULT(p_mode).tested(site) = True is updated in Read_DVFM_To_GradeVDD if product voltage of p_mode was fused in previous testJob.
                If VBIN_RESULT(p_mode).tested(site) = True Then
                    '''init
                    find_out_flag = False
                    
                    '''//Use step-loop to find the matched step in Dynamic_IDS_zone.
                    For idx_step = 0 To DYNAMIC_VBIN_IDS_ZONE(p_mode).Max_Step(site)
                        If CDec(DYNAMIC_VBIN_IDS_ZONE(p_mode).Product_Voltage(idx_step)) = CDec(VBIN_RESULT(p_mode).GRADEVDD) Then
                            '''//Check if p_mode for BinCut search has the dedicated Efuse category in the current testJob.
                            '''20210730: C651 Toby revised the rule of adjust_VddBinning that "If CP1 has only PCPU bin cut search, then we only need PCPU for PrintOut_VddBinning and Adjust_VddBinning. The other domains can be skipped at CP1."
                            '''//Get Efuse category of Efuse product voltage(GradeVDD) for p_mode.
                            str_Efuse_write_pmode = get_Efuse_category_by_BinCut_testJob("write", VddBinName(p_mode))
                            
                            '''//If p_mode has Efuse category for the current BinCut testJob, get Efuse product voltage for P_mode.
                            '''Note: As per discussion with TSMC ZYLINI and TER Jeff, for those p_mode to be fused in the current testJob, update Grade by step, do not use "product - [insertion]GB".
                            '''
                            ''' Since WLFT2 is mapping to WLFT1 as well, so we need to skip WLFT1 to align CP grade vbin result
                            ''' During Bring up, we need to remove it first, since we don't have WLFT2 and flow will be MP like
                            '''
                            If str_Efuse_write_pmode <> "" Then
                            'If str_Efuse_write_pmode <> "" And Not LCase(TheExec.CurrentJob) = "wlft1" Then     ''' for wlft2 request, if no wlft2,should comment
                                '''//Update PassBin, Pass step, flag"VBIN_Result(p_mode).tested", and voltage to VBIN_Result by the step in Dynamic_IDS_Zone.
                                '''20210813: Modified to use Set_VBinResult_by_Step for updating PassBin, Pass step, and voltage to VBIN_Result.
                                Call Set_VBinResult_by_Step(site, p_mode, idx_step)
                            Else
                                VBIN_RESULT(p_mode).step_in_IDS_Zone = idx_step
                                VBIN_RESULT(p_mode).passBinCut = DYNAMIC_VBIN_IDS_ZONE(p_mode).passBinCut(idx_step)
                                VBIN_RESULT(p_mode).step_in_BinCut = DYNAMIC_VBIN_IDS_ZONE(p_mode).EQ_Num(idx_step) - 1 '''step_in_BinCut=EQN-1
                            End If
                            
                            '''//Print start Bin and EQN, as requested by C651 Si.
                            If Flag_Remove_Printing_BV_voltages = False Then
                                TheExec.Datalog.WriteComment "Site:" & site & "," & VddBinName(p_mode) & "," & _
                                                                "PassBin=" & VBIN_RESULT(p_mode).passBinCut & ",EQN=" & VBIN_RESULT(p_mode).step_in_BinCut + 1 & ", already searched"
                            End If
                            
                            find_out_flag = True
                            Exit For
                        End If
                    Next idx_step
                    
                    '''//If p_mode can't find start_step, bin out the failed DUT.
                    If find_out_flag = False Then
                        TheExec.sites.item(site).FlagState(strGlb_Flag_Vddbinning_Fail_Stop) = logicTrue
                        Call Print_Error_Message(Error_Info, "VBT_LIB_VDD_Binning", "align_startStep_to_GradeVDD", "Site:" & site & "," & VddBinName(p_mode) & ", align_startStep_to_GradeVDD can't find start_step for Efuse product_identifier and product voltage. Error!!!")
                        ''TheExec.ErrorLogMessage "Site:" & site & "," & VddBinName(p_mode) & ", align_startStep_to_GradeVDD can't find start_step for Efuse product_identifier and product voltage. Error!!!"    ''20230807: Removed by sanity check
                    End If
                Else '''If product voltage of p_mode isn't fused, adjust the correct step to align passbincut of P_mode with CurrentPassBinCutNum, and update Grade/GradeVDD for p_mode.
                    '''//Print start Bin and EQN, as requested by C651 Si.
                    '''20210813: Modified to revise the format of the info about p_mode, as suggested by TSMC ZYLINI and TER Jeff.
                    If Flag_Remove_Printing_BV_voltages = False Then
                        TheExec.Datalog.WriteComment "Site:" & site & "," & VddBinName(p_mode) & ", no fused product voltage, to be evaluated"
                    End If
                    
                    If VBIN_RESULT(p_mode).passBinCut <> CurrentPassBinCutNum(site) Then
                        '''20210810: Modified to skip printing the info about the step-adjusted voltage for p_mode, requested by C651 Si and TSMC ZYLINI.
                        Adjust_Multi_PassBinCut_Per_Site p_mode, site, CurrentPassBinCutNum(site), True
                    End If
                End If '''If VBIN_RESULT(p_mode).tested(site) = True
            End If '''If DYNAMIC_VBIN_IDS_ZONE(p_mode).Used = True
        Next p_mode
    Next site
    
    TheExec.Datalog.WriteComment "***** End of align_startStep_to_GradeVDD *****"
    TheExec.Datalog.WriteComment ""
    TheExec.Datalog.WriteComment ""
    TheExec.Datalog.WriteComment vbCrLf
Exit Function
errHandler: 'Add ErrHandler 2023/08/01
    Call Print_Error_Message(Error_Info, "VBT_LIB_VDD_Binning", "align_startStep_to_GradeVDD") 'Add ErrHandler 2023/08/01
    If AbortTest Then Exit Function Else Resume Next 'Add ErrHandler 2023/08/01
End Function

Public Function Pre_adjust_VddBinning(Adjust_Max_Enable As Boolean, Adjust_Min_Enable As Boolean, Optional Adjust_Power_Max_list As String, Optional Adjust_Power_Min_list As String)
On Error GoTo errHandler 'Add ErrHandler 2023/08/01
    Dim site As Variant
    Dim p_mode As Integer
    Dim i As Integer
    Dim BinNumber As Long
    Dim Active_site_count As Long
    Dim printflag As Boolean

    '''20231017: Modified to use the rule from Adjust_Vddbinning
    '''//Check if BinCut testJob is for BinCut search.
    '''//The flag "is_BinCutJob_for_StepSearch" is True if any testCondition from the table "Non_Binning_Rail" has the keyword "*Evaluate*Bin*".
    If is_BinCutJob_for_StepSearch = False Then
        '''//Align the format of Judge_PF and Judge_PF_func in the datalog.
        TheExec.Datalog.WriteComment ""
        Exit Function
    End If
    
    '''//BinCut adjust_VddBinning is for CP1 only.
    'If bincutJobName = "cp1" Then
        printflag = False
        
        For BinNumber = 1 To Total_Bincut_Num
            For Each site In TheExec.sites
                If BinNumber = CurrentPassBinCutNum Then
                    '''//Check all Pmodes bin num and adjust
                    For p_mode = 0 To MaxPerformanceModeCount - 1
                        '''//If PassBinCut of P_mode doesn't match CurrentPassBinCutNum, adjust step inDynamic_IDS_Zone to CurrentPassBinCutNum.
                        '''20200317: Modified for SearchByPmode.
                        printflag = False

                        ''20231017: Modified to use the rule from PrintOut VddBinning. Added for VDD_Sram pins
                        If VBIN_RESULT(p_mode).passBinCut <> CurrentPassBinCutNum And _
                            (AllBinCut(p_mode).Used = True Or (dict_IsCorePower(AllBinCut(p_mode).powerPin) = True And UCase(AllBinCut(p_mode).powerPin) Like "*VDD*_SRAM*")) Then
                            Adjust_Multi_PassBinCut_Per_Site p_mode, site, BinNumber
                        End If
                    Next p_mode
                
                    '''//Check Adjust_Max and Adjust_Min from the arguments.
                    Call Check_Adjust_Max_Min(Adjust_Max_Enable, Adjust_Min_Enable, Adjust_Power_Max_list, Adjust_Power_Min_list)
            
                    For i = 0 To UBound(pinGroup_CorePower)
                        '20231003: Added to support SRAM Voltage calculate with equation for pre-adjust vddbinning
                        If dict_IsCorePower(pinGroup_CorePower(i)) = True And UCase(pinGroup_CorePower(i)) Like "*VDD*_SRAM*" Then
                            Call check_voltageInheritance_for_SrampowerDomain(pinGroup_CorePower(i))
                        Else
                            Call check_pmode_for_pre_adjust_VddBinning(site, BinCut_Power_Seq(VddBinStr2Enum(pinGroup_CorePower(i))).Power_Seq, IDS_for_BinCut(VddBinStr2Enum(pinGroup_CorePower(i))).Real)
                        End If
                    Next i
                End If
            Next site
        Next
    'Else
        'TheExec.Datalog.WriteComment "The Job selection is incorrect for Pre_adjust_VddBinning. Error!!!"
        'TheExec.ErrorLogMessage "The Job selection is incorrect for Pre_adjust_VddBinning. Error!!!"
        'TheExec.Flow.TestLimit resultVal:=999, lowVal:=1, hiVal:=1, Tname:="The Job selection is incorrect for Pre_adjust_VddBinning. Error!!!" '''Modified for BinOut, 20180904
    'End If
Exit Function
errHandler:
    TheExec.Datalog.WriteComment "Error encountered in VBT Function of adjust_VddBinning"
    If AbortTest Then Exit Function Else Resume Next
End Function
