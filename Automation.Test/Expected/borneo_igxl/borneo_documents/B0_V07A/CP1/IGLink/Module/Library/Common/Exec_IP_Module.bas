Attribute VB_Name = "Exec_IP_Module"
#Const isUFP = True
Option Explicit
'Public write_SPIROM_CheckSum As Integer

Public write_spirom As New SiteBoolean
Public Mbist_Repair_CompareType As Variant  'for Mbist finger print

Public DoAll_save As Boolean
Public OverRide_FailStop As Boolean




' Immediately at the conclusion of the initialization process.
' Do not program test system hardware from this function.
Function OnTesterInitialized()
On Error GoTo errHandler 'Add ErrHandler 2023/08/18

    ' Put code here
    
    
    Exit Function
errHandler: 'Add ErrHandler 2023/08/18
    MsgBox "Error encountered in Exec Interpose Function OnTesterInitialized" + vbCrLf + _
        "VBT Error # " + Trim(str(err.number)) + ": " + err.Description
    If AbortTest Then Exit Function Else Resume Next 'Add ErrHandler 2023/08/18
End Function
 
' Immediately at the conclusion of the load process.
' Do not program test system hardware from this function.
Function OnProgramLoaded()
On Error GoTo errHandler 'Add ErrHandler 2023/08/18

    


    'set ture to get if there is any halt in pattern when we use pattern burst yes
    'TheHdw.Digital.CheckAllModulesExecuted = True
    
    ' Put code here
    #If IGXL8p30 Then
    #Else
    thehdw.Digital.LevelSets.OptimizeAllocation = True      'add to avoid levelsets over 255 limitation 2018/01/09
    Call tl_activateallusersheets
    #End If
    
    glb_TesterType = thehdw.Tester.type
    
    ' Put code here
    'for plus LogInTestExecutionOrder need to set in OnProgramLoaded to fix datalog print sequence
    If glb_TesterType = "UltraFLEXplus" Then
        TheExec.Datalog.Setup.DatalogSetup.LogInTestExecutionOrder = True ''''' Add for UFP
        TheExec.Datalog.ApplySetup  'must need to apply after datalog setup
    End If
        
    'for TERA1 encryption need, add following code
    If Not TheExec.SoftwareVersion Like "8.10.90_uflx*" Then
        m_cpcmodule.SuppressCheckForUnProtectedPatterns = True
    End If
    
    'for 8.30 encryption need, add following code
    m_STDSvcClient.CPCModule.SuppressCheckForUnProtectedPatterns = True
    
    If TheExec.SoftwareVersion Like "*9.10*" Then
        CallByName TheExec.TestProgram, "MemoryLimitCheckEnabled", VbLet, False
    End If
    
    'for relax reference clock to lower frequency
    'Enable the full frequency range for the nWire PA clock
''    TheHdw.Digital.Timing.FullPAClockFrequencyRange = True
''
    'turnning on the simulator
    'note if so, the test time of simulation would increase
    'TheExec.Simulator.ForceAllSimulation (tlSimForce)
    
    TheExec.DataManager.MaxSheetValidationErrorEnabled = False
    thehdw.Digital.EnablePinRespecification = True
    
    If is_reference_installed("Scripting") = False Then
        Application.ActiveWorkbook.VBProject.References.AddFromFile "C:\WINDOWS\system32\scrrun.dll"
    End If
    
    If is_reference_installed("VBScript_RegExp_55") = False Then
        Application.VBE.ActiveVBProject.References.AddFromFile "C:\WINDOWS\system32\vbscript.dll\3"
    End If
    
    
  '  If is_reference_installed("PATTERNDATAMANAGERLib") = False Then
  '      If Application.OperatingSystem Like "*NT 6*" Then
  '          Application.ActiveWorkbook.VBProject.References.AddFromFile "C:\Program Files (x86)\Teradyne\IG-XL\8.10.12_uflx\bin\PatternDataManager.dll"
  '      Else
  '          Application.ActiveWorkbook.VBProject.References.AddFromFile "C:\Program Files\Teradyne\IG-XL\8.10.90_uflx\bin\PatternDataManager.dll"
  '      End If
  '  End If
  '
   'nWire PA for XI0
   '20210416, add for Ufp
    If glb_TesterType = "Jaguar" Then
   
        Call thehdw.Protocol.Families("nWire").Types.Add(".\xml_Files\FreeRunClk_TDR_TRUE_32Clk_8Idle.xml", "Clock")
        Call thehdw.Protocol.Families("nWire").Types.Add(".\xml_Files\FreeRunClk_differential.xml", "Clock_Diff")
   
   
        Call thehdw.Protocol.Families("FRC").Types.Add(".\xml_Files\FRC.pa")
        Call thehdw.Protocol.Families("FRC").Types.Add(".\xml_Files\FRCRef.pa")
        Call thehdw.Protocol.Families("FRC").Types.Add(".\xml_Files\FRCRef_differential.pa", "FRC_Clock_Diff")
    End If
   
    thehdw.Digital.EnableSharedsiteSupportCheck = True
    TheHdw.PinLevels.SharedSiteCheckEnabled = True
    
#If AP = True Then
    Call thehdw.Protocol.Families("nWire").Types.Add(".\xml_Files\UART_x3_RX.xml", "UART_PA_RX") ' 20160322
    Call thehdw.Protocol.Families("nWire").Types.Add(".\xml_Files\UART_x3_TX.xml", "UART_PA_TX") ' 20160603 Leslie
#ElseIf LCD = True Then
    Call LCD_OnProgramLoad
#End If



    'nWire PA for UART receiver need to take care if other pattern use this PA pin.
    'Call TheHdw.Protocol.Families("nWire").Types.Add(".\xml_Files\UART_x3_RX.xml", "UART") ' 20160322
 

    '''''''''20180628 add to prevent shmoo can't read sheet
    Dim CZ_Activate_Sheet As Worksheet
    For Each CZ_Activate_Sheet In ThisWorkbook.Sheets
        If UCase(CZ_Activate_Sheet.name) Like "*FLOW_DCTEST*" Or UCase(CZ_Activate_Sheet.name) Like "*FLOW_HARDIP*" Then
            CZ_Activate_Sheet.Activate
        End If
    Next CZ_Activate_Sheet
    ''''''''

    '----------Debug.Print "Force Compilation"
    Dim VBACompile As Object
    Dim CommandBar As Object
    Set CommandBar = Application.VBE.CommandBars.FindControl(msoControlButton, 578)
    Set VBACompile = CommandBar.Control
    VBACompile.Execute
    '----------

    If glb_TesterType = "Jaguar" Then
    Else
        CallByName TheExec.Flow, "WaitForOrphanAlarmsEnabled", VbLet, True
    End If

    Exit Function
    ''20240227 Initial Syncup_Dict
    Syncup_Dict.RemoveAll
errHandler: 'Add ErrHandler 2023/08/18
    HandleExecIPError "OnProgramLoaded"
    If AbortTest Then Exit Function Else Resume Next 'Add ErrHandler 2023/08/18
End Function
 
' Immediately at the conclusion of the validate process. Called only if validation succeeds.
' [20230907][All][Doris] Modify for enable word gating
Function OnProgramValidated()
On Error GoTo errHandler 'Add ErrHandler 2023/08/18
    
   ' write_SPIROM_CheckSum = 0

    ''''flag initilize
    Call InitVariableOnValidated
    
    currentJobName = LCase(TheExec.CurrentJob)
    CurrentChannelMap = LCase(TheExec.CurrentChanMap)
    
    thehdw.Patterns.EnableExplicitFileNames = True  'to fix dssc can not recognize pat.gz format
    TheExec.Flow.HighParallelMode = True            '140501 pre-shut down in parallel, TTR purpose, false to check if nWire not stop by site fail
    Ignore_nWire_Error                              '140328 2D-shmoo nWire debug
    
#If AP = True Then
    Call CheckPatAndCZmappingTable
#ElseIf LCD = True Then
    Call LCD_OnProgramValidate
#End If
        
    Init_Datalog_Setup          'datalog settings for the requirement of product
    
    ''202107 - UFP issue
    ''move the command from "OnProgramLoad"
    ''due to need to get pin / channel info. IGXL can't get the info from "OnProgramLoad".
    'TheHdw.Digital.CheckAllModulesExecuted = True
    
    Call TDR_ExcludedPin_WriteTraceLen("Cal_Excluded")
    If glb_TesterType = "Jaguar" Then
        thehdw.Digital.Pins("Cal_Excluded").Calibration.Excluded = True 'bypass TDR
    ElseIf glb_TesterType = "UltraFLEXplus" Then
        thehdw.Calibration.Pins("Cal_Excluded").Excluded = True
    End If
 
    If glb_TesterType = "Jaguar" Then
        thehdw.Digital.CheckAllModulesExecuted = True
        thehdw.Digital.Calibration.Validate
    ElseIf glb_TesterType = "UltraFLEXplus" Then
        thehdw.Calibration.Validate
    'Call SetInitialTrace("Cal_Excluded")
    ' Add for Hatkar suggestion for CMEM and HRAM 220517
        thehdw.TimeOut.Results = 100
    End If

 '   TheExec.Flow.EnableWord("Read_EEPROM_DIBID") = True 'for DIB Board ID read out
    
    '*** protect efuse sheets ***
'    gL_1st_FuseSheetRead = 0 ''''20150624
'    Call UnProtect_eFuse_Sheet
'    Call autoArrange
'    Call autoArrange("UDR_compare_ChkList_appA")
'    Call Protect_eFuse_Sheet ''''MUST be here
    
    ''''---------Start of Mbist ChkList---------------
'    gL_1st_MbistSheetRead = 0 ''''20151020
    ''Call UnProtect_Mbist_Sheet
    ''Call Protect_Mbist_Sheet
    ''''---------  End of Mbist ChkList---------------
     
     ' 20150121 for SPI ROM auto fuse
    TheExec.Flow.enableWord("Write_SPIROM") = True 'trigger SPI ROM write
    
    Call ReloadUARTModules
    
    ' 20150128 - Load PA UART RX module
        'only enable it during SPI debug.
    'TheHdw.Protocol.Ports("UART_PA").ModuleFiles.UnloadAll
    'Call PreLoad_PA_Modules("READ", 15000, "UART_PA")
    'only enable it during SPI debug.
    
    
    write_spirom = True '20160324 central lib

    ' 20160422 - Write Bin Name to STDF
    Call Bintable_initial
    If glb_TesterType = "Jaguar" Then
        thehdw.Digital.Alarm(tlHSDMAlarmAll) = tlAlarmForceBin
    End If
    Call Process_DSSC_Dictionary ' Added By Oscar 180523 For Read DigSrc assignment from Sheet(Function is in LIB_HardIP)
'    Call HardIP_OnProgramStarted_Process
    
     ''' to expand fail cycle number more than 100 million (from 10 million) - 190408
''' 20210709 for datalog format alignment
'''    TheExec.Datalog.Setup.Shared.ascii.Columns.EnableCustomWidths = True
'''    TheExec.Datalog.Setup.Shared.ascii.Columns.Functional.Cycle.Width = 9
'''    TheExec.Datalog.ApplySetup
    
        ' for License Check, initial global variable
    gL_License_check = 0
    ENG_SweepPin = True
    dic_mapCritcalPin.RemoveAll
    
    If TheExec.RunMode = runModeDebug Then isDebugMode = True

    'CharStoreResultsUntilNextRun, clear shmoo momory to prevent crash
    If isDebugMode = False Then
        TheExec.DevChar.Configuration.Features.item(tlDevCharFeature_StoreResultsUntilNextRun).Enabled = False
        m_STDSvcClient.SelfTest.MemoryCollectRunInterval = 1
    End If
    
    '20230907 add for gating enable word
    Call EnableWdGating(OnProgValidation)
    Enable_CZsetup_Chk = True

    glb_CheckHarvestInst = False

    Exit Function
errHandler: 'Add ErrHandler 2023/08/18
    HandleExecIPError "OnProgramValidated"
    If AbortTest Then Exit Function Else Resume Next 'Add ErrHandler 2023/08/18
End Function
 
' Immediately at the conclusion of the validate process. Called only if validation fails.
Function OnProgramFailedValidation()
On Error GoTo errHandler 'Add ErrHandler 2023/08/18

    ' Put code here
    
    
    Exit Function
errHandler: 'Add ErrHandler 2023/08/18
    HandleExecIPError "OnProgramFailedValidation"
    If AbortTest Then Exit Function Else Resume Next 'Add ErrHandler 2023/08/18
End Function
 
' Immediately at the conclusion of the user DIB calibration process (previously
' known as the TDR calibration process). Called only if user DIB calibration succeeds.
Function OnTDRCalibrated()
On Error GoTo errHandler 'Add ErrHandler 2023/08/18

    ' Put code here
    'To fix freerunclock unstable issue used for  8.10.12 is fixed in 8.30
    'm_stdsvcclient.TesterSupport.TimingCalChannel.DriveEnable
    
    Dim sbCurrentSiteStart As New SiteBoolean
    
    
    sbCurrentSiteStart = TheExec.sites.Starting     'save current site status
    
    TheExec.sites.Starting = True
    
    Call SetPowerAndIOPin_Volt_0v("All_Power", "All_Digital")       'set power and digital pin to 0v
    
    TheExec.sites.Starting = sbCurrentSiteStart     'restore site status
        
    Exit Function
errHandler: 'Add ErrHandler 2023/08/18
    HandleExecIPError "OnTDRCalibrated"
    If AbortTest Then Exit Function Else Resume Next 'Add ErrHandler 2023/08/18
End Function
 
 
' Immediately after "pre-job reset" when the test program starts.
' Note that "first run" actions can be enclosed in
' If TheExec.ExecutionCount = 0 Then...
' (see online help for ExecutionCount)
' [20230424][All][Tank] add run function Get_CurrentProfile to parsing current profile sheet
' [20230524][All][Si] add TTR enable word "TTR_Disable_Alarm"
' [20230620][T-All][James] fix Call ParseExecutionProfileSheet logic error
' [20230906][All][Carter] Modify fix process powerdown when Conti binout
' [20231108][All][Tank] Add ParseIDSMappingTable in OnProgramStarted first
' [20231228][T-Tah][Tank] Parsing SFCPatterns sheet
Function OnProgramStarted()
On Error GoTo errHandler 'Add ErrHandler 2023/08/18
    Dim i As Integer, j As Integer
    
    Call InitVariableOnStarted
    
    glb_TesterType = thehdw.Tester.type
    
    TestProgram_Day_Code = CStr(Year(Now)) & right("0" & CStr(Month(Now)), 2) & right("0" & CStr(Day(Now)), 2)
    TestProgram_Day_Code = TestProgram_Day_Code & right("0" & CStr(Hour(Now)), 2) & right("0" & CStr(Minute(Now)), 2) & right("0" & CStr(Second(Now)), 2)
    
    currentJobName = LCase(TheExec.CurrentJob) ''Carter, 20191115
    CurrentChannelMap = LCase(TheExec.CurrentChanMap)
    
    '20211007 If CurrentProfile enable
    ''Save DoAll and Override Fail-stop settings
    ''Set DoAll = Enable, Override Fail-stop = Enable
    PowerPinCnt_mapping = 0
    If TheExec.Flow.enableWord("CurrentProfile") Or TheExec.Flow.enableWord("VoltageProfile") Then
        TheExec.Datalog.WriteComment "CurrentProfile is Working!!"
        create_folderName = False
        MsgBox "CurrentProfile is Working!!"  ''mop for execution time
        glb_DisablePlot_IProfile = TheExec.Flow.enableWord("DisablePlot_IProfile")
        glb_Print_IProfileValue = TheExec.Flow.enableWord("Print_IProfileValue")
    Else
        glb_DisablePlot_IProfile = False
        glb_Print_IProfileValue = False
    End If
        
    gl_EnableCurrentProfile = False
    gl_EnableVoltageProfile = False
    glb_ProfileFilter = False
    gl_EnableCurrentProfile = TheExec.Flow.enableWord("CurrentProfile")
    gl_EnableVoltageProfile = TheExec.Flow.enableWord("VoltageProfile")
    glb_ProfileFilter = TheExec.Flow.enableWord("Profile_Filter")
    Profile_byflow = False
#If AP = True Then
    If Flag_CZMappingTable_Check = False Then CheckPatAndCZmappingTable
#ElseIf LCD = True Then
    Call LCD_OnProgramStarted
#End If


#If AP Then
        '20211008 add for settle time
    If glb_TesterType = "UltraFLEXplus" Then
        If Flag_DC_ParsingRecoveryTable = False Then
            ''' clear data in array
            ReDim DC_SettleTime_UVS256HP(MaxLoadCount)
            ReDim DC_SettleTime_UVS64(MaxLoadCount)
            ReDim C_code_UVS256HP(MaxCcodeCount_256)
            ReDim C_code_UVS64(MaxCcodeCount_64)
            Dict_Dibcap_PowerPinName.RemoveAll
            
            Initial_UVS256HP_UVS64_settleTime_CCode
            Parsing_dibcap_table
        End If
    End If
    
    'lesson learn from Tahiti/Tupai -> adjust HexVS 0.1A current range from 10ms to 10.3ms
    'Find_Range_WaitTime will only execute when first touch down, If you want to apply this wait time you need to set Flag_GetChannelType = False
    glb_HexVs_0p1range_waittime = 0.3 * ms
    Call Print_Error_Message(Warning_Info, "LIB_DC.bas", "Find_Range_WaitTime", "hexvs 0.1 Current Range settle time change from 10ms to " & (10 * ms + glb_HexVs_0p1range_waittime) * 1000 & "ms")
    Call Get_Channel_Type
    Call Get_PowerSeq
    Call Get_CurrentProfile
    ' Put code here
    Call RemoveAllStored

    'Add for RAK
    Call HardIP_OnProgramStarted_Process
    
    
    '20240124: Added for new MFSTP sheet
    Call Parsing_UserFunction_Sheet
    
    'use "Parse_SELSRM_Mapping_Table" for parsing selsrm table
    'Call Parsing_SELSRM_Mapping_Table '"Parsing" is for Bincut, we call it in initVDDbinning

''''' Modify by TY
'    Call Parse_SELSRM_Mapping_Table '"Parse" is for CHAR/function_T
'    Call Parse_EMA_DigSrcInfo ''Carter, 20191115
''''' Modify by TY
    Call Parse_DSSCPat_DigSrcInfo
    
    Call ParseIDSMappingTable   'just parsing IDS_Mapping_Table

    '20231017 T0TX modified
    If TheExec.enableWord("T0TX_ROOM") = True And TheExec.enableWord("T0TX_HOT") = True Then
        TheExec.Datalog.WriteComment "Conflict! EnableWord T0TX_ROOM and T0TX_Hot are both selected!!!"
        MsgBox ("Conflict! EnableWord T0TX_ROOM and T0TX_Hot are both selected!!!")
    End If
   
    Call initVddBinning
    
    
    If TheExec.Flow.enableWord("ENG_IDS_SweepPin") Then ParsePreCondition
    
    ''Export UnExistPins
    If TheExec.enableWord("aExportUnExistPins") = True Then
        Call Search_UnExistPin
    End If
#End If

    
    If glb_TesterType = "Jaguar" Then
        thehdw.Digital.CheckContextExclusion = True 'nWire context check flag, need to think about how to switch between engineering mode and production mode.
    End If
    'Init_DIB_Power  'reset DIB power ''move to onProgramLoad
    
    ''20220308 move to OnProgramLoaded Init_DIB_Power  'reset DIB power
    
    If thehdw.DIB.powerOn = False Then
        Init_DIB_Power  'reset DIB power
    End If
    
    gL_ProductionTemp = vbNullString  'init temparature
    TheExec.Flow.FlowFlagMode = tlFlowFlagLatchTestResult 'prevent next pass overwrite previous  fail
    thehdw.DSP.ExecutionMode = tlDSPModeHostDebug   'use host computer to cal dsp function
    
     
 #If AP Then
        'Get differential pair
    Call RetrieveDictionaryOfDiffPairs
#End If
    '*** eFuse initial
    ''''20150624 move to Flow_Table_Main_Init_Flows
''    Call auto_eFuse_Initialize
    
    '' 20151029 - Compensate DIB impedence for RAK
    Dim ws_def As Worksheet
    Dim wb As Workbook
    Dim siteIdx As Integer
    Dim SiteNum1 As Integer
    Dim SiteNum2 As Integer
    Dim DACInitialFlag As Boolean

    Set wb = Application.ActiveWorkbook

    If TheExec.Flow.enableWord("DebugPrintFlag") = True Then
        DebugPrintFlag_Chk = True
    Else
        DebugPrintFlag_Chk = False
    End If

    '20231120: Parse SSN mapping table and harvest table
    If Flag_HarvPinFlag_Mapping_Table_Parsed = False Then
        Parsing_HarvPinFlag_Mapping_Table
    End If
    
    '20231220: Initial SSN core mask
    If ssnPatternsDict.Count <> 0 Then
        thehdw.Digital.ScanNetworks.ClearAllMasks
        TheExec.Datalog.WriteComment "SSN clear all mask!!"
    End If

    'Check whether to Enable Harvest Flag, if EnableWord "EnableCoreHarvest" = TRUE, and Harvest fail, then IGXL will Set Harvest Flag to TRUE
    If TheExec.Flow.enableWord("EnableCoreHarvest") = True Then
        EnableCoreHarvest = True
    Else
        EnableCoreHarvest = False
    End If
    
    'Check whether to Enable DisableCompare, if EnableWord "EnableCoreMask" = TRUE, and Harvest Fail at previous stage, then IGXL will DisableCompare the related Harvest pin
    If TheExec.Flow.enableWord("EnableCoreMask") = True Then
        EnableCoreMask = True
    Else
        EnableCoreMask = False
    End If
    DACInitialFlag = False  '''added for NonAP flag initial

    Find_nWire_Pin   '''update for multiple nWire CLK, 2017/07/18
''
Dim site As Variant
For Each site In TheExec.sites
    If write_spirom = True Then
        TheExec.Flow.enableWord("Write_SPIROM") = True
    End If

Next site
''
#If AP = True Then
    gl_UseStandardTestName_Flag = True
    If gl_UseStandardTestName_Flag Then
        Call SetupDatalogFormat(TestNameW:=90, PatternW:=100)
    End If
    
    If TheExec.Flow.enableWord("production") = True Then
        TheExec.Datalog.WriteComment "[HIP TTR EnableWord: Production]"
    ElseIf TheExec.Flow.enableWord("monitoring") = True Then
        TheExec.Datalog.WriteComment "[HIP TTR EnableWord: Monitoring]"
    ElseIf TheExec.Flow.enableWord("char") = True Then
        TheExec.Datalog.WriteComment "[HIP TTR EnableWord: Char]"
    Else
        TheExec.Datalog.WriteComment "[HIP TTR EnableWord: None]"
    End If
    
    ''''' 20180710 Add initialize value ''''''''''''
    CHAR_USL_HVCC = 9999
    CHAR_USL_LVCC = 9999
    CHAR_LSL_HVCC = 9999
    CHAR_LSL_LVCC = 9999

    'for Mbist finger print
    Mbist_Repair_CompareType = "Cycle"
    If TheExec.Flow.enableWord("Mbist_FingerPrint_Vector") = True Then Mbist_Repair_CompareType = "Vector"
    
    ''20210504 add autoZ
    If UCase(RegKeyRead_autoZ("AutoZEnable")) = "TRUE" Then 'For Auto try Z
        TheExec.enableWord("AutoZOnly") = True
    Else
        TheExec.enableWord("AutoZOnly") = False
    End If
        
        gl_Flag_1st_contact_R = False
#End If
    If ENG_SweepPin And TheExec.Flow.enableWord("ENG_IDS_SweepPin") Then
        ENG_SweepPin = True
    Else
        ENG_SweepPin = False
    End If

        glb_EVS_Disable_Printout = TheExec.Flow.enableWord("EVS_Disable_Printout")

    If ((TheExec.RunOptions.Profiler.OutputExecutionProfileSheet = False Or TheExec.RunOptions.Profiler.Enabled = False) And (gl_EnableCurrentProfile Or gl_EnableVoltageProfile)) Then 'add () because RF/LCD meet
        
        Call ParseExecutionProfileSheet
        
    End If
    
    
    'CharStoreResultsUntilNextRun, For OI debug mode
    If isDebugMode = False And TheExec.Flow.enableWord("Disable_MemoryClr") = True Then
        TheExec.DevChar.Configuration.Features.item(tlDevCharFeature_StoreResultsUntilNextRun).Enabled = True
        m_STDSvcClient.SelfTest.MemoryCollectRunInterval = 1
    End If
    
    If TheExec.Flow.enableWord("TTR_Disable_Alarm") Then        'T-Col TTR approve by Si -- 230413
        gl_bTTRDisableAlarm = True
    Else
        gl_bTTRDisableAlarm = False
    End If
    
    If g_AllDCVSPin = "" Then g_AllDCVSPin = "All_Power"
    
    glb_isSFC_Enabled = TheExec.enableWord("Enable_SFC")
    ' Parsing SFCPatterns
    Call Read_SFC_Table
    Get_CurrentProfile
    If gl_EnableCurrentProfile Or gl_EnableVoltageProfile Then

        Call Export_All
        glb_Boolean_export = True
    End If
    Exit Function
errHandler: 'Add ErrHandler 2023/08/18
    HandleExecIPError "OnProgramStarted"
    If AbortTest Then Exit Function Else Resume Next 'Add ErrHandler 2023/08/18
End Function
 

' Immediately before "post-job reset" when the test program completes.
' Note that any actions taken here with respect to modification of binning
' will affect the binning sent to the Operator Interface, but will not affect
' the binning reported in Datalog.
Function OnProgramEnded()
On Error GoTo errHandler 'Add ErrHandler 2023/08/18
    glb_FlagCheckingFRCClock = True
    ' Put code here
    ''TheHdw.DIB.powerOn = False  'reset DIB power, prevent hot switch
        ''20220119, Avoid different behavior between UP and UFP after click "Debug_stop"
    ''' Check DC limit on first TD only. Check function => CheckTestInst_HiLoLimit()'''
    If gl_isCheckClampLimit = ContiClampCheckType.CheckInit Then
        gl_isCheckClampLimit = ContiClampCheckType.CheckPass
    End If

    Exit Function
errHandler: 'Add ErrHandler 2023/08/18
    HandleExecIPError "OnProgramEnded"
    If AbortTest Then Exit Function Else Resume Next 'Add ErrHandler 2023/08/18
End Function
 
' Immediately before a site is disconnected.
' Use TheExec.Sites.SiteNumber to determine which site is being disconnected.
' [20230906][All][Carter] Modify fix process powerdown when Conti binout
Function OnPreShutDownSite()
On Error GoTo errHandler 'Add ErrHandler 2023/08/18
    ' Put code here
    Dim v_site As Variant
    Dim powerDownEnable As Boolean
    
    If TheExec.RunOptions.DoAll = False And ATE_STR_Summary_Table_Parse_Flag = True Then
        Call ATE_STR_Summary_Flag_Operate
    End If
    ' Put code here
    TheExec.Flow.TestLimit resultVal:=0, lowVal:=0, hiVal:=0, Tname:="Dummy", ForceResults:=tlForceNone, tNum:="100"  '' to effective SetDynamicTestName setting
    
    For Each v_site In TheExec.sites
        If powerUpDone(v_site) Then
            powerDownEnable = True
            powerUpDone(v_site) = False
        End If
    Next v_site
    
    If powerDownEnable Then
        PowerDown_Parallel powerUpPins, All_DigitalPinlist_Disc, , True
    End If
    
    If TheExec.sites.Active.Count = 0 Then
        If gl_EnableCurrentProfile Or gl_EnableVoltageProfile Then
            TheExec.RunOptions.DoAll = DoAll_save
            TheExec.RunOptions.OverrideFailStop = OverRide_FailStop
        End If
    End If
        
    If glb_Boolean_export = True Then
        glb_Boolean_export = False
        Close #13
    End If
    

#If LCD = True Then
    myUtility.PreShutDownPinConditionSetting
#End If
    
    '230328 Print Harv
    For Each v_site In TheExec.sites
        If (TheExec.Flow.SiteFlag(v_site, "F_PrintHarvReport") = 1) Then

            Call TheExec.Datalog.SetDynamicTestName("ATE_STR_Summary", False)                                   '' ssign Instance name to be "ATE_STR_Summary"
            'TheExec.Flow.TestLimit resultVal:=0, lowVal:=0, hiVal:=0, Tname:="Dummy", ForceResults:=tlForceNone, TNum:="100"  '' 230518 move to the beginning of OnPreShutDownSite. this line will be bypass if run TheHdw.Alarms.Check before
            TheExec.sites.item(v_site).TestNumber = 1000000    '' should same as test number of ATR_STR_Summary
    
            TheExec.Datalog.WriteComment "Start to Print out Harvest Summary at PowerDown."
            TheExec.Datalog.WriteComment "<ATE_STR_Summary>"
            Call TheExec.Flow.instance("ATE_STR_Summary").Execute
            TheExec.Datalog.WriteComment "End Print out Harvest Summary at PowerDown."
            
            TheExec.sites(v_site).FlagState("F_PrintHarvReport") = logicFalse
        End If
    Next v_site

    
    
    Exit Function
errHandler: 'Add ErrHandler 2023/08/18
    HandleExecIPError "OnPreShutDownSite"
    If AbortTest Then Exit Function Else Resume Next 'Add ErrHandler 2023/08/18
End Function
 
' Use TheExec.Sites.SiteNumber to determine which site is being disconnected.
' Immediately after a site is disconnected.
Function OnPostShutDownSite()
On Error GoTo errHandler 'Add ErrHandler 2023/08/18
        
    ' Put code here
    ''TheHdw.DIB.powerOn = False 'For Debug DIB power alarm issue.
        ''20220119, Avoid different behavior between UP and UFP after click "Debug_stop"
        
    Exit Function
errHandler: 'Add ErrHandler 2023/08/18
    HandleExecIPError "OnPostShutDownSite"
    If AbortTest Then Exit Function Else Resume Next 'Add ErrHandler 2023/08/18
End Function
 
' Immediately befoe any new calibration factors are loaded
' or new calibrations run.  Not called if no action is taken during AutoCal.
Function OnAutoCalStarted()
On Error GoTo errHandler 'Add ErrHandler 2023/08/18

    ' Put code here
    
    
    Exit Function
errHandler: 'Add ErrHandler 2023/08/18
    HandleExecIPError "OnAutoCalStarted"
    If AbortTest Then Exit Function Else Resume Next 'Add ErrHandler 2023/08/18
End Function

' Immediately after AutoCal has completed.
' Not called no action has been taken (new factors loaded, or cal performed).
Function OnAutoCalCompleted()
On Error GoTo errHandler 'Add ErrHandler 2023/08/18

    ' Put code here
    
    
    Exit Function
errHandler: 'Add ErrHandler 2023/08/18
    HandleExecIPError "OnAutoCalCompleted"
    If AbortTest Then Exit Function Else Resume Next 'Add ErrHandler 2023/08/18
End Function


' Called right before an alarm is reported
' The alarmList is a tab delimited string of alarm error messages
Function OnAlarmOccurred(alarmList As String)
On Error GoTo errHandler 'Add ErrHandler 2023/08/18
    
'    UNCOMMENT TO THE FOLLOWING LINES TO PARSE ALARMS
    Dim i As Long
    Dim j As Long
    Dim alarmArray() As String
    Dim S As Long
    Dim Count As Long
    Dim CaptureSiteLong As Long
    Dim alarmArray_Split() As String
    Dim site As Variant 'Carter, 20240304
    alarmArray = Split(alarmList, vbTab)
    Count = TheExec.sites.Existing.Count
    CaptureSiteLong = Len(CStr(Count - 1))
    ' This will loop through all the alarms
    For i = LBound(alarmArray) To UBound(alarmArray)
        alarmArray(i) = LCase(alarmArray(i))
        If InStr(1, alarmArray(i), "site") <> 0 Then   '''' We need to add
            alarmArray(i) = Replace(alarmArray(i), "]", "")
            alarmArray(i) = Replace(alarmArray(i), " ", "")
            S = CLng(mid(alarmArray(i), InStr(1, alarmArray(i), "site") + 4, CaptureSiteLong))   ''''alarm message = "pwr1.site 0"
            If S >= 0 And S < Count Then
                alarmFail(S) = True
            Else
                TheExec.Datalog.WriteComment "<WARNING> OnAlarmOccurred without define site, setup all site alarm"
                If isDebugMode Then TheExec.AddOutput "<WARNING> OnAlarmOccurred without define site, setup all site alarm"
            
                For Each site In TheExec.sites
                    TheExec.sites.item(site).SortNumber = 999
                    TheExec.sites.item(site).BinNumber = 9
                    TheExec.sites.item(site).testResult = siteFail
                Next site
            End If
        Else
            TheExec.Datalog.WriteComment "<WARNING> OnAlarmOccurred without define site, setup all site alarm"
            If isDebugMode Then TheExec.AddOutput "<WARNING> OnAlarmOccurred without define site, setup all site alarm"
        
            For Each site In TheExec.sites
                TheExec.sites.item(site).SortNumber = 999
                TheExec.sites.item(site).BinNumber = 9
                TheExec.sites.item(site).testResult = siteFail
            Next site
        End If  '''' We need to add
    Next i

    Exit Function
errHandler: 'Add ErrHandler 2023/08/18
    HandleExecIPError "OnAlarmOccurred"
    If AbortTest Then Exit Function Else Resume Next 'Add ErrHandler 2023/08/18
End Function
' When the user pressed the VB Stop button, this interpose function would be called after OnPostShutDownSite was called.
' The user would put code here to make sure global variable are created and contain the correct data.
Function OnGlobalVariableReset()
On Error GoTo errHandler 'Add ErrHandler 2023/08/18

    ' Put code here
    Call Process_DSSC_Dictionary ' Added By Oscar 180523 For Read DigSrc assignment from Sheet(Function is in LIB_HardIP)
    'Call HardIP_OnProgramStarted_Process
    
    Exit Function
errHandler: 'Add ErrHandler 2023/08/18
    HandleExecIPError "OnGlobalVariableReset"
    If AbortTest Then Exit Function Else Resume Next 'Add ErrHandler 2023/08/18
End Function

' Immediately once Vaildation get started
' [20230809][All][Oliver] initial current profile variables to avoid execute when validating
' [20230907][All][Rigo] Add LCD_OnValidationStart
Function OnValidationStart()
On Error GoTo errHandler 'Add ErrHandler 2023/08/18

    gl_EnableCurrentProfile = False
    gl_EnableVoltageProfile = False
    glb_ProfileFilter = False
    
    Call Check_Program_Name
    
#If LCD = True Then
    Call LCD_OnValidationStart
#End If


    Exit Function
errHandler: 'Add ErrHandler 2023/08/18
    HandleExecIPError "OnValidationStart"
    If AbortTest Then Exit Function Else Resume Next 'Add ErrHandler 2023/08/18
End Function
' Immediately at the conclusion of the workbook close process. The function is called in any of the following options,
' File->Close
' File->Exit
' Directly triggered the close (?X?) button of the workbook.
Function OnProgramClose()
On Error GoTo errHandler 'Add ErrHandler 2023/08/18

    ' Put code here
    thehdw.DIB.powerOn = False  'reset DIB power


    Exit Function
errHandler: 'Add ErrHandler 2023/08/18
    HandleExecIPError "OnProgramClose"
    If AbortTest Then Exit Function Else Resume Next 'Add ErrHandler 2023/08/18
End Function


Public Function Ignore_nWire_Error()
On Error GoTo errHandler 'Add ErrHandler 2023/08/18
    If glb_TesterType = "Jaguar" Then
        TheExec.Error.Behavior("HSDMPI:1335") = tlErrorIgnore
        TheExec.Error.Behavior("HSDMPI:0109") = tlErrorIgnore
    End If
    TheExec.Error.Behavior("NWirePI:0074") = tlErrorIgnore
Exit Function 'Add ErrHandler 2023/08/18
errHandler: 'Add ErrHandler 2023/08/18
    Call Print_Error_Message(Error_Info, "Exec_IP_Module", "Ignore_nWire_Error") 'Add ErrHandler 2023/08/18
    If AbortTest Then Exit Function Else Resume Next 'Add ErrHandler 2023/08/18
End Function


Public Function Init_DIB_Power()
On Error GoTo errHandler 'Add ErrHandler 2023/08/18
    'TheHdw.DIB.powerOn = False  'reset DIB power
    thehdw.DIB.Power.item("12V").State = tlOn
    thehdw.DIB.Power.item("5V_1").State = tlOn
    thehdw.DIB.Power.item("5V_2").State = tlOn
    thehdw.DIB.Power.item("3.3V").State = tlOn
Exit Function 'Add ErrHandler 2023/08/18
errHandler: 'Add ErrHandler 2023/08/18
    Call Print_Error_Message(Error_Info, "Exec_IP_Module", "Init_DIB_Power") 'Add ErrHandler 2023/08/18
    If AbortTest Then Exit Function Else Resume Next 'Add ErrHandler 2023/08/18
End Function
