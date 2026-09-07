Attribute VB_Name = "Exec_IP_Module"
#Const isUFP = False
#Const LCD = False
#Const RF = False
Option Explicit
'Public write_SPIROM_CheckSum As Integer
#Const AP = True
Public write_spirom As New SiteBoolean
Public Mbist_Repair_CompareType As Variant  'for Mbist finger print

Public DoAll_save As Boolean
Public OverRide_FailStop As Boolean




' Immediately at the conclusion of the initialization process.
' Do not program test system hardware from this function.
Function OnTesterInitialized()
    On Error GoTo errHandler

    ' Put code here
    
    
    Exit Function
errHandler:
    ' OnTesterInitialized executes before TheExec is even established so nothing
    ' better to do then msgbox in this case.  Note that unhandled errors can allow the
    ' user to press "End" which will result in a DataTool crash.  Errors in this routine
    ' need to be debugged carefully.
    MsgBox "Error encountered in Exec Interpose Function OnTesterInitialized" + vbCrLf + _
        "VBT Error # " + Trim(str(err.number)) + ": " + err.Description
                If AbortTest Then Exit Function Else Resume Next
End Function
 
' Immediately at the conclusion of the load process.
' Do not program test system hardware from this function.
Function OnProgramLoaded()

    On Error GoTo errHandler
    
'    glb_TesterType = TheHdw.Tester.type
'
'    'for TERA1 encryption need, add following code
'    If Not TheExec.SoftwareVersion Like "8.10.90_uflx*" Then
'        m_cpcmodule.SuppressCheckForUnProtectedPatterns = True
'    End If
'
'    'for 8.30 encryption need, add following code
'    m_STDSvcClient.CPCModule.SuppressCheckForUnProtectedPatterns = True
'
'    If TheExec.SoftwareVersion Like "*9.10*" Then
'        CallByName TheExec.TestProgram, "MemoryLimitCheckEnabled", VbLet, False
'    End If
'
'
'    If is_reference_installed("Scripting") = False Then
'        Application.ActiveWorkbook.VBProject.References.AddFromFile "C:\WINDOWS\system32\scrrun.dll"
'    End If
'
'     If is_reference_installed("VBScript_RegExp_55") = False Then
'        Application.VBE.ActiveVBProject.References.AddFromFile "C:\WINDOWS\system32\vbscript.dll\3"
'    End If

    Exit Function
errHandler:
    HandleExecIPError "OnProgramLoaded"
    If AbortTest Then Exit Function Else Resume Next
End Function
 
' Immediately at the conclusion of the validate process. Called only if validation succeeds.
Function OnProgramValidated()

    On Error GoTo errHandler
    
'    currentJobName = LCase(TheExec.CurrentJob)
'    CurrentChannelMap = LCase(TheExec.CurrentChanMap)
'
'    'CharStoreResultsUntilNextRun, clear shmoo momory to prevent crash
'    If LCase(TheExec.CurrentJob) Like "char*" Or TheExec.Flow.enableWord("Shmoo_BringUp") Then
'        TheExec.DevChar.Configuration.Features.item(tlDevCharFeature_StoreResultsUntilNextRun).Enabled = False
'        m_STDSvcClient.SelfTest.MemoryCollectRunInterval = 1
'    End If
        
    Exit Function
errHandler:
    HandleExecIPError "OnProgramValidated"
    If AbortTest Then Exit Function Else Resume Next
End Function
 
' Immediately at the conclusion of the validate process. Called only if validation fails.
Function OnProgramFailedValidation()
    On Error GoTo errHandler

    ' Put code here
    
    
    Exit Function
errHandler:
    HandleExecIPError "OnProgramFailedValidation"
    If AbortTest Then Exit Function Else Resume Next
End Function
 
' Immediately at the conclusion of the user DIB calibration process (previously
' known as the TDR calibration process). Called only if user DIB calibration succeeds.
Function OnTDRCalibrated()
    On Error GoTo errHandler
    
       
    Exit Function
errHandler:
    HandleExecIPError "OnTDRCalibrated"
    If AbortTest Then Exit Function Else Resume Next
End Function
 
' Immediately after "pre-job reset" when the test program starts.
' Note that "first run" actions can be enclosed in
' If TheExec.ExecutionCount = 0 Then...
' (see online help for ExecutionCount)
Function OnProgramStarted()
    On Error GoTo errHandler
    Dim i As Integer, j As Integer
    
    
    'For STC mini correlation
    ' TheExec.Flow.enableWord("BypRetestCheck") = True
    ' TheExec.Flow.enableWord("DRAM_POPB") = True
    ' TheExec.Flow.enableWord("eFuse_All_Enable") = True
    ' TheExec.Flow.enableWord("eFuse_FT_Wafer_PseudoFuse") = True
    ' TheExec.Flow.enableWord("MbistEfuse") = True
    ' TheExec.Flow.enableWord("Pgm2File") = True
    ' TheExec.Flow.enableWord("RUN_CP_IN_FT") = True
    ' TheExec.Flow.enableWord("CFG_Early_Enable") = False
    
    ' TheExec.Flow.enableWord("LogLevel_NPI") = True
    ' setXY 11, 11
    'For STC mini correlation
        
'    glb_TesterType = TheHdw.Tester.type
'    TestProgram_Day_Code = CStr(Year(Now)) & right("0" & CStr(Month(Now)), 2) & right("0" & CStr(Day(Now)), 2)
'    TestProgram_Day_Code = TestProgram_Day_Code & right("0" & CStr(Hour(Now)), 2) & right("0" & CStr(Minute(Now)), 2) & right("0" & CStr(Second(Now)), 2)
'
'    currentJobName = LCase(TheExec.CurrentJob) ''Carter, 20191115
'    CurrentChannelMap = LCase(TheExec.CurrentChanMap)
'
'    Find_nWire_Pin   '''update for multiple nWire CLK, 2017/07/18

    Exit Function
errHandler:
    HandleExecIPError "OnProgramStarted"
    If AbortTest Then Exit Function Else Resume Next
End Function
 

' Immediately before "post-job reset" when the test program completes.
' Note that any actions taken here with respect to modification of binning
' will affect the binning sent to the Operator Interface, but will not affect
' the binning reported in Datalog.
Function OnProgramEnded()
    On Error GoTo errHandler

    Exit Function
errHandler:
    HandleExecIPError "OnProgramEnded"
    If AbortTest Then Exit Function Else Resume Next
End Function
 
' Immediately before a site is disconnected.
' Use TheExec.Sites.SiteNumber to determine which site is being disconnected.
Function OnPreShutDownSite()
    On Error GoTo errHandler
'    Dim v_site As Variant
'    Dim powerDownEnable As Boolean
'
'    If TheExec.RunOptions.DoAll = False And ATE_STR_Summary_Table_Parse_Flag = True Then
'        Call ATE_STR_Summary_Flag_Operate
'    End If
'    ' Put code here
'    TheExec.Flow.TestLimit resultVal:=0, lowVal:=0, hiVal:=0, Tname:="Dummy", ForceResults:=tlForceNone, tNum:="100"  '' to effective SetDynamicTestName setting
'
'    '230328 Print Harv
'    For Each Site In TheExec.Sites
'        If (TheExec.Flow.SiteFlag(Site, "F_PrintHarvReport") = 1) Then
'
'            Call TheExec.Datalog.SetDynamicTestName("ATE_STR_Summary", False)                                   '' ssign Instance name to be "ATE_STR_Summary"
'            'TheExec.Flow.TestLimit resultVal:=0, lowVal:=0, hiVal:=0, Tname:="Dummy", ForceResults:=tlForceNone, TNum:="100"  '' 230518 move to the beginning of OnPreShutDownSite. this line will be bypass if run TheHdw.Alarms.Check before
'            TheExec.Sites.item(Site).TestNumber = 1000000    '' should same as test number of ATR_STR_Summary
'
'            TheExec.Datalog.WriteComment "Start to Print out Harvest Summary at PowerDown."
'            TheExec.Datalog.WriteComment "<ATE_STR_Summary>"
'            Call TheExec.Flow.instance("ATE_STR_Summary").Execute
'            TheExec.Datalog.WriteComment "End Print out Harvest Summary at PowerDown."
'
'            TheExec.Sites(Site).FlagState("F_PrintHarvReport") = logicFalse
'        End If
'   Next Site
'
'
'    Exit Function
'errHandler:
'    HandleExecIPError "OnPreShutDownSite"
'    If AbortTest Then Exit Function Else Resume Next
'End Function
'
'' Use TheExec.Sites.SiteNumber to determine which site is being disconnected.
'' Immediately after a site is disconnected.
'Function OnPostShutDownSite()
'    On Error GoTo errHandler
'
'    ' Put code here
'    ''TheHdw.DIB.powerOn = False 'For Debug DIB power alarm issue.
'        ''20220119, Avoid different behavior between UP and UFP after click "Debug_stop"
'
    Exit Function
errHandler:
    HandleExecIPError "OnPostShutDownSite"
    If AbortTest Then Exit Function Else Resume Next
End Function
 
' Immediately befoe any new calibration factors are loaded
' or new calibrations run.  Not called if no action is taken during AutoCal.
Function OnAutoCalStarted()
    On Error GoTo errHandler

    ' Put code here
    
    
    Exit Function
errHandler:
    HandleExecIPError "OnAutoCalStarted"
    If AbortTest Then Exit Function Else Resume Next
End Function

' Immediately after AutoCal has completed.
' Not called no action has been taken (new factors loaded, or cal performed).
Function OnAutoCalCompleted()
    On Error GoTo errHandler

    ' Put code here
    
    
    Exit Function
errHandler:
    HandleExecIPError "OnAutoCalCompleted"
    If AbortTest Then Exit Function Else Resume Next
End Function


' Called right before an alarm is reported
' The alarmList is a tab delimited string of alarm error messages
Function OnAlarmOccurred(alarmList As String)
    On Error GoTo errHandler
    
                
    Exit Function
errHandler:
    HandleExecIPError "OnAlarmOccurred"
    If AbortTest Then Exit Function Else Resume Next
End Function
' When the user pressed the VB Stop button, this interpose function would be called after OnPostShutDownSite was called.
' The user would put code here to make sure global variable are created and contain the correct data.
Function OnGlobalVariableReset()
    On Error GoTo errHandler

    Exit Function
errHandler:
    HandleExecIPError "OnGlobalVariableReset"
    If AbortTest Then Exit Function Else Resume Next
End Function

' Immediately once Vaildation get started
Function OnValidationStart()
    On Error GoTo errHandler

    Exit Function
errHandler:
    HandleExecIPError "OnValidationStart"
    If AbortTest Then Exit Function Else Resume Next
End Function
' Immediately at the conclusion of the workbook close process. The function is called in any of the following options,
' File->Close
' File->Exit
' Directly triggered the close (?X?) button of the workbook.
Function OnProgramClose()
    On Error GoTo errHandler

    ' Put code here
    TheHdw.DIB.powerOn = False  'reset DIB power


    Exit Function
errHandler:

    HandleExecIPError "OnProgramClose"
    If AbortTest Then Exit Function Else Resume Next

End Function


'Public Function Init_DIB_Power()
'On Error GoTo errHandler 'Add ErrHandler 2023/08/18
'    'TheHdw.DIB.powerOn = False  'reset DIB power
'    TheHdw.DIB.power.item("12V").State = tlOn
'    TheHdw.DIB.power.item("5V_1").State = tlOn
'    TheHdw.DIB.power.item("5V_2").State = tlOn
'    TheHdw.DIB.power.item("3.3V").State = tlOn
'Exit Function 'Add ErrHandler 2023/08/18
'errHandler: 'Add ErrHandler 2023/08/18
'    Call Print_Error_Message(Error_Info, "Exec_IP_Module", "Init_DIB_Power") 'Add ErrHandler 2023/08/18
'    If AbortTest Then Exit Function Else Resume Next 'Add ErrHandler 2023/08/18
'End Function
'
'
'Public Sub SetInitialTrace(PinGroup As String)
'
'Dim SplitPins() As String
'
'Dim PinCnt As Long
'Dim m As Integer
'Dim Channels As String
'Dim Site As Variant
'TheExec.DataManager.ReturnSignalNames = True
'''thehdw.Digital.Calibration.DIB.InitializeCalData
'For Each Site In TheExec.Sites.Existing
'
'    TheExec.DataManager.DecomposePinList PinGroup, SplitPins(), PinCnt
'
'    For m = 0 To PinCnt - 1
'        TheExec.DataManager.GetChannelStringFromPinAndSite SplitPins(m), Site, Channels
'        If Channels <> "" Then
'        '''-----------------UFP-----------------
'        If glb_TesterType = "Jaguar" Then
'            TheHdw.Digital.Calibration.Channels(Channels).DIB.trace = 0.000000002
'#If isUFP = True Then
'        ElseIf glb_TesterType = "UltraFLEXplus" Then
'            TheHdw.Calibration.DIB.Channels(Channels).trace = 0.000000002
'                        '''211221 Follow Hatkar's finding that add this command to prevent trace value re-write issue
''''''                        TheHdw.Calibration.DIB.Traces.Apply
'#End If
'        End If
'        '''-----------------UFP-----------------
'        End If
'    Next m
'
'Next Site
'
'End Sub
