VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "IExecInterposeFunctions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' Immediately at the conclusion of the initialization process.
' Do not program test system hardware from this function.
Function OnTesterInitialized()
  
End Function
 ' Immediately at the conclusion of the load process.
' Do not program test system hardware from this function.
Function OnProgramLoaded()

End Function
 
' Immediately at the conclusion of the validate process. Called only if validation succeeds.
Function OnProgramValidated()
  
End Function
 
' Immediately at the conclusion of the validate process. Called only if validation fails.
Function OnProgramFailedValidation()

End Function
 
' Immediately at the conclusion of the user DIB calibration process (previously
' known as the TDR calibration process). Called only if user DIB calibration succeeds.
Function OnTDRCalibrated()
  
End Function
 
' Immediately after "pre-job reset" when the test program starts.
' Note that "first run" actions can be enclosed in
' If TheExec.ExecutionCount = 0 Then...
' (see online help for ExecutionCount)
Function OnProgramStarted()
   
End Function
 
' Immediately before "post-job reset" when the test program completes.
' Note that any actions taken here with respect to modification of binning
' will affect the binning sent to the Operator Interface, but will not affect
' the binning reported in Datalog.
Function OnProgramEnded()
    
End Function
 
' Immediately before a site is disconnected.
' Use TheExec.Sites.SiteNumber to determine which site is being disconnected.
Function OnPreShutDownSite()
   
End Function
 
' Use TheExec.Sites.SiteNumber to determine which site is being disconnected.
' Immediately after a site is disconnected.
Function OnPostShutDownSite()
 
End Function
 
' Immediately befoe any new calibration factors are loaded
' or new calibrations run.  Not called if no action is taken during AutoCal.
Function OnAutoCalStarted()
    
End Function

' Immediately after AutoCal has completed.
' Not called no action has been taken (new factors loaded, or cal performed).
Function OnAutoCalCompleted()
   
End Function


' Called right before an alarm is reported
' The alarmList is a tab delimited string of alarm error messages
Function OnAlarmOccurred(alarmList As String)

   
End Function

' When the user pressed the VB Stop button, this interpose function would be called after OnPostShutDownSite was called.
' The user would put code here to make sure global variable are created and contain the correct data.
Function OnGlobalVariableReset()
   
End Function

' Immediately once Vaildation get started
Function OnValidationStart()
    
End Function
' Immediately at the conclusion of the workbook close process. The function is called in any of the following options,
' File->Close
' File->Exit
' Directly triggered the close (“X”) button of the workbook.
Function OnProgramClose()
   
End Function