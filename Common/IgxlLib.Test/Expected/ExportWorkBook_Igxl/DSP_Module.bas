Attribute VB_Name = "DSP_Module"
Option Explicit
Private gDB_SerialType_ As New SiteBoolean
Private gDL_BitsPerRow_ As New SiteLong
Private gDL_ReadCycles_ As New SiteLong
Private gDL_BitsPerCycle_ As New SiteLong
Private gDL_BitsPerBlock_ As New SiteLong
Private gDL_TotalBits_ As New SiteLong
Private gDL_DigSrcRepeatN_ As New SiteLong
Private gDD_BaseVoltage_ As New SiteDouble
Private gDD_BaseStepVoltage_ As New SiteDouble
Private gDL_CRC_EndBit_ As New SiteLong
' This module should be used only for DSP Procedure code.  Functions in this
' module will be available to be called to perform DSP in all DSP modes.
' Additional modules may be added as needed (all starting with "DSP_").
'
' The required signature for a DSP Procedure is:
'
' Public Function FuncName(<arglist>) as Long
'   where <arglist> is any list of arguments supported by DSP code.
'
' See online help for supported types and other restrictions.
Public Function test1(ByVal Init_Pat1 As Pattern, ByVal Init_Pat2 As Pattern, _
                             ByVal Init_Pat3 As Pattern, ByVal Init_Pat4 As Pattern, _
                             ByVal Init_Pat5 As Pattern, ByVal Init_Pat6 As Pattern, _
                             ByVal Init_Pat7 As Pattern, ByVal Init_Pat8 As Pattern, _
                             ByVal Init_Pat9 As Pattern, ByVal Init_Pat10 As Pattern, _
                             ByVal CGVoltage As String, _
                             SRMType As String, PatternTimeout As String) As Long
On Error Resume Next
   
End Function
Public Function test2(ByVal Init_Pat1 As Pattern, ByVal Init_Pat2 As Pattern, _
                             ByVal Init_Pat3 As Pattern, ByVal Init_Pat4 As Pattern, _
                             ByVal Init_Pat5 As Pattern, ByVal Init_Pat6 As Pattern, _
                             ByVal Init_Pat7 As Pattern, ByVal Init_Pat8 As Pattern, _
                             ByVal Init_Pat9 As Pattern, ByVal Init_Pat10 As Pattern, _
                             ByVal CGVoltage As String, _
                             SRMType As String, PatternTimeout As String) As Long
End Function
