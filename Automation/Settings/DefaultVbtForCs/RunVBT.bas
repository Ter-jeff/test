Attribute VB_Name = "RunVBT"
' This ALWAYS GENERATED file contains wrappers for VBT tests.
' Do not edit.

Public Sub COD_Check()

End Sub

Private Sub HandleUntrappedError()
    ' Sanity clause
    If TheExec Is Nothing Then
        MsgBox "IG-XL is not running!  VBT tests cannot execute unless IG-XL is running."
        Exit Sub
    End If
    ' If the last site has failed out, let's ignore the error
    If TheExec.Sites.Active.Count = 0 Then Exit Sub  ' don't log the error
    ' If in a legacy site loop, make sure to complete it. (For-Each site syntax in IG-XL 6.10 aborts gracefully.)
    Do While TheExec.Sites.InSiteLoop
        Call TheExec.Sites.SelectNext(loopTop) '  Legacy syntax (hidden)
    Loop
    ' Select all active sites in case a subset of sites was selected when error occurred.
    TheExec.Sites.Selected = TheExec.Sites.Active
    ' Log the error to the IG-XL Error logging mechanism (tells Flow to fail the test)
    AbortTest
End Sub

Public Function Empty_T__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    Dim p1 As New InterposeName
    p1.value = v(0)
    Dim p2 As New InterposeName
    p2.value = v(1)
    Dim p3 As New InterposeName
    p3.value = v(2)
    Dim p4 As New InterposeName
    p4.value = v(3)
    Dim p5 As New InterposeName
    p5.value = v(4)
    Dim p6 As New InterposeName
    p6.value = v(5)
    Dim p7 As New PinList
    p7.value = v(12)
    Dim p8 As New PinList
    p8.value = v(13)
    Dim p9 As New PinList
    p9.value = v(14)
    Dim p10 As New PinList
    p10.value = v(15)
    Dim p11 As New PinList
    p11.value = v(16)
    Dim p12 As New PinList
    p12.value = v(17)
    Dim p13 As New PinList
    p13.value = v(18)
    Dim pStep As SubType
    pStep = TheExec.flow.StepType
    Empty_T__ = Template.VBT_Empty_T.Empty_T(p1, p2, p3, p4, p5, p6, CStr(v(6)), CStr(v(7)), CStr(v(8)), CStr(v(9)), CStr(v(10)), CStr(v(11)), p7, p8, p9, p10, p11, p12, p13, pStep, CBool(v(19)))
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function









































Public Function Functional_T__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    Dim p1 As New Pattern
    p1.value = v(0)
    Dim p2 As New InterposeName
    p2.value = v(1)
    Dim p3 As New InterposeName
    p3.value = v(2)
    Dim p4 As New InterposeName
    p4.value = v(3)
    Dim p5 As New InterposeName
    p5.value = v(4)
    Dim p6 As New InterposeName
    p6.value = v(5)
    Dim p7 As New InterposeName
    p7.value = v(6)
    Dim p8 As PFType
    p8 = v(7)
    Dim p9 As tlResultMode
    p9 = v(8)
    Dim p10 As New PinList
    p10.value = v(9)
    Dim p11 As New PinList
    p11.value = v(10)
    Dim p12 As New PinList
    p12.value = v(11)
    Dim p13 As New PinList
    p13.value = v(12)
    Dim p14 As New PinList
    p14.value = v(13)
    Dim p15 As New PinList
    p15.value = v(20)
    Dim p16 As New PinList
    p16.value = v(21)
    Dim p17 As New InterposeName
    p17.value = v(22)
    Dim p18 As tlRelayMode
    p18 = v(24)
    Dim p19 As tlWaitVal
    p19 = v(27)
    Dim p20 As tlWaitVal
    p20 = v(28)
    Dim p21 As tlWaitVal
    p21 = v(29)
    Dim p22 As tlWaitVal
    p22 = v(30)
    Dim pStep As SubType
    pStep = TheExec.flow.StepType
    Dim p23 As tlPatConcurrentMode
    p23 = v(34)
    Dim p24 As tlTemplateScanFailDataLogging
    p24 = v(35)
    Dim p25 As tlDigitalCMEMCaptureLimitMode
    p25 = v(36)
    Dim p26 As tlTemplateScanPinListSource
    p26 = v(38)
    Dim p27 As New PinList
    p27.value = v(39)
    Dim p28 As tlTemplateScanCaptureFormat
    p28 = v(40)
    Dim p29 As tlTemplateScanCaptureDataType
    p29 = v(41)
    Dim p30 As tlTemplateScanUserCommentSource
    p30 = v(42)
    Dim p31 As tlTemplateATPGPinMapSource
    p31 = v(44)
    Functional_T__ = Template.VBT_Functional_T.Functional_T(p1, p2, p3, p4, p5, p6, p7, p8, p9, p10, p11, p12, p13, p14, CStr(v(14)), CStr(v(15)), CStr(v(16)), CStr(v(17)), CStr(v(18)), CStr(v(19)), p15, p16, p17, CStr(v(23)), p18, CBool(v(25)), CBool(v(26)), p19, p20, p21, p22, CBool(v(UBound(v))), CStr(v(32)), pStep, CStr(v(33)), p23, p24, p25, CLng(v(37)), p26, p27, p28, p29, p30, CStr(v(43)), p31, CStr(v(45)), CStr(v(46)), CBool(v(47)))
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function









































Public Function PinPMU_T__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    Dim p1 As New InterposeName
    p1.value = v(1)
    Dim p2 As New InterposeName
    p2.value = v(2)
    Dim p3 As New InterposeName
    p3.value = v(3)
    Dim p4 As New InterposeName
    p4.value = v(4)
    Dim p5 As New InterposeName
    p5.value = v(5)
    Dim p6 As New InterposeName
    p6.value = v(6)
    Dim p7 As New Pattern
    p7.value = v(7)
    Dim p8 As New Pattern
    p8.value = v(8)
    Dim p9 As New PinList
    p9.value = v(10)
    Dim p10 As New PinList
    p10.value = v(11)
    Dim p11 As New PinList
    p11.value = v(12)
    Dim p12 As New PinList
    p12.value = v(13)
    Dim p13 As New PinList
    p13.value = v(14)
    Dim p14 As New PinList
    p14.value = v(15)
    Dim p15 As tlPPMUMode
    p15 = v(16)
    Dim p16 As New FormulaArg
    p16.value = v(18)
    Dim p17 As New FormulaArg
    p17.value = v(19)
    Dim p18 As tlPPMURelayMode
    p18 = v(20)
    Dim p19 As New PinList
    p19.value = v(36)
    Dim p20 As New PinList
    p20.value = v(37)
    Dim p21 As tlWaitVal
    p21 = v(38)
    Dim p22 As tlWaitVal
    p22 = v(39)
    Dim p23 As tlWaitVal
    p23 = v(40)
    Dim p24 As tlWaitVal
    p24 = v(41)
    Dim p25 As tlPPMUMode
    p25 = v(49)
    Dim p26 As New FormulaArg
    p26.value = v(52)
    Dim pStep As SubType
    pStep = TheExec.flow.StepType
    Dim p27 As New PinList
    p27.value = v(53)
    Dim p28 As tlPPMUMode
    p28 = v(54)
    Dim p29 As New FormulaArg
    p29.value = v(55)
    PinPMU_T__ = Template.VBT_PinPmu_T.PinPMU_T(CStr(v(0)), p1, p2, p3, p4, p5, p6, p7, p8, CStr(v(9)), p9, p10, p11, p12, p13, p14, p15, CDbl(v(17)), p16, p17, p18, CStr(v(21)), CStr(v(22)), CStr(v(23)), CStr(v(24)), CStr(v(25)), CStr(v(26)), CStr(v(27)), CStr(v(28)), CStr(v(29)), CStr(v(30)), CDbl(v(31)), CLng(v(32)), CBool(v(33)), CStr(v(34)), CStr(v(35)), p19, p20, p21, p22, p23, p24, CBool(v(UBound(v))), CStr(v(43)), CStr(v(44)), , CStr(v(45)), CBool(v(46)), CBool(v(47)), CBool(v(48)), p25, CStr(v(50)), CStr(v(51)), p26, pStep, p27, p28, p29, CStr(v(56)), CStr(v(57)))
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function









































Public Function DCVSPowerSupply_T__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    Dim p1 As New Pattern
    p1.value = v(0)
    Dim p2 As New InterposeName
    p2.value = v(1)
    Dim p3 As New InterposeName
    p3.value = v(2)
    Dim p4 As New InterposeName
    p4.value = v(3)
    Dim p5 As New InterposeName
    p5.value = v(4)
    Dim p6 As New InterposeName
    p6.value = v(5)
    Dim p7 As New InterposeName
    p7.value = v(6)
    Dim p8 As New Pattern
    p8.value = v(7)
    Dim p9 As New PinList
    p9.value = v(8)
    Dim p10 As New PinList
    p10.value = v(9)
    Dim p11 As New PinList
    p11.value = v(10)
    Dim p12 As New PinList
    p12.value = v(11)
    Dim p13 As New PinList
    p13.value = v(12)
    Dim p14 As New PinList
    p14.value = v(16)
    Dim p15 As tlPSSource
    p15 = v(17)
    Dim p16 As tlCommonRelayMode
    p16 = v(31)
    Dim p17 As New PinList
    p17.value = v(32)
    Dim p18 As New PinList
    p18.value = v(33)
    Dim p19 As tlPSTestControl
    p19 = v(34)
    Dim p20 As tlWaitVal
    p20 = v(35)
    Dim p21 As tlWaitVal
    p21 = v(36)
    Dim p22 As tlWaitVal
    p22 = v(37)
    Dim p23 As tlWaitVal
    p23 = v(38)
    Dim p24 As New FormulaArg
    p24.value = v(40)
    Dim p25 As New FormulaArg
    p25.value = v(41)
    Dim p26 As New FormulaArg
    p26.value = v(42)
    Dim p27 As New FormulaArg
    p27.value = v(43)
    Dim pStep As SubType
    pStep = TheExec.flow.StepType
    Dim p28 As New FormulaArg
    p28.value = v(47)
    DCVSPowerSupply_T__ = Template.VBT_DCVSPowerSupply_T.DCVSPowerSupply_T(p1, p2, p3, p4, p5, p6, p7, p8, p9, p10, p11, p12, p13, CDbl(v(13)), CLng(v(14)), CStr(v(15)), p14, p15, CStr(v(18)), CStr(v(19)), CStr(v(20)), CStr(v(21)), CStr(v(22)), CStr(v(23)), CStr(v(24)), CStr(v(25)), CStr(v(26)), CStr(v(27)), CStr(v(28)), CStr(v(29)), CBool(v(30)), p16, p17, p18, p19, p20, p21, p22, p23, CBool(v(UBound(v))), p24, p25, p26, p27, , CStr(v(44)), CBool(v(45)), CBool(v(46)), pStep, p28, CBool(v(48)))
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function









































Public Function CTO_ADC_AC_T__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    Dim p1 As New PinList
    p1.value = v(0)
    Dim p2 As New PinList
    p2.value = v(1)
    Dim p3 As New PinList
    p3.value = v(2)
    Dim p4 As New PinList
    p4.value = v(3)
    Dim p5 As tlACEncoding
    p5 = v(7)
    Dim p6 As tlACVoltageRange
    p6 = v(8)
    Dim p7 As tlRelayMode
    p7 = v(9)
    Dim p8 As tlCTOSourceBandwidth
    p8 = v(12)
    Dim p9 As New Pattern
    p9.value = v(20)
    Dim p10 As New Pattern
    p10.value = v(42)
    Dim p11 As New PinList
    p11.value = v(47)
    Dim p12 As New PinList
    p12.value = v(48)
    Dim p13 As New PinList
    p13.value = v(49)
    Dim p14 As New PinList
    p14.value = v(50)
    Dim p15 As New PinList
    p15.value = v(51)
    Dim p16 As New PinList
    p16.value = v(52)
    Dim p17 As New PinList
    p17.value = v(53)
    Dim p18 As New InterposeName
    p18.value = v(54)
    Dim ExtraArgs(0 To 10) As Variant
    Dim i As Integer
    For i = 0 To 10
        ExtraArgs(i) = v(56 + i)
    Next i
    Dim pStep As SubType
    pStep = TheExec.flow.StepType
    CTO_ADC_AC_T__ = Template.VBT_CTO_ADC_AC_T.CTO_ADC_AC_T(p1, p2, p3, p4, CDbl(v(4)), CDbl(v(5)), CLng(v(6)), p5, p6, p7, CBool(v(10)), CDbl(v(11)), p8, CBool(v(13)), CDbl(v(14)), CLng(v(15)), CDbl(v(16)), CDbl(v(17)), CDbl(v(18)), CBool(v(19)), p9, CBool(v(21)), CStr(v(22)), CBool(v(23)), CStr(v(24)), CBool(v(25)), CStr(v(26)), CBool(v(27)), CStr(v(28)), CBool(v(29)), CStr(v(30)), CBool(v(31)), CStr(v(32)), CBool(v(33)), CStr(v(34)), CBool(v(35)), CStr(v(36)), CBool(v(37)), CStr(v(38)), CBool(v(39)), CStr(v(40)), CLng(v(41)), p10, CStr(v(43)), CStr(v(44)), CBool(v(45)), CBool(v(46)), p11, p12, p13, p14, p15, p16, p17, p18, ExtraArgs, CBool(v(UBound(v))), pStep, CStr(v(56)))
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function









































Public Function CTO_ADC_DC_T__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    Dim p1 As New PinList
    p1.value = v(0)
    Dim p2 As New PinList
    p2.value = v(1)
    Dim p3 As New PinList
    p3.value = v(2)
    Dim p4 As New PinList
    p4.value = v(3)
    Dim p5 As tlDCEncoding
    p5 = v(7)
    Dim p6 As tlDCTransitionPoint
    p6 = v(8)
    Dim p7 As tlDCVoltageRange
    p7 = v(9)
    Dim p8 As tlRelayMode
    p8 = v(10)
    Dim p9 As tlCTOSourceBandwidth
    p9 = v(13)
    Dim p10 As tlDCInputSignal
    p10 = v(15)
    Dim p11 As tlDCAnalysisAlgorithm
    p11 = v(17)
    Dim p12 As tlNormalizationMethod
    p12 = v(21)
    Dim p13 As New Pattern
    p13.value = v(23)
    Dim p14 As tlLimitUnit
    p14 = v(28)
    Dim p15 As tlLimitUnit
    p15 = v(33)
    Dim p16 As tlLimitUnit
    p16 = v(38)
    Dim p17 As tlLimitUnit
    p17 = v(43)
    Dim p18 As New Pattern
    p18.value = v(48)
    Dim ExtraArgs(0 To 21) As Variant
    Dim i As Integer
    For i = 0 To 21
        ExtraArgs(i) = v(51 + i)
    Next i
    Dim pStep As SubType
    pStep = TheExec.flow.StepType
    CTO_ADC_DC_T__ = Template.VBT_CTO_ADC_DC_T.CTO_ADC_DC_T(p1, p2, p3, p4, CDbl(v(4)), CDbl(v(5)), CLng(v(6)), p5, p6, p7, p8, CBool(v(11)), CDbl(v(12)), p9, CBool(v(14)), p10, CStr(v(16)), p11, CDbl(v(18)), CLng(v(19)), CLng(v(20)), p12, CBool(v(22)), p13, CBool(v(24)), CStr(v(25)), CBool(v(26)), CStr(v(27)), p14, CBool(v(29)), CStr(v(30)), CBool(v(31)), CStr(v(32)), p15, CBool(v(34)), CStr(v(35)), CBool(v(36)), CStr(v(37)), p16, CBool(v(39)), CStr(v(40)), CBool(v(41)), CStr(v(42)), p17, CBool(v(44)), CStr(v(45)), CBool(v(46)), CStr(v(47)), p18, CStr(v(49)), ExtraArgs, CBool(v(UBound(v))), pStep, CStr(v(51)))
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function









































Public Function DCVIPowerSupply_T__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    Dim p1 As New Pattern
    p1.value = v(0)
    Dim p2 As New InterposeName
    p2.value = v(1)
    Dim p3 As New InterposeName
    p3.value = v(2)
    Dim p4 As New InterposeName
    p4.value = v(3)
    Dim p5 As New InterposeName
    p5.value = v(4)
    Dim p6 As New InterposeName
    p6.value = v(5)
    Dim p7 As New InterposeName
    p7.value = v(6)
    Dim p8 As New Pattern
    p8.value = v(7)
    Dim p9 As New PinList
    p9.value = v(8)
    Dim p10 As New PinList
    p10.value = v(9)
    Dim p11 As New PinList
    p11.value = v(10)
    Dim p12 As New PinList
    p12.value = v(11)
    Dim p13 As New PinList
    p13.value = v(17)
    Dim p14 As New PinList
    p14.value = v(18)
    Dim p15 As tlPSSource
    p15 = v(19)
    Dim p16 As tlRelayMode
    p16 = v(34)
    Dim p17 As New PinList
    p17.value = v(35)
    Dim p18 As New PinList
    p18.value = v(36)
    Dim p19 As tlPSTestControl
    p19 = v(37)
    Dim p20 As New InterposeName
    p20.value = v(39)
    Dim p21 As tlWaitVal
    p21 = v(41)
    Dim p22 As tlWaitVal
    p22 = v(42)
    Dim p23 As tlWaitVal
    p23 = v(43)
    Dim p24 As tlWaitVal
    p24 = v(44)
    Dim pStep As SubType
    pStep = TheExec.flow.StepType
    DCVIPowerSupply_T__ = Template.VBT_DCVIPowerSupply_T.DCVIPowerSupply_T(p1, p2, p3, p4, p5, p6, p7, p8, p9, p10, p11, p12, CDbl(v(12)), CLng(v(13)), CStr(v(14)), CDbl(v(15)), CDbl(v(16)), p13, p14, p15, CStr(v(20)), CStr(v(21)), CStr(v(22)), CStr(v(23)), CStr(v(24)), CStr(v(25)), CStr(v(26)), CStr(v(27)), CStr(v(28)), CStr(v(29)), CStr(v(30)), CDbl(v(31)), CStr(v(32)), CBool(v(33)), p16, p17, p18, p19, CBool(v(38)), p20, CStr(v(40)), p21, p22, p23, p24, CBool(v(UBound(v))), CStr(v(46)), , CStr(v(47)), CBool(v(48)), CBool(v(49)), pStep, CStr(v(50)), CBool(v(51)))
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function









































Public Function ATC_ReadTemp__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    ATC_ReadTemp__ = VBAProject.VBT_GPIB.ATC_ReadTemp()
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function ATC_ControlTemp__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    ATC_ControlTemp__ = VBAProject.VBT_GPIB.ATC_ControlTemp(CStr(v(0)))
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function ATC_GetPowerTableParameter__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    ATC_GetPowerTableParameter__ = VBAProject.VBT_GPIB.ATC_GetPowerTableParameter(CStr(v(0)))
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function Get_Voltage__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    Get_Voltage__ = VBAProject.VBT_GPIB.Get_Voltage()
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function Read_Tj__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    Read_Tj__ = VBAProject.VBT_GPIB.Read_Tj()
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function Get_SlopeOffset__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    Get_SlopeOffset__ = VBAProject.VBT_GPIB.Get_SlopeOffset()
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function ATC_SetControlMode__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    ATC_SetControlMode__ = VBAProject.VBT_GPIB.ATC_SetControlMode(CStr(v(0)))
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function Set_SlopeOffset__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    Set_SlopeOffset__ = VBAProject.VBT_GPIB.Set_SlopeOffset(CStr(v(0)))
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function ATC_SetPowerTableParameter__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    ATC_SetPowerTableParameter__ = VBAProject.VBT_GPIB.ATC_SetPowerTableParameter(CStr(v(0)), CStr(v(1)))
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function ATC_Connect__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    ATC_Connect__ = VBAProject.VBT_GPIB.ATC_Connect()
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function Tempread__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    Tempread__ = VBAProject.VBT_GPIB.Tempread()
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function TestIPD__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    TestIPD__ = VBAProject.VBT_GPIB.TestIPD(CInt(v(0)), CStr(v(1)))
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function PrintOutFormat__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    tl_dt_ErrorMsg "Unsupported type at function PrintOutFormat"

    ' Call VBAProject.VBT_GPIB.PrintOutFormat(*One or more unsupported types in argument list or non Long/Integer return type*)
    PrintOutFormat__ = TL_SUCCESS
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function CalculateResult__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    tl_dt_ErrorMsg "Unsupported type at function CalculateResult"

    ' Call VBAProject.VBT_GPIB.CalculateResult(*One or more unsupported types in argument list or non Long/Integer return type*)
    CalculateResult__ = TL_SUCCESS
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function InitInstrument__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    InitInstrument__ = VBAProject.VBT_GPIB.InitInstrument(CInt(v(0)))
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function CheckNum__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    tl_dt_ErrorMsg "Unsupported type at function CheckNum"

    ' CheckNum__ = VBAProject.VBT_GPIB.CheckNum(*One or more unsupported types in argument list or non Long/Integer return type*)
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function timestamp__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    tl_dt_ErrorMsg "Unsupported type at function timestamp"

    ' timestamp__ = VBAProject.VBT_GPIB.timestamp(*One or more unsupported types in argument list or non Long/Integer return type*)
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function GPIB_Ctrl_Test_CS__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    GPIB_Ctrl_Test_CS__ = VBAProject.VBT_GPIB.GPIB_Ctrl_Test_CS(CStr(v(0)), CStr(v(1)))
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function GPIB_Ctrl_Test_forloop__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    GPIB_Ctrl_Test_forloop__ = VBAProject.VBT_GPIB.GPIB_Ctrl_Test_forloop(CStr(v(0)))
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function GPIB_Ctrl_Test__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    GPIB_Ctrl_Test__ = VBAProject.VBT_GPIB.GPIB_Ctrl_Test()
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function GPIB_Ctrl_Test_CS2__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    GPIB_Ctrl_Test_CS2__ = VBAProject.VBT_GPIB.GPIB_Ctrl_Test_CS2()
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function create_folder__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    Call VBAProject.VBT_GPIB.create_folder(v(0))
    create_folder__ = TL_SUCCESS
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function ATC_set_tempoffset__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    ATC_set_tempoffset__ = VBAProject.VBT_GPIB.ATC_set_tempoffset(CStr(v(0)))
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function ATC_rollback_oritemp__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    ATC_rollback_oritemp__ = VBAProject.VBT_GPIB.ATC_rollback_oritemp()
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function GPIB_PID_Ctrl__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    GPIB_PID_Ctrl__ = VBAProject.VBT_GPIB.GPIB_PID_Ctrl(CStr(v(0)))
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function GPIB_PID_rollback__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    GPIB_PID_rollback__ = VBAProject.VBT_GPIB.GPIB_PID_rollback(CStr(v(0)))
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function PPMU_Continuity_diode__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    PPMU_Continuity_diode__ = VBAProject.VBT_GPIB.PPMU_Continuity_diode()
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function Pre_Tjmode_setup__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    Pre_Tjmode_setup__ = VBAProject.VBT_GPIB.Pre_Tjmode_setup()
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function Set_ControlMode_VI__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    Set_ControlMode_VI__ = VBAProject.VBT_GPIB.Set_ControlMode_VI(CBool(v(0)))
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function thermal_control_time__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    thermal_control_time__ = VBAProject.VBT_GPIB.thermal_control_time()
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function Set_VI_table__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    Set_VI_table__ = VBAProject.VBT_GPIB.Set_VI_table()
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function Rollback_VI_table__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    Rollback_VI_table__ = VBAProject.VBT_GPIB.Rollback_VI_table()
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function Fake_VIsignal__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    Fake_VIsignal__ = VBAProject.VBT_GPIB.Fake_VIsignal()
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function ATCS_set_tempoffset__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    ATCS_set_tempoffset__ = VBAProject.VBT_GPIB.ATCS_set_tempoffset(CStr(v(0)))
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function ATCS_TCR__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    ATCS_TCR__ = VBAProject.VBT_GPIB.ATCS_TCR()
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function ATCS_TJCAL__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    ATCS_TJCAL__ = VBAProject.VBT_GPIB.ATCS_TJCAL()
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function ATCS_TCS__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    ATCS_TCS__ = VBAProject.VBT_GPIB.ATCS_TCS()
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function Printout_ATC_info__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    Printout_ATC_info__ = VBAProject.VBT_GPIB.Printout_ATC_info(CStr(v(0)))
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function Random_power_presetup__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    Random_power_presetup__ = VBAProject.VBT_GPIB.Random_power_presetup()
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function ATCS_TJVOL__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    ATCS_TJVOL__ = VBAProject.VBT_GPIB.ATCS_TJVOL()
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function ATCS_TJCH__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    ATCS_TJCH__ = VBAProject.VBT_GPIB.ATCS_TJCH(CStr(v(0)), CStr(v(1)))
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function ATCS_TJSLOPE__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    ATCS_TJSLOPE__ = VBAProject.VBT_GPIB.ATCS_TJSLOPE(CStr(v(0)), CStr(v(1)), CStr(v(2)))
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function









































Public Function ValidateSystemSetup_T__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    ValidateSystemSetup_T__ = OasisXLA.VBT_ConfigCheck.ValidateSystemSetup_T(CStr(v(0)))
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function









































Public Function IGSim_Functional_T__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    Dim p1 As New Pattern
    p1.value = v(0)
    Dim p2 As New InterposeName
    p2.value = v(1)
    Dim p3 As New InterposeName
    p3.value = v(2)
    Dim p4 As New InterposeName
    p4.value = v(3)
    Dim p5 As New InterposeName
    p5.value = v(4)
    Dim p6 As New InterposeName
    p6.value = v(5)
    Dim p7 As New InterposeName
    p7.value = v(6)
    Dim p8 As PFType
    p8 = v(7)
    Dim p9 As tlResultMode
    p9 = v(8)
    Dim p10 As New PinList
    p10.value = v(9)
    Dim p11 As New PinList
    p11.value = v(10)
    Dim p12 As New PinList
    p12.value = v(11)
    Dim p13 As New PinList
    p13.value = v(12)
    Dim p14 As New PinList
    p14.value = v(13)
    Dim p15 As New PinList
    p15.value = v(20)
    Dim p16 As New PinList
    p16.value = v(21)
    Dim p17 As New InterposeName
    p17.value = v(22)
    Dim p18 As tlRelayMode
    p18 = v(24)
    Dim p19 As tlWaitVal
    p19 = v(27)
    Dim p20 As tlWaitVal
    p20 = v(28)
    Dim p21 As tlWaitVal
    p21 = v(29)
    Dim p22 As tlWaitVal
    p22 = v(30)
    Dim pStep As SubType
    pStep = TheExec.flow.StepType
    Dim p23 As tlPatConcurrentMode
    p23 = v(34)
    Dim p24 As tlTemplateScanFailDataLogging
    p24 = v(35)
    Dim p25 As tlDigitalCMEMCaptureLimitMode
    p25 = v(36)
    Dim p26 As tlTemplateScanPinListSource
    p26 = v(38)
    Dim p27 As New PinList
    p27.value = v(39)
    Dim p28 As tlTemplateScanCaptureFormat
    p28 = v(40)
    Dim p29 As tlTemplateScanCaptureDataType
    p29 = v(41)
    Dim p30 As tlTemplateScanUserCommentSource
    p30 = v(42)
    Dim p31 As tlTemplateATPGPinMapSource
    p31 = v(44)
    IGSim_Functional_T__ = OasisXLA.VBT_IGSIM_FUNCTIONAL_T.IGSim_Functional_T(p1, p2, p3, p4, p5, p6, p7, p8, p9, p10, p11, p12, p13, p14, CStr(v(14)), CStr(v(15)), CStr(v(16)), CStr(v(17)), CStr(v(18)), CStr(v(19)), p15, p16, p17, CStr(v(23)), p18, CBool(v(25)), CBool(v(26)), p19, p20, p21, p22, CBool(v(UBound(v))), CStr(v(32)), pStep, CStr(v(33)), p23, p24, p25, CLng(v(37)), p26, p27, p28, p29, p30, CStr(v(43)), p31, CStr(v(45)), CStr(v(46)), CBool(v(47)), CStr(v(48)))
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function DatalogType__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    DatalogType__ = OasisXLA.VBT_IGSIM_FUNCTIONAL_T.DatalogType()
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function PostTest__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    tl_dt_ErrorMsg "Unsupported type at function PostTest"

    ' Call OasisXLA.VBT_IGSIM_FUNCTIONAL_T.PostTest(*One or more unsupported types in argument list or non Long/Integer return type*)
    PostTest__ = TL_SUCCESS
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function getdefaults__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    tl_dt_ErrorMsg "Unsupported type at function getdefaults"

    ' getdefaults__ = OasisXLA.VBT_IGSIM_FUNCTIONAL_T.getdefaults(*One or more unsupported types in argument list or non Long/Integer return type*)
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function









































Public Function IGSIM_PinPMU_T__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    Dim p1 As New InterposeName
    p1.value = v(1)
    Dim p2 As New InterposeName
    p2.value = v(2)
    Dim p3 As New InterposeName
    p3.value = v(3)
    Dim p4 As New InterposeName
    p4.value = v(4)
    Dim p5 As New InterposeName
    p5.value = v(5)
    Dim p6 As New InterposeName
    p6.value = v(6)
    Dim p7 As New Pattern
    p7.value = v(7)
    Dim p8 As New Pattern
    p8.value = v(8)
    Dim p9 As New PinList
    p9.value = v(10)
    Dim p10 As New PinList
    p10.value = v(11)
    Dim p11 As New PinList
    p11.value = v(12)
    Dim p12 As New PinList
    p12.value = v(13)
    Dim p13 As New PinList
    p13.value = v(14)
    Dim p14 As New PinList
    p14.value = v(15)
    Dim p15 As tlPPMUMode
    p15 = v(16)
    Dim p16 As New FormulaArg
    p16.value = v(18)
    Dim p17 As New FormulaArg
    p17.value = v(19)
    Dim p18 As tlPPMURelayMode
    p18 = v(20)
    Dim p19 As New PinList
    p19.value = v(36)
    Dim p20 As New PinList
    p20.value = v(37)
    Dim p21 As tlWaitVal
    p21 = v(38)
    Dim p22 As tlWaitVal
    p22 = v(39)
    Dim p23 As tlWaitVal
    p23 = v(40)
    Dim p24 As tlWaitVal
    p24 = v(41)
    Dim p25 As tlPPMUMode
    p25 = v(49)
    Dim p26 As New FormulaArg
    p26.value = v(52)
    Dim pStep As SubType
    pStep = TheExec.flow.StepType
    Dim p27 As New PinList
    p27.value = v(53)
    Dim p28 As tlPPMUMode
    p28 = v(54)
    Dim p29 As New FormulaArg
    p29.value = v(55)
    IGSIM_PinPMU_T__ = OasisXLA.VBT_IGSIM_PinPMU_T.IGSIM_PinPMU_T(CStr(v(0)), p1, p2, p3, p4, p5, p6, p7, p8, CStr(v(9)), p9, p10, p11, p12, p13, p14, p15, CDbl(v(17)), p16, p17, p18, CStr(v(21)), CStr(v(22)), CStr(v(23)), CStr(v(24)), CStr(v(25)), CStr(v(26)), CStr(v(27)), CStr(v(28)), CStr(v(29)), CStr(v(30)), CDbl(v(31)), CLng(v(32)), CBool(v(33)), CStr(v(34)), CStr(v(35)), p19, p20, p21, p22, p23, p24, CBool(v(UBound(v))), CStr(v(43)), CStr(v(44)), , CStr(v(45)), CBool(v(46)), CBool(v(47)), CBool(v(48)), p25, CStr(v(50)), CStr(v(51)), p26, pStep, p27, p28, p29, CStr(v(56)), CStr(v(57)))
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function
Public Function GetBarcode__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    GetBarcode__ = VBAProject.VBT_LIB_Common.GetBarcode()
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function EEPROM_Write_Hw__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    EEPROM_Write_Hw__ = VBAProject.VBT_LIB_Common.EEPROM_Write_Hw(CStr(v(0)))
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function EEPROM_init__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    EEPROM_init__ = VBAProject.VBT_LIB_Common.EEPROM_init()
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function Read_EEPROM_proberTemp_compare__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    Read_EEPROM_proberTemp_compare__ = VBAProject.VBT_LIB_Common.Read_EEPROM_proberTemp_compare()
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function EEPROM_readBack__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    tl_dt_ErrorMsg "Unsupported type at function EEPROM_readBack"

    ' EEPROM_readBack__ = VBAProject.VBT_LIB_Common.EEPROM_readBack(*One or more unsupported types in argument list or non Long/Integer return type*)
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function EEPROM_read__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    EEPROM_read__ = VBAProject.VBT_LIB_Common.EEPROM_read()
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function GetStoredData__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    tl_dt_ErrorMsg "Unsupported type at function GetStoredData"

    ' GetStoredData__ = VBAProject.VBT_LIB_Common.GetStoredData(*One or more unsupported types in argument list or non Long/Integer return type*)
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function









































