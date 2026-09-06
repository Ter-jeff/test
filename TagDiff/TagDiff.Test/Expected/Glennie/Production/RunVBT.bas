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
    If TheExec.sites.Active.Count = 0 Then Exit Sub  ' don't log the error
    ' If in a legacy site loop, make sure to complete it. (For-Each site syntax in IG-XL 6.10 aborts gracefully.)
    Do While TheExec.sites.InSiteLoop
        Call TheExec.sites.SelectNext(loopTop) '  Legacy syntax (hidden)
    Loop
    ' Select all active sites in case a subset of sites was selected when error occurred.
    TheExec.sites.Selected = TheExec.sites.Active
    ' Log the error to the IG-XL Error logging mechanism (tells Flow to fail the test)
    AbortTest
End Sub

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
    DCVIPowerSupply_T__ = Template.VBT_DCVIPowerSupply_T.DCVIPowerSupply_T(p1, p2, p3, p4, p5, p6, p7, p8, p9, p10, p11, p12, CDbl(v(12)), CLng(v(13)), CStr(v(14)), CDbl(v(15)), CDbl(v(16)), p13, p14, p15, CStr(v(20)), CStr(v(21)), CStr(v(22)), CStr(v(23)), CStr(v(24)), CStr(v(25)), CStr(v(26)), CStr(v(27)), CStr(v(28)), CStr(v(29)), CStr(v(30)), CDbl(v(31)), CStr(v(32)), CBool(v(33)), p16, p17, p18, p19, CBool(v(38)), p20, CStr(v(40)), p21, p22, p23, p24, CBool(v(UBound(v))), CStr(v(46)), , CStr(v(47)), CBool(v(48)), CBool(v(49)), pStep)
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
    Dim p16 As tlRelayMode
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
    DCVSPowerSupply_T__ = Template.VBT_DCVSPowerSupply_T.DCVSPowerSupply_T(p1, p2, p3, p4, p5, p6, p7, p8, p9, p10, p11, p12, p13, CDbl(v(13)), CLng(v(14)), CStr(v(15)), p14, p15, CStr(v(18)), CStr(v(19)), CStr(v(20)), CStr(v(21)), CStr(v(22)), CStr(v(23)), CStr(v(24)), CStr(v(25)), CStr(v(26)), CStr(v(27)), CStr(v(28)), CStr(v(29)), CBool(v(30)), p16, p17, p18, p19, p20, p21, p22, p23, CBool(v(UBound(v))), p24, p25, p26, p27, , CStr(v(44)), CBool(v(45)), CBool(v(46)), pStep)
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function









































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









































Public Function MtoMemory_T__(v As Variant) As Long
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
    p14.value = v(19)
    Dim p15 As New PinList
    p15.value = v(20)
    Dim p16 As New InterposeName
    p16.value = v(21)
    Dim p17 As tlRelayMode
    p17 = v(24)
    Dim pStep As SubType
    pStep = TheExec.flow.StepType
    Dim ExtraArgs(0 To 49) As Variant
    Dim i As Integer
    For i = 0 To 49
        ExtraArgs(i) = v(51 + i)
    Next i
    MtoMemory_T__ = Template.VBT_MTOMemory_T.MtoMemory_T(p1, p2, p3, p4, p5, p6, p7, p8, p9, p10, p11, p12, p13, CStr(v(13)), CStr(v(14)), CStr(v(15)), CStr(v(16)), CStr(v(17)), CStr(v(18)), p14, p15, p16, CStr(v(22)), CBool(v(23)), p17, CStr(v(25)), CStr(v(26)), CStr(v(27)), CStr(v(28)), CLng(v(29)), CStr(v(30)), CStr(v(31)), CStr(v(32)), CStr(v(33)), CLng(v(34)), CStr(v(35)), CStr(v(36)), CStr(v(37)), CStr(v(38)), CLng(v(39)), CLng(v(40)), CBool(v(UBound(v))), pStep, ExtraArgs, CStr(v(42)), CStr(v(43)), CStr(v(44)), CStr(v(45)), CStr(v(46)), CStr(v(47)), CStr(v(48)), CStr(v(49)), CStr(v(50)), CBool(v(51)))
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function









































Public Function ShmooDIfferentPatterns__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    Dim p1 As New Pattern
    p1.value = v(0)
    ShmooDIfferentPatterns__ = VBAProject.VBT_AI_Modules.ShmooDIfferentPatterns(p1, CStr(v(1)), CStr(v(2)), CStr(v(3)), CStr(v(4)), CStr(v(5)), CStr(v(6)), CStr(v(7)), CStr(v(8)), CStr(v(9)), CBool(v(UBound(v))))
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function RepeatPatternToGetFailLog__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    Dim p1 As New Pattern
    p1.value = v(0)
    Dim p2 As New PinList
    p2.value = v(13)
    Dim p3 As pinTypes
    p3 = v(16)
    Dim p4 As volTypes
    p4 = v(17)
    Dim p5 As loopTypes
    p5 = v(18)
    RepeatPatternToGetFailLog__ = VBAProject.VBT_AI_Modules.RepeatPatternToGetFailLog(p1, CStr(v(1)), CStr(v(2)), CStr(v(3)), CDbl(v(4)), CStr(v(5)), CStr(v(6)), CStr(v(7)), CDbl(v(8)), CLng(v(9)), CLng(v(10)), CLng(v(11)), CLng(v(12)), p2, CStr(v(14)), CStr(v(15)), p3, p4, p5, CStr(v(19)), CStr(v(20)), CStr(v(21)), CBool(v(UBound(v))))
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function MeasFrequencyOnPins__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    Dim p1 As New PinList
    p1.value = v(0)
    MeasFrequencyOnPins__ = VBAProject.VBT_AI_Modules.MeasFrequencyOnPins(p1)
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function charSetup__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    charSetup__ = VBAProject.VBT_AI_Modules.charSetup()
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function CovCond__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    tl_dt_ErrorMsg "Unsupported type at function CovCond"

    ' CovCond__ = VBAProject.VBT_AI_Modules.CovCond(*One or more unsupported types in argument list or non Long/Integer return type*)
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function AddTestNumber__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    Call VBAProject.VBT_AI_Modules.AddTestNumber
    AddTestNumber__ = TL_SUCCESS
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function buildTestNumber__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    buildTestNumber__ = VBAProject.VBT_AI_Modules.buildTestNumber(CStr(v(0)), CLng(v(1)))
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function CorrectBincut__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    CorrectBincut__ = VBAProject.VBT_AI_Modules.CorrectBincut(CInt(v(0)), CBool(v(1)))
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function CorrectBinout__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    CorrectBinout__ = VBAProject.VBT_AI_Modules.CorrectBinout(CStr(v(0)))
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function GenChar__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    tl_dt_ErrorMsg "Unsupported type at function GenChar"

    ' GenChar__ = VBAProject.VBT_AI_Modules.GenChar(*One or more unsupported types in argument list or non Long/Integer return type*)
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function PatRemove__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    PatRemove__ = VBAProject.VBT_AI_Modules.PatRemove(CStr(v(0)), CStr(v(1)), CStr(v(2)))
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function ExportModule__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    ExportModule__ = VBAProject.VBT_AI_Modules.ExportModule(CStr(v(0)), CStr(v(1)))
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function









































Public Function ParsingCorrelationIDSTableSheet__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    ParsingCorrelationIDSTableSheet__ = VBAProject.VBT_Correlation_AP.ParsingCorrelationIDSTableSheet()
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function CorrelationBankBlankChk__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    CorrelationBankBlankChk__ = VBAProject.VBT_Correlation_AP.CorrelationBankBlankChk(CStr(v(0)))
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function CorrelationBankVarChk__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    CorrelationBankVarChk__ = VBAProject.VBT_Correlation_AP.CorrelationBankVarChk(CStr(v(0)))
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function CheckBankBlank__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    tl_dt_ErrorMsg "Unsupported type at function CheckBankBlank"

    ' CheckBankBlank__ = VBAProject.VBT_Correlation_AP.CheckBankBlank(*One or more unsupported types in argument list or non Long/Integer return type*)
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function IDS_Correlation_Delta__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    IDS_Correlation_Delta__ = VBAProject.VBT_Correlation_AP.IDS_Correlation_Delta(CStr(v(0)))
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function SwitchPseudoFuseEnableFlag__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    Call VBAProject.VBT_Correlation_AP.SwitchPseudoFuseEnableFlag
    SwitchPseudoFuseEnableFlag__ = TL_SUCCESS
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function









































Public Function ExecuteIndex__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    ExecuteIndex__ = VBAProject.VBT_DIBC_Info.ExecuteIndex()
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









































Public Function BinOutCalc_CleanFlagStatus__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    BinOutCalc_CleanFlagStatus__ = VBAProject.VBT_LIB_BinOutCalc.BinOutCalc_CleanFlagStatus(CStr(v(0)), CStr(v(1)), CStr(v(2)), CBool(v(3)))
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function BinOutCalc_ReadBlockFlagState__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    BinOutCalc_ReadBlockFlagState__ = VBAProject.VBT_LIB_BinOutCalc.BinOutCalc_ReadBlockFlagState(CBool(v(0)))
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function BinOutCalc_AssignBinOutFlag__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    BinOutCalc_AssignBinOutFlag__ = VBAProject.VBT_LIB_BinOutCalc.BinOutCalc_AssignBinOutFlag(CBool(v(0)))
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function









































Public Function Relay_Control__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    Dim p1 As New PinList
    p1.value = v(0)
    Dim p2 As New PinList
    p2.value = v(1)
    Relay_Control__ = VBAProject.VBT_LIB_Common.Relay_Control(p1, p2, CDbl(v(2)))
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function StartSBClock__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    StartSBClock__ = VBAProject.VBT_LIB_Common.StartSBClock(CDbl(v(0)))
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function StopSBClock__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    StopSBClock__ = VBAProject.VBT_LIB_Common.StopSBClock()
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function FreeRunClk_Disable__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    FreeRunClk_Disable__ = VBAProject.VBT_LIB_Common.FreeRunClk_Disable(CStr(v(0)), CBool(v(1)))
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function Start_Profile_DebugOnly__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    Dim p1 As New PinList
    p1.value = v(0)
    Start_Profile_DebugOnly__ = VBAProject.VBT_LIB_Common.Start_Profile_DebugOnly(p1, CStr(v(1)), CDbl(v(2)), CLng(v(3)), CStr(v(4)))
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function start_profile_DCVI__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    start_profile_DCVI__ = VBAProject.VBT_LIB_Common.start_profile_DCVI(CStr(v(0)), CStr(v(1)), CDbl(v(2)), CLng(v(3)))
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function Plot_Profile__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    Dim p1 As New PinList
    p1.value = v(0)
    Plot_Profile__ = VBAProject.VBT_LIB_Common.Plot_Profile(p1, CStr(v(1)), CBool(v(2)), CBool(v(3)), CBool(v(4)), CDbl(v(5)))
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function Plot_profile_DCVI__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    Plot_profile_DCVI__ = VBAProject.VBT_LIB_Common.Plot_profile_DCVI(CStr(v(0)))
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function Start_Profile_AutoResolution__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    Start_Profile_AutoResolution__ = VBAProject.VBT_LIB_Common.Start_Profile_AutoResolution(CStr(v(0)), CStr(v(1)), CStr(v(2)), CDbl(v(3)), CBool(v(4)), CLng(v(5)))
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function Print_Footer__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    Print_Footer__ = VBAProject.VBT_LIB_Common.Print_Footer(CStr(v(0)))
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function Print_Header__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    Print_Header__ = VBAProject.VBT_LIB_Common.Print_Header(CStr(v(0)))
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function Print_PgmInfo__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    Print_PgmInfo__ = VBAProject.VBT_LIB_Common.Print_PgmInfo(CBool(v(0)))
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function Write_DIB_EEPROM__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    Write_DIB_EEPROM__ = VBAProject.VBT_LIB_Common.Write_DIB_EEPROM(CStr(v(0)))
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function Read_DIB_EEPROM__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    Read_DIB_EEPROM__ = VBAProject.VBT_LIB_Common.Read_DIB_EEPROM()
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function ReadProberTemp__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    ReadProberTemp__ = VBAProject.VBT_LIB_Common.ReadProberTemp(CDbl(v(0)), CDbl(v(1)))
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function Set_PPMU_Clamp__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    Dim p1 As New PinList
    p1.value = v(0)
    Dim p2 As New PinList
    p2.value = v(2)
    Dim p3 As New PinList
    p3.value = v(4)
    Dim p4 As New PinList
    p4.value = v(6)
    Dim p5 As New PinList
    p5.value = v(8)
    Dim p6 As New PinList
    p6.value = v(10)
    Dim p7 As New PinList
    p7.value = v(12)
    Dim p8 As New PinList
    p8.value = v(14)
    Dim p9 As New PinList
    p9.value = v(16)
    Dim p10 As New PinList
    p10.value = v(18)
    Set_PPMU_Clamp__ = VBAProject.VBT_LIB_Common.Set_PPMU_Clamp(p1, CDbl(v(1)), p2, CDbl(v(3)), p3, CDbl(v(5)), p4, CDbl(v(7)), p5, CDbl(v(9)), p6, CDbl(v(11)), p7, CDbl(v(13)), p8, CDbl(v(15)), p9, CDbl(v(17)), p10, CDbl(v(19)))
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function Setup_ATPG_Output_Format__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    Setup_ATPG_Output_Format__ = VBAProject.VBT_LIB_Common.Setup_ATPG_Output_Format(CBool(v(0)))
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function Disable_compare__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    Dim p1 As New PinList
    p1.value = v(0)
    Disable_compare__ = VBAProject.VBT_LIB_Common.Disable_compare(p1)
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function Enble_compare__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    Dim p1 As New PinList
    p1.value = v(0)
    Enble_compare__ = VBAProject.VBT_LIB_Common.Enble_compare(p1)
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function FreeRunclk_Enable__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    FreeRunclk_Enable__ = VBAProject.VBT_LIB_Common.FreeRunclk_Enable(CStr(v(0)), CDbl(v(1)), CStr(v(2)))
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function PowerDown_Parallel__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    PowerDown_Parallel__ = VBAProject.VBT_LIB_Common.PowerDown_Parallel(CStr(v(0)), CStr(v(1)), CDbl(v(2)), CBool(v(3)))
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function PowerUp_Parallel__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    PowerUp_Parallel__ = VBAProject.VBT_LIB_Common.PowerUp_Parallel(CStr(v(0)), CStr(v(1)), CStr(v(2)), CDbl(v(3)), CBool(v(4)))
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function Set_Power_Alarm__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    Set_Power_Alarm__ = VBAProject.VBT_LIB_Common.Set_Power_Alarm(CStr(v(0)), CDbl(v(1)), CDbl(v(2)), CDbl(v(3)), CStr(v(4)), CDbl(v(5)), CStr(v(6)), CDbl(v(7)), CDbl(v(8)), CDbl(v(9)), CBool(v(10)))
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function Search_UnExistPin__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    Search_UnExistPin__ = VBAProject.VBT_LIB_Common.Search_UnExistPin()
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function auto_ReadHandlerOCRData__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    auto_ReadHandlerOCRData__ = VBAProject.VBT_LIB_Common.auto_ReadHandlerOCRData()
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function Init_Datalog_Setup__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    Init_Datalog_Setup__ = VBAProject.VBT_LIB_Common.Init_Datalog_Setup()
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function Customize_Datalog_Setup__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    Customize_Datalog_Setup__ = VBAProject.VBT_LIB_Common.Customize_Datalog_Setup(CStr(v(0)), CLng(v(1)), CLng(v(2)), CLng(v(3)), CLng(v(4)), CLng(v(5)), CLng(v(6)), CLng(v(7)), CLng(v(8)))
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function Print_Version__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    Print_Version__ = VBAProject.VBT_LIB_Common.Print_Version()
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function Print_Tsense_Block_Start__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    Print_Tsense_Block_Start__ = VBAProject.VBT_LIB_Common.Print_Tsense_Block_Start()
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function Print_Tsense_Block_End__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    Print_Tsense_Block_End__ = VBAProject.VBT_LIB_Common.Print_Tsense_Block_End()
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function show_FRC__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    show_FRC__ = VBAProject.VBT_LIB_Common.show_FRC()
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function FixedTuneCRWT__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    FixedTuneCRWT__ = VBAProject.VBT_LIB_Common.FixedTuneCRWT(CStr(v(0)), CStr(v(1)), CDbl(v(2)), CStr(v(3)), CStr(v(4)), CStr(v(5)), CStr(v(6)), CDbl(v(7)), CDbl(v(8)), CInt(v(9)), CBool(v(10)), CBool(v(11)))
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function LoopTuneCRWT__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    LoopTuneCRWT__ = VBAProject.VBT_LIB_Common.LoopTuneCRWT(CStr(v(0)), CStr(v(1)), CDbl(v(2)), CDbl(v(3)), CDbl(v(4)), CDbl(v(5)), CStr(v(6)), CStr(v(7)), CDbl(v(8)), CDbl(v(9)), CInt(v(10)), CBool(v(11)), CBool(v(12)))
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function onProgramStartedBinOutFunction__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    onProgramStartedBinOutFunction__ = VBAProject.VBT_LIB_Common.onProgramStartedBinOutFunction()
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function JudgeForDramConti__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    JudgeForDramConti__ = VBAProject.VBT_LIB_Common.JudgeForDramConti(CStr(v(0)))
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function Apply_Timing_Level_PCPU__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    Apply_Timing_Level_PCPU__ = VBAProject.VBT_LIB_Common.Apply_Timing_Level_PCPU(CStr(v(0)))
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

Public Function EEPROM_delete_Hw__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    EEPROM_delete_Hw__ = VBAProject.VBT_LIB_Common.EEPROM_delete_Hw()
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

Public Function EEPROM_read__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    EEPROM_read__ = VBAProject.VBT_LIB_Common.EEPROM_read()
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









































Public Function VBT_IEDA_Registry__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    VBT_IEDA_Registry__ = VBAProject.VBT_LIB_Common_AP.VBT_IEDA_Registry(CStr(v(0)), CBool(v(1)), CBool(v(2)))
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function Harvest_eFuse_Write__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    Harvest_eFuse_Write__ = VBAProject.VBT_LIB_Common_AP.Harvest_eFuse_Write(v(0), v(1), v(2))
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function Harvest_eFuse_Read__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    Harvest_eFuse_Read__ = VBAProject.VBT_LIB_Common_AP.Harvest_eFuse_Read(v(0), v(1), v(2))
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function Harvest_SUBFLOW__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    Harvest_SUBFLOW__ = VBAProject.VBT_LIB_Common_AP.Harvest_SUBFLOW()
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function Harvest_FailedCoreCount__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    Harvest_FailedCoreCount__ = VBAProject.VBT_LIB_Common_AP.Harvest_FailedCoreCount(v(0), v(1), v(2), v(3), v(4), v(5))
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function Harvest_Summary__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    Harvest_Summary__ = VBAProject.VBT_LIB_Common_AP.Harvest_Summary(v(0), v(1), v(2), v(3), v(4), CStr(v(5)), CStr(v(6)), CStr(v(7)))
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function Harvest_ReadValue__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    Harvest_ReadValue__ = VBAProject.VBT_LIB_Common_AP.Harvest_ReadValue(v(0), v(1))
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function HarvFlagPrinting__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    HarvFlagPrinting__ = VBAProject.VBT_LIB_Common_AP.HarvFlagPrinting()
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function Harvest_CustomFlagJudge__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    Harvest_CustomFlagJudge__ = VBAProject.VBT_LIB_Common_AP.Harvest_CustomFlagJudge(CStr(v(0)), CStr(v(1)), CStr(v(2)), CLng(v(3)))
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function Harvest_Print_Flag_Status__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    Harvest_Print_Flag_Status__ = VBAProject.VBT_LIB_Common_AP.Harvest_Print_Flag_Status(CStr(v(0)), CInt(v(1)))
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function Harvest_Postcheck__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    Harvest_Postcheck__ = VBAProject.VBT_LIB_Common_AP.Harvest_Postcheck(v(0), v(1), v(2))
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function









































Public Function Enable_SubFlowFlag_ByDieMap__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    Enable_SubFlowFlag_ByDieMap__ = VBAProject.VBT_LIB_Customized.Enable_SubFlowFlag_ByDieMap()
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function Enable_SubFlowFlag_ByFlagMap__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    Enable_SubFlowFlag_ByFlagMap__ = VBAProject.VBT_LIB_Customized.Enable_SubFlowFlag_ByFlagMap()
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function Parse_SubFlowControlTable__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    Parse_SubFlowControlTable__ = VBAProject.VBT_LIB_Customized.Parse_SubFlowControlTable()
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function SetMFP_Judgement__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    SetMFP_Judgement__ = VBAProject.VBT_LIB_Customized.SetMFP_Judgement(CBool(v(0)), CBool(v(1)))
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function Create_Overlay__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    Create_Overlay__ = VBAProject.VBT_LIB_Customized.Create_Overlay(CStr(v(0)), CStr(v(1)), CStr(v(2)))
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function Create_Overlay_CPM__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    Create_Overlay_CPM__ = VBAProject.VBT_LIB_Customized.Create_Overlay_CPM(CStr(v(0)), CStr(v(1)), CStr(v(2)), CDbl(v(3)))
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function CalcCPM_FailingPatCount__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    CalcCPM_FailingPatCount__ = VBAProject.VBT_LIB_Customized.CalcCPM_FailingPatCount(CStr(v(0)), CLng(v(1)))
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function Create_overlay_PostBV__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    Create_overlay_PostBV__ = VBAProject.VBT_LIB_Customized.Create_overlay_PostBV(CStr(v(0)), CStr(v(1)), CStr(v(2)))
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function Create_Overlay_MultiSite__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    tl_dt_ErrorMsg "Unsupported type at function Create_Overlay_MultiSite"

    ' Create_Overlay_MultiSite__ = VBAProject.VBT_LIB_Customized.Create_Overlay_MultiSite(*One or more unsupported types in argument list or non Long/Integer return type*)
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function ClearOverlay__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    ClearOverlay__ = VBAProject.VBT_LIB_Customized.ClearOverlay(CStr(v(0)))
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function SetFuse_baseOnFlowVariant__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    SetFuse_baseOnFlowVariant__ = VBAProject.VBT_LIB_Customized.SetFuse_baseOnFlowVariant(CStr(v(0)), CStr(v(1)), CStr(v(2)))
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function









































Public Function DAC_Trim_Initial__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    Call VBAProject.VBT_LIB_DACTrim_AP.DAC_Trim_Initial
    DAC_Trim_Initial__ = TL_SUCCESS
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function FuseUpdate__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    Call VBAProject.VBT_LIB_DACTrim_AP.FuseUpdate(CStr(v(0)), CStr(v(1)))
    FuseUpdate__ = TL_SUCCESS
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function FuseUpdate2__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    tl_dt_ErrorMsg "Unsupported type at function FuseUpdate2"

    ' Call VBAProject.VBT_LIB_DACTrim_AP.FuseUpdate2(*One or more unsupported types in argument list or non Long/Integer return type*)
    FuseUpdate2__ = TL_SUCCESS
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function SearchPathSplit__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    tl_dt_ErrorMsg "Unsupported type at function SearchPathSplit"

    ' Call VBAProject.VBT_LIB_DACTrim_AP.SearchPathSplit(*One or more unsupported types in argument list or non Long/Integer return type*)
    SearchPathSplit__ = TL_SUCCESS
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function









































Public Function IDS_eFuse_Write__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    IDS_eFuse_Write__ = VBAProject.VBT_LIB_DC_AP.IDS_eFuse_Write(CStr(v(0)), CStr(v(1)))
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function DCVS_IDS_main_current_Delta__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    Dim p1 As New PinList
    p1.value = v(0)
    DCVS_IDS_main_current_Delta__ = VBAProject.VBT_LIB_DC_AP.DCVS_IDS_main_current_Delta(p1, CStr(v(1)))
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function









































Public Function PPMU_Continuity__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    Dim p1 As tlLimitForceResults
    p1 = v(4)
    PPMU_Continuity__ = VBAProject.VBT_LIB_DC_Conti.PPMU_Continuity(CStr(v(0)), CDbl(v(1)), CDbl(v(2)), CDbl(v(3)), p1, CBool(v(5)), CStr(v(6)), CStr(v(7)), CStr(v(8)), CDbl(v(9)), CDbl(v(10)), CBool(v(11)))
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function UVI80_Continuity__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    Dim p1 As New PinList
    p1.value = v(0)
    Dim p2 As tlLimitForceResults
    p2 = v(4)
    Dim p3 As New PinList
    p3.value = v(9)
    UVI80_Continuity__ = VBAProject.VBT_LIB_DC_Conti.UVI80_Continuity(p1, CDbl(v(1)), CDbl(v(2)), CDbl(v(3)), p2, CBool(v(5)), CBool(v(6)), CDbl(v(7)), CDbl(v(8)), p3, CStr(v(10)), CStr(v(11)))
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function Conti_WalkingZ__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    Dim p1 As New Pattern
    p1.value = v(0)
    Dim p2 As New PinList
    p2.value = v(1)
    Conti_WalkingZ__ = VBAProject.VBT_LIB_DC_Conti.Conti_WalkingZ(p1, p2, CStr(v(2)))
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function PowerSensePins_continuity__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    Dim p1 As New PinList
    p1.value = v(0)
    PowerSensePins_continuity__ = VBAProject.VBT_LIB_DC_Conti.PowerSensePins_continuity(p1, CDbl(v(1)), CDbl(v(2)), CDbl(v(3)), CStr(v(4)), CBool(v(5)))
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function p2p_short_Power_FVMI__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    Dim p1 As tlLimitForceResults
    p1 = v(4)
    Dim p2 As New PinList
    p2.value = v(6)
    p2p_short_Power_FVMI__ = VBAProject.VBT_LIB_DC_Conti.p2p_short_Power_FVMI(CStr(v(0)), CStr(v(1)), CStr(v(2)), CStr(v(3)), p1, CBool(v(5)), p2, CStr(v(7)), CStr(v(8)))
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function PPMU_Continuity_IV_Curve__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    Dim p1 As New PinList
    p1.value = v(0)
    Dim p2 As tlLimitForceResults
    p2 = v(6)
    PPMU_Continuity_IV_Curve__ = VBAProject.VBT_LIB_DC_Conti.PPMU_Continuity_IV_Curve(p1, CDbl(v(1)), CDbl(v(2)), CDbl(v(3)), CDbl(v(4)), CDbl(v(5)), p2, CBool(v(7)), CBool(v(8)))
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function GndSensePins_continuity__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    GndSensePins_continuity__ = VBAProject.VBT_LIB_DC_Conti.GndSensePins_continuity(CStr(v(0)), CStr(v(1)), CDbl(v(2)), CDbl(v(3)), CDbl(v(4)), CDbl(v(5)))
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function p2p_short_Power_FVMI_VI_Curve__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    Dim p1 As New PinList
    p1.value = v(6)
    p2p_short_Power_FVMI_VI_Curve__ = VBAProject.VBT_LIB_DC_Conti.p2p_short_Power_FVMI_VI_Curve(CStr(v(0)), CDbl(v(1)), CDbl(v(2)), CDbl(v(3)), CStr(v(4)), CStr(v(5)), p1, CStr(v(7)))
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function RetrieveDictionaryOfDiffPairs__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    RetrieveDictionaryOfDiffPairs__ = VBAProject.VBT_LIB_DC_Conti.RetrieveDictionaryOfDiffPairs()
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function PPMU_Measure_Contact_Resistance_Corner_Vss_bk__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    Dim p1 As New PinList
    p1.value = v(0)
    Dim p2 As tlLimitForceResults
    p2 = v(6)
    PPMU_Measure_Contact_Resistance_Corner_Vss_bk__ = VBAProject.VBT_LIB_DC_Conti.PPMU_Measure_Contact_Resistance_Corner_Vss_bk(p1, CDbl(v(1)), CDbl(v(2)), CDbl(v(3)), CDbl(v(4)), CDbl(v(5)), p2, CDbl(v(7)), CBool(v(8)), CBool(v(9)))
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function PPMU_Measure_Contact_Resistance_Corner_Vss__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    Dim p1 As New PinList
    p1.value = v(0)
    Dim p2 As tlLimitForceResults
    p2 = v(6)
    PPMU_Measure_Contact_Resistance_Corner_Vss__ = VBAProject.VBT_LIB_DC_Conti.PPMU_Measure_Contact_Resistance_Corner_Vss(p1, CDbl(v(1)), CDbl(v(2)), CDbl(v(3)), CDbl(v(4)), CDbl(v(5)), p2, CDbl(v(7)), CBool(v(8)), CBool(v(9)))
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function p2p_short_Power_FVMI_Parallel__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    Dim p1 As tlLimitForceResults
    p1 = v(7)
    Dim p2 As New PinList
    p2.value = v(9)
    p2p_short_Power_FVMI_Parallel__ = VBAProject.VBT_LIB_DC_Conti.p2p_short_Power_FVMI_Parallel(CStr(v(0)), CStr(v(1)), CStr(v(2)), CStr(v(3)), CDbl(v(4)), CDbl(v(5)), CDbl(v(6)), p1, CBool(v(8)), p2, CStr(v(10)))
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function AutoZ_Continuity__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    Dim p1 As tlLimitForceResults
    p1 = v(4)
    AutoZ_Continuity__ = VBAProject.VBT_LIB_DC_Conti.AutoZ_Continuity(CStr(v(0)), CStr(v(1)), CStr(v(2)), CStr(v(3)), p1, CBool(v(5)), CStr(v(6)), CStr(v(7)), CStr(v(8)), CBool(v(9)), CStr(v(10)), CStr(v(11)))
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function PPMU_Impedance_Function__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    PPMU_Impedance_Function__ = VBAProject.VBT_LIB_DC_Conti.PPMU_Impedance_Function(CStr(v(0)), CStr(v(1)), CDbl(v(2)), CDbl(v(3)), CDbl(v(4)), CDbl(v(5)))
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function PowerPin_CRES_Calculation__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    Dim p1 As New PinList
    p1.value = v(0)
    Dim p2 As tlLimitForceResults
    p2 = v(7)
    PowerPin_CRES_Calculation__ = VBAProject.VBT_LIB_DC_Conti.PowerPin_CRES_Calculation(p1, CDbl(v(1)), CDbl(v(2)), CDbl(v(3)), CDbl(v(4)), CDbl(v(5)), CDbl(v(6)), p2, CStr(v(8)), CStr(v(9)), CStr(v(10)), CStr(v(11)))
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function DC_conti_WalkingZ_PN_Disconnect__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    Dim p1 As New Pattern
    p1.value = v(0)
    Dim p2 As New PinList
    p2.value = v(1)
    Dim p3 As New PinList
    p3.value = v(2)
    Dim p4 As New PinList
    p4.value = v(3)
    DC_conti_WalkingZ_PN_Disconnect__ = VBAProject.VBT_LIB_DC_Conti.DC_conti_WalkingZ_PN_Disconnect(p1, p2, p3, p4)
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function









































Public Function DC_Func_WriteFuncResult__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    DC_Func_WriteFuncResult__ = VBAProject.VBT_LIB_DC_Func.DC_Func_WriteFuncResult(CBool(v(0)))
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function IO_HardIP_PPMU_Measure_I_TTR__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    tl_dt_ErrorMsg "Unsupported type at function IO_HardIP_PPMU_Measure_I_TTR"

    ' IO_HardIP_PPMU_Measure_I_TTR__ = VBAProject.VBT_LIB_DC_Func.IO_HardIP_PPMU_Measure_I_TTR(*One or more unsupported types in argument list or non Long/Integer return type*)
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function IO_HardIP_PPMU_Measure_V_TTR__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    Dim p1 As InstrumentSpecialSetup
    p1 = v(11)
    Dim p2 As Enum_RAK
    p2 = v(12)
    tl_dt_ErrorMsg "Unsupported type at function IO_HardIP_PPMU_Measure_V_TTR"

    ' IO_HardIP_PPMU_Measure_V_TTR__ = VBAProject.VBT_LIB_DC_Func.IO_HardIP_PPMU_Measure_V_TTR(*One or more unsupported types in argument list or non Long/Integer return type*)
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function Meas_VIR_IO_Universal_func_GPIO_TTR__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    Dim p1 As New Pattern
    p1.value = v(0)
    Dim p2 As New PinList
    p2.value = v(3)
    Dim p3 As New PinList
    p3.value = v(4)
    Dim p4 As New PinList
    p4.value = v(11)
    Dim p5 As CalculateMethodSetup
    p5 = v(14)
    Dim p6 As New PinList
    p6.value = v(15)
    Dim p7 As InstrumentSpecialSetup
    p7 = v(21)
    Dim p8 As Enum_RAK
    p8 = v(23)
    Dim p9 As New InterposeName
    p9.value = v(29)
    Dim p10 As New InterposeName
    p10.value = v(31)
    Meas_VIR_IO_Universal_func_GPIO_TTR__ = VBAProject.VBT_LIB_DC_Func.Meas_VIR_IO_Universal_func_GPIO_TTR(p1, CStr(v(1)), CBool(v(2)), p2, p3, CBool(v(5)), CStr(v(6)), CStr(v(7)), CStr(v(8)), CStr(v(9)), CStr(v(10)), p4, CLng(v(12)), CLng(v(13)), p5, p6, CLng(v(16)), CLng(v(17)), CStr(v(18)), CStr(v(19)), CStr(v(20)), p7, CBool(v(22)), p8, CStr(v(24)), CStr(v(25)), CStr(v(26)), CBool(v(27)), CStr(v(28)), p9, CStr(v(30)), p10, CStr(v(32)), CStr(v(33)), CBool(v(34)), CStr(v(35)), CStr(v(36)), CStr(v(37)), CBool(v(38)), CStr(v(39)), CInt(v(40)), CStr(v(41)), CStr(v(42)), CBool(v(UBound(v))))
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function EvaluateEachBlock__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    tl_dt_ErrorMsg "Unsupported type at function EvaluateEachBlock"

    ' EvaluateEachBlock__ = VBAProject.VBT_LIB_DC_Func.EvaluateEachBlock(*One or more unsupported types in argument list or non Long/Integer return type*)
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function GetFlowSingleUseLimit_KeepEmpty__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    tl_dt_ErrorMsg "Unsupported type at function GetFlowSingleUseLimit_KeepEmpty"

    ' GetFlowSingleUseLimit_KeepEmpty__ = VBAProject.VBT_LIB_DC_Func.GetFlowSingleUseLimit_KeepEmpty(*One or more unsupported types in argument list or non Long/Integer return type*)
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function EVS_Static_Power_Ramp__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    EVS_Static_Power_Ramp__ = VBAProject.VBT_LIB_DC_Func.EVS_Static_Power_Ramp(CStr(v(0)), CDbl(v(1)), CStr(v(2)), CInt(v(3)), CDbl(v(4)), CBool(v(5)), CDbl(v(6)), CStr(v(7)), CStr(v(8)), CBool(v(9)), CBool(v(10)), CInt(v(11)), CBool(v(12)), CDbl(v(13)), CDbl(v(14)), CBool(v(15)), CStr(v(16)))
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function EVS_Pre_Setting__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    tl_dt_ErrorMsg "Unsupported type at function EVS_Pre_Setting"

    ' EVS_Pre_Setting__ = VBAProject.VBT_LIB_DC_Func.EVS_Pre_Setting(*One or more unsupported types in argument list or non Long/Integer return type*)
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function Evs_Ramp_UPorDown__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    tl_dt_ErrorMsg "Unsupported type at function Evs_Ramp_UPorDown"

    ' Evs_Ramp_UPorDown__ = VBAProject.VBT_LIB_DC_Func.Evs_Ramp_UPorDown(*One or more unsupported types in argument list or non Long/Integer return type*)
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function Test_time_breakdown_End__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    Test_time_breakdown_End__ = VBAProject.VBT_LIB_DC_Func.Test_time_breakdown_End(CDbl(v(0)), CBool(v(1)), CStr(v(2)), CInt(v(3)), CDbl(v(4)))
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function Test_time_breakdown_Start__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    Test_time_breakdown_Start__ = VBAProject.VBT_LIB_DC_Func.Test_time_breakdown_Start(CDbl(v(0)), CBool(v(1)), CStr(v(2)))
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function auto_Checkboard_EVS_Probe_Location__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    auto_Checkboard_EVS_Probe_Location__ = VBAProject.VBT_LIB_DC_Func.auto_Checkboard_EVS_Probe_Location()
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function EVS_Static_Power_Ramp_PAEVS__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    EVS_Static_Power_Ramp_PAEVS__ = VBAProject.VBT_LIB_DC_Func.EVS_Static_Power_Ramp_PAEVS(CStr(v(0)), CDbl(v(1)), CStr(v(2)), CInt(v(3)), CDbl(v(4)), CBool(v(5)), CDbl(v(6)), CStr(v(7)), CStr(v(8)), CBool(v(9)), CBool(v(10)), CInt(v(11)), CBool(v(12)), CDbl(v(13)), CStr(v(14)), CStr(v(15)), CStr(v(16)), CBool(v(17)))
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function SetCurrentRange_From_Ifold__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    SetCurrentRange_From_Ifold__ = VBAProject.VBT_LIB_DC_Func.SetCurrentRange_From_Ifold(CStr(v(0)))
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function EVS_Pre_Setting_wPACAVS__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    tl_dt_ErrorMsg "Unsupported type at function EVS_Pre_Setting_wPACAVS"

    ' EVS_Pre_Setting_wPACAVS__ = VBAProject.VBT_LIB_DC_Func.EVS_Pre_Setting_wPACAVS(*One or more unsupported types in argument list or non Long/Integer return type*)
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function Evs_Ramp_UPorDown_PAEVS__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    tl_dt_ErrorMsg "Unsupported type at function Evs_Ramp_UPorDown_PAEVS"

    ' Evs_Ramp_UPorDown_PAEVS__ = VBAProject.VBT_LIB_DC_Func.Evs_Ramp_UPorDown_PAEVS(*One or more unsupported types in argument list or non Long/Integer return type*)
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function Test_time_breakdown_PAEVS_End__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    Test_time_breakdown_PAEVS_End__ = VBAProject.VBT_LIB_DC_Func.Test_time_breakdown_PAEVS_End(CDbl(v(0)), CBool(v(1)), CStr(v(2)), CInt(v(3)), CDbl(v(4)), CDbl(v(5)), CStr(v(6)))
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function EVS_Cooling__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    EVS_Cooling__ = VBAProject.VBT_LIB_DC_Func.EVS_Cooling(CDbl(v(0)))
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function EVS_TestFunction__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    EVS_TestFunction__ = VBAProject.VBT_LIB_DC_Func.EVS_TestFunction(CLng(v(0)))
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function EVS_setbinnum__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    EVS_setbinnum__ = VBAProject.VBT_LIB_DC_Func.EVS_setbinnum()
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function









































Public Function DCVS_IDS_main_auto_range_and_measure__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    tl_dt_ErrorMsg "Unsupported type at function DCVS_IDS_main_auto_range_and_measure"

    ' DCVS_IDS_main_auto_range_and_measure__ = VBAProject.VBT_LIB_DC_IDS.DCVS_IDS_main_auto_range_and_measure(*One or more unsupported types in argument list or non Long/Integer return type*)
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function DCVI_IDS_main_auto_range_and_measure__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    tl_dt_ErrorMsg "Unsupported type at function DCVI_IDS_main_auto_range_and_measure"

    ' DCVI_IDS_main_auto_range_and_measure__ = VBAProject.VBT_LIB_DC_IDS.DCVI_IDS_main_auto_range_and_measure(*One or more unsupported types in argument list or non Long/Integer return type*)
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function IDS_main_current__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    Dim p1 As New Pattern
    p1.value = v(0)
    Dim p2 As New PinList
    p2.value = v(24)
    IDS_main_current__ = VBAProject.VBT_LIB_DC_IDS.IDS_main_current(p1, CStr(v(1)), CStr(v(2)), CStr(v(3)), CLng(v(4)), CBool(v(5)), CStr(v(6)), CBool(v(7)), CBool(v(8)), CBool(v(9)), CStr(v(10)), CStr(v(11)), CStr(v(12)), CStr(v(13)), CStr(v(14)), CStr(v(15)), CBool(v(16)), CStr(v(17)), CDbl(v(18)), CBool(v(19)), CStr(v(20)), CBool(v(21)), CStr(v(22)), CStr(v(23)), p2, CLng(v(25)), CLng(v(26)), CStr(v(27)), CStr(v(28)), CStr(v(29)), CStr(v(30)), CStr(v(31)), CBool(v(32)), CStr(v(33)), CStr(v(34)), CBool(v(UBound(v))), CBool(v(36)), CStr(v(37)), CBool(v(38)), CStr(v(39)), CStr(v(40)), CStr(v(41)))
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function IDS_Delta_calc__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    IDS_Delta_calc__ = VBAProject.VBT_LIB_DC_IDS.IDS_Delta_calc(CStr(v(0)), CStr(v(1)), CStr(v(2)))
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function DCVS_IDS_main_current_ratio__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    DCVS_IDS_main_current_ratio__ = VBAProject.VBT_LIB_DC_IDS.DCVS_IDS_main_current_ratio(CStr(v(0)), CStr(v(1)), CStr(v(2)))
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function IDS_main_current_MathFunc__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    Dim p1 As New PinList
    p1.value = v(0)
    IDS_main_current_MathFunc__ = VBAProject.VBT_LIB_DC_IDS.IDS_main_current_MathFunc(p1, CStr(v(1)), CStr(v(2)), CBool(v(3)), CStr(v(4)), CStr(v(5)))
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function Parse_IDS_Mapping_Table__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    Parse_IDS_Mapping_Table__ = VBAProject.VBT_LIB_DC_IDS.Parse_IDS_Mapping_Table()
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function Check_IDS_Result_With_BinCut__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    Dim p1 As New PinList
    p1.value = v(0)
    Check_IDS_Result_With_BinCut__ = VBAProject.VBT_LIB_DC_IDS.Check_IDS_Result_With_BinCut(p1, CStr(v(1)), CStr(v(2)))
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function IDS_main_current_Binning_new__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    IDS_main_current_Binning_new__ = VBAProject.VBT_LIB_DC_IDS.IDS_main_current_Binning_new(CStr(v(0)), CStr(v(1)), CStr(v(2)), CBool(v(3)), CStr(v(4)), CLng(v(5)))
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function IDS_main_current_Binning__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    IDS_main_current_Binning__ = VBAProject.VBT_LIB_DC_IDS.IDS_main_current_Binning(CStr(v(0)), CStr(v(1)), CStr(v(2)))
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function Clear_IDS_Dic_for_each_die__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    Clear_IDS_Dic_for_each_die__ = VBAProject.VBT_LIB_DC_IDS.Clear_IDS_Dic_for_each_die()
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function DCVS_IDS_main_current_Read__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    Dim p1 As New PinList
    p1.value = v(0)
    DCVS_IDS_main_current_Read__ = VBAProject.VBT_LIB_DC_IDS.DCVS_IDS_main_current_Read(p1, CStr(v(1)))
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function DCVS_IDS_main_current_Read_CP2__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    Dim p1 As New PinList
    p1.value = v(0)
    DCVS_IDS_main_current_Read_CP2__ = VBAProject.VBT_LIB_DC_IDS.DCVS_IDS_main_current_Read_CP2(p1, CStr(v(1)))
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function









































Public Function HiZ_Leakage_Parallel__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    Dim p1 As New Pattern
    p1.value = v(0)
    Dim p2 As New PinList
    p2.value = v(4)
    HiZ_Leakage_Parallel__ = VBAProject.VBT_LIB_DC_Leak.HiZ_Leakage_Parallel(p1, CDbl(v(1)), CDbl(v(2)), CDbl(v(3)), p2)
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function HiZ_Leakage__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    Dim p1 As New Pattern
    p1.value = v(0)
    Dim p2 As New PinList
    p2.value = v(4)
    HiZ_Leakage__ = VBAProject.VBT_LIB_DC_Leak.HiZ_Leakage(p1, CDbl(v(1)), CDbl(v(2)), CDbl(v(3)), p2, CDbl(v(5)), CDbl(v(6)), CStr(v(7)), CBool(v(8)), CBool(v(9)))
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function HiZ_Leakage_MeasPower__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    Dim p1 As New Pattern
    p1.value = v(0)
    Dim p2 As New PinList
    p2.value = v(4)
    Dim p3 As New PinList
    p3.value = v(7)
    HiZ_Leakage_MeasPower__ = VBAProject.VBT_LIB_DC_Leak.HiZ_Leakage_MeasPower(p1, CDbl(v(1)), CDbl(v(2)), CDbl(v(3)), p2, CDbl(v(5)), CDbl(v(6)), p3, CDbl(v(8)), CDbl(v(9)), CStr(v(10)), CBool(v(11)), CStr(v(12)))
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function HiZ_Leakage_Parallel_GPIO__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    Dim p1 As New Pattern
    p1.value = v(0)
    Dim p2 As New PinList
    p2.value = v(4)
    HiZ_Leakage_Parallel_GPIO__ = VBAProject.VBT_LIB_DC_Leak.HiZ_Leakage_Parallel_GPIO(p1, CDbl(v(1)), CDbl(v(2)), CDbl(v(3)), p2, CDbl(v(5)), CDbl(v(6)), CBool(v(7)))
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function









































Public Function CP_DIBC_Start__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    Call VBAProject.VBT_LIB_DIBC_CP.CP_DIBC_Start
    CP_DIBC_Start__ = TL_SUCCESS
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function DIBC_PPMU_DisConnect__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    Call VBAProject.VBT_LIB_DIBC_CP.DIBC_PPMU_DisConnect(CStr(v(0)))
    DIBC_PPMU_DisConnect__ = TL_SUCCESS
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function DIBC_PPMU_Set0V__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    Call VBAProject.VBT_LIB_DIBC_CP.DIBC_PPMU_Set0V(CStr(v(0)))
    DIBC_PPMU_Set0V__ = TL_SUCCESS
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function DIBC_DCVI_DisConnect__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    Call VBAProject.VBT_LIB_DIBC_CP.DIBC_DCVI_DisConnect(CStr(v(0)), CStr(v(1)))
    DIBC_DCVI_DisConnect__ = TL_SUCCESS
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function DIBC_DCVI_Set0V__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    Call VBAProject.VBT_LIB_DIBC_CP.DIBC_DCVI_Set0V(CStr(v(0)), CStr(v(1)))
    DIBC_DCVI_Set0V__ = TL_SUCCESS
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function DIBC_DCVS_DisConnect__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    Call VBAProject.VBT_LIB_DIBC_CP.DIBC_DCVS_DisConnect(CStr(v(0)))
    DIBC_DCVS_DisConnect__ = TL_SUCCESS
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function DIBC_DCVS_Set0V__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    Call VBAProject.VBT_LIB_DIBC_CP.DIBC_DCVS_Set0V(CStr(v(0)))
    DIBC_DCVS_Set0V__ = TL_SUCCESS
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function









































Public Function FT_DIBC_Start__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    Call VBAProject.VBT_LIB_DIBC_FT.FT_DIBC_Start
    FT_DIBC_Start__ = TL_SUCCESS
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function









































Public Function WLFT_DIBC_Start__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    Call VBAProject.VBT_LIB_DIBC_WLFT.WLFT_DIBC_Start
    WLFT_DIBC_Start__ = TL_SUCCESS
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function









































Public Function Functional_T_updated__(v As Variant) As Long
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
    Dim p19 As CusWaitVal
    p19 = v(27)
    Dim p20 As CusWaitVal
    p20 = v(28)
    Dim p21 As CusWaitVal
    p21 = v(29)
    Dim p22 As CusWaitVal
    p22 = v(30)
    Dim pStep As SubType
    pStep = TheExec.flow.StepType
    Dim p23 As tlPatConcurrentMode
    p23 = v(34)
    Dim p24 As tlTemplateScanFailDataLogging
    p24 = v(43)
    Dim p25 As tlDigitalCMEMCaptureLimitMode
    p25 = v(44)
    Dim p26 As tlTemplateScanPinListSource
    p26 = v(46)
    Dim p27 As New PinList
    p27.value = v(47)
    Dim p28 As tlTemplateScanCaptureFormat
    p28 = v(48)
    Dim p29 As tlTemplateScanCaptureDataType
    p29 = v(49)
    Dim p30 As tlTemplateScanUserCommentSource
    p30 = v(50)
    Dim p31 As tlTemplateATPGPinMapSource
    p31 = v(52)
    Functional_T_updated__ = VBAProject.VBT_LIB_Digital_Functional_T.Functional_T_updated(p1, p2, p3, p4, p5, p6, p7, p8, p9, p10, p11, p12, p13, p14, CStr(v(14)), CStr(v(15)), CStr(v(16)), CStr(v(17)), CStr(v(18)), CStr(v(19)), p15, p16, p17, CStr(v(23)), p18, CBool(v(25)), CBool(v(26)), p19, p20, p21, p22, CBool(v(UBound(v))), CStr(v(32)), pStep, CStr(v(33)), p23, CStr(v(35)), CBool(v(36)), CBool(v(37)), CBool(v(38)), CStr(v(39)), CStr(v(40)), CStr(v(41)), CLng(v(42)), p24, p25, CLng(v(45)), p26, p27, p28, p29, p30, CStr(v(51)), p31, CStr(v(53)), CStr(v(54)), CStr(v(55)), CStr(v(56)), CStr(v(57)), CBool(v(58)))
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function DatalogType__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    DatalogType__ = VBAProject.VBT_LIB_Digital_Functional_T.DatalogType()
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function PostTest__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    tl_dt_ErrorMsg "Unsupported type at function PostTest"

    ' Call VBAProject.VBT_LIB_Digital_Functional_T.PostTest(*One or more unsupported types in argument list or non Long/Integer return type*)
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

    ' getdefaults__ = VBAProject.VBT_LIB_Digital_Functional_T.getdefaults(*One or more unsupported types in argument list or non Long/Integer return type*)
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function pattern_module_test__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    Dim p1 As PFType
    p1 = v(3)
    Dim p2 As tlResultMode
    p2 = v(5)
    Dim p3 As tlPatConcurrentMode
    p3 = v(6)
    tl_dt_ErrorMsg "Unsupported type at function pattern_module_test"

    ' pattern_module_test__ = VBAProject.VBT_LIB_Digital_Functional_T.pattern_module_test(*One or more unsupported types in argument list or non Long/Integer return type*)
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function auto_Mbist_Block_loop_inst_match__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    auto_Mbist_Block_loop_inst_match__ = VBAProject.VBT_LIB_Digital_Functional_T.auto_Mbist_Block_loop_inst_match(CStr(v(0)), CStr(v(1)), CLng(v(2)), CStr(v(3)), CBool(v(4)), CBool(v(5)))
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function auto_Mbist_Block_loop_inst_non_match__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    auto_Mbist_Block_loop_inst_non_match__ = VBAProject.VBT_LIB_Digital_Functional_T.auto_Mbist_Block_loop_inst_non_match(CStr(v(0)), CStr(v(1)), CLng(v(2)), CStr(v(3)), CBool(v(4)), CBool(v(5)))
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function Sub_SourceEMA_SelSRM__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    Sub_SourceEMA_SelSRM__ = VBAProject.VBT_LIB_Digital_Functional_T.Sub_SourceEMA_SelSRM(CStr(v(0)), CStr(v(1)), CStr(v(2)))
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function SSN_CoreMask__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    SSN_CoreMask__ = VBAProject.VBT_LIB_Digital_Functional_T.SSN_CoreMask(CStr(v(0)))
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function









































Public Function Init_RSCR__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    Init_RSCR__ = VBAProject.VBT_LIB_Digital_Mbist.Init_RSCR()
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function Mbist_RSCR__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    Dim p1 As New Pattern
    p1.value = v(0)
    Mbist_RSCR__ = VBAProject.VBT_LIB_Digital_Mbist.Mbist_RSCR(p1, CStr(v(1)), CStr(v(2)), CStr(v(3)))
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function TurnOnEfusePwrPins_Mbist__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    TurnOnEfusePwrPins_Mbist__ = VBAProject.VBT_LIB_Digital_Mbist.TurnOnEfusePwrPins_Mbist(CStr(v(0)))
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function TurnOffEfusePwrPins_Mbist__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    TurnOffEfusePwrPins_Mbist__ = VBAProject.VBT_LIB_Digital_Mbist.TurnOffEfusePwrPins_Mbist(CStr(v(0)))
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function MbistRetentionLevelWait_and_lowDown_power__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    Dim p1 As New PinList
    p1.value = v(1)
    MbistRetentionLevelWait_and_lowDown_power__ = VBAProject.VBT_LIB_Digital_Mbist.MbistRetentionLevelWait_and_lowDown_power(CDbl(v(0)), p1, CDbl(v(2)))
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function MbistRetentionLevelWait_new__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    Dim p1 As New PinList
    p1.value = v(6)
    Dim p2 As New PinList
    p2.value = v(7)
    MbistRetentionLevelWait_new__ = VBAProject.VBT_LIB_Digital_Mbist.MbistRetentionLevelWait_new(CDbl(v(0)), CDbl(v(1)), CDbl(v(2)), CBool(v(3)), CBool(v(4)), CBool(v(5)), p1, p2, CStr(v(8)), CStr(v(9)))
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function MbistRetentionLevelWait__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    MbistRetentionLevelWait__ = VBAProject.VBT_LIB_Digital_Mbist.MbistRetentionLevelWait(CDbl(v(0)), CDbl(v(1)), CDbl(v(2)), CBool(v(3)), CBool(v(4)), CStr(v(5)))
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function Init_MBISTFailBlock__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    Init_MBISTFailBlock__ = VBAProject.VBT_LIB_Digital_Mbist.Init_MBISTFailBlock()
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function GetFlagInfoArrIndex__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    GetFlagInfoArrIndex__ = VBAProject.VBT_LIB_Digital_Mbist.GetFlagInfoArrIndex(CStr(v(0)))
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function MbistRetentionWait__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    MbistRetentionWait__ = VBAProject.VBT_LIB_Digital_Mbist.MbistRetentionWait(CDbl(v(0)))
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function









































Public Function TPmode_Char_on__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    TPmode_Char_on__ = VBAProject.VBT_LIB_Digital_Shmoo.TPmode_Char_on()
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function TPmode_Char_off__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    TPmode_Char_off__ = VBAProject.VBT_LIB_Digital_Shmoo.TPmode_Char_off()
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function freerunclk_set_XY__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    tl_dt_ErrorMsg "Unsupported type at function freerunclk_set_XY"

    ' freerunclk_set_XY__ = VBAProject.VBT_LIB_Digital_Shmoo.freerunclk_set_XY(*One or more unsupported types in argument list or non Long/Integer return type*)
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function freerunclk_stop__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    tl_dt_ErrorMsg "Unsupported type at function freerunclk_stop"

    ' freerunclk_stop__ = VBAProject.VBT_LIB_Digital_Shmoo.freerunclk_stop(*One or more unsupported types in argument list or non Long/Integer return type*)
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function CharStoreResultsUntilNextRun__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    CharStoreResultsUntilNextRun__ = VBAProject.VBT_LIB_Digital_Shmoo.CharStoreResultsUntilNextRun()
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function run_shmoo__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    run_shmoo__ = VBAProject.VBT_LIB_Digital_Shmoo.run_shmoo(CStr(v(0)))
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function Functional_T_char__(v As Variant) As Long
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
    Dim p7 As PFType
    p7 = v(6)
    Dim p8 As tlResultMode
    p8 = v(7)
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
    p14.value = v(19)
    Dim p15 As New PinList
    p15.value = v(20)
    Dim p16 As New InterposeName
    p16.value = v(21)
    Dim pStep As SubType
    pStep = TheExec.flow.StepType
    Dim p17 As tlPatConcurrentMode
    p17 = v(26)
    Dim p18 As New Pattern
    p18.value = v(28)
    Dim p19 As New Pattern
    p19.value = v(29)
    Functional_T_char__ = VBAProject.VBT_LIB_Digital_Shmoo.Functional_T_char(p1, p2, p3, p4, p5, p6, p7, p8, p9, p10, p11, p12, p13, CStr(v(13)), CStr(v(14)), CStr(v(15)), CStr(v(16)), CStr(v(17)), CStr(v(18)), p14, p15, p16, CStr(v(22)), CBool(v(UBound(v))), CStr(v(24)), pStep, CStr(v(25)), p17, CStr(v(27)), p18, p19, CStr(v(30)), CStr(v(31)), CStr(v(32)), CStr(v(33)), CStr(v(34)), CStr(v(35)), CStr(v(36)), CBool(v(37)), CStr(v(38)), CStr(v(39)), CStr(v(40)), CBool(v(41)), CBool(v(42)), CStr(v(43)), CStr(v(44)), CBool(v(45)), CStr(v(46)), CStr(v(47)), CBool(v(48)))
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function PrintShmooInfo__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    tl_dt_ErrorMsg "Unsupported type at function PrintShmooInfo"

    ' PrintShmooInfo__ = VBAProject.VBT_LIB_Digital_Shmoo.PrintShmooInfo(*One or more unsupported types in argument list or non Long/Integer return type*)
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function Flow_Shmoo_Setup__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    Flow_Shmoo_Setup__ = VBAProject.VBT_LIB_Digital_Shmoo.Flow_Shmoo_Setup()
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function CheckCharErrorCount__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    CheckCharErrorCount__ = VBAProject.VBT_LIB_Digital_Shmoo.CheckCharErrorCount(CStr(v(0)), CDbl(v(1)))
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function EnableShmooAbnormalCounter__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    EnableShmooAbnormalCounter__ = VBAProject.VBT_LIB_Digital_Shmoo.EnableShmooAbnormalCounter()
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function DisableShmooAbnormalCounter__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    DisableShmooAbnormalCounter__ = VBAProject.VBT_LIB_Digital_Shmoo.DisableShmooAbnormalCounter()
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function Re_PowerOn_WhileSweep__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    tl_dt_ErrorMsg "Unsupported type at function Re_PowerOn_WhileSweep"

    ' Re_PowerOn_WhileSweep__ = VBAProject.VBT_LIB_Digital_Shmoo.Re_PowerOn_WhileSweep(*One or more unsupported types in argument list or non Long/Integer return type*)
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function DCVS_PowerUp_Parallel_Interpose__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    DCVS_PowerUp_Parallel_Interpose__ = VBAProject.VBT_LIB_Digital_Shmoo.DCVS_PowerUp_Parallel_Interpose(CStr(v(0)), CStr(v(1)), CDbl(v(2)), CBool(v(3)))
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function DCVS_PowerDown_Parallel_Interpose__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    DCVS_PowerDown_Parallel_Interpose__ = VBAProject.VBT_LIB_Digital_Shmoo.DCVS_PowerDown_Parallel_Interpose(CStr(v(0)), CDbl(v(1)), CBool(v(2)))
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function StoreMaxNum__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    tl_dt_ErrorMsg "Unsupported type at function StoreMaxNum"

    ' StoreMaxNum__ = VBAProject.VBT_LIB_Digital_Shmoo.StoreMaxNum(*One or more unsupported types in argument list or non Long/Integer return type*)
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function InverStr__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    tl_dt_ErrorMsg "Unsupported type at function InverStr"

    ' InverStr__ = VBAProject.VBT_LIB_Digital_Shmoo.InverStr(*One or more unsupported types in argument list or non Long/Integer return type*)
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function CharacterizationDigSrcPrePoint__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    tl_dt_ErrorMsg "Unsupported type at function CharacterizationDigSrcPrePoint"

    ' CharacterizationDigSrcPrePoint__ = VBAProject.VBT_LIB_Digital_Shmoo.CharacterizationDigSrcPrePoint(*One or more unsupported types in argument list or non Long/Integer return type*)
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function ApplySafeVoltage_W_WaitTime__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    ApplySafeVoltage_W_WaitTime__ = VBAProject.VBT_LIB_Digital_Shmoo.ApplySafeVoltage_W_WaitTime(CDbl(v(0)), CBool(v(1)))
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function Adjust_IndexbyFlag__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    Adjust_IndexbyFlag__ = VBAProject.VBT_LIB_Digital_Shmoo.Adjust_IndexbyFlag(CStr(v(0)), CStr(v(1)))
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function









































Public Function UARTTest_CMEM_T__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    Dim p1 As New Pattern
    p1.value = v(0)
    UARTTest_CMEM_T__ = VBAProject.VBT_LIB_Digital_UART.UARTTest_CMEM_T(p1)
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function UART_LoopBackTest__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    UART_LoopBackTest__ = VBAProject.VBT_LIB_Digital_UART.UART_LoopBackTest()
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function ReadCMEM__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    ReadCMEM__ = VBAProject.VBT_LIB_Digital_UART.ReadCMEM()
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function UART_read_n_byte_DSP__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    UART_read_n_byte_DSP__ = VBAProject.VBT_LIB_Digital_UART.UART_read_n_byte_DSP(CLng(v(0)))
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function UART_write_n_byte_DSP__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    UART_write_n_byte_DSP__ = VBAProject.VBT_LIB_Digital_UART.UART_write_n_byte_DSP(CLng(v(0)), CStr(v(1)))
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function PreLoad_PA_Modules__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    PreLoad_PA_Modules__ = VBAProject.VBT_LIB_Digital_UART.PreLoad_PA_Modules(CStr(v(0)), CLng(v(1)), CStr(v(2)), CStr(v(3)))
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function VaryFreq_PA_UART__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    VaryFreq_PA_UART__ = VBAProject.VBT_LIB_Digital_UART.VaryFreq_PA_UART(CStr(v(0)), CDbl(v(1)), CStr(v(2)))
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function UARTReadRegDSP__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    UARTReadRegDSP__ = VBAProject.VBT_LIB_Digital_UART.UARTReadRegDSP()
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function UARTWriteRegDSP__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    UARTWriteRegDSP__ = VBAProject.VBT_LIB_Digital_UART.UARTWriteRegDSP(CLng(v(0)))
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function UARTTest_CMEM_T_Update__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    Dim p1 As New Pattern
    p1.value = v(0)
    UARTTest_CMEM_T_Update__ = VBAProject.VBT_LIB_Digital_UART.UARTTest_CMEM_T_Update(p1)
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function ReStartFRC__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    Call VBAProject.VBT_LIB_Digital_UART.ReStartFRC
    ReStartFRC__ = TL_SUCCESS
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function









































Public Function ECID_Dict_Buildup_S__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    ECID_Dict_Buildup_S__ = VBAProject.VBT_LIB_eFuse_ECID_Sorting_FT.ECID_Dict_Buildup_S()
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function BinSorting_Compare_FT_ECID_S__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    BinSorting_Compare_FT_ECID_S__ = VBAProject.VBT_LIB_eFuse_ECID_Sorting_FT.BinSorting_Compare_FT_ECID_S()
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function









































Public Function auto_CleanRegData_New__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    auto_CleanRegData_New__ = VBAProject.VBT_LIB_eFuse_Func.auto_CleanRegData_New()
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function auto_isBinaryString__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    tl_dt_ErrorMsg "Unsupported type at function auto_isBinaryString"

    ' auto_isBinaryString__ = VBAProject.VBT_LIB_eFuse_Func.auto_isBinaryString(*One or more unsupported types in argument list or non Long/Integer return type*)
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function UpdateDLogColumns__False__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    UpdateDLogColumns__False__ = VBAProject.VBT_LIB_eFuse_Func.UpdateDLogColumns__False()
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function auto_eFuse_CheckPatternFlag__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    auto_eFuse_CheckPatternFlag__ = VBAProject.VBT_LIB_eFuse_Func.auto_eFuse_CheckPatternFlag(CStr(v(0)))
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function EFUSE_Resistance__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    Dim p1 As New Pattern
    p1.value = v(0)
    Dim p2 As New Pattern
    p2.value = v(3)
    EFUSE_Resistance__ = VBAProject.VBT_LIB_eFuse_Func.EFUSE_Resistance(p1, CStr(v(1)), CDbl(v(2)), p2, CStr(v(4)), CBool(v(UBound(v))))
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function auto_UDR_UFP__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    Dim p1 As New Pattern
    p1.value = v(0)
    Dim p2 As New Pattern
    p2.value = v(3)
    auto_UDR_UFP__ = VBAProject.VBT_LIB_eFuse_Func.auto_UDR_UFP(p1, CStr(v(1)), CDbl(v(2)), p2, CStr(v(4)), CBool(v(UBound(v))))
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function auto_UDR_UFR__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    Dim p1 As New Pattern
    p1.value = v(0)
    Dim p2 As New Pattern
    p2.value = v(1)
    auto_UDR_UFR__ = VBAProject.VBT_LIB_eFuse_Func.auto_UDR_UFR(p1, p2, CStr(v(2)), CBool(v(UBound(v))))
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function auto_ShowECIDData__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    auto_ShowECIDData__ = VBAProject.VBT_LIB_eFuse_Func.auto_ShowECIDData()
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function auto_Function_Test__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    Dim p1 As New Pattern
    p1.value = v(0)
    Dim p2 As New Pattern
    p2.value = v(1)
    auto_Function_Test__ = VBAProject.VBT_LIB_eFuse_Func.auto_Function_Test(p1, p2, CStr(v(2)), CBool(v(UBound(v))))
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function auto_ConfigWrite_CFG_DV__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    Dim p1 As New Pattern
    p1.value = v(0)
    auto_ConfigWrite_CFG_DV__ = VBAProject.VBT_LIB_eFuse_Func.auto_ConfigWrite_CFG_DV(p1, CStr(v(1)), CDbl(v(2)), CBool(v(UBound(v))))
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function DRAM_eFuse_Write__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    DRAM_eFuse_Write__ = VBAProject.VBT_LIB_eFuse_Func.DRAM_eFuse_Write(CStr(v(0)))
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function DRAM_eFuse_CheckByHip__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    DRAM_eFuse_CheckByHip__ = VBAProject.VBT_LIB_eFuse_Func.DRAM_eFuse_CheckByHip(CStr(v(0)))
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function BinXYCheck__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    BinXYCheck__ = VBAProject.VBT_LIB_eFuse_Func.BinXYCheck()
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function CPM_eFuse_Write__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    CPM_eFuse_Write__ = VBAProject.VBT_LIB_eFuse_Func.CPM_eFuse_Write()
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function CPM_eFuse_Read__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    CPM_eFuse_Read__ = VBAProject.VBT_LIB_eFuse_Func.CPM_eFuse_Read()
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function ECID_ReadChk__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    ECID_ReadChk__ = VBAProject.VBT_LIB_eFuse_Func.ECID_ReadChk()
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function BankVarChk__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    BankVarChk__ = VBAProject.VBT_LIB_eFuse_Func.BankVarChk(CStr(v(0)))
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function Efuse_bincheck__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    Efuse_bincheck__ = VBAProject.VBT_LIB_eFuse_Func.Efuse_bincheck(CBool(v(0)))
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function Temp_eFuse_setWrite__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    Temp_eFuse_setWrite__ = VBAProject.VBT_LIB_eFuse_Func.Temp_eFuse_setWrite(CStr(v(0)), CStr(v(1)), CStr(v(2)))
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function BinFuse_eFuse_setWrite__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    BinFuse_eFuse_setWrite__ = VBAProject.VBT_LIB_eFuse_Func.BinFuse_eFuse_setWrite()
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function Temp_eFuse_setWriteZero__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    Temp_eFuse_setWriteZero__ = VBAProject.VBT_LIB_eFuse_Func.Temp_eFuse_setWriteZero(CStr(v(0)), CStr(v(1)))
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function ForceSepvmFusing_mod__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    ForceSepvmFusing_mod__ = VBAProject.VBT_LIB_eFuse_Func.ForceSepvmFusing_mod(CStr(v(0)), CStr(v(1)))
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function BinCut_PreSetAndCheck__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    BinCut_PreSetAndCheck__ = VBAProject.VBT_LIB_eFuse_Func.BinCut_PreSetAndCheck()
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function EFUSE_Flat_Pattern_Check__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    Dim p1 As New Pattern
    p1.value = v(0)
    Dim p2 As New Pattern
    p2.value = v(2)
    EFUSE_Flat_Pattern_Check__ = VBAProject.VBT_LIB_eFuse_Func.EFUSE_Flat_Pattern_Check(p1, CStr(v(1)), p2, CStr(v(3)), CBool(v(UBound(v))))
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function DRAM_eFuse_WriteByHip__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    DRAM_eFuse_WriteByHip__ = VBAProject.VBT_LIB_eFuse_Func.DRAM_eFuse_WriteByHip(CStr(v(0)))
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function DRAM_eFuse_CheckByEnableWd__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    DRAM_eFuse_CheckByEnableWd__ = VBAProject.VBT_LIB_eFuse_Func.DRAM_eFuse_CheckByEnableWd(CStr(v(0)))
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function Bincut_eFuse_OtherSetWrite__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    Bincut_eFuse_OtherSetWrite__ = VBAProject.VBT_LIB_eFuse_Func.Bincut_eFuse_OtherSetWrite(CStr(v(0)), CStr(v(1)), CStr(v(2)), CLng(v(3)))
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function EfuseExecInitPattern__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    tl_dt_ErrorMsg "Unsupported type at function EfuseExecInitPattern"

    ' EfuseExecInitPattern__ = VBAProject.VBT_LIB_eFuse_Func.EfuseExecInitPattern(*One or more unsupported types in argument list or non Long/Integer return type*)
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function DTS_AddStoredData__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    DTS_AddStoredData__ = VBAProject.VBT_LIB_eFuse_Func.DTS_AddStoredData(CStr(v(0)))
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function ECID_DTS__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    ECID_DTS__ = VBAProject.VBT_LIB_eFuse_Func.ECID_DTS()
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function PKGType_setWrite__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    PKGType_setWrite__ = VBAProject.VBT_LIB_eFuse_Func.PKGType_setWrite()
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function EVSsorting_setWrite__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    EVSsorting_setWrite__ = VBAProject.VBT_LIB_eFuse_Func.EVSsorting_setWrite(CStr(v(0)), CStr(v(1)), CStr(v(2)))
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function EVSsorting_setFlag__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    EVSsorting_setFlag__ = VBAProject.VBT_LIB_eFuse_Func.EVSsorting_setFlag(CStr(v(0)), CStr(v(1)), CStr(v(2)))
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function DieCorner_setWrite__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    DieCorner_setWrite__ = VBAProject.VBT_LIB_eFuse_Func.DieCorner_setWrite(CStr(v(0)), CStr(v(1)), CStr(v(2)), CStr(v(3)))
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function DieCorner_setFlag__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    DieCorner_setFlag__ = VBAProject.VBT_LIB_eFuse_Func.DieCorner_setFlag(CStr(v(0)), CStr(v(1)), CStr(v(2)))
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function ECID_Dict_Buildup__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    ECID_Dict_Buildup__ = VBAProject.VBT_LIB_eFuse_Func.ECID_Dict_Buildup()
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function BinSorting_Compare__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    BinSorting_Compare__ = VBAProject.VBT_LIB_eFuse_Func.BinSorting_Compare()
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function









































Public Function auto_Dummy_Item__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    Dim p1 As New Pattern
    p1.value = v(0)
    auto_Dummy_Item__ = VBAProject.VBT_LIB_EVS.auto_Dummy_Item(p1)
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function IFold_Max_Cal__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    tl_dt_ErrorMsg "Unsupported type at function IFold_Max_Cal"

    ' IFold_Max_Cal__ = VBAProject.VBT_LIB_EVS.IFold_Max_Cal(*One or more unsupported types in argument list or non Long/Integer return type*)
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function PrintCorePowerGate__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    tl_dt_ErrorMsg "Unsupported type at function PrintCorePowerGate"

    ' PrintCorePowerGate__ = VBAProject.VBT_LIB_EVS.PrintCorePowerGate(*One or more unsupported types in argument list or non Long/Integer return type*)
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function CheckCorePowerGate__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    tl_dt_ErrorMsg "Unsupported type at function CheckCorePowerGate"

    ' CheckCorePowerGate__ = VBAProject.VBT_LIB_EVS.CheckCorePowerGate(*One or more unsupported types in argument list or non Long/Integer return type*)
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function EVS_Check_AlarmTimeOut__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    EVS_Check_AlarmTimeOut__ = VBAProject.VBT_LIB_EVS.EVS_Check_AlarmTimeOut(CStr(v(0)))
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function EVS_Power_Ramp__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    EVS_Power_Ramp__ = VBAProject.VBT_LIB_EVS.EVS_Power_Ramp(CStr(v(0)), CStr(v(1)), CDbl(v(2)), CDbl(v(3)), CStr(v(4)), CBool(v(5)), CLng(v(6)), CDbl(v(7)), CDbl(v(8)), CBool(v(9)), CBool(v(10)), CStr(v(11)), CStr(v(12)), CDbl(v(13)), CBool(v(14)), CDbl(v(15)), CBool(v(16)), CStr(v(17)), CLng(v(18)), CBool(v(19)), CLng(v(20)))
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function EVS_Pins_Index__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    tl_dt_ErrorMsg "Unsupported type at function EVS_Pins_Index"

    ' EVS_Pins_Index__ = VBAProject.VBT_LIB_EVS.EVS_Pins_Index(*One or more unsupported types in argument list or non Long/Integer return type*)
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function DVS_Power_Ramp__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    Dim p1 As New Pattern
    p1.value = v(5)
    Dim p2 As tlResultMode
    p2 = v(7)
    DVS_Power_Ramp__ = VBAProject.VBT_LIB_EVS.DVS_Power_Ramp(CStr(v(0)), CStr(v(1)), CDbl(v(2)), CDbl(v(3)), CStr(v(4)), p1, CLng(v(6)), p2)
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function Start_Profile_EVS__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    Dim p1 As New PinList
    p1.value = v(0)
    Start_Profile_EVS__ = VBAProject.VBT_LIB_EVS.Start_Profile_EVS(p1, CStr(v(1)), CDbl(v(2)), CLng(v(3)), CStr(v(4)))
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function Plot_Profile_EVS__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    Dim p1 As New PinList
    p1.value = v(0)
    Plot_Profile_EVS__ = VBAProject.VBT_LIB_EVS.Plot_Profile_EVS(p1, CStr(v(1)), CStr(v(2)), CStr(v(3)), CBool(v(4)), CStr(v(5)))
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function









































Public Function Meas_Vdiff_func__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    Dim p1 As New Pattern
    p1.value = v(0)
    Dim p2 As New PinList
    p2.value = v(1)
    Dim p3 As New PinList
    p3.value = v(2)
    Dim p4 As New PinList
    p4.value = v(7)
    Dim p5 As New PinList
    p5.value = v(10)
    Meas_Vdiff_func__ = VBAProject.VBT_LIB_HardIP.Meas_Vdiff_func(p1, p2, p3, CStr(v(3)), CBool(v(4)), CStr(v(5)), CStr(v(6)), p4, CLng(v(8)), CLng(v(9)), p5, CStr(v(11)), CStr(v(12)), CStr(v(13)), CStr(v(14)), CStr(v(15)), CStr(v(16)), CStr(v(17)), CStr(v(18)), CStr(v(19)), CStr(v(20)), CStr(v(21)), CStr(v(22)), CStr(v(23)), CStr(v(24)), CStr(v(25)), CStr(v(26)), CStr(v(27)), CStr(v(28)), CStr(v(29)), CBool(v(UBound(v))))
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function Meas_VIR_IO_Universal_func__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    Dim p1 As New Pattern
    p1.value = v(0)
    Dim p2 As New PinList
    p2.value = v(3)
    Dim p3 As New PinList
    p3.value = v(4)
    Dim p4 As New PinList
    p4.value = v(11)
    Dim p5 As CalculateMethodSetup
    p5 = v(14)
    Dim p6 As New PinList
    p6.value = v(15)
    Dim p7 As InstrumentSpecialSetup
    p7 = v(21)
    Dim p8 As CalculateMethodSetup
    p8 = v(22)
    Dim p9 As Enum_RAK
    p9 = v(23)
    Meas_VIR_IO_Universal_func__ = VBAProject.VBT_LIB_HardIP.Meas_VIR_IO_Universal_func(p1, CStr(v(1)), CBool(v(2)), p2, p3, CBool(v(5)), CStr(v(6)), CStr(v(7)), CStr(v(8)), CStr(v(9)), CStr(v(10)), p4, CLng(v(12)), CLng(v(13)), p5, p6, CLng(v(16)), CLng(v(17)), CStr(v(18)), CStr(v(19)), CStr(v(20)), p7, p8, p9, CStr(v(24)), CStr(v(25)), CStr(v(26)), CBool(v(27)), CStr(v(28)), CBool(v(29)), CStr(v(30)), CStr(v(31)), CStr(v(32)), CStr(v(33)), CStr(v(34)), CStr(v(35)), CBool(v(UBound(v))))
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function Meas_FreqVoltCurr_Universal_func__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    Dim p1 As New Pattern
    p1.value = v(0)
    Dim p2 As New PinList
    p2.value = v(3)
    Dim p3 As New PinList
    p3.value = v(4)
    Dim p4 As EventSourceWithTerminationMode
    p4 = v(10)
    Dim p5 As New PinList
    p5.value = v(17)
    Dim p6 As New PinList
    p6.value = v(20)
    Dim p7 As CalculateMethodSetup
    p7 = v(26)
    Dim p8 As InstrumentSpecialSetup
    p8 = v(27)
    Dim p9 As Enum_RAK
    p9 = v(52)
    Meas_FreqVoltCurr_Universal_func__ = VBAProject.VBT_LIB_HardIP.Meas_FreqVoltCurr_Universal_func(p1, CStr(v(1)), CBool(v(2)), p2, p3, CBool(v(5)), CStr(v(6)), CStr(v(7)), CStr(v(8)), CStr(v(9)), p4, CBool(v(11)), CDbl(v(12)), CStr(v(13)), CStr(v(14)), CStr(v(15)), CStr(v(16)), p5, CLng(v(18)), CLng(v(19)), p6, CStr(v(21)), CStr(v(22)), CStr(v(23)), CStr(v(24)), CStr(v(25)), p7, p8, CStr(v(28)), CStr(v(29)), CStr(v(30)), CBool(v(31)), CStr(v(32)), CStr(v(33)), CBool(v(34)), CBool(v(35)), CDbl(v(36)), CDbl(v(37)), CDbl(v(38)), CDbl(v(39)), CDbl(v(40)), CDbl(v(41)), CStr(v(42)), CStr(v(43)), CStr(v(44)), CStr(v(45)), CStr(v(46)), CStr(v(47)), CStr(v(48)), CStr(v(49)), CStr(v(50)), CStr(v(51)), p9, CStr(v(53)), CBool(v(54)), CBool(v(55)), CStr(v(56)), CBool(v(UBound(v))))
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function Opt_DdrLpBkFunc2__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    Dim p1 As New Pattern
    p1.value = v(0)
    Dim p2 As New Pattern
    p2.value = v(1)
    Dim p3 As New PinList
    p3.value = v(2)
    Dim p4 As New PinList
    p4.value = v(3)
    Dim p5 As New PinList
    p5.value = v(4)
    Dim p6 As New PinList
    p6.value = v(9)
    Dim p7 As CalculateMethodSetup
    p7 = v(16)
    Opt_DdrLpBkFunc2__ = VBAProject.VBT_LIB_HardIP.Opt_DdrLpBkFunc2(p1, p2, p3, p4, p5, CLng(v(5)), CStr(v(6)), CStr(v(7)), CBool(v(8)), p6, CLng(v(10)), CStr(v(11)), CStr(v(12)), CStr(v(13)), CStr(v(14)), CStr(v(15)), p7, CStr(v(17)), CLng(v(18)), CStr(v(19)), CInt(v(20)), CStr(v(21)), CStr(v(22)), CStr(v(23)), CStr(v(24)), CStr(v(25)), CBool(v(UBound(v))))
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function Eye_Diagram__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    Eye_Diagram__ = VBAProject.VBT_LIB_HardIP.Eye_Diagram(CLng(v(0)))
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function PPMU_Impedance_Function_bk__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    PPMU_Impedance_Function_bk__ = VBAProject.VBT_LIB_HardIP.PPMU_Impedance_Function_bk(CStr(v(0)), CStr(v(1)), CDbl(v(2)), CDbl(v(3)), CDbl(v(4)), CDbl(v(5)))
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function TrimCodeDig_BK__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    Dim p1 As New Pattern
    p1.value = v(0)
    Dim p2 As New PinList
    p2.value = v(3)
    Dim p3 As New PinList
    p3.value = v(4)
    Dim p4 As New PinList
    p4.value = v(8)
    TrimCodeDig_BK__ = VBAProject.VBT_LIB_HardIP.TrimCodeDig_BK(p1, CStr(v(1)), CBool(v(2)), p2, p3, CStr(v(5)), CStr(v(6)), CStr(v(7)), p4, CStr(v(9)), CLng(v(10)), CDbl(v(11)), CDbl(v(12)), CStr(v(13)), CStr(v(14)), CLng(v(15)), CStr(v(16)), CStr(v(17)), CStr(v(18)), CStr(v(19)), CBool(v(20)), CBool(v(21)), CBool(v(22)), CBool(v(23)), CBool(v(UBound(v))))
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function TrimCodeDig__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    Dim p1 As New Pattern
    p1.value = v(0)
    Dim p2 As New PinList
    p2.value = v(3)
    Dim p3 As New PinList
    p3.value = v(4)
    Dim p4 As New PinList
    p4.value = v(8)
    TrimCodeDig__ = VBAProject.VBT_LIB_HardIP.TrimCodeDig(p1, CStr(v(1)), CBool(v(2)), p2, p3, CStr(v(5)), CStr(v(6)), CStr(v(7)), p4, CStr(v(9)), CLng(v(10)), CDbl(v(11)), CDbl(v(12)), CStr(v(13)), CStr(v(14)), CLng(v(15)), CStr(v(16)), CStr(v(17)), CStr(v(18)), CStr(v(19)), CBool(v(20)), CBool(v(21)), CBool(v(22)), CBool(v(23)), CBool(v(24)), CBool(v(UBound(v))))
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function TrimCodeDig_Tahiti__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    Dim p1 As New Pattern
    p1.value = v(0)
    Dim p2 As New PinList
    p2.value = v(3)
    Dim p3 As New PinList
    p3.value = v(4)
    Dim p4 As New PinList
    p4.value = v(18)
    TrimCodeDig_Tahiti__ = VBAProject.VBT_LIB_HardIP.TrimCodeDig_Tahiti(p1, CStr(v(1)), CBool(v(2)), p2, p3, CStr(v(5)), CStr(v(6)), CBool(v(7)), CBool(v(8)), CBool(v(9)), CDbl(v(10)), CDbl(v(11)), CStr(v(12)), CStr(v(13)), CStr(v(14)), CStr(v(15)), CStr(v(16)), CStr(v(17)), p4, CLng(v(19)), CBool(v(UBound(v))), CStr(v(21)))
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function TrimCodeDig_Reverse_Ellis__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    Dim p1 As New Pattern
    p1.value = v(0)
    Dim p2 As New PinList
    p2.value = v(3)
    Dim p3 As New PinList
    p3.value = v(4)
    Dim p4 As New PinList
    p4.value = v(18)
    TrimCodeDig_Reverse_Ellis__ = VBAProject.VBT_LIB_HardIP.TrimCodeDig_Reverse_Ellis(p1, CStr(v(1)), CBool(v(2)), p2, p3, CLng(v(5)), CStr(v(6)), CBool(v(7)), CBool(v(8)), CBool(v(9)), CDbl(v(10)), CDbl(v(11)), CStr(v(12)), CStr(v(13)), CStr(v(14)), CStr(v(15)), CStr(v(16)), CStr(v(17)), p4, CLng(v(19)), CBool(v(UBound(v))), CStr(v(21)))
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function TrimCodeBasicDig__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    Dim p1 As New Pattern
    p1.value = v(0)
    Dim p2 As New PinList
    p2.value = v(1)
    Dim p3 As New PinList
    p3.value = v(11)
    TrimCodeBasicDig__ = VBAProject.VBT_LIB_HardIP.TrimCodeBasicDig(p1, p2, CStr(v(2)), CDbl(v(3)), CStr(v(4)), CStr(v(5)), CStr(v(6)), CBool(v(7)), CBool(v(8)), CBool(v(9)), CStr(v(10)), p3, CLng(v(12)), CBool(v(UBound(v))), CStr(v(14)), CStr(v(15)), CStr(v(16)), CStr(v(17)), CStr(v(18)))
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function Enable_HIP_Datalog_Format__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    Enable_HIP_Datalog_Format__ = VBAProject.VBT_LIB_HardIP.Enable_HIP_Datalog_Format()
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function LDO_Calibration__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    Dim p1 As New Pattern
    p1.value = v(0)
    Dim p2 As New PinList
    p2.value = v(4)
    Dim p3 As Enum_RAK
    p3 = v(22)
    Dim p4 As New PinList
    p4.value = v(23)
    LDO_Calibration__ = VBAProject.VBT_LIB_HardIP.LDO_Calibration(p1, CStr(v(1)), CStr(v(2)), CStr(v(3)), p2, CStr(v(5)), CStr(v(6)), CStr(v(7)), CStr(v(8)), CDbl(v(9)), CLng(v(10)), CLng(v(11)), CStr(v(12)), CStr(v(13)), CDbl(v(14)), CStr(v(15)), CStr(v(16)), CStr(v(17)), CStr(v(18)), CStr(v(19)), CStr(v(20)), CStr(v(21)), p3, p4, CLng(v(24)), CLng(v(25)), CStr(v(26)), CStr(v(27)), CBool(v(28)), CBool(v(29)), CStr(v(30)), CBool(v(UBound(v))))
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function HIP_TTR_Enable_Control__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    HIP_TTR_Enable_Control__ = VBAProject.VBT_LIB_HardIP.HIP_TTR_Enable_Control()
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function Disable_HIP_Datalog_Format__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    Disable_HIP_Datalog_Format__ = VBAProject.VBT_LIB_HardIP.Disable_HIP_Datalog_Format()
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function VTHSENSOR_Calibration__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    Dim p1 As New Pattern
    p1.value = v(0)
    Dim p2 As New PinList
    p2.value = v(13)
    VTHSENSOR_Calibration__ = VBAProject.VBT_LIB_HardIP.VTHSENSOR_Calibration(p1, CStr(v(1)), CBool(v(2)), CStr(v(3)), CStr(v(4)), CStr(v(5)), CStr(v(6)), CStr(v(7)), CStr(v(8)), CStr(v(9)), CStr(v(10)), CBool(v(11)), CStr(v(12)), p2, CStr(v(14)), CStr(v(15)), CStr(v(16)), CStr(v(17)), CStr(v(18)), CStr(v(19)), CStr(v(20)), CStr(v(21)), CStr(v(22)), CStr(v(23)), CStr(v(24)), CBool(v(25)), CBool(v(UBound(v))))
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function LDO_Calibration_AMPLP5X_UCSDM__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    Dim p1 As New Pattern
    p1.value = v(0)
    Dim p2 As New PinList
    p2.value = v(4)
    Dim p3 As Enum_RAK
    p3 = v(22)
    Dim p4 As New PinList
    p4.value = v(23)
    LDO_Calibration_AMPLP5X_UCSDM__ = VBAProject.VBT_LIB_HardIP.LDO_Calibration_AMPLP5X_UCSDM(p1, CStr(v(1)), CStr(v(2)), CStr(v(3)), p2, CStr(v(5)), CStr(v(6)), CStr(v(7)), CStr(v(8)), CDbl(v(9)), CLng(v(10)), CLng(v(11)), CStr(v(12)), CStr(v(13)), CDbl(v(14)), CStr(v(15)), CStr(v(16)), CStr(v(17)), CStr(v(18)), CStr(v(19)), CStr(v(20)), CStr(v(21)), p3, p4, CLng(v(24)), CLng(v(25)), CStr(v(26)), CStr(v(27)), CBool(v(28)), CBool(v(29)), CStr(v(30)), CBool(v(UBound(v))))
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function LDO_Calibration_CaptureAndFreq__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    Dim p1 As New Pattern
    p1.value = v(0)
    Dim p2 As New PinList
    p2.value = v(4)
    Dim p3 As Enum_RAK
    p3 = v(22)
    Dim p4 As New PinList
    p4.value = v(23)
    LDO_Calibration_CaptureAndFreq__ = VBAProject.VBT_LIB_HardIP.LDO_Calibration_CaptureAndFreq(p1, CStr(v(1)), CStr(v(2)), CStr(v(3)), p2, CStr(v(5)), CStr(v(6)), CStr(v(7)), CStr(v(8)), CDbl(v(9)), CLng(v(10)), CLng(v(11)), CStr(v(12)), CStr(v(13)), CDbl(v(14)), CStr(v(15)), CStr(v(16)), CStr(v(17)), CStr(v(18)), CStr(v(19)), CStr(v(20)), CStr(v(21)), p3, p4, CLng(v(24)), CLng(v(25)), CStr(v(26)), CStr(v(27)), CBool(v(28)), CBool(v(29)), CStr(v(30)), CStr(v(31)), CStr(v(32)), CBool(v(UBound(v))))
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function









































Public Function ADC_BandGap_Calibration__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    Dim p1 As New Pattern
    p1.value = v(0)
    Dim p2 As New PinList
    p2.value = v(1)
    Dim p3 As New PinList
    p3.value = v(4)
    ADC_BandGap_Calibration__ = VBAProject.VBT_LIB_HardIP_ADC_AP.ADC_BandGap_Calibration(p1, p2, CLng(v(2)), CLng(v(3)), p3, CStr(v(5)), CStr(v(6)), CStr(v(7)), CStr(v(8)), CStr(v(9)), CStr(v(10)), CDbl(v(11)), CLng(v(12)), CLng(v(13)), CStr(v(14)), CStr(v(15)), CStr(v(16)), CBool(v(17)), CBool(v(UBound(v))))
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function ADC_BandGap_Calibration_Reverse__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    Dim p1 As New Pattern
    p1.value = v(0)
    Dim p2 As New PinList
    p2.value = v(1)
    Dim p3 As New PinList
    p3.value = v(4)
    ADC_BandGap_Calibration_Reverse__ = VBAProject.VBT_LIB_HardIP_ADC_AP.ADC_BandGap_Calibration_Reverse(p1, p2, CLng(v(2)), CLng(v(3)), p3, CStr(v(5)), CStr(v(6)), CStr(v(7)), CStr(v(8)), CStr(v(9)), CStr(v(10)), CDbl(v(11)), CLng(v(12)), CLng(v(13)), CStr(v(14)), CStr(v(15)), CStr(v(16)), CBool(v(UBound(v))))
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function ADC_LDO_Calibration__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    Dim p1 As New Pattern
    p1.value = v(0)
    Dim p2 As New PinList
    p2.value = v(4)
    ADC_LDO_Calibration__ = VBAProject.VBT_LIB_HardIP_ADC_AP.ADC_LDO_Calibration(p1, CStr(v(1)), CStr(v(2)), CStr(v(3)), p2, CStr(v(5)), CStr(v(6)), CStr(v(7)), CStr(v(8)), CStr(v(9)), CDbl(v(10)), CLng(v(11)), CLng(v(12)), CStr(v(13)), CDbl(v(14)), CBool(v(UBound(v))))
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function MetrologyICTS_DAP_PN_AutoTRIM__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    Dim p1 As New Pattern
    p1.value = v(0)
    Dim p2 As New PinList
    p2.value = v(1)
    Dim p3 As New PinList
    p3.value = v(4)
    MetrologyICTS_DAP_PN_AutoTRIM__ = VBAProject.VBT_LIB_HardIP_ADC_AP.MetrologyICTS_DAP_PN_AutoTRIM(p1, p2, CLng(v(2)), CLng(v(3)), p3, CStr(v(5)), CStr(v(6)), CStr(v(7)), CStr(v(8)), CStr(v(9)), CStr(v(10)), CDbl(v(11)), CStr(v(12)), CLng(v(13)), CStr(v(14)), CStr(v(15)), CStr(v(16)), CBool(v(17)), CBool(v(UBound(v))), CStr(v(19)))
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function









































Public Function HIP_eFuse_Write__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    HIP_eFuse_Write__ = VBAProject.VBT_LIB_HardIP_AP.HIP_eFuse_Write(CStr(v(0)), CStr(v(1)), CStr(v(2)), CStr(v(3)), CBool(v(4)), CStr(v(5)))
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function TrimUVI80Code_VFI_ADC__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    Dim p1 As New PinList
    p1.value = v(5)
    Dim p2 As EventSourceWithTerminationMode
    p2 = v(7)
    Dim p3 As New PinList
    p3.value = v(11)
    Dim p4 As New PinList
    p4.value = v(16)
    TrimUVI80Code_VFI_ADC__ = VBAProject.VBT_LIB_HardIP_AP.TrimUVI80Code_VFI_ADC(CStr(v(0)), CStr(v(1)), CStr(v(2)), CStr(v(3)), CDbl(v(4)), p1, CStr(v(6)), p2, CDbl(v(8)), CLng(v(9)), CStr(v(10)), p3, CLng(v(12)), CLng(v(13)), CStr(v(14)), CStr(v(15)), p4, CLng(v(17)), CLng(v(18)), CStr(v(19)), CStr(v(20)), CDbl(v(21)), CLng(v(22)), CStr(v(23)), CStr(v(24)), CStr(v(25)), CBool(v(26)), CBool(v(UBound(v))), CStr(v(28)), CBool(v(29)), CStr(v(30)), CBool(v(31)), CDbl(v(32)), CBool(v(33)))
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function TrimCodeFreq__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    Dim p1 As New Pattern
    p1.value = v(0)
    Dim p2 As New PinList
    p2.value = v(4)
    Dim p3 As New PinList
    p3.value = v(20)
    TrimCodeFreq__ = VBAProject.VBT_LIB_HardIP_AP.TrimCodeFreq(p1, CStr(v(1)), CBool(v(2)), CStr(v(3)), p2, CStr(v(5)), CBool(v(6)), CBool(v(7)), CBool(v(8)), CDbl(v(9)), CDbl(v(10)), CStr(v(11)), CStr(v(12)), CStr(v(13)), CStr(v(14)), CStr(v(15)), CStr(v(16)), CStr(v(17)), CStr(v(18)), CStr(v(19)), p3, CLng(v(21)), CLng(v(22)), CStr(v(23)), CStr(v(24)), CStr(v(25)), CBool(v(26)), CBool(v(27)), CBool(v(UBound(v))))
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function HIP_eFuse_Read__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    HIP_eFuse_Read__ = VBAProject.VBT_LIB_HardIP_AP.HIP_eFuse_Read(CStr(v(0)), CStr(v(1)), CStr(v(2)), CLng(v(3)), CBool(v(4)), CStr(v(5)), CStr(v(6)), CStr(v(7)))
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function TrimCodeFreq_New_ALG__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    Dim p1 As New Pattern
    p1.value = v(0)
    Dim p2 As New PinList
    p2.value = v(3)
    Dim p3 As New PinList
    p3.value = v(4)
    TrimCodeFreq_New_ALG__ = VBAProject.VBT_LIB_HardIP_AP.TrimCodeFreq_New_ALG(p1, CStr(v(1)), CBool(v(2)), p2, p3, CStr(v(5)), CBool(v(6)), CBool(v(7)), CBool(v(8)), CDbl(v(9)), CDbl(v(10)), CStr(v(11)), CStr(v(12)), CStr(v(13)), CStr(v(14)), CStr(v(15)), CStr(v(16)), CInt(v(17)), CStr(v(18)), CBool(v(UBound(v))))
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function TMPS__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    Dim p1 As New Pattern
    p1.value = v(0)
    Dim p2 As New PinList
    p2.value = v(2)
    Dim p3 As New PinList
    p3.value = v(7)
    Dim p4 As CalculateMethodSetup
    p4 = v(10)
    TMPS__ = VBAProject.VBT_LIB_HardIP_AP.TMPS(p1, CBool(v(1)), p2, CLng(v(3)), CLng(v(4)), CStr(v(5)), CStr(v(6)), p3, CLng(v(8)), CLng(v(9)), p4, CStr(v(11)), CStr(v(12)), CStr(v(13)), CBool(v(UBound(v))))
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function









































Public Function CUS_VIR_MainProgram_MeasV_CalR__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    tl_dt_ErrorMsg "Unsupported type at function CUS_VIR_MainProgram_MeasV_CalR"

    ' CUS_VIR_MainProgram_MeasV_CalR__ = VBAProject.VBT_LIB_HardIP_Customize.CUS_VIR_MainProgram_MeasV_CalR(*One or more unsupported types in argument list or non Long/Integer return type*)
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function AnalyzeCusStrToCalcR__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    tl_dt_ErrorMsg "Unsupported type at function AnalyzeCusStrToCalcR"

    ' AnalyzeCusStrToCalcR__ = VBAProject.VBT_LIB_HardIP_Customize.AnalyzeCusStrToCalcR(*One or more unsupported types in argument list or non Long/Integer return type*)
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function Cust_Sweep_V__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    Cust_Sweep_V__ = VBAProject.VBT_LIB_HardIP_Customize.Cust_Sweep_V()
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function VOLH_Sweep__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    VOLH_Sweep__ = VBAProject.VBT_LIB_HardIP_Customize.VOLH_Sweep(CStr(v(0)), CStr(v(1)))
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function MTR_UVI80_Setup__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    MTR_UVI80_Setup__ = VBAProject.VBT_LIB_HardIP_Customize.MTR_UVI80_Setup()
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function CUS_DDR_Emulate_Const_Res_Loading__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    Dim p1 As Enum_RAK
    p1 = v(4)
    tl_dt_ErrorMsg "Unsupported type at function CUS_DDR_Emulate_Const_Res_Loading"

    ' CUS_DDR_Emulate_Const_Res_Loading__ = VBAProject.VBT_LIB_HardIP_Customize.CUS_DDR_Emulate_Const_Res_Loading(*One or more unsupported types in argument list or non Long/Integer return type*)
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function CUS_DDR_DCS_PrintOut__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    CUS_DDR_DCS_PrintOut__ = VBAProject.VBT_LIB_HardIP_Customize.CUS_DDR_DCS_PrintOut()
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function MEAS_I_ABS__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    tl_dt_ErrorMsg "Unsupported type at function MEAS_I_ABS"

    ' MEAS_I_ABS__ = VBAProject.VBT_LIB_HardIP_Customize.MEAS_I_ABS(*One or more unsupported types in argument list or non Long/Integer return type*)
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function CUS_RREF_Rak_Calc__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    tl_dt_ErrorMsg "Unsupported type at function CUS_RREF_Rak_Calc"

    ' CUS_RREF_Rak_Calc__ = VBAProject.VBT_LIB_HardIP_Customize.CUS_RREF_Rak_Calc(*One or more unsupported types in argument list or non Long/Integer return type*)
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function CUS_AMP_SDLL_SWP__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    tl_dt_ErrorMsg "Unsupported type at function CUS_AMP_SDLL_SWP"

    ' CUS_AMP_SDLL_SWP__ = VBAProject.VBT_LIB_HardIP_Customize.CUS_AMP_SDLL_SWP(*One or more unsupported types in argument list or non Long/Integer return type*)
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function D2D_Flag_Sync_Master_Slave__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    D2D_Flag_Sync_Master_Slave__ = VBAProject.VBT_LIB_HardIP_Customize.D2D_Flag_Sync_Master_Slave(CStr(v(0)), CStr(v(1)))
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function Aciphy_Switch__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    Dim p1 As New PinList
    p1.value = v(0)
    Dim p2 As New PinList
    p2.value = v(2)
    Aciphy_Switch__ = VBAProject.VBT_LIB_HardIP_Customize.Aciphy_Switch(p1, CDbl(v(1)), p2, CDbl(v(3)), CBool(v(4)))
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function Create_Pseudo_Dictionary__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    Create_Pseudo_Dictionary__ = VBAProject.VBT_LIB_HardIP_Customize.Create_Pseudo_Dictionary(CStr(v(0)), CStr(v(1)), CStr(v(2)), CStr(v(3)))
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function









































Public Function MetrologyTMPS_Calibration__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    Dim p1 As New Pattern
    p1.value = v(0)
    Dim p2 As New PinList
    p2.value = v(4)
    MetrologyTMPS_Calibration__ = VBAProject.VBT_LIB_HardIP_Metrology_AP.MetrologyTMPS_Calibration(p1, CStr(v(1)), CStr(v(2)), CStr(v(3)), p2, CStr(v(5)), CStr(v(6)), CStr(v(7)), CStr(v(8)), CDbl(v(9)), CLng(v(10)), CLng(v(11)), CLng(v(12)), CStr(v(13)), CStr(v(14)), CBool(v(UBound(v))))
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function MetrologyGR_Calibration__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    Dim p1 As New Pattern
    p1.value = v(0)
    Dim p2 As New PinList
    p2.value = v(4)
    MetrologyGR_Calibration__ = VBAProject.VBT_LIB_HardIP_Metrology_AP.MetrologyGR_Calibration(p1, CStr(v(1)), CStr(v(2)), CStr(v(3)), p2, CStr(v(5)), CStr(v(6)), CStr(v(7)), CStr(v(8)), CDbl(v(9)), CLng(v(10)), CLng(v(11)), CLng(v(12)), CStr(v(13)), CStr(v(14)), CStr(v(15)), CBool(v(UBound(v))))
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function MTR_TDiode_Test__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    MTR_TDiode_Test__ = VBAProject.VBT_LIB_HardIP_Metrology_AP.MTR_TDiode_Test(CStr(v(0)), CStr(v(1)), CStr(v(2)), CStr(v(3)), CStr(v(4)), CStr(v(5)), CStr(v(6)), CStr(v(7)), CStr(v(8)), CStr(v(9)), CStr(v(10)), CStr(v(11)), CBool(v(UBound(v))))
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function MetrologyICTS_DAP_PN_TRIM__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    Dim p1 As New Pattern
    p1.value = v(0)
    Dim p2 As New PinList
    p2.value = v(1)
    Dim p3 As New PinList
    p3.value = v(4)
    MetrologyICTS_DAP_PN_TRIM__ = VBAProject.VBT_LIB_HardIP_Metrology_AP.MetrologyICTS_DAP_PN_TRIM(p1, p2, CLng(v(2)), CLng(v(3)), p3, CStr(v(5)), CStr(v(6)), CStr(v(7)), CStr(v(8)), CStr(v(9)), CStr(v(10)), CDbl(v(11)), CStr(v(12)), CLng(v(13)), CStr(v(14)), CStr(v(15)), CStr(v(16)), CBool(v(17)), CBool(v(UBound(v))))
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function









































Public Function SaveEfuseDecimal_to_Dictionary__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    tl_dt_ErrorMsg "Unsupported type at function SaveEfuseDecimal_to_Dictionary"

    ' SaveEfuseDecimal_to_Dictionary__ = VBAProject.VBT_LIB_K_ADDRIO_AP.SaveEfuseDecimal_to_Dictionary(*One or more unsupported types in argument list or non Long/Integer return type*)
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function TrimImpedance_MeasR__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    tl_dt_ErrorMsg "Unsupported type at function TrimImpedance_MeasR"

    ' TrimImpedance_MeasR__ = VBAProject.VBT_LIB_K_ADDRIO_AP.TrimImpedance_MeasR(*One or more unsupported types in argument list or non Long/Integer return type*)
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function UP1600_PPMU_Measure_R_SingleEnd__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    Dim p1 As Enum_RAK
    p1 = v(3)
    tl_dt_ErrorMsg "Unsupported type at function UP1600_PPMU_Measure_R_SingleEnd"

    ' UP1600_PPMU_Measure_R_SingleEnd__ = VBAProject.VBT_LIB_K_ADDRIO_AP.UP1600_PPMU_Measure_R_SingleEnd(*One or more unsupported types in argument list or non Long/Integer return type*)
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function UP1600_PPMU_Measure_R_Differential__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    Dim p1 As Enum_RAK
    p1 = v(4)
    tl_dt_ErrorMsg "Unsupported type at function UP1600_PPMU_Measure_R_Differential"

    ' UP1600_PPMU_Measure_R_Differential__ = VBAProject.VBT_LIB_K_ADDRIO_AP.UP1600_PPMU_Measure_R_Differential(*One or more unsupported types in argument list or non Long/Integer return type*)
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function MeasurePins_to_Differential_Pair__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    Dim p1 As New PinList
    p1.value = v(0)
    MeasurePins_to_Differential_Pair__ = VBAProject.VBT_LIB_K_ADDRIO_AP.MeasurePins_to_Differential_Pair(p1, CStr(v(1)), CStr(v(2)))
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function K_Set_SpecialSettingSequence__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    Call VBAProject.VBT_LIB_K_ADDRIO_AP.K_Set_SpecialSettingSequence(CStr(v(0)))
    K_Set_SpecialSettingSequence__ = TL_SUCCESS
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function K_TrimImpedance_Turks_TTR__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    Dim p1 As New Pattern
    p1.value = v(0)
    Dim p2 As New PinList
    p2.value = v(2)
    Dim p3 As New PinList
    p3.value = v(3)
    Dim p4 As New PinList
    p4.value = v(6)
    Dim p5 As New PinList
    p5.value = v(11)
    K_TrimImpedance_Turks_TTR__ = VBAProject.VBT_LIB_K_ADDRIO_AP.K_TrimImpedance_Turks_TTR(p1, CBool(v(1)), p2, p3, CDbl(v(4)), CStr(v(5)), p4, CLng(v(7)), CLng(v(8)), CLng(v(9)), CLng(v(10)), p5, CStr(v(12)), CStr(v(13)), CStr(v(14)), CBool(v(15)), CDbl(v(16)), CLng(v(17)), CDbl(v(18)), CDbl(v(19)), CInt(v(20)), CStr(v(21)), CBool(v(UBound(v))))
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function K_TrimImpedance_Turks_V1__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    Dim p1 As New Pattern
    p1.value = v(0)
    Dim p2 As New PinList
    p2.value = v(2)
    Dim p3 As New PinList
    p3.value = v(3)
    Dim p4 As New PinList
    p4.value = v(6)
    Dim p5 As New PinList
    p5.value = v(11)
    K_TrimImpedance_Turks_V1__ = VBAProject.VBT_LIB_K_ADDRIO_AP.K_TrimImpedance_Turks_V1(p1, CBool(v(1)), p2, p3, CDbl(v(4)), CStr(v(5)), p4, CLng(v(7)), CLng(v(8)), CLng(v(9)), CLng(v(10)), p5, CStr(v(12)), CStr(v(13)), CStr(v(14)), CBool(v(15)), CDbl(v(16)), CLng(v(17)), CDbl(v(18)), CDbl(v(19)), CInt(v(20)), CStr(v(21)), CBool(v(UBound(v))))
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function K_TrimImpedance_CodeCheck_V1__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    Dim p1 As New Pattern
    p1.value = v(0)
    Dim p2 As New PinList
    p2.value = v(2)
    Dim p3 As New PinList
    p3.value = v(3)
    Dim p4 As New PinList
    p4.value = v(6)
    Dim p5 As New PinList
    p5.value = v(11)
    K_TrimImpedance_CodeCheck_V1__ = VBAProject.VBT_LIB_K_ADDRIO_AP.K_TrimImpedance_CodeCheck_V1(p1, CBool(v(1)), p2, p3, CDbl(v(4)), CStr(v(5)), p4, CLng(v(7)), CStr(v(8)), CLng(v(9)), CLng(v(10)), p5, CStr(v(12)), CStr(v(13)), CStr(v(14)), CBool(v(15)), CDbl(v(16)), CStr(v(17)), CStr(v(18)), CBool(v(19)), CStr(v(20)), CStr(v(21)), CBool(v(UBound(v))))
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function K_Str_DumpDspWave__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    tl_dt_ErrorMsg "Unsupported type at function K_Str_DumpDspWave"

    ' K_Str_DumpDspWave__ = VBAProject.VBT_LIB_K_ADDRIO_AP.K_Str_DumpDspWave(*One or more unsupported types in argument list or non Long/Integer return type*)
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function CZ_Style_TName_InstanceInfo_Clear__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    CZ_Style_TName_InstanceInfo_Clear__ = VBAProject.VBT_LIB_K_ADDRIO_AP.CZ_Style_TName_InstanceInfo_Clear()
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function









































Public Function Protect_Mbist_Sheet__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    Protect_Mbist_Sheet__ = VBAProject.VBT_LIB_MBIST.Protect_Mbist_Sheet()
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function UnProtect_Mbist_Sheet__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    UnProtect_Mbist_Sheet__ = VBAProject.VBT_LIB_MBIST.UnProtect_Mbist_Sheet()
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function auto_Mbist_SetLoopCNT_BM__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    auto_Mbist_SetLoopCNT_BM__ = VBAProject.VBT_LIB_MBIST.auto_Mbist_SetLoopCNT_BM(CStr(v(0)), CStr(v(1)), CStr(v(2)), CLng(v(3)))
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function Mbist_Initialize__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    Mbist_Initialize__ = VBAProject.VBT_LIB_MBIST.Mbist_Initialize()
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function init_MBIST_ChkList_block_loop__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    init_MBIST_ChkList_block_loop__ = VBAProject.VBT_LIB_MBIST.init_MBIST_ChkList_block_loop(CStr(v(0)))
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function Separate_nu_char__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    Separate_nu_char__ = VBAProject.VBT_LIB_MBIST.Separate_nu_char(CStr(v(0)), CLng(v(1)), CStr(v(2)))
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function









































Public Function RTOS_Capture_Data__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    tl_dt_ErrorMsg "Unsupported type at function RTOS_Capture_Data"

    ' RTOS_Capture_Data__ = VBAProject.VBT_LIB_RTOS.RTOS_Capture_Data(*One or more unsupported types in argument list or non Long/Integer return type*)
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function RTOS_Command__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    RTOS_Command__ = VBAProject.VBT_LIB_RTOS.RTOS_Command(CStr(v(0)), CStr(v(1)), CStr(v(2)), CStr(v(3)), CStr(v(4)), CDbl(v(5)), CDbl(v(6)), CDbl(v(7)), CDbl(v(8)), CDbl(v(9)))
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function RTOS_eFuse_Read__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    RTOS_eFuse_Read__ = VBAProject.VBT_LIB_RTOS.RTOS_eFuse_Read(CStr(v(0)), CStr(v(1)), CStr(v(2)), CLng(v(3)), CBool(v(4)), CStr(v(5)), CStr(v(6)))
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function RTOS_eFuse_Write__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    RTOS_eFuse_Write__ = VBAProject.VBT_LIB_RTOS.RTOS_eFuse_Write(CStr(v(0)), CStr(v(1)), CStr(v(2)), CStr(v(3)), CBool(v(4)), CStr(v(5)))
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function RTOS_IDS__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    tl_dt_ErrorMsg "Unsupported type at function RTOS_IDS"

    ' RTOS_IDS__ = VBAProject.VBT_LIB_RTOS.RTOS_IDS(*One or more unsupported types in argument list or non Long/Integer return type*)
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function RTOS_Boot__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    Dim p1 As New Pattern
    p1.value = v(1)
    Dim p2 As New PinList
    p2.value = v(19)
    Dim p3 As New PinList
    p3.value = v(20)
    RTOS_Boot__ = VBAProject.VBT_LIB_RTOS.RTOS_Boot(CBool(v(0)), p1, CBool(v(2)), CBool(v(3)), CBool(v(4)), CStr(v(5)), CStr(v(6)), CStr(v(7)), CStr(v(8)), CStr(v(9)), CBool(v(10)), CBool(v(11)), CStr(v(12)), CStr(v(13)), CBool(v(14)), CStr(v(15)), CBool(v(16)), CStr(v(17)), CBool(v(18)), p2, p3, CStr(v(21)), CDbl(v(22)), CDbl(v(23)), CDbl(v(24)), CStr(v(25)))
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function RTOS_Boot_CZ__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    tl_dt_ErrorMsg "Unsupported type at function RTOS_Boot_CZ"

    ' RTOS_Boot_CZ__ = VBAProject.VBT_LIB_RTOS.RTOS_Boot_CZ(*One or more unsupported types in argument list or non Long/Integer return type*)
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function RTOS_IOState_Checker__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    RTOS_IOState_Checker__ = VBAProject.VBT_LIB_RTOS.RTOS_IOState_Checker(CStr(v(0)))
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function RTOS_Meas_FRC__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    RTOS_Meas_FRC__ = VBAProject.VBT_LIB_RTOS.RTOS_Meas_FRC(CStr(v(0)))
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function RTOS_RunMetrology__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    RTOS_RunMetrology__ = VBAProject.VBT_LIB_RTOS.RTOS_RunMetrology(CStr(v(0)), CStr(v(1)), CStr(v(2)), CStr(v(3)), CStr(v(4)), CStr(v(5)), CStr(v(6)), CStr(v(7)), CDbl(v(8)), CDbl(v(9)), CDbl(v(10)), CDbl(v(11)), CStr(v(12)))
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function RTOS_Shmoo_Reboot__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    tl_dt_ErrorMsg "Unsupported type at function RTOS_Shmoo_Reboot"

    ' RTOS_Shmoo_Reboot__ = VBAProject.VBT_LIB_RTOS.RTOS_Shmoo_Reboot(*One or more unsupported types in argument list or non Long/Integer return type*)
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function RTOS_RunScenario__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    RTOS_RunScenario__ = VBAProject.VBT_LIB_RTOS.RTOS_RunScenario(CStr(v(0)), CStr(v(1)), CStr(v(2)), CStr(v(3)), CStr(v(4)), CStr(v(5)), CDbl(v(6)), CDbl(v(7)), CDbl(v(8)), CDbl(v(9)), CDbl(v(10)), CStr(v(11)), CStr(v(12)), CStr(v(13)), CInt(v(14)), CDbl(v(15)), CBool(v(16)))
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function RTOS_UART_Print__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    tl_dt_ErrorMsg "Unsupported type at function RTOS_UART_Print"

    ' RTOS_UART_Print__ = VBAProject.VBT_LIB_RTOS.RTOS_UART_Print(*One or more unsupported types in argument list or non Long/Integer return type*)
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function SendCmd__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    tl_dt_ErrorMsg "Unsupported type at function SendCmd"

    ' SendCmd__ = VBAProject.VBT_LIB_RTOS.SendCmd(*One or more unsupported types in argument list or non Long/Integer return type*)
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function LogDUTResponse__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    tl_dt_ErrorMsg "Unsupported type at function LogDUTResponse"

    ' LogDUTResponse__ = VBAProject.VBT_LIB_RTOS.LogDUTResponse(*One or more unsupported types in argument list or non Long/Integer return type*)
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function SendCmdOnly__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    Call VBAProject.VBT_LIB_RTOS.SendCmdOnly(CStr(v(0)))
    SendCmdOnly__ = TL_SUCCESS
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function WriteToOutputWindow__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    tl_dt_ErrorMsg "Unsupported type at function WriteToOutputWindow"

    ' Call VBAProject.VBT_LIB_RTOS.WriteToOutputWindow(*One or more unsupported types in argument list or non Long/Integer return type*)
    WriteToOutputWindow__ = TL_SUCCESS
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function WriteToDatalog__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    tl_dt_ErrorMsg "Unsupported type at function WriteToDatalog"

    ' Call VBAProject.VBT_LIB_RTOS.WriteToDatalog(*One or more unsupported types in argument list or non Long/Integer return type*)
    WriteToDatalog__ = TL_SUCCESS
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function ReloadUARTModules__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    ReloadUARTModules__ = VBAProject.VBT_LIB_RTOS.ReloadUARTModules()
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function RTOS_Voltage_Rampdown__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    RTOS_Voltage_Rampdown__ = VBAProject.VBT_LIB_RTOS.RTOS_Voltage_Rampdown()
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function RTOS_Voltage_RampUp__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    RTOS_Voltage_RampUp__ = VBAProject.VBT_LIB_RTOS.RTOS_Voltage_RampUp()
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function Decide_Switching_Bit_RTOS__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    tl_dt_ErrorMsg "Unsupported type at function Decide_Switching_Bit_RTOS"

    ' Decide_Switching_Bit_RTOS__ = VBAProject.VBT_LIB_RTOS.Decide_Switching_Bit_RTOS(*One or more unsupported types in argument list or non Long/Integer return type*)
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function RTOS_Freq_Measurement__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    RTOS_Freq_Measurement__ = VBAProject.VBT_LIB_RTOS.RTOS_Freq_Measurement(CStr(v(0)), CDbl(v(1)), CDbl(v(2)))
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function RTOS_DigCap_Setting__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    RTOS_DigCap_Setting__ = VBAProject.VBT_LIB_RTOS.RTOS_DigCap_Setting(CStr(v(0)), CStr(v(1)), CStr(v(2)), CLng(v(3)))
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function RTOS_RunScenario_T__(v As Variant) As Long
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
    RTOS_RunScenario_T__ = VBAProject.VBT_LIB_RTOS.RTOS_RunScenario_T(p1, p2, p3, p4, p5, p6, CStr(v(6)), CStr(v(7)), CStr(v(8)), CStr(v(9)), CStr(v(10)), CStr(v(11)), p7, p8, p9, p10, p11, p12, p13, pStep, CBool(v(19)), CBool(v(20)), CStr(v(21)), CStr(v(22)), CStr(v(23)), CStr(v(24)), CStr(v(25)), CStr(v(26)), CStr(v(27)), CDbl(v(28)), CDbl(v(29)), CDbl(v(30)), CDbl(v(31)), CDbl(v(32)), CStr(v(33)), CStr(v(34)), CStr(v(35)), CStr(v(36)), CInt(v(37)), CDbl(v(38)), CBool(v(39)), CStr(v(40)))
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function RTOS_Command_UID_EFUSE__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    RTOS_Command_UID_EFUSE__ = VBAProject.VBT_LIB_RTOS.RTOS_Command_UID_EFUSE(CStr(v(0)), CDbl(v(1)), CStr(v(2)), CDbl(v(3)))
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function









































Public Function Print_Shmoo_ND_Summary__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    Print_Shmoo_ND_Summary__ = VBAProject.VBT_LIB_Shmoo_ND.Print_Shmoo_ND_Summary(CStr(v(0)))
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function Setup_FlowLoop__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    Setup_FlowLoop__ = VBAProject.VBT_LIB_Shmoo_ND.Setup_FlowLoop(CStr(v(0)), CStr(v(1)))
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function Read_Shmoo_ND_Table__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    Read_Shmoo_ND_Table__ = VBAProject.VBT_LIB_Shmoo_ND.Read_Shmoo_ND_Table()
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function Calc_Pass_List__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    Calc_Pass_List__ = VBAProject.VBT_LIB_Shmoo_ND.Calc_Pass_List()
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function Skip_All_Fail_Start__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    Skip_All_Fail_Start__ = VBAProject.VBT_LIB_Shmoo_ND.Skip_All_Fail_Start()
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function Skip_All_Fail_End__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    Skip_All_Fail_End__ = VBAProject.VBT_LIB_Shmoo_ND.Skip_All_Fail_End()
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function









































Public Function Functional_T_SRM__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    Dim pStep As SubType
    pStep = TheExec.flow.StepType
    Dim p1 As New Pattern
    p1.value = v(1)
    Dim p2 As New Pattern
    p2.value = v(2)
    Dim p3 As New Pattern
    p3.value = v(3)
    Dim p4 As New Pattern
    p4.value = v(4)
    Dim p5 As New Pattern
    p5.value = v(5)
    Dim p6 As New Pattern
    p6.value = v(6)
    Dim p7 As New Pattern
    p7.value = v(7)
    Dim p8 As New Pattern
    p8.value = v(8)
    Dim p9 As New Pattern
    p9.value = v(9)
    Dim p10 As New Pattern
    p10.value = v(10)
    Dim p11 As New Pattern
    p11.value = v(11)
    Dim p12 As New Pattern
    p12.value = v(12)
    Dim p13 As New Pattern
    p13.value = v(13)
    Dim p14 As New Pattern
    p14.value = v(14)
    Dim p15 As New Pattern
    p15.value = v(15)
    Functional_T_SRM__ = VBAProject.VBT_LIB_SRM.Functional_T_SRM(CStr(v(0)), pStep, p1, p2, p3, p4, p5, p6, p7, p8, p9, p10, p11, p12, p13, p14, p15, CStr(v(16)), CStr(v(17)), CStr(v(18)), CStr(v(19)), CStr(v(20)), CStr(v(21)), CStr(v(22)), CBool(v(UBound(v))), CBool(v(24)))
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function SRM_CoustomSetting__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    SRM_CoustomSetting__ = VBAProject.VBT_LIB_SRM.SRM_CoustomSetting()
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function SRM_ParseChkList__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    SRM_ParseChkList__ = VBAProject.VBT_LIB_SRM.SRM_ParseChkList(CStr(v(0)))
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function SRM_ArgsToStr__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    tl_dt_ErrorMsg "Unsupported type at function SRM_ArgsToStr"

    ' SRM_ArgsToStr__ = VBAProject.VBT_LIB_SRM.SRM_ArgsToStr(*One or more unsupported types in argument list or non Long/Integer return type*)
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function SRM_CreatePatSetAllDic__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    SRM_CreatePatSetAllDic__ = VBAProject.VBT_LIB_SRM.SRM_CreatePatSetAllDic()
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function SRM_PatPreCheck__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    tl_dt_ErrorMsg "Unsupported type at function SRM_PatPreCheck"

    ' SRM_PatPreCheck__ = VBAProject.VBT_LIB_SRM.SRM_PatPreCheck(*One or more unsupported types in argument list or non Long/Integer return type*)
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function SRM_InitDatalogSetup__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    SRM_InitDatalogSetup__ = VBAProject.VBT_LIB_SRM.SRM_InitDatalogSetup()
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function SRM_InitArrComma__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    tl_dt_ErrorMsg "Unsupported type at function SRM_InitArrComma"

    ' SRM_InitArrComma__ = VBAProject.VBT_LIB_SRM.SRM_InitArrComma(*One or more unsupported types in argument list or non Long/Integer return type*)
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function SRM_ReDimArrSize__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    tl_dt_ErrorMsg "Unsupported type at function SRM_ReDimArrSize"

    ' SRM_ReDimArrSize__ = VBAProject.VBT_LIB_SRM.SRM_ReDimArrSize(*One or more unsupported types in argument list or non Long/Integer return type*)
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function SRM_InitPayloadArr__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    tl_dt_ErrorMsg "Unsupported type at function SRM_InitPayloadArr"

    ' SRM_InitPayloadArr__ = VBAProject.VBT_LIB_SRM.SRM_InitPayloadArr(*One or more unsupported types in argument list or non Long/Integer return type*)
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function SRM_FindSrmPatIndex__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    tl_dt_ErrorMsg "Unsupported type at function SRM_FindSrmPatIndex"

    ' SRM_FindSrmPatIndex__ = VBAProject.VBT_LIB_SRM.SRM_FindSrmPatIndex(*One or more unsupported types in argument list or non Long/Integer return type*)
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function SRM_DecideType__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    SRM_DecideType__ = VBAProject.VBT_LIB_SRM.SRM_DecideType(CStr(v(0)))
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function SRM_GetInfo__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    SRM_GetInfo__ = VBAProject.VBT_LIB_SRM.SRM_GetInfo(CStr(v(0)))
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function SRM_InitLoopCnt__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    SRM_InitLoopCnt__ = VBAProject.VBT_LIB_SRM.SRM_InitLoopCnt(CStr(v(0)))
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function SRM_AutoSetLoopCnt__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    SRM_AutoSetLoopCnt__ = VBAProject.VBT_LIB_SRM.SRM_AutoSetLoopCnt(CLng(v(0)), CLng(v(1)))
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function SRM_GetChnType__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    SRM_GetChnType__ = VBAProject.VBT_LIB_SRM.SRM_GetChnType()
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function SRM_FindRangeWaitTime__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    tl_dt_ErrorMsg "Unsupported type at function SRM_FindRangeWaitTime"

    ' SRM_FindRangeWaitTime__ = VBAProject.VBT_LIB_SRM.SRM_FindRangeWaitTime(*One or more unsupported types in argument list or non Long/Integer return type*)
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function SRM_FindRangeAccuracy__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    tl_dt_ErrorMsg "Unsupported type at function SRM_FindRangeAccuracy"

    ' SRM_FindRangeAccuracy__ = VBAProject.VBT_LIB_SRM.SRM_FindRangeAccuracy(*One or more unsupported types in argument list or non Long/Integer return type*)
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function SRM_FindChnMergeCase__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    tl_dt_ErrorMsg "Unsupported type at function SRM_FindChnMergeCase"

    ' SRM_FindChnMergeCase__ = VBAProject.VBT_LIB_SRM.SRM_FindChnMergeCase(*One or more unsupported types in argument list or non Long/Integer return type*)
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function









































Public Function initVddBinning__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    initVddBinning__ = VBAProject.VBT_LIB_VDD_Binning.initVddBinning()
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function PrintOut_VDD_BIN__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    PrintOut_VDD_BIN__ = VBAProject.VBT_LIB_VDD_Binning.PrintOut_VDD_BIN(CBool(v(0)))
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function Read_DVFM_To_GradeVDD__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    Read_DVFM_To_GradeVDD__ = VBAProject.VBT_LIB_VDD_Binning.Read_DVFM_To_GradeVDD()
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function ReGenerate_IDS_ZONE_Voltage_Per_Site_ver2__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    ReGenerate_IDS_ZONE_Voltage_Per_Site_ver2__ = VBAProject.VBT_LIB_VDD_Binning.ReGenerate_IDS_ZONE_Voltage_Per_Site_ver2(CStr(v(0)))
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function Print_DYNAMIC_VBIN_IDS_ZONE_to_sheet__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    Print_DYNAMIC_VBIN_IDS_ZONE_to_sheet__ = VBAProject.VBT_LIB_VDD_Binning.Print_DYNAMIC_VBIN_IDS_ZONE_to_sheet(CStr(v(0)))
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function UpdateDLogColumns_Bincut__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    UpdateDLogColumns_Bincut__ = VBAProject.VBT_LIB_VDD_Binning.UpdateDLogColumns_Bincut(CLng(v(0)))
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function check_IDS__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    check_IDS__ = VBAProject.VBT_LIB_VDD_Binning.check_IDS(CBool(v(0)))
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function adjust_VddBinning__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    adjust_VddBinning__ = VBAProject.VBT_LIB_VDD_Binning.adjust_VddBinning(CBool(v(0)), CBool(v(1)), CStr(v(2)), CStr(v(3)), CBool(v(4)))
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function GradeSearch_VT__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    Dim p1 As New Pattern
    p1.value = v(0)
    Dim p2 As tlResultMode
    p2 = v(2)
    Dim p3 As New Pattern
    p3.value = v(5)
    GradeSearch_VT__ = VBAProject.VBT_LIB_VDD_Binning.GradeSearch_VT(p1, CStr(v(1)), p2, CStr(v(3)), CBool(v(4)), p3, CInt(v(6)), CBool(v(UBound(v))), CStr(v(8)), CLng(v(9)), CStr(v(10)), CBool(v(11)), CStr(v(12)), CBool(v(13)), CBool(v(14)), CLng(v(15)), CBool(v(16)), CLng(v(17)), CStr(v(18)))
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function Power_Binning_Calculation__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    Power_Binning_Calculation__ = VBAProject.VBT_LIB_VDD_Binning.Power_Binning_Calculation(CBool(v(0)))
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function GradeSearch_HVCC_VT__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    Dim p1 As New Pattern
    p1.value = v(0)
    Dim p2 As tlResultMode
    p2 = v(2)
    Dim p3 As New Pattern
    p3.value = v(6)
    GradeSearch_HVCC_VT__ = VBAProject.VBT_LIB_VDD_Binning.GradeSearch_HVCC_VT(p1, CStr(v(1)), p2, CStr(v(3)), CBool(v(4)), CStr(v(5)), p3, CInt(v(7)), CBool(v(8)), CBool(v(UBound(v))), CStr(v(10)), CStr(v(11)), CBool(v(12)), CStr(v(13)), CStr(v(14)), CLng(v(15)), CStr(v(16)))
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function GradeSearch_postBinCut_VT__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    Dim p1 As New Pattern
    p1.value = v(0)
    Dim p2 As tlResultMode
    p2 = v(2)
    Dim p3 As New Pattern
    p3.value = v(6)
    GradeSearch_postBinCut_VT__ = VBAProject.VBT_LIB_VDD_Binning.GradeSearch_postBinCut_VT(p1, CStr(v(1)), p2, CStr(v(3)), CBool(v(4)), CStr(v(5)), p3, CInt(v(7)), CBool(v(8)), CBool(v(UBound(v))), CBool(v(10)), CStr(v(11)), CBool(v(12)), CDbl(v(13)), CStr(v(14)), CLng(v(15)), CStr(v(16)), CBool(v(17)))
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function BV_Init_Datalog_Setup__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    BV_Init_Datalog_Setup__ = VBAProject.VBT_LIB_VDD_Binning.BV_Init_Datalog_Setup(CLng(v(0)))
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function Restore_BV_DataLog_SetUp__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    Restore_BV_DataLog_SetUp__ = VBAProject.VBT_LIB_VDD_Binning.Restore_BV_DataLog_SetUp()
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function Print_BinCut_config__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    Print_BinCut_config__ = VBAProject.VBT_LIB_VDD_Binning.Print_BinCut_config(CStr(v(0)))
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function Set_VBinResult_without_Test__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    Set_VBinResult_without_Test__ = VBAProject.VBT_LIB_VDD_Binning.Set_VBinResult_without_Test(CBool(v(0)))
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function GradeSearch_CallInstance_VT__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    Dim p1 As tlResultMode
    p1 = v(1)
    GradeSearch_CallInstance_VT__ = VBAProject.VBT_LIB_VDD_Binning.GradeSearch_CallInstance_VT(CStr(v(0)), p1, CStr(v(2)), CBool(v(3)), CStr(v(4)), CBool(v(UBound(v))), CStr(v(6)), CLng(v(7)), CStr(v(8)), CBool(v(9)), CStr(v(10)))
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function GradeSearch_HVCC_CallInstance_VT__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    Dim p1 As tlResultMode
    p1 = v(1)
    GradeSearch_HVCC_CallInstance_VT__ = VBAProject.VBT_LIB_VDD_Binning.GradeSearch_HVCC_CallInstance_VT(CStr(v(0)), p1, CStr(v(2)), CBool(v(3)), CStr(v(4)), CBool(v(UBound(v))), CStr(v(6)))
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function Overwrite_PassBinNum_by_ForcedBin__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    Overwrite_PassBinNum_by_ForcedBin__ = VBAProject.VBT_LIB_VDD_Binning.Overwrite_PassBinNum_by_ForcedBin(CBool(v(0)), CLng(v(1)))
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function save_siteMask_for_MultiFSTP__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    save_siteMask_for_MultiFSTP__ = VBAProject.VBT_LIB_VDD_Binning.save_siteMask_for_MultiFSTP(CStr(v(0)))
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function restore_siteMask_for_MultiFSTP__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    restore_siteMask_for_MultiFSTP__ = VBAProject.VBT_LIB_VDD_Binning.restore_siteMask_for_MultiFSTP(CStr(v(0)), CStr(v(1)))
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function Check_flagstate_for_failflag__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    Check_flagstate_for_failflag__ = VBAProject.VBT_LIB_VDD_Binning.Check_flagstate_for_failflag(CStr(v(0)))
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function align_startStep_to_GradeVDD__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    align_startStep_to_GradeVDD__ = VBAProject.VBT_LIB_VDD_Binning.align_startStep_to_GradeVDD()
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function Pre_adjust_VddBinning__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    Pre_adjust_VddBinning__ = VBAProject.VBT_LIB_VDD_Binning.Pre_adjust_VddBinning(CBool(v(0)), CBool(v(1)), CStr(v(2)), CStr(v(3)))
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function Set_E1_Voltage_ForPmode__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    Set_E1_Voltage_ForPmode__ = VBAProject.VBT_LIB_VDD_Binning.Set_E1_Voltage_ForPmode(CStr(v(0)), CStr(v(1)), CStr(v(2)))
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function SetHarvestFlagForBV__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    SetHarvestFlagForBV__ = VBAProject.VBT_LIB_VDD_Binning.SetHarvestFlagForBV()
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function SetHarvestFlagForBV_Coll__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    SetHarvestFlagForBV_Coll__ = VBAProject.VBT_LIB_VDD_Binning.SetHarvestFlagForBV_Coll()
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function SetHarvestFlagForBV_Tahiti__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    SetHarvestFlagForBV_Tahiti__ = VBAProject.VBT_LIB_VDD_Binning.SetHarvestFlagForBV_Tahiti()
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function SetHarvestFlagForBV_Thera__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    SetHarvestFlagForBV_Thera__ = VBAProject.VBT_LIB_VDD_Binning.SetHarvestFlagForBV_Thera()
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function JudgeBINXForHarvesting__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    JudgeBINXForHarvesting__ = VBAProject.VBT_LIB_VDD_Binning.JudgeBINXForHarvesting()
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function JudgeBINXForWLFT__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    JudgeBINXForWLFT__ = VBAProject.VBT_LIB_VDD_Binning.JudgeBINXForWLFT()
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function









































Public Function SetFlag_PassFail_for_MbistBlocks__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    tl_dt_ErrorMsg "Unsupported type at function SetFlag_PassFail_for_MbistBlocks"

    ' SetFlag_PassFail_for_MbistBlocks__ = VBAProject.VBT_MBIST_Binout.SetFlag_PassFail_for_MbistBlocks(*One or more unsupported types in argument list or non Long/Integer return type*)
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function CPU_MBist_BinOut__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    CPU_MBist_BinOut__ = VBAProject.VBT_MBIST_Binout.CPU_MBist_BinOut()
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function SOC_MBist_BinOut__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    SOC_MBist_BinOut__ = VBAProject.VBT_MBIST_Binout.SOC_MBist_BinOut()
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function MBist_BinOutFlag__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    MBist_BinOutFlag__ = VBAProject.VBT_MBIST_Binout.MBist_BinOutFlag(CLng(v(0)), CLng(v(1)), CLng(v(2)))
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function BinOutFlag_BlockPassRate__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    BinOutFlag_BlockPassRate__ = VBAProject.VBT_MBIST_Binout.BinOutFlag_BlockPassRate(CLng(v(0)), CLng(v(1)), CLng(v(2)), CStr(v(3)))
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function WriteRepairInfo_in_STDF__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    WriteRepairInfo_in_STDF__ = VBAProject.VBT_MBIST_Binout.WriteRepairInfo_in_STDF(CBool(v(0)))
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function









































Public Function Parsing_Limit_Table__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    Parsing_Limit_Table__ = VBAProject.VBT_PC_Diagnosis.Parsing_Limit_Table()
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function Trace_Resistance_Meas__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    Trace_Resistance_Meas__ = VBAProject.VBT_PC_Diagnosis.Trace_Resistance_Meas(CStr(v(0)), CStr(v(1)))
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function Digital_Open_Voltage_Meas__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    Digital_Open_Voltage_Meas__ = VBAProject.VBT_PC_Diagnosis.Digital_Open_Voltage_Meas(CStr(v(0)), CStr(v(1)), CDbl(v(2)), CDbl(v(3)))
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function Power_Trace_Resistance_Meas__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    Dim p1 As New PinList
    p1.value = v(0)
    Power_Trace_Resistance_Meas__ = VBAProject.VBT_PC_Diagnosis.Power_Trace_Resistance_Meas(p1, CDbl(v(1)))
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function UVI80_Power_Trace_Resistance_Meas__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    Dim p1 As New PinList
    p1.value = v(0)
    UVI80_Power_Trace_Resistance_Meas__ = VBAProject.VBT_PC_Diagnosis.UVI80_Power_Trace_Resistance_Meas(p1, CStr(v(1)))
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function PC_Leakage__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    Dim p1 As New PinList
    p1.value = v(3)
    Dim p2 As New PinList
    p2.value = v(6)
    PC_Leakage__ = VBAProject.VBT_PC_Diagnosis.PC_Leakage(CDbl(v(0)), CDbl(v(1)), CDbl(v(2)), p1, CDbl(v(4)), CDbl(v(5)), p2)
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function Loopback_IO_Leakage__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    Dim p1 As New PinList
    p1.value = v(3)
    Dim p2 As New PinList
    p2.value = v(6)
    Loopback_IO_Leakage__ = VBAProject.VBT_PC_Diagnosis.Loopback_IO_Leakage(CDbl(v(0)), CDbl(v(1)), CDbl(v(2)), p1, CDbl(v(4)), CDbl(v(5)), p2)
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function UVS256_HexVs_Leakage__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    Dim p1 As New PinList
    p1.value = v(3)
    UVS256_HexVs_Leakage__ = VBAProject.VBT_PC_Diagnosis.UVS256_HexVs_Leakage(CDbl(v(0)), CDbl(v(1)), CDbl(v(2)), p1, CDbl(v(4)), CDbl(v(5)))
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function UVI80_Leakage__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    Dim p1 As New PinList
    p1.value = v(3)
    UVI80_Leakage__ = VBAProject.VBT_PC_Diagnosis.UVI80_Leakage(CDbl(v(0)), CDbl(v(1)), CDbl(v(2)), p1, CDbl(v(4)), CDbl(v(5)))
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function Relay_Circuit_check__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    Dim p1 As New PinList
    p1.value = v(6)
    Dim p2 As New PinList
    p2.value = v(8)
    Dim p3 As New PinList
    p3.value = v(9)
    Relay_Circuit_check__ = VBAProject.VBT_PC_Diagnosis.Relay_Circuit_check(CDbl(v(0)), CDbl(v(1)), CDbl(v(2)), CDbl(v(3)), CDbl(v(4)), CDbl(v(5)), p1, CDbl(v(7)), p2, p3)
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function Relay_Circuit_check_short__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    Dim p1 As New PinList
    p1.value = v(6)
    Dim p2 As New PinList
    p2.value = v(8)
    Dim p3 As New PinList
    p3.value = v(9)
    Relay_Circuit_check_short__ = VBAProject.VBT_PC_Diagnosis.Relay_Circuit_check_short(CDbl(v(0)), CDbl(v(1)), CDbl(v(2)), CDbl(v(3)), CDbl(v(4)), CDbl(v(5)), p1, CDbl(v(7)), p2, p3)
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function DPDT_Relay_check__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    Dim p1 As New PinList
    p1.value = v(5)
    Dim p2 As New PinList
    p2.value = v(7)
    Dim p3 As New PinList
    p3.value = v(8)
    DPDT_Relay_check__ = VBAProject.VBT_PC_Diagnosis.DPDT_Relay_check(CDbl(v(0)), CDbl(v(1)), CDbl(v(2)), CDbl(v(3)), CDbl(v(4)), p1, CDbl(v(6)), p2, p3)
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function DCVS_MeasureCapacitor_2__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    DCVS_MeasureCapacitor_2__ = VBAProject.VBT_PC_Diagnosis.DCVS_MeasureCapacitor_2(CStr(v(0)), CStr(v(1)), CDbl(v(2)), CDbl(v(3)), CDbl(v(4)), CDbl(v(5)), CDbl(v(6)), CDbl(v(7)), CStr(v(8)), CBool(v(9)))
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function UVI80_DPDT_Relay_check__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    Dim p1 As New PinList
    p1.value = v(5)
    Dim p2 As New PinList
    p2.value = v(7)
    Dim p3 As New PinList
    p3.value = v(8)
    UVI80_DPDT_Relay_check__ = VBAProject.VBT_PC_Diagnosis.UVI80_DPDT_Relay_check(CDbl(v(0)), CDbl(v(1)), CDbl(v(2)), CDbl(v(3)), CDbl(v(4)), p1, CDbl(v(6)), p2, p3)
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function Rly_Normal_Close_check__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    Dim p1 As New PinList
    p1.value = v(5)
    Dim p2 As New PinList
    p2.value = v(7)
    Dim p3 As New PinList
    p3.value = v(8)
    Rly_Normal_Close_check__ = VBAProject.VBT_PC_Diagnosis.Rly_Normal_Close_check(CDbl(v(0)), CDbl(v(1)), CDbl(v(2)), CDbl(v(3)), CDbl(v(4)), p1, CDbl(v(6)), p2, p3)
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function UVI80_Relay_check_Short__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    Dim p1 As New PinList
    p1.value = v(5)
    Dim p2 As New PinList
    p2.value = v(7)
    Dim p3 As New PinList
    p3.value = v(8)
    UVI80_Relay_check_Short__ = VBAProject.VBT_PC_Diagnosis.UVI80_Relay_check_Short(CDbl(v(0)), CDbl(v(1)), CDbl(v(2)), CDbl(v(3)), CDbl(v(4)), p1, CDbl(v(6)), p2, p3)
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function K1_3P01KR_check__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    Dim p1 As New PinList
    p1.value = v(0)
    Dim p2 As New PinList
    p2.value = v(2)
    Dim p3 As New PinList
    p3.value = v(3)
    K1_3P01KR_check__ = VBAProject.VBT_PC_Diagnosis.K1_3P01KR_check(p1, CDbl(v(1)), p2, p3)
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function DCVS_MeasureCapacitor__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    DCVS_MeasureCapacitor__ = VBAProject.VBT_PC_Diagnosis.DCVS_MeasureCapacitor(CStr(v(0)), CStr(v(1)), CDbl(v(2)), CDbl(v(3)), CDbl(v(4)), CDbl(v(5)), CDbl(v(6)), CDbl(v(7)), CStr(v(8)), CBool(v(9)))
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function DCVS_MeasureCapacitor_check__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    DCVS_MeasureCapacitor_check__ = VBAProject.VBT_PC_Diagnosis.DCVS_MeasureCapacitor_check(CStr(v(0)), CStr(v(1)), CDbl(v(2)), CDbl(v(3)), CDbl(v(4)), CDbl(v(5)), CDbl(v(6)), CStr(v(7)), CBool(v(8)))
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function MEM0_cap__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    Dim p1 As New PinList
    p1.value = v(0)
    Dim p2 As New PinList
    p2.value = v(2)
    Dim p3 As New PinList
    p3.value = v(3)
    MEM0_cap__ = VBAProject.VBT_PC_Diagnosis.MEM0_cap(p1, CDbl(v(1)), p2, p3)
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function DIB_ReadTDR_Value__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    DIB_ReadTDR_Value__ = VBAProject.VBT_PC_Diagnosis.DIB_ReadTDR_Value()
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function Check_User_Power__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    Check_User_Power__ = VBAProject.VBT_PC_Diagnosis.Check_User_Power()
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function DIBCheckerTestFunction_CheckUDBState__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    DIBCheckerTestFunction_CheckUDBState__ = VBAProject.VBT_PC_Diagnosis.DIBCheckerTestFunction_CheckUDBState()
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function CheckUdbState__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    tl_dt_ErrorMsg "Unsupported type at function CheckUdbState"

    ' CheckUdbState__ = VBAProject.VBT_PC_Diagnosis.CheckUdbState(*One or more unsupported types in argument list or non Long/Integer return type*)
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function Read_Device_ID__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    Read_Device_ID__ = VBAProject.VBT_PC_Diagnosis.Read_Device_ID(CBool(v(UBound(v))))
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function dib_nwire_test__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    dib_nwire_test__ = VBAProject.VBT_PC_Diagnosis.dib_nwire_test(CLng(v(0)))
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function FreeRunClk_Disable_dibchecker__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    FreeRunClk_Disable_dibchecker__ = VBAProject.VBT_PC_Diagnosis.FreeRunClk_Disable_dibchecker(CStr(v(0)))
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function StartSupportBoardClock_dibchecker__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    StartSupportBoardClock_dibchecker__ = VBAProject.VBT_PC_Diagnosis.StartSupportBoardClock_dibchecker()
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function MeasFreq_dibchecker__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    MeasFreq_dibchecker__ = VBAProject.VBT_PC_Diagnosis.MeasFreq_dibchecker(CStr(v(0)), CDbl(v(1)), CStr(v(2)), CBool(v(3)))
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function FreeRunclk_Enable_dibchecker__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    FreeRunclk_Enable_dibchecker__ = VBAProject.VBT_PC_Diagnosis.FreeRunclk_Enable_dibchecker(CStr(v(0)), CStr(v(1)), CStr(v(2)), CStr(v(3)))
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function get_status_reg__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    tl_dt_ErrorMsg "Unsupported type at function get_status_reg"

    ' get_status_reg__ = VBAProject.VBT_PC_Diagnosis.get_status_reg(*One or more unsupported types in argument list or non Long/Integer return type*)
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function get_data__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    tl_dt_ErrorMsg "Unsupported type at function get_data"

    ' Call VBAProject.VBT_PC_Diagnosis.get_data(*One or more unsupported types in argument list or non Long/Integer return type*)
    get_data__ = TL_SUCCESS
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function SetupHRam__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    SetupHRam__ = VBAProject.VBT_PC_Diagnosis.SetupHRam()
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function ReadEEPROM_Block__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    ReadEEPROM_Block__ = VBAProject.VBT_PC_Diagnosis.ReadEEPROM_Block()
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function MCC_Cap_meas__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    MCC_Cap_meas__ = VBAProject.VBT_PC_Diagnosis.MCC_Cap_meas(CStr(v(0)), CStr(v(1)), CStr(v(2)), CDbl(v(3)), CDbl(v(4)), CDbl(v(5)), CBool(v(UBound(v))))
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function LoopBack_1MHz_VOH_check__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    LoopBack_1MHz_VOH_check__ = VBAProject.VBT_PC_Diagnosis.LoopBack_1MHz_VOH_check(CBool(v(UBound(v))))
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function LoopBack_1MHz_VOH_check_2__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    LoopBack_1MHz_VOH_check_2__ = VBAProject.VBT_PC_Diagnosis.LoopBack_1MHz_VOH_check_2(CBool(v(UBound(v))))
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function FreqSweptTest_1MHz__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    Dim p1 As New PinList
    p1.value = v(0)
    FreqSweptTest_1MHz__ = VBAProject.VBT_PC_Diagnosis.FreqSweptTest_1MHz(p1, CDbl(v(1)), CStr(v(2)), CDbl(v(3)), CDbl(v(4)), CDbl(v(5)), CDbl(v(6)), CInt(v(7)))
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function MeasFreq_VohSwept__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    tl_dt_ErrorMsg "Unsupported type at function MeasFreq_VohSwept"

    ' MeasFreq_VohSwept__ = VBAProject.VBT_PC_Diagnosis.MeasFreq_VohSwept(*One or more unsupported types in argument list or non Long/Integer return type*)
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function MeasFreqFunc__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    tl_dt_ErrorMsg "Unsupported type at function MeasFreqFunc"

    ' MeasFreqFunc__ = VBAProject.VBT_PC_Diagnosis.MeasFreqFunc(*One or more unsupported types in argument list or non Long/Integer return type*)
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function CheckStopCondition__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    tl_dt_ErrorMsg "Unsupported type at function CheckStopCondition"

    ' CheckStopCondition__ = VBAProject.VBT_PC_Diagnosis.CheckStopCondition(*One or more unsupported types in argument list or non Long/Integer return type*)
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function InitPinListData__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    tl_dt_ErrorMsg "Unsupported type at function InitPinListData"

    ' InitPinListData__ = VBAProject.VBT_PC_Diagnosis.InitPinListData(*One or more unsupported types in argument list or non Long/Integer return type*)
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function InitExecData__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    tl_dt_ErrorMsg "Unsupported type at function InitExecData"

    ' InitExecData__ = VBAProject.VBT_PC_Diagnosis.InitExecData(*One or more unsupported types in argument list or non Long/Integer return type*)
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function PrintDebugLog__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    tl_dt_ErrorMsg "Unsupported type at function PrintDebugLog"

    ' PrintDebugLog__ = VBAProject.VBT_PC_Diagnosis.PrintDebugLog(*One or more unsupported types in argument list or non Long/Integer return type*)
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function ExtendStr__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    tl_dt_ErrorMsg "Unsupported type at function ExtendStr"

    ' ExtendStr__ = VBAProject.VBT_PC_Diagnosis.ExtendStr(*One or more unsupported types in argument list or non Long/Integer return type*)
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function FindNextExecDef__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    tl_dt_ErrorMsg "Unsupported type at function FindNextExecDef"

    ' FindNextExecDef__ = VBAProject.VBT_PC_Diagnosis.FindNextExecDef(*One or more unsupported types in argument list or non Long/Integer return type*)
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function FillSimulateData__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    tl_dt_ErrorMsg "Unsupported type at function FillSimulateData"

    ' FillSimulateData__ = VBAProject.VBT_PC_Diagnosis.FillSimulateData(*One or more unsupported types in argument list or non Long/Integer return type*)
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function AssignPinMeasData__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    tl_dt_ErrorMsg "Unsupported type at function AssignPinMeasData"

    ' AssignPinMeasData__ = VBAProject.VBT_PC_Diagnosis.AssignPinMeasData(*One or more unsupported types in argument list or non Long/Integer return type*)
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function LoopBack_Freq__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    LoopBack_Freq__ = VBAProject.VBT_PC_Diagnosis.LoopBack_Freq(CBool(v(UBound(v))))
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function LoopBack_Freq_2__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    LoopBack_Freq_2__ = VBAProject.VBT_PC_Diagnosis.LoopBack_Freq_2(CBool(v(UBound(v))))
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function Read_Status_Reg__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    Read_Status_Reg__ = VBAProject.VBT_PC_Diagnosis.Read_Status_Reg(CBool(v(UBound(v))))
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function get_status_reg2__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    tl_dt_ErrorMsg "Unsupported type at function get_status_reg2"

    ' get_status_reg2__ = VBAProject.VBT_PC_Diagnosis.get_status_reg2(*One or more unsupported types in argument list or non Long/Integer return type*)
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function Relay_Parasitic_Cap__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    Relay_Parasitic_Cap__ = VBAProject.VBT_PC_Diagnosis.Relay_Parasitic_Cap(CStr(v(0)), CStr(v(1)), CDbl(v(2)), CStr(v(3)), CStr(v(4)), CDbl(v(5)), CDbl(v(6)), CDbl(v(7)), CDbl(v(8)), CBool(v(UBound(v))))
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function DCVI_Cap_meas__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    Dim p1 As New PinList
    p1.value = v(7)
    DCVI_Cap_meas__ = VBAProject.VBT_PC_Diagnosis.DCVI_Cap_meas(CStr(v(0)), CStr(v(1)), CDbl(v(2)), CDbl(v(3)), CDbl(v(4)), CDbl(v(5)), CDbl(v(6)), p1)
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function DCVI_Cap_Meas_2__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    Dim p1 As New PinList
    p1.value = v(7)
    DCVI_Cap_Meas_2__ = VBAProject.VBT_PC_Diagnosis.DCVI_Cap_Meas_2(CStr(v(0)), CStr(v(1)), CStr(v(2)), CDbl(v(3)), CDbl(v(4)), CDbl(v(5)), CDbl(v(6)), p1)
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function Relay_Parasitic_Cap_2__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    Dim p1 As New PinList
    p1.value = v(8)
    Relay_Parasitic_Cap_2__ = VBAProject.VBT_PC_Diagnosis.Relay_Parasitic_Cap_2(CStr(v(0)), CStr(v(1)), CStr(v(2)), CDbl(v(3)), CDbl(v(4)), CBool(v(UBound(v))), CDbl(v(6)), CDbl(v(7)), p1)
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function PPMU_Relay_FIMV_rising__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    PPMU_Relay_FIMV_rising__ = VBAProject.VBT_PC_Diagnosis.PPMU_Relay_FIMV_rising(CStr(v(0)), CStr(v(1)), CDbl(v(2)), CDbl(v(3)))
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function DCVI_RLY_Rising_Meas__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    Dim p1 As New PinList
    p1.value = v(5)
    DCVI_RLY_Rising_Meas__ = VBAProject.VBT_PC_Diagnosis.DCVI_RLY_Rising_Meas(CStr(v(0)), CStr(v(1)), CDbl(v(2)), CDbl(v(3)), CStr(v(4)), p1)
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function measureRelayTDRDelta__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    measureRelayTDRDelta__ = VBAProject.VBT_PC_Diagnosis.measureRelayTDRDelta(CStr(v(0)), CStr(v(1)), CDbl(v(2)), CDbl(v(3)), CBool(v(UBound(v))))
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function searchTDRRisingTime__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    tl_dt_ErrorMsg "Unsupported type at function searchTDRRisingTime"

    ' searchTDRRisingTime__ = VBAProject.VBT_PC_Diagnosis.searchTDRRisingTime(*One or more unsupported types in argument list or non Long/Integer return type*)
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function Test_Summary__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    Test_Summary__ = VBAProject.VBT_PC_Diagnosis.Test_Summary()
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function DIB_RegKeyRead__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    DIB_RegKeyRead__ = VBAProject.VBT_PC_Diagnosis.DIB_RegKeyRead(CStr(v(0)))
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function









































Public Function UART_write_pmgr__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    UART_write_pmgr__ = VBAProject.VBT_UART_RX_Module.UART_write_pmgr()
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function









































Public Function UART_read_response__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    UART_read_response__ = VBAProject.VBT_UART_TX_Module.UART_read_response()
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function UART_read_response_extended__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    UART_read_response_extended__ = VBAProject.VBT_UART_TX_Module.UART_read_response_extended()
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function UART_boot__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    UART_boot__ = VBAProject.VBT_UART_TX_Module.UART_boot()
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function









































Public Function SPIROM_Continuity__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    Dim p1 As New PinList
    p1.value = v(0)
    SPIROM_Continuity__ = VBAProject.VBT_Write_SPIROM.SPIROM_Continuity(p1)
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function Get_RomSize__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    Get_RomSize__ = VBAProject.VBT_Write_SPIROM.Get_RomSize(CStr(v(0)))
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function PageProgram__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    Dim p1 As New Pattern
    p1.value = v(0)
    Dim p2 As New PinList
    p2.value = v(2)
    Dim p3 As New PinList
    p3.value = v(3)
    PageProgram__ = VBAProject.VBT_Write_SPIROM.PageProgram(p1, CStr(v(1)), p2, p3)
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function PageProgram_4Byte__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    Dim p1 As New Pattern
    p1.value = v(0)
    Dim p2 As New PinList
    p2.value = v(2)
    Dim p3 As New PinList
    p3.value = v(3)
    PageProgram_4Byte__ = VBAProject.VBT_Write_SPIROM.PageProgram_4Byte(p1, CStr(v(1)), p2, p3)
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function SPIROM_p2p_short_Power__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    Dim p1 As New PinList
    p1.value = v(0)
    SPIROM_p2p_short_Power__ = VBAProject.VBT_Write_SPIROM.SPIROM_p2p_short_Power(p1, CDbl(v(1)))
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function SPIROM_Erase_Universal__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    Dim p1 As New Pattern
    p1.value = v(0)
    Dim p2 As New PinList
    p2.value = v(1)
    Dim p3 As New PinList
    p3.value = v(2)
    SPIROM_Erase_Universal__ = VBAProject.VBT_Write_SPIROM.SPIROM_Erase_Universal(p1, p2, p3, CLng(v(3)))
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function SPIROM_SectorErase_Universal__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    Dim p1 As New Pattern
    p1.value = v(0)
    Dim p2 As New PinList
    p2.value = v(1)
    Dim p3 As New PinList
    p3.value = v(2)
    SPIROM_SectorErase_Universal__ = VBAProject.VBT_Write_SPIROM.SPIROM_SectorErase_Universal(p1, p2, p3)
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function SPIROM_InitialRead_Universal__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    Dim p1 As New Pattern
    p1.value = v(0)
    Dim p2 As New PinList
    p2.value = v(1)
    Dim p3 As New PinList
    p3.value = v(2)
    Dim p4 As New PinList
    p4.value = v(3)
    Dim p5 As New PinList
    p5.value = v(4)
    SPIROM_InitialRead_Universal__ = VBAProject.VBT_Write_SPIROM.SPIROM_InitialRead_Universal(p1, p2, p3, p4, p5, CLng(v(5)))
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function Read_RomCode__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    Read_RomCode__ = VBAProject.VBT_Write_SPIROM.Read_RomCode(CStr(v(0)), CLng(v(1)), CStr(v(2)))
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function InitRead__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    Dim p1 As New Pattern
    p1.value = v(0)
    Dim p2 As New PinList
    p2.value = v(1)
    Dim p3 As New PinList
    p3.value = v(2)
    Dim p4 As New PinList
    p4.value = v(6)
    InitRead__ = VBAProject.VBT_Write_SPIROM.InitRead(p1, p2, p3, CLng(v(3)), CStr(v(4)), CInt(v(5)), p4)
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function SPIROM_CheckSum_Universal__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    Dim p1 As New Pattern
    p1.value = v(0)
    Dim p2 As New PinList
    p2.value = v(1)
    Dim p3 As New PinList
    p3.value = v(2)
    Dim p4 As New PinList
    p4.value = v(3)
    Dim p5 As New PinList
    p5.value = v(4)
    SPIROM_CheckSum_Universal__ = VBAProject.VBT_Write_SPIROM.SPIROM_CheckSum_Universal(p1, p2, p3, p4, p5, CLng(v(5)))
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function SPIROM_Program_Universal__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    Dim p1 As New Pattern
    p1.value = v(0)
    Dim p2 As New PinList
    p2.value = v(1)
    Dim p3 As New PinList
    p3.value = v(2)
    SPIROM_Program_Universal__ = VBAProject.VBT_Write_SPIROM.SPIROM_Program_Universal(p1, p2, p3)
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function SPIROM_Program_4ByteAddr_Universal__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    Dim p1 As New Pattern
    p1.value = v(0)
    Dim p2 As New PinList
    p2.value = v(1)
    Dim p3 As New PinList
    p3.value = v(2)
    SPIROM_Program_4ByteAddr_Universal__ = VBAProject.VBT_Write_SPIROM.SPIROM_Program_4ByteAddr_Universal(p1, p2, p3)
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function SPIROM_Device_Check_Universal__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    Dim p1 As New Pattern
    p1.value = v(0)
    Dim p2 As New PinList
    p2.value = v(1)
    Dim p3 As New PinList
    p3.value = v(2)
    SPIROM_Device_Check_Universal__ = VBAProject.VBT_Write_SPIROM.SPIROM_Device_Check_Universal(p1, p2, p3, CLng(v(3)))
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function CheckSum__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    Dim p1 As New Pattern
    p1.value = v(0)
    Dim p2 As New PinList
    p2.value = v(1)
    Dim p3 As New PinList
    p3.value = v(2)
    Dim p4 As New PinList
    p4.value = v(5)
    CheckSum__ = VBAProject.VBT_Write_SPIROM.CheckSum(p1, p2, p3, CStr(v(3)), CLng(v(4)), p4)
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function SPIROM_ROMFile_Check__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    tl_dt_ErrorMsg "Unsupported type at function SPIROM_ROMFile_Check"

    ' SPIROM_ROMFile_Check__ = VBAProject.VBT_Write_SPIROM.SPIROM_ROMFile_Check(*One or more unsupported types in argument list or non Long/Integer return type*)
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function RTOS_Parse_Info__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    RTOS_Parse_Info__ = VBAProject.VBT_Write_SPIROM.RTOS_Parse_Info()
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function









































Public Function RunDspSet__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    Call VBAProject.VBT_ZeDspOperation.RunDspSet
    RunDspSet__ = TL_SUCCESS
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function Ze_SingleDoubleCal__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    tl_dt_ErrorMsg "Unsupported type at function Ze_SingleDoubleCal"

    ' Call VBAProject.VBT_ZeDspOperation.Ze_SingleDoubleCal(*One or more unsupported types in argument list or non Long/Integer return type*)
    Ze_SingleDoubleCal__ = TL_SUCCESS
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function Ze_DspwaveToParallelMsb__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    tl_dt_ErrorMsg "Unsupported type at function Ze_DspwaveToParallelMsb"

    ' Call VBAProject.VBT_ZeDspOperation.Ze_DspwaveToParallelMsb(*One or more unsupported types in argument list or non Long/Integer return type*)
    Ze_DspwaveToParallelMsb__ = TL_SUCCESS
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function Ze_XExpandBitVector__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    tl_dt_ErrorMsg "Unsupported type at function Ze_XExpandBitVector"

    ' Call VBAProject.VBT_ZeDspOperation.Ze_XExpandBitVector(*One or more unsupported types in argument list or non Long/Integer return type*)
    Ze_XExpandBitVector__ = TL_SUCCESS
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function Ze_DspwaveParallel8ToParallel32__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    tl_dt_ErrorMsg "Unsupported type at function Ze_DspwaveParallel8ToParallel32"

    ' Call VBAProject.VBT_ZeDspOperation.Ze_DspwaveParallel8ToParallel32(*One or more unsupported types in argument list or non Long/Integer return type*)
    Ze_DspwaveParallel8ToParallel32__ = TL_SUCCESS
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function Ze_DspwaveBlankCheck__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    tl_dt_ErrorMsg "Unsupported type at function Ze_DspwaveBlankCheck"

    ' Call VBAProject.VBT_ZeDspOperation.Ze_DspwaveBlankCheck(*One or more unsupported types in argument list or non Long/Integer return type*)
    Ze_DspwaveBlankCheck__ = TL_SUCCESS
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function Ze_SplitWaveSerial32__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    tl_dt_ErrorMsg "Unsupported type at function Ze_SplitWaveSerial32"

    ' Call VBAProject.VBT_ZeDspOperation.Ze_SplitWaveSerial32(*One or more unsupported types in argument list or non Long/Integer return type*)
    Ze_SplitWaveSerial32__ = TL_SUCCESS
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function Ze_SplitWave32__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    tl_dt_ErrorMsg "Unsupported type at function Ze_SplitWave32"

    ' Call VBAProject.VBT_ZeDspOperation.Ze_SplitWave32(*One or more unsupported types in argument list or non Long/Integer return type*)
    Ze_SplitWave32__ = TL_SUCCESS
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function Ze_SplitWave8__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    tl_dt_ErrorMsg "Unsupported type at function Ze_SplitWave8"

    ' Call VBAProject.VBT_ZeDspOperation.Ze_SplitWave8(*One or more unsupported types in argument list or non Long/Integer return type*)
    Ze_SplitWave8__ = TL_SUCCESS
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function Ze_GetSingleBitWave__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    tl_dt_ErrorMsg "Unsupported type at function Ze_GetSingleBitWave"

    ' Call VBAProject.VBT_ZeDspOperation.Ze_GetSingleBitWave(*One or more unsupported types in argument list or non Long/Integer return type*)
    Ze_GetSingleBitWave__ = TL_SUCCESS
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function Ze_GetSingleBitWave_Jtag__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    tl_dt_ErrorMsg "Unsupported type at function Ze_GetSingleBitWave_Jtag"

    ' Call VBAProject.VBT_ZeDspOperation.Ze_GetSingleBitWave_Jtag(*One or more unsupported types in argument list or non Long/Integer return type*)
    Ze_GetSingleBitWave_Jtag__ = TL_SUCCESS
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function Ze_GetDspWaveExMask__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    tl_dt_ErrorMsg "Unsupported type at function Ze_GetDspWaveExMask"

    ' Call VBAProject.VBT_ZeDspOperation.Ze_GetDspWaveExMask(*One or more unsupported types in argument list or non Long/Integer return type*)
    Ze_GetDspWaveExMask__ = TL_SUCCESS
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function Ze_CalOneInAverage__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    Call VBAProject.VBT_ZeDspOperation.Ze_CalOneInAverage(v(0), CDbl(v(1)), v(2))
    Ze_CalOneInAverage__ = TL_SUCCESS
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function Ze_TwoDspWaveCompare__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    tl_dt_ErrorMsg "Unsupported type at function Ze_TwoDspWaveCompare"

    ' Call VBAProject.VBT_ZeDspOperation.Ze_TwoDspWaveCompare(*One or more unsupported types in argument list or non Long/Integer return type*)
    Ze_TwoDspWaveCompare__ = TL_SUCCESS
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function Ze_XExpandDblBit__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    tl_dt_ErrorMsg "Unsupported type at function Ze_XExpandDblBit"

    ' Call VBAProject.VBT_ZeDspOperation.Ze_XExpandDblBit(*One or more unsupported types in argument list or non Long/Integer return type*)
    Ze_XExpandDblBit__ = TL_SUCCESS
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function Ze_CombineShortWave__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    tl_dt_ErrorMsg "Unsupported type at function Ze_CombineShortWave"

    ' Call VBAProject.VBT_ZeDspOperation.Ze_CombineShortWave(*One or more unsupported types in argument list or non Long/Integer return type*)
    Ze_CombineShortWave__ = TL_SUCCESS
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function









































Public Function eFuseInitialization__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    eFuseInitialization__ = VBAProject.VBT_ZeFuse_Glb.eFuseInitialization()
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function ObtainTrimOverWriteItems__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    ObtainTrimOverWriteItems__ = VBAProject.VBT_ZeFuse_Glb.ObtainTrimOverWriteItems()
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function DecideEarlyFuseAlgorithm__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    Call VBAProject.VBT_ZeFuse_Glb.DecideEarlyFuseAlgorithm
    DecideEarlyFuseAlgorithm__ = TL_SUCCESS
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function eFuse_GetReadDecimal__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    tl_dt_ErrorMsg "Unsupported type at function eFuse_GetReadDecimal"

    ' eFuse_GetReadDecimal__ = VBAProject.VBT_ZeFuse_Glb.eFuse_GetReadDecimal(*One or more unsupported types in argument list or non Long/Integer return type*)
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function EfuseOnPgmValidate__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    tl_dt_ErrorMsg "Unsupported type at function EfuseOnPgmValidate"

    ' EfuseOnPgmValidate__ = VBAProject.VBT_ZeFuse_Glb.EfuseOnPgmValidate(*One or more unsupported types in argument list or non Long/Integer return type*)
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function auto_eFuse_GetIDSResolution__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    tl_dt_ErrorMsg "Unsupported type at function auto_eFuse_GetIDSResolution"

    ' auto_eFuse_GetIDSResolution__ = VBAProject.VBT_ZeFuse_Glb.auto_eFuse_GetIDSResolution(*One or more unsupported types in argument list or non Long/Integer return type*)
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function GetEfuseHipValue__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    tl_dt_ErrorMsg "Unsupported type at function GetEfuseHipValue"

    ' GetEfuseHipValue__ = VBAProject.VBT_ZeFuse_Glb.GetEfuseHipValue(*One or more unsupported types in argument list or non Long/Integer return type*)
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function SetEfuseHipValue__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    Call VBAProject.VBT_ZeFuse_Glb.SetEfuseHipValue(CStr(v(0)), CStr(v(1)), v(2), v(3))
    SetEfuseHipValue__ = TL_SUCCESS
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function GetBdfBank__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    tl_dt_ErrorMsg "Unsupported type at function GetBdfBank"

    ' GetBdfBank__ = VBAProject.VBT_ZeFuse_Glb.GetBdfBank(*One or more unsupported types in argument list or non Long/Integer return type*)
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function DicRemoved__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    tl_dt_ErrorMsg "Unsupported type at function DicRemoved"

    ' Call VBAProject.VBT_ZeFuse_Glb.DicRemoved(*One or more unsupported types in argument list or non Long/Integer return type*)
    DicRemoved__ = TL_SUCCESS
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function DicSameKept__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    tl_dt_ErrorMsg "Unsupported type at function DicSameKept"

    ' Call VBAProject.VBT_ZeFuse_Glb.DicSameKept(*One or more unsupported types in argument list or non Long/Integer return type*)
    DicSameKept__ = TL_SUCCESS
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function DicCloned__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    tl_dt_ErrorMsg "Unsupported type at function DicCloned"

    ' Call VBAProject.VBT_ZeFuse_Glb.DicCloned(*One or more unsupported types in argument list or non Long/Integer return type*)
    DicCloned__ = TL_SUCCESS
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function CfgEnableCheck__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    tl_dt_ErrorMsg "Unsupported type at function CfgEnableCheck"

    ' CfgEnableCheck__ = VBAProject.VBT_ZeFuse_Glb.CfgEnableCheck(*One or more unsupported types in argument list or non Long/Integer return type*)
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function eFuse_DSSC_CapSetup__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    Dim p1 As ReadTestType
    p1 = v(4)
    tl_dt_ErrorMsg "Unsupported type at function eFuse_DSSC_CapSetup"

    ' eFuse_DSSC_CapSetup__ = VBAProject.VBT_ZeFuse_Glb.eFuse_DSSC_CapSetup(*One or more unsupported types in argument list or non Long/Integer return type*)
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function eFuse_DSSC_SrcSetup__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    tl_dt_ErrorMsg "Unsupported type at function eFuse_DSSC_SrcSetup"

    ' eFuse_DSSC_SrcSetup__ = VBAProject.VBT_ZeFuse_Glb.eFuse_DSSC_SrcSetup(*One or more unsupported types in argument list or non Long/Integer return type*)
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function GenerateEcid__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    tl_dt_ErrorMsg "Unsupported type at function GenerateEcid"

    ' Call VBAProject.VBT_ZeFuse_Glb.GenerateEcid(*One or more unsupported types in argument list or non Long/Integer return type*)
    GenerateEcid__ = TL_SUCCESS
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function EfuseTrimData_ATETrimValue_BitWidth_Read__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    tl_dt_ErrorMsg "Unsupported type at function EfuseTrimData_ATETrimValue_BitWidth_Read"

    ' EfuseTrimData_ATETrimValue_BitWidth_Read__ = VBAProject.VBT_ZeFuse_Glb.EfuseTrimData_ATETrimValue_BitWidth_Read(*One or more unsupported types in argument list or non Long/Integer return type*)
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function EfuseTrimData_ATETrimValue_Write__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    tl_dt_ErrorMsg "Unsupported type at function EfuseTrimData_ATETrimValue_Write"

    ' Call VBAProject.VBT_ZeFuse_Glb.EfuseTrimData_ATETrimValue_Write(*One or more unsupported types in argument list or non Long/Integer return type*)
    EfuseTrimData_ATETrimValue_Write__ = TL_SUCCESS
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function EfuseTrimData_ATETrimValue_Read__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    tl_dt_ErrorMsg "Unsupported type at function EfuseTrimData_ATETrimValue_Read"

    ' EfuseTrimData_ATETrimValue_Read__ = VBAProject.VBT_ZeFuse_Glb.EfuseTrimData_ATETrimValue_Read(*One or more unsupported types in argument list or non Long/Integer return type*)
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function UpdateFullCfg2BdfEfuse__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    UpdateFullCfg2BdfEfuse__ = VBAProject.VBT_ZeFuse_Glb.UpdateFullCfg2BdfEfuse()
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function UpdateCfg2BdfEfuse__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    Dim p1 As DataRetrieve
    p1 = v(1)
    UpdateCfg2BdfEfuse__ = VBAProject.VBT_ZeFuse_Glb.UpdateCfg2BdfEfuse(CStr(v(0)), p1, CStr(v(2)))
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function IsInstMatch__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    tl_dt_ErrorMsg "Unsupported type at function IsInstMatch"

    ' IsInstMatch__ = VBAProject.VBT_ZeFuse_Glb.IsInstMatch(*One or more unsupported types in argument list or non Long/Integer return type*)
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function DoubleBitChecks__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    tl_dt_ErrorMsg "Unsupported type at function DoubleBitChecks"

    ' DoubleBitChecks__ = VBAProject.VBT_ZeFuse_Glb.DoubleBitChecks(*One or more unsupported types in argument list or non Long/Integer return type*)
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function UpdateCmpFields__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    Call VBAProject.VBT_ZeFuse_Glb.UpdateCmpFields(CStr(v(0)))
    UpdateCmpFields__ = TL_SUCCESS
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function AllBanksPatFailCnt__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    tl_dt_ErrorMsg "Unsupported type at function AllBanksPatFailCnt"

    ' AllBanksPatFailCnt__ = VBAProject.VBT_ZeFuse_Glb.AllBanksPatFailCnt(*One or more unsupported types in argument list or non Long/Integer return type*)
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function PatFailCnt__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    tl_dt_ErrorMsg "Unsupported type at function PatFailCnt"

    ' PatFailCnt__ = VBAProject.VBT_ZeFuse_Glb.PatFailCnt(*One or more unsupported types in argument list or non Long/Integer return type*)
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function PutData2RegKeySave__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    tl_dt_ErrorMsg "Unsupported type at function PutData2RegKeySave"

    ' Call VBAProject.VBT_ZeFuse_Glb.PutData2RegKeySave(*One or more unsupported types in argument list or non Long/Integer return type*)
    PutData2RegKeySave__ = TL_SUCCESS
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function BKM_Update__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    BKM_Update__ = VBAProject.VBT_ZeFuse_Glb.BKM_Update(CStr(v(0)), CStr(v(1)), CBool(v(2)), CLng(v(3)))
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function CheckFileExist__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    tl_dt_ErrorMsg "Unsupported type at function CheckFileExist"

    ' CheckFileExist__ = VBAProject.VBT_ZeFuse_Glb.CheckFileExist(*One or more unsupported types in argument list or non Long/Integer return type*)
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function GetFusedBKMData__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    GetFusedBKMData__ = VBAProject.VBT_ZeFuse_Glb.GetFusedBKMData(CStr(v(0)), CStr(v(1)))
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function WriteIEDARegistry__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    WriteIEDARegistry__ = VBAProject.VBT_ZeFuse_Glb.WriteIEDARegistry(CStr(v(0)), CStr(v(1)), CStr(v(2)))
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function Glb_Setup__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    Call VBAProject.VBT_ZeFuse_Glb.Glb_Setup
    Glb_Setup__ = TL_SUCCESS
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function DebugPrintSingleDoubleBitCheck__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    tl_dt_ErrorMsg "Unsupported type at function DebugPrintSingleDoubleBitCheck"

    ' DebugPrintSingleDoubleBitCheck__ = VBAProject.VBT_ZeFuse_Glb.DebugPrintSingleDoubleBitCheck(*One or more unsupported types in argument list or non Long/Integer return type*)
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function GetCompareFailBit__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    tl_dt_ErrorMsg "Unsupported type at function GetCompareFailBit"

    ' GetCompareFailBit__ = VBAProject.VBT_ZeFuse_Glb.GetCompareFailBit(*One or more unsupported types in argument list or non Long/Integer return type*)
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function ShadowAndNonShadowSameStage__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    ShadowAndNonShadowSameStage__ = VBAProject.VBT_ZeFuse_Glb.ShadowAndNonShadowSameStage()
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function









































Public Function FakeBankDsscResult_Serial__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    tl_dt_ErrorMsg "Unsupported type at function FakeBankDsscResult_Serial"

    ' FakeBankDsscResult_Serial__ = VBAProject.VBT_ZeFuseOffline.FakeBankDsscResult_Serial(*One or more unsupported types in argument list or non Long/Integer return type*)
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function FakeBankDsscResult__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    tl_dt_ErrorMsg "Unsupported type at function FakeBankDsscResult"

    ' FakeBankDsscResult__ = VBAProject.VBT_ZeFuseOffline.FakeBankDsscResult(*One or more unsupported types in argument list or non Long/Integer return type*)
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function









































Public Function PseudoFuse_WriteToFile__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    PseudoFuse_WriteToFile__ = VBAProject.VBT_ZeFuseOthers.PseudoFuse_WriteToFile(CStr(v(0)))
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function PseudoFuse_ReadFromFile__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    PseudoFuse_ReadFromFile__ = VBAProject.VBT_ZeFuseOthers.PseudoFuse_ReadFromFile(CStr(v(0)))
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function SwitchFlag__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    Call VBAProject.VBT_ZeFuseOthers.SwitchFlag
    SwitchFlag__ = TL_SUCCESS
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function SwitchEfuseDecodeFlag__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    Call VBAProject.VBT_ZeFuseOthers.SwitchEfuseDecodeFlag
    SwitchEfuseDecodeFlag__ = TL_SUCCESS
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function CheckValueToDec__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    tl_dt_ErrorMsg "Unsupported type at function CheckValueToDec"

    ' CheckValueToDec__ = VBAProject.VBT_ZeFuseOthers.CheckValueToDec(*One or more unsupported types in argument list or non Long/Integer return type*)
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function ParseBinChk__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    tl_dt_ErrorMsg "Unsupported type at function ParseBinChk"

    ' ParseBinChk__ = VBAProject.VBT_ZeFuseOthers.ParseBinChk(*One or more unsupported types in argument list or non Long/Integer return type*)
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function cmplfuse_createrealfielmask_setfuse__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    tl_dt_ErrorMsg "Unsupported type at function cmplfuse_createrealfielmask_setfuse"

    ' cmplfuse_createrealfielmask_setfuse__ = VBAProject.VBT_ZeFuseOthers.cmplfuse_createrealfielmask_setfuse(*One or more unsupported types in argument list or non Long/Integer return type*)
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function cmplfuse_modifydspwave_avoiddoublefusing__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    tl_dt_ErrorMsg "Unsupported type at function cmplfuse_modifydspwave_avoiddoublefusing"

    ' cmplfuse_modifydspwave_avoiddoublefusing__ = VBAProject.VBT_ZeFuseOthers.cmplfuse_modifydspwave_avoiddoublefusing(*One or more unsupported types in argument list or non Long/Integer return type*)
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function ParseBincut__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    tl_dt_ErrorMsg "Unsupported type at function ParseBincut"

    ' ParseBincut__ = VBAProject.VBT_ZeFuseOthers.ParseBincut(*One or more unsupported types in argument list or non Long/Integer return type*)
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function EvsFuse_WriteToFuse__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    EvsFuse_WriteToFuse__ = VBAProject.VBT_ZeFuseOthers.EvsFuse_WriteToFuse(CStr(v(0)), CStr(v(1)), CStr(v(2)), CStr(v(3)), CStr(v(4)))
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function EvsFuse_TestedWriteToFuse__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    EvsFuse_TestedWriteToFuse__ = VBAProject.VBT_ZeFuseOthers.EvsFuse_TestedWriteToFuse(CStr(v(0)), CStr(v(1)))
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function ParserEnableWdGatingTable__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    tl_dt_ErrorMsg "Unsupported type at function ParserEnableWdGatingTable"

    ' ParserEnableWdGatingTable__ = VBAProject.VBT_ZeFuseOthers.ParserEnableWdGatingTable(*One or more unsupported types in argument list or non Long/Integer return type*)
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function EnableWdGating__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    Dim p1 As EnableWdGatingPhase
    p1 = v(0)
    EnableWdGating__ = VBAProject.VBT_ZeFuseOthers.EnableWdGating(p1)
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function GattingOIConfigEnableWD__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    tl_dt_ErrorMsg "Unsupported type at function GattingOIConfigEnableWD"

    ' Call VBAProject.VBT_ZeFuseOthers.GattingOIConfigEnableWD(*One or more unsupported types in argument list or non Long/Integer return type*)
    GattingOIConfigEnableWD__ = TL_SUCCESS
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function JudgeDRAMType_T__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    JudgeDRAMType_T__ = VBAProject.VBT_ZeFuseOthers.JudgeDRAMType_T()
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function CheckDRAMEnableWordExist__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    tl_dt_ErrorMsg "Unsupported type at function CheckDRAMEnableWordExist"

    ' CheckDRAMEnableWordExist__ = VBAProject.VBT_ZeFuseOthers.CheckDRAMEnableWordExist(*One or more unsupported types in argument list or non Long/Integer return type*)
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function CheckCurrentJob__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    tl_dt_ErrorMsg "Unsupported type at function CheckCurrentJob"

    ' CheckCurrentJob__ = VBAProject.VBT_ZeFuseOthers.CheckCurrentJob(*One or more unsupported types in argument list or non Long/Integer return type*)
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function UpdateReadPatternResult__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    tl_dt_ErrorMsg "Unsupported type at function UpdateReadPatternResult"

    ' UpdateReadPatternResult__ = VBAProject.VBT_ZeFuseOthers.UpdateReadPatternResult(*One or more unsupported types in argument list or non Long/Integer return type*)
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function









































Public Function Bank_Read__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    Dim p1 As New Pattern
    p1.value = v(0)
    Dim p2 As New PinList
    p2.value = v(1)
    Dim p3 As New PinList
    p3.value = v(12)
    Dim p4 As New PinList
    p4.value = v(13)
    Dim p5 As New PinList
    p5.value = v(14)
    Dim p6 As New Pattern
    p6.value = v(15)
    Bank_Read__ = VBAProject.VBT_ZeFuseRead.Bank_Read(p1, p2, CStr(v(2)), CBool(v(3)), CBool(v(4)), CBool(v(5)), CBool(v(6)), CBool(v(7)), CBool(v(8)), CBool(v(9)), CBool(v(10)), CBool(v(11)), p3, p4, p5, p6, CStr(v(16)), CStr(v(17)), CBool(v(UBound(v))))
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function Bank_ApbRead__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    Dim p1 As New Pattern
    p1.value = v(0)
    Dim p2 As New PinList
    p2.value = v(1)
    Dim p3 As New PinList
    p3.value = v(3)
    Dim p4 As New PinList
    p4.value = v(4)
    Dim p5 As New PinList
    p5.value = v(5)
    Bank_ApbRead__ = VBAProject.VBT_ZeFuseRead.Bank_ApbRead(p1, p2, CStr(v(2)), p3, p4, p5, CBool(v(UBound(v))))
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function Bank_TapRead__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    Dim p1 As New Pattern
    p1.value = v(0)
    Dim p2 As New PinList
    p2.value = v(1)
    Dim p3 As New PinList
    p3.value = v(3)
    Dim p4 As New PinList
    p4.value = v(4)
    Dim p5 As New PinList
    p5.value = v(5)
    Dim p6 As New Pattern
    p6.value = v(6)
    Bank_TapRead__ = VBAProject.VBT_ZeFuseRead.Bank_TapRead(p1, p2, CStr(v(2)), p3, p4, p5, p6, CStr(v(7)), CBool(v(UBound(v))))
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function Bank_Udrcmp__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    Dim p1 As New Pattern
    p1.value = v(0)
    Dim p2 As New PinList
    p2.value = v(1)
    Dim p3 As New PinList
    p3.value = v(3)
    Dim p4 As New PinList
    p4.value = v(4)
    Dim p5 As New PinList
    p5.value = v(5)
    Bank_Udrcmp__ = VBAProject.VBT_ZeFuseRead.Bank_Udrcmp(p1, p2, CStr(v(2)), p3, p4, p5, CBool(v(UBound(v))))
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function Bank_SingleDoubleBitsCheck__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    tl_dt_ErrorMsg "Unsupported type at function Bank_SingleDoubleBitsCheck"

    ' Bank_SingleDoubleBitsCheck__ = VBAProject.VBT_ZeFuseRead.Bank_SingleDoubleBitsCheck(*One or more unsupported types in argument list or non Long/Integer return type*)
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function Bank_IsBlank__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    Bank_IsBlank__ = VBAProject.VBT_ZeFuseRead.Bank_IsBlank()
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function Bank_SyntaxCheck__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    Bank_SyntaxCheck__ = VBAProject.VBT_ZeFuseRead.Bank_SyntaxCheck(CStr(v(0)), CBool(v(1)), CBool(v(2)), CBool(v(3)), CBool(v(4)), CBool(v(5)), CBool(v(6)), CBool(v(7)), CStr(v(8)), CBool(v(UBound(v))))
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function Bank_ReadWaferData__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    Bank_ReadWaferData__ = VBAProject.VBT_ZeFuseRead.Bank_ReadWaferData(CStr(v(0)), CBool(v(UBound(v))))
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function PassBinCheck__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    PassBinCheck__ = VBAProject.VBT_ZeFuseRead.PassBinCheck()
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function Bank_CompareWRData__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    Bank_CompareWRData__ = VBAProject.VBT_ZeFuseRead.Bank_CompareWRData(CStr(v(0)), CBool(v(1)), CBool(v(2)), CStr(v(3)), CBool(v(UBound(v))))
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function SplitDspWave__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    Dim p1 As New Pattern
    p1.value = v(0)
    Dim p2 As New PinList
    p2.value = v(1)
    Dim p3 As New PinList
    p3.value = v(4)
    Dim p4 As New PinList
    p4.value = v(5)
    Dim p5 As New PinList
    p5.value = v(6)
    SplitDspWave__ = VBAProject.VBT_ZeFuseRead.SplitDspWave(p1, p2, CStr(v(2)), CLng(v(3)), p3, p4, p5, CBool(v(UBound(v))))
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function Bank_Decode_Print__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    tl_dt_ErrorMsg "Unsupported type at function Bank_Decode_Print"

    ' Bank_Decode_Print__ = VBAProject.VBT_ZeFuseRead.Bank_Decode_Print(*One or more unsupported types in argument list or non Long/Integer return type*)
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function









































Public Function ReadEfuseDataFromBinCut__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    Call VBAProject.VBT_ZeFuseVddBin_Ids_Glb.ReadEfuseDataFromBinCut(CBool(v(0)))
    ReadEfuseDataFromBinCut__ = TL_SUCCESS
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function GetIdsValues__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    Call VBAProject.VBT_ZeFuseVddBin_Ids_Glb.GetIdsValues(CBool(v(0)))
    GetIdsValues__ = TL_SUCCESS
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function PutIdsCodes__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    Call VBAProject.VBT_ZeFuseVddBin_Ids_Glb.PutIdsCodes
    PutIdsCodes__ = TL_SUCCESS
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function ObtainIdsValues__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    tl_dt_ErrorMsg "Unsupported type at function ObtainIdsValues"

    ' ObtainIdsValues__ = VBAProject.VBT_ZeFuseVddBin_Ids_Glb.ObtainIdsValues(*One or more unsupported types in argument list or non Long/Integer return type*)
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function GetBinCutValuesAndLimits__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    Call VBAProject.VBT_ZeFuseVddBin_Ids_Glb.GetBinCutValuesAndLimits(CBool(v(0)))
    GetBinCutValuesAndLimits__ = TL_SUCCESS
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function Ze_Read_DVFM_To_GradeVDD__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    Ze_Read_DVFM_To_GradeVDD__ = VBAProject.VBT_ZeFuseVddBin_Ids_Glb.Ze_Read_DVFM_To_GradeVDD()
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function PutBinCutCodes__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    Call VBAProject.VBT_ZeFuseVddBin_Ids_Glb.PutBinCutCodes
    PutBinCutCodes__ = TL_SUCCESS
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function GenVddBinFuseCode__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    tl_dt_ErrorMsg "Unsupported type at function GenVddBinFuseCode"

    ' GenVddBinFuseCode__ = VBAProject.VBT_ZeFuseVddBin_Ids_Glb.GenVddBinFuseCode(*One or more unsupported types in argument list or non Long/Integer return type*)
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function Ids_BinCut_PostCheck__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    Ids_BinCut_PostCheck__ = VBAProject.VBT_ZeFuseVddBin_Ids_Glb.Ids_BinCut_PostCheck()
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function GetVddBinEqu__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    tl_dt_ErrorMsg "Unsupported type at function GetVddBinEqu"

    ' GetVddBinEqu__ = VBAProject.VBT_ZeFuseVddBin_Ids_Glb.GetVddBinEqu(*One or more unsupported types in argument list or non Long/Integer return type*)
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function GetVddBinPmodeMap__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    Call VBAProject.VBT_ZeFuseVddBin_Ids_Glb.GetVddBinPmodeMap
    GetVddBinPmodeMap__ = TL_SUCCESS
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function









































Public Function Bank_Write__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    Dim p1 As New Pattern
    p1.value = v(0)
    Dim p2 As New PinList
    p2.value = v(1)
    Dim p3 As New PinList
    p3.value = v(10)
    Dim p4 As New PinList
    p4.value = v(11)
    Dim p5 As New PinList
    p5.value = v(12)
    Dim p6 As New Pattern
    p6.value = v(13)
    Bank_Write__ = VBAProject.VBT_ZeFuseWrite.Bank_Write(p1, p2, CStr(v(2)), CBool(v(3)), CBool(v(4)), CStr(v(5)), CDbl(v(6)), CBool(v(7)), CBool(v(8)), CBool(v(9)), p3, p4, p5, p6, CStr(v(14)), CStr(v(15)), CBool(v(UBound(v))))
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function Bank_LoadDefaultValues__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    Bank_LoadDefaultValues__ = VBAProject.VBT_ZeFuseWrite.Bank_LoadDefaultValues(CStr(v(0)), CBool(v(1)), CBool(v(UBound(v))))
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function Bank_TrimmedUpdate__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    Bank_TrimmedUpdate__ = VBAProject.VBT_ZeFuseWrite.Bank_TrimmedUpdate(CStr(v(0)), CBool(v(1)), CBool(v(2)), CBool(v(3)), CStr(v(4)), CBool(v(UBound(v))))
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function ObtainCatDictionary__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    tl_dt_ErrorMsg "Unsupported type at function ObtainCatDictionary"

    ' ObtainCatDictionary__ = VBAProject.VBT_ZeFuseWrite.ObtainCatDictionary(*One or more unsupported types in argument list or non Long/Integer return type*)
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function

Public Function DicTrimmedExistCrc__(v As Variant) As Long
    m_STDSvcClient.ProfileService.OverrideEnabled = True
    If TheExec.RunMode = runModeProduction Or tl_IsRunningSynchronus Or errDestLogfile = TheExec.ErrorOutputMode Then On Error GoTo errpt
    m_STDSvcClient.ProfileService.OverrideEnabled = False
    tl_dt_ErrorMsg "Unsupported type at function DicTrimmedExistCrc"

    ' DicTrimmedExistCrc__ = VBAProject.VBT_ZeFuseWrite.DicTrimmedExistCrc(*One or more unsupported types in argument list or non Long/Integer return type*)
    Exit Function
errpt:     ' Untrapped VB error in production.  Fail the test.
    HandleUntrappedError
End Function









































