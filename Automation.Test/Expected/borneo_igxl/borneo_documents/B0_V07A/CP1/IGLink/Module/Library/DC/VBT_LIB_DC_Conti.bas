Attribute VB_Name = "VBT_LIB_DC_Conti"
#Const isUFP = True
Option Explicit

'20210507 new format
'------------------------------------------------------------------------------------------------------------------------------------
'The argument different setting for AP/RF/LCD
'AP / RF proj---has only one set H/L limit - use LowLimit / HiLimit.
'LCD proj-------has two set H/L limit - one for open, one for short. - LowLimit/HiLimit -> open case, LoLimit2/HiLimit2 -> short case
'-------------------------------------------------------------------------------------------------------------------------------------
' [20230801][All][James] flag need to judge use default or input argument
' [20231003][All][Tank] modify after Chihome review
' [20231003][All][HY] Just test WalkingZ fail pins
' [20231228][T-All][Tank] Check 1st digital pin NC
Public Function PPMU_Continuity(digital_pins As String, force_i As Double, LowLimit As Double, hiLimit As Double, TestLimitMode As tlLimitForceResults, _
                                Optional PN_Disconnect As Boolean = False, Optional Flag_Open As String = "F_open", Optional Flag_Short As String = "F_short", _
                                Optional connect_all_pins As String, Optional LoLimit2 As Double, Optional HiLimit2 As Double, Optional isTestWalkingZResult As Boolean = False) As Long
'Public Function PPMU_Continuity(digital_pins As PinList, force_i As Double, LowLimit As Double, HiLimit As Double, TestLimitMode As tlLimitForceResults, Optional PN_Disconnect As Boolean = False, Optional Flag_Open As String = "F_open", Optional Flag_Short As String = "F_short", Optional connect_all_pins As PinList) As Long
On Error GoTo errHandler

    Dim PPMUMeasure As New PinListData
    Dim PinGroup As IPinListData
    Dim DUTPin As Variant
    Dim Pins() As String, Pin_Cnt As Long
    Dim Power_conti_volt As New PinListData
    Dim PPMUMeas_HexVs As New PinListData
    Dim i As Long
    Dim Tname As String
    Dim site As Variant
    Dim PinStr As String
    '////////////////////////////////////////////////
    Dim FlowLimitObj As IFlowLimitsInfo
    Dim Lolimit_new() As Double
    Dim HiLimit_new() As Double
    Dim Lolimit_str() As String
    Dim Hilimit_str() As String
    
    Dim SiteStore As New SiteBoolean
    
    Dim pinsequence As Integer: pinsequence = 0
    Dim CustomizeTname As String
    Dim PinNameCustomize As String
    Dim TmpTname As String
    
    Dim b_isGetFailPin As Boolean
    Dim s_ErrorMsg As String
    Dim s_TestFailPin As String
    Dim s_MeasurePin As String
    Dim s_TmpDigital As String
    
    TheExec.DataManager.DecomposePinList digital_pins, Pins(), Pin_Cnt
    TheHdw.Digital.ApplyLevelsTiming True, True, True, tlPowered
    
    Set GlbCustomizeSet = New CustomizeSetting
    If gB_Enable_AP Then CustomizeTname = GlbCustomizeSet.ReportTname(Type_AP, Type_Conti, digital_pins)
    If gB_Enable_LCD Then CustomizeTname = GlbCustomizeSet.ReportTname(Type_LCD, Type_Conti, , force_i)
    If gB_Enable_RF Then CustomizeTname = GlbCustomizeSet.ReportTname(Type_RF, Type_Conti)

    
'''    '' ------  20191218 -- JIM,CCPANJ
'''    Dim Enable_Conti_Tname_Format As Boolean ''Continuity Tname "HAC" Format
'''    Enable_Conti_Tname_Format = True '' We can control by enable word or something else later.
'''    If Enable_Conti_Tname_Format Then
'''        Dim TnameFormatStr As String
'''        Dim Tname_Prefix As String
'''        Dim Tname_Suffix As String
'''        Call Continuity_Tname_Format(digital_pins, Tname_Prefix, Tname_Suffix)
'''    End If
'''    '' ------  20191218 -- JIM,CCPANJ
    
    ''--------------------------------------------------------------------------
    ''Get flow H/L testlimit
    ''--------------------------------------------------------------------------
    If TestLimitMode = tlForceFlow Then
        Call TheExec.Flow.GetTestLimits(FlowLimitObj)
        Call FlowLimitObj.GetLowLimits(Lolimit_str)
        Call FlowLimitObj.GetHighLimits(Hilimit_str)

        ReDim Lolimit_new(UBound(Lolimit_str))
        ReDim HiLimit_new(UBound(Hilimit_str))
        
        For i = 0 To UBound(Lolimit_str)
            If Lolimit_str(i) <> "" Then Lolimit_new(i) = CDbl(Lolimit_str(i))
        Next i
        For i = 0 To UBound(Hilimit_str)
            If Hilimit_str(i) <> "" Then HiLimit_new(i) = CDbl(Hilimit_str(i))
        Next i
    
    End If
    '////////////////////////////////////////////////
    
    '---------------------------------------------------------'
    '--- Lowlimit value should be more than ClampLow value ---'
    '--- Highlimit value should be less than ClampHigh value -'
    '---------------------------------------------------------'
    If gl_isCheckClampLimit <> ContiClampCheckType.CheckPass Then
        If True = CheckTestInst_HiLoLimit(LowLimit, hiLimit, Lolimit_new(), HiLimit_new(), Pins(), TestLimitMode) Then Exit Function
    End If
    
    'Judge just test walkingZ test fail pin or not
    If isTestWalkingZResult Then
        s_MeasurePin = digital_pins
        b_isGetFailPin = False
        s_TestFailPin = vbNullString
        
        If gldic_ComposeFailPins.Count > 0 Then
            b_isGetFailPin = isContiTestFailPinList(s_MeasurePin, gldic_ComposeFailPins, s_TestFailPin)
        Else
            b_isGetFailPin = False
        End If
    
        If b_isGetFailPin = False Then
            TheExec.Flow.TestLimit resultVal:=1, lowVal:=-1, hiVal:=-1, unit:=unitNone, Tname:=TheExec.DataManager.instancename
            s_ErrorMsg = "Doesn't get fail pins when isTestWalkingZResult = True"
            Call Print_Error_Message(Error_Warning_Info.Warning_Info, "VBT_LIB_DC_Conti", "PPMU_Continuity", s_ErrorMsg)
            Exit Function
        End If
    End If
    
    If Flag_Open Like "" Then Flag_Open = "F_open"
    If Flag_Short Like "" Then Flag_Short = "F_short"
''
''    TheExec.DataManager.DecomposePinList digital_pins, Pins(), Pin_Cnt
''    TheHdw.Digital.ApplyLevelsTiming True, True, True, tlPowered

    ''''Disconnect os_Pins Pin Electronics from Pins in order to connect PPMU's''''
    ''20210507 UltraPacPin do nothing.
    If gl_GetInstrument_Dic.Exists(LCase(Pins(0))) Then
        If UCase(gl_GetInstrument_Dic(LCase(Pins(0)))) = glbConstIns_UP2200 Or UCase(gl_GetInstrument_Dic(LCase(Pins(0)))) = glbConstIns_UP1600 Then
            TheHdw.Digital.Pins(digital_pins).Disconnect
        End If
    Else
        Call SortAllPinInstrumentType(digital_pins, sDigital_Pin:=s_TmpDigital)
        TheHdw.Digital.Pins(s_TmpDigital).Disconnect
    End If
    
    If connect_all_pins <> "" Then
        Call PPMUInit(connect_all_pins, 0#, "v")       '20230406 add sub function to init ppmu pin
    End If
    '''''' Connect all os_Pins to ppmu and ppmu force 0v for each one
    If Pin_Cnt > 0 Then
        Call PPMUInit(digital_pins, 0#, "v")       '20230406 add sub function to init ppmu pin
    End If
    
    If isTestWalkingZResult = True And s_TestFailPin <> "" Then
        TheExec.DataManager.DecomposePinList s_TestFailPin, Pins(), Pin_Cnt
    End If

    For Each DUTPin In Pins
        PPMUMeasure.AddPin (DUTPin)

        PPMUMeasure.Pins(DUTPin) = Measure_Pin_FIMV(CStr(DUTPin), force_i, PN_Disconnect)
        
        'offline mode simulation
        If TheExec.TesterMode = testModeOffline Then
            For Each site In TheExec.sites
                If LCase(TheExec.DataManager.instancename) Like "*neg*" Then PPMUMeasure.Pins(DUTPin).value(site) = -0.5 ' Use fixed value for easily checking offline mode.
                If LCase(TheExec.DataManager.instancename) Like "*pos*" Then PPMUMeasure.Pins(DUTPin).value(site) = 0.5 ' Use fixed value for easily checking offline mode.
            Next site
        End If
    Next DUTPin
        '//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
'''        Dim pinsequence As Integer
'''        pinsequence = 0

        If TestLimitMode = tlForceNone Then CustomizeTname = Replace(CustomizeTname, "_<PinSequence>", "")

        For Each DUTPin In Pins
            'TnameFormatStr = Tname_Prefix & Replace(CStr(DUTPin), "_", "") & Tname_Suffix & CStr(PinSequence) '' 20191218 -- JIM,CCPANJ
            If TheExec.DataManager.ChannelType(DUTPin) <> "N/C" Then 'if N/C jump next Pin
            
                PinNameCustomize = CStr(DUTPin)
                If gB_Enable_AP Or gB_Enable_RF Then
                    PinNameCustomize = Replace(PinNameCustomize, "_", "")
                ElseIf gB_Enable_LCD Then
                    PinNameCustomize = Replace(PinNameCustomize, "_", "-")
                End If
                
                TmpTname = Replace(CustomizeTname, "<PinName>", PinNameCustomize)
                
                ''--------------------------------------------------------------------------
                ''The test limit is from flow
                ''--------------------------------------------------------------------------
                If TestLimitMode = tlForceFlow Then
                    CustomizeTname = Replace(CustomizeTname, "<PinSequence>", CStr(pinsequence))
                    TheExec.Flow.TestLimit resultVal:=PPMUMeasure.Pins(DUTPin), scaletype:=scaleNone, unit:=unitVolt, formatStr:="%.3f", Tname:=TmpTname, ForceVal:=force_i, ForceUnit:=unitAmp, ForceResults:=tlForceFlow
                    'Judge failed open or failed short for tlForceFlow
                    Call GlbCustomizeSet.SetContiFlagOnelimit(HiLimit_new(pinsequence), Lolimit_new(pinsequence), PPMUMeasure.Pins(DUTPin), Flag_Open, Flag_Short)
                        
                    ''--------------------------------------------------------------------------
                    ''The test limit is from argument.
                    ''AP\RF - one test limit set(HiLimit\LowLimit of H\L.
                    ''LCD - one set test limit for "OPEN". one set test limit for "SHORT"
                    ''--------------------------------------------------------------------------
                ElseIf TestLimitMode = tlForceNone Then
                    TmpTname = TmpTname + "_" + CStr(pinsequence)
                    If gB_Enable_AP Or gB_Enable_RF Then
                        TheExec.Flow.TestLimit resultVal:=PPMUMeasure.Pins(DUTPin), lowVal:=LowLimit, hiVal:=hiLimit, scaletype:=scaleNone, unit:=unitVolt, formatStr:="%.3f", Tname:=TmpTname, ForceVal:=force_i, ForceUnit:=unitAmp, ForceResults:=tlForceNone
                        'Judge failed open or failed short for tlForceNone
                        Call GlbCustomizeSet.SetContiFlagOnelimit(hiLimit, LowLimit, PPMUMeasure.Pins(DUTPin), Flag_Open, Flag_Short)
                        
                    ElseIf gB_Enable_LCD Then
                        ''Open case
                        Tname = Replace(TmpTname, "<OpenShort>", "OPEN")
                        If Name_Flag = 1 Then Tname = Replace(Tname, "OPEN", "OPEN-EOP")
                        TheExec.Flow.TestLimit resultVal:=PPMUMeasure.Pins(DUTPin), lowVal:=LowLimit, hiVal:=hiLimit, scaletype:=scaleNone, unit:=unitVolt, formatStr:="%.3f", Tname:=Tname, ForceVal:=force_i, ForceUnit:=unitAmp, ForceResults:=tlForceNone
                        Call GlbCustomizeSet.SetContiFlagTwoLimit(hiLimit, LowLimit, PPMUMeasure.Pins(DUTPin), Flag_Open)
                    End If
                End If
                
            End If

            pinsequence = pinsequence + 1
        Next DUTPin
        '//////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
        

    
    If gB_Enable_LCD And TestLimitMode = tlForceNone Then
        For Each DUTPin In Pins
            If TheExec.DataManager.ChannelType(DUTPin) <> "N/C" Then 'if N/C jump next Pin
                PinNameCustomize = CStr(DUTPin)
                If gB_Enable_AP Or gB_Enable_RF Then
                    PinNameCustomize = Replace(PinNameCustomize, "_", "")
                ElseIf gB_Enable_LCD Then
                    PinNameCustomize = Replace(PinNameCustomize, "_", "-")
                End If
                    
                TmpTname = Replace(CustomizeTname, "<PinName>", PinNameCustomize)
                ''Short case
                Tname = Replace(TmpTname, "<OpenShort>", "SHORT")
                If Name_Flag = 1 Then Tname = Replace(Tname, "SHORT", "SHORT-EOP")
                    TheExec.Flow.TestLimit resultVal:=PPMUMeasure.Pins(DUTPin), lowVal:=LoLimit2, hiVal:=HiLimit2, scaletype:=scaleNone, unit:=unitVolt, formatStr:="%.3f", Tname:=Tname, ForceVal:=force_i, ForceUnit:=unitAmp, ForceResults:=tlForceNone
                Call GlbCustomizeSet.SetContiFlagTwoLimit(HiLimit2, LoLimit2, PPMUMeasure.Pins(DUTPin), Flag_Short)

            End If
            pinsequence = pinsequence + 1
        Next DUTPin
    End If
    If connect_all_pins <> "" Then
        With TheHdw.PPMU.Pins(connect_all_pins)
            '.ForceV 0#
            .Gate = tlOff
            .Disconnect
        End With
    End If
    
    DebugPrintFunc ""
    
    Exit Function
errHandler:
        Call Print_Error_Message(Error_Info, "VBT_LIB_DC_Conti", "PPMU_Continuity")
        If AbortTest Then Exit Function Else Resume Next
End Function

' [20230620][T-All][KC] mark DCVI tlAlarmOff
' [20230620][T-Tah] DCVI disable .mode = tlDCVIModeVoltage
' [20231003][All][Tank] modify after Chihome review
' [20240315][T-All][Clyde] add force 0 voltage before measuring, and disconnect after measuring
Public Function UVI80_Continuity(digital_pins As String, force_i As Double, LowLimit As Double, hiLimit As Double, TestLimitMode As tlLimitForceResults, Optional PN_Disconnect As Boolean = False, _
                                        Optional Separate_limit As Boolean = False, Optional LowLimit2 As Double, Optional HiLimit2 As Double, Optional connect_all_pins As String, _
                                        Optional Flag_Open As String = "F_open", Optional Flag_Short As String = "F_short") As Long
On Error GoTo errHandler 'Add ErrHandler 2023/08/18
    Dim PPMUMeasure As New PinListData
    Dim PinGroup As IPinListData
    Dim DUTPin As Variant
    Dim Pins() As String, Pin_Cnt As Long
    Dim Power_conti_volt As New PinListData
    Dim PPMUMeas_HexVs As New PinListData
    Dim i As Long
    Dim Tname As String
    Dim site As Variant
    Dim pinsequence As Long
    
    Dim CustomizeTname As String
    Dim Hlimit As Double, Llimit As Double
    Dim PinNameCustomize As String
    Dim TmpTname As String
    Dim instSet As New InstrumentUtility
    
    If Flag_Short = "" And force_i > 0 Then Flag_Short = "F_short"
    If Flag_Open = "" And force_i > 0 Then Flag_Open = "F_open"
    If Flag_Short = "" And force_i < 0 Then Flag_Short = "F_Neg_short"
    If Flag_Open = "" And force_i < 0 Then Flag_Open = "F_Neg_open"

'''    On Error GoTo errHandler
    
    TheHdw.Digital.ApplyLevelsTiming True, True, True, tlPowered
    TheHdw.Wait 0.001
    
    ''--------------------------------------------------------------------------
    ''Customize datalog format setting
    ''--------------------------------------------------------------------------
    Set GlbCustomizeSet = New CustomizeSetting
    If gB_Enable_AP Then CustomizeTname = GlbCustomizeSet.ReportTname(Type_AP, Type_Conti, CStr(digital_pins))
    If gB_Enable_LCD Then CustomizeTname = GlbCustomizeSet.ReportTname(Type_LCD, Type_Conti, , force_i)
    If gB_Enable_RF Then CustomizeTname = GlbCustomizeSet.ReportTname(Type_RF, Type_Conti)


    '*************   Test digital channel continuity  ******************************
    Call PPMUInit(connect_all_pins, 0#, "v")       '20230406 add sub function to init ppmu pin
    instSet.Initialize digital_pins
    ' force 0 voltage and connect gate on for measured pin
    instSet.ApplyPower 0, fVoltage, gOn, Connect
    
    TheExec.DataManager.DecomposePinList digital_pins, Pins(), Pin_Cnt

    For Each DUTPin In Pins
        PPMUMeasure.AddPin (DUTPin)

        'FIMV and meter read
        If gl_GetInstrument_Dic.Exists(LCase(DUTPin)) Then
            PPMUMeasure.Pins(DUTPin) = Measure_Pin_FIMV(CStr(DUTPin), force_i)
        End If

        
        'offline mode simulation
        If TheExec.TesterMode = testModeOffline Then
            For Each site In TheExec.sites
                If LCase(TheExec.DataManager.instancename) Like "*neg*" Then PPMUMeasure.Pins(DUTPin).value(site) = -0.5 + Rnd() * 0.1
                If LCase(TheExec.DataManager.instancename) Like "*pos*" Then PPMUMeasure.Pins(DUTPin).value(site) = 0.5 + Rnd() * 0.1
            Next site
        End If
    Next DUTPin
    
    ' Disconnect measured pins
    instSet.ApplyPower 0, fNone, gOff, Disconnect
    If connect_all_pins <> "" Then
        With TheHdw.PPMU.Pins(connect_all_pins)
            '.ForceV 0#
            .Gate = tlOff
            .Disconnect
        End With
    End If
    
    If TestLimitMode = tlForceNone Then CustomizeTname = Replace(CustomizeTname, "_<PinSequence>", "")

        pinsequence = 0
        For Each DUTPin In Pins
            If gl_GetInstrument_Dic.Exists(LCase(DUTPin)) Then
                PinNameCustomize = CStr(DUTPin)
                If gB_Enable_AP Or gB_Enable_RF Then
                    PinNameCustomize = Replace(PinNameCustomize, "_", "")
                ElseIf gB_Enable_LCD Then
                    PinNameCustomize = Replace(PinNameCustomize, "_", "-")
                End If
            
                TmpTname = Replace(CustomizeTname, "<PinName>", PinNameCustomize)
                Tname = TmpTname + "_" + CStr(pinsequence)
                ''--------------------------------------------------------------------------
                ''The test limit is from flow
                ''--------------------------------------------------------------------------
                If TestLimitMode = tlForceFlow Then
                
                    TheExec.Flow.TestLimit resultVal:=PPMUMeasure.Pins(DUTPin), scaletype:=scaleNone, unit:=unitVolt, formatStr:="%.3f", Tname:=CustomizeTname, ForceVal:=force_i, ForceUnit:=unitAmp, ForceResults:=tlForceFlow
            
                ''--------------------------------------------------------------------------------------------------------------------
                ''The test limit is from argument.
                ''AP  - one limit set(LowLimit/HiLimit) for Flag limitcheck. TestLimit uses another one limit set(HiLimit2/LowLimit2).
                ''RF  - one limit set(LowLimit/HiLimit) for testlimit and flag check.
                ''LCD - one set test limit for "OPEN". one set test limit for "SHORT"
                ''--------------------------------------------------------------------------------------------------------------------
                ElseIf TestLimitMode = tlForceNone Then
                    If gB_Enable_AP Or gB_Enable_RF Then
                        TheExec.Flow.TestLimit resultVal:=PPMUMeasure.Pins(DUTPin), lowVal:=LowLimit, hiVal:=hiLimit, scaletype:=scaleNone, unit:=unitVolt, formatStr:="%.3f", Tname:=Tname, ForceVal:=force_i, ForceUnit:=unitAmp, ForceResults:=tlForceNone
                        'Judge failed open or failed short for tlForceNone
                        Call GlbCustomizeSet.SetContiFlagOnelimit(hiLimit, LowLimit, PPMUMeasure.Pins(DUTPin), Flag_Open, Flag_Short)
                    
                    ElseIf gB_Enable_LCD Then
                        ''Open case
                        Tname = Replace(TmpTname, "<OpenShort>", "OPEN")
                        TheExec.Flow.TestLimit resultVal:=PPMUMeasure.Pins(DUTPin), lowVal:=LowLimit, hiVal:=hiLimit, scaletype:=scaleNone, unit:=unitVolt, formatStr:="%.3f", Tname:=Tname, ForceVal:=force_i, ForceUnit:=unitAmp, ForceResults:=tlForceNone
                        Call GlbCustomizeSet.SetContiFlagTwoLimit(hiLimit, LowLimit, PPMUMeasure.Pins(DUTPin), Flag_Open)

                        ''Short case
                        Tname = Replace(TmpTname, "<OpenShort>", "SHORT")
                        TheExec.Flow.TestLimit resultVal:=PPMUMeasure.Pins(DUTPin), lowVal:=LowLimit2, hiVal:=HiLimit2, scaletype:=scaleNone, unit:=unitVolt, formatStr:="%.3f", Tname:=Tname, ForceVal:=force_i, ForceUnit:=unitAmp, ForceResults:=tlForceNone
                        Call GlbCustomizeSet.SetContiFlagTwoLimit(HiLimit2, LowLimit2, PPMUMeasure.Pins(DUTPin), Flag_Short)

                    End If
                End If

                '''-----------------UF-----------------
                If glb_TesterType = "Jaguar" Then
                    'reset alarm
                    TheHdw.DCVI.Pins(DUTPin).Alarm(tlDCVIAlarmOverRange) = tlAlarmForceFail
                    TheHdw.DCVI.Pins(DUTPin).Alarm(tlDCVIAlarmMode) = tlAlarmForceFail
                    TheHdw.DCVI.Pins(DUTPin).Alarm(tlDCVIAlarmCapture) = tlAlarmForceFail
                End If
                '''-----------------UF-----------------
     
            End If
            pinsequence = pinsequence + 1
       
        Next DUTPin

    DebugPrintFunc ""

Exit Function
errHandler: 'Add ErrHandler 2023/08/18
    Call Print_Error_Message(Error_Info, "VBT_LIB_DC_Conti", "UVI80_Continuity") 'Add ErrHandler 2023/08/18
    If isDebugMode Then TheExec.AddOutput "Error in Continuity"
    If AbortTest Then Exit Function Else Resume Next 'Add ErrHandler 2023/08/18
End Function

' [20231003][All][HY] Just test WalkingZ fail pins
Public Function Conti_WalkingZ(patset As Pattern, digital_pins As PinList, PN_Disconnect As String)
On Error GoTo errHandler 'Add ErrHandler 2023/08/18
    Dim FailPinsGroup() As String
    Dim pin As Variant
    Dim site As Variant
    TheHdw.PPMU.Pins(digital_pins).Disconnect
    
    TheHdw.Digital.ApplyLevelsTiming True, True, True, tlPowered
    TheHdw.Wait 0.001
    If PN_Disconnect <> "" Then TheHdw.Digital.Pins(PN_Disconnect).Disconnect
    TheHdw.Wait 0.002
'' float Pin before fuction run
    If (digital_pins <> "") Then
'        Thehdw.Patterns(PatSet).Load
        TheHdw.Patterns(patset).test pfAlways, 0
    End If
   
    Dim WalkingZFailPins As String
        
    For Each site In TheExec.sites
        
        FailPinsGroup = TheHdw.Digital.FailedPins(site)
        For Each pin In FailPinsGroup
            TheExec.Flow.TestLimit 1, 0, 0, PinName:=pin
            
            If gldic_ComposeFailPins.Exists(LCase(pin)) = False Then
                gldic_ComposeFailPins.Add LCase(pin), 0
            End If
        Next pin
        
    Next site
    
    DebugPrintFunc patset.value

Exit Function 'Add ErrHandler 2023/08/18
errHandler: 'Add ErrHandler 2023/08/18
    Call Print_Error_Message(Error_Info, "VBT_LIB_DC_Conti", "Conti_WalkingZ") 'Add ErrHandler 2023/08/18
    If isDebugMode Then TheExec.AddOutput "Error in DC_Conti_pattern"
    If AbortTest Then Exit Function Else Resume Next 'Add ErrHandler 2023/08/18
End Function

' [20231003][All][Tank] modify after Chihome review
' [20231003][All][HY] Add Parallel test
' [20231003][All][HY] Modify sense pin name to input by each power pin
Public Function PowerSensePins_continuity(PowerPins As PinList, LowLimit As Double, hiLimit As Double, Force_V As Double, SensePins As String, Optional isUseParallel As Boolean = False) As Long
    On Error GoTo errHandler

    Dim sFuncName As String:: sFuncName = "PowerSensePins_continuity"
    Dim ResultPower As New PinListData
    Dim i As Integer
    Dim sa_PowerPinsAry() As String
    Dim sa_SensePinsAry() As String
    Dim digital_pins_array() As String
    Dim Ts As Variant
    Dim PowerSeqNum As Integer
    Dim DigitalSeqNum As Integer
    Dim seqnum As Integer
    Dim seqnum_check As Integer
    Dim n_SensePin_Cnt As Long
    Dim n_PowerPin_Cnt As Long
    Dim TestNameInput As String
    Dim CustomizeTname As String
    Dim site As Variant 'Carter, 20240304
    TheHdw.Digital.ApplyLevelsTiming True, True, True, tlPowered
    TheExec.DataManager.DecomposePinList PowerPins, sa_PowerPinsAry(), n_PowerPin_Cnt ''160725 Turn string to Pinlist
    TheExec.DataManager.DecomposePinList SensePins, sa_SensePinsAry(), n_SensePin_Cnt
    
    'Judge measure pin exist
    If (n_PowerPin_Cnt < 1) Or (n_SensePin_Cnt < 1) Then
        TheExec.Flow.TestLimit resultVal:=-1, lowVal:=0, hiVal:=1, unit:=unitNone, Tname:=TheExec.DataManager.instancename
        Call Print_Error_Message(Warning_Info, "VBT_LIB_DC_Conti", sFuncName, "Power pin or sense pin count = 0 !!")
        Exit Function
    End If
    
    '''Power pins connect ground'''
    TheHdw.DCVS.Pins(PowerPins).Voltage.Main = 0

    '''Serial Test Method'''
    If False = isUseParallel Then
        '''The numbers of Power Pins and Measure pins should be the same by serial test method'''
        If n_PowerPin_Cnt <> n_SensePin_Cnt Then
            TheExec.Flow.TestLimit resultVal:=-1, lowVal:=0, hiVal:=1, unit:=unitNone, Tname:=TheExec.DataManager.instancename
            Call Print_Error_Message(Warning_Info, "VBT_LIB_DC_Conti", sFuncName, "The Numbers of Power(" & n_PowerPin_Cnt & " pins" & ")" & " and Measure" & "(" & n_SensePin_Cnt & " pins" & ")" & " Pins aren't the same!!")
            TheExec.Datalog.WriteComment "The number of Power pins: " & n_PowerPin_Cnt & " pins" & " ,The number of Sense pins: "
            TheExec.Datalog.WriteComment "The numbers of Power Pins and Measure pins should be the same by serial test method."
            Exit Function
        End If

        For i = 0 To n_PowerPin_Cnt - 1
        
            Call PPMUInit(sa_SensePinsAry(i), 0#, "i")       '20230406 add sub function to init ppmu pin
            
            TheHdw.DCVS.Pins(sa_PowerPinsAry(i)).Voltage.Main = Force_V

            TheHdw.Wait 0.05

            ResultPower.AddPin (sa_SensePinsAry(i))
            ResultPower.Pins(sa_SensePinsAry(i)).value = TheHdw.PPMU.Pins(sa_SensePinsAry(i)).Read
            '''offline mode simulation'''
            If TheExec.TesterMode = testModeOffline Then
                For Each site In TheExec.sites
                    ResultPower.Pins(sa_SensePinsAry(i)).value(site) = hiLimit - Rnd() * 0.01
                Next site
            End If
            TheHdw.DCVS.Pins(sa_PowerPinsAry(i)).Voltage.Main = 0
            TheHdw.PPMU.Pins(sa_SensePinsAry(i)).Gate = tlOff
            TheHdw.PPMU.Pins(sa_SensePinsAry(i)).Disconnect
        Next i
    '''Parallel Test Method'''
    Else
        Call PPMUInit(SensePins, 0#, "i")
        
        TheHdw.DCVS.Pins(PowerPins).Voltage.Main = Force_V

        TheHdw.Wait 0.05

        ResultPower = TheHdw.PPMU.Pins(SensePins).Read
        '''offline mode simulation'''
        If TheExec.TesterMode = testModeOffline Then
            For Each site In TheExec.sites
                ResultPower = hiLimit - Rnd() * 0.01
            Next site
        End If
        TheHdw.DCVS.Pins(PowerPins).Voltage.Main = 0
        TheHdw.PPMU.Pins(SensePins).Disconnect
        TheHdw.PPMU.Pins(SensePins).Gate = tlOff
    End If

    '''-----------------------HAC Printing Format-----------------------'''
    Set GlbCustomizeSet = New CustomizeSetting
    If gB_Enable_AP Then CustomizeTname = GlbCustomizeSet.ReportTname(Type_AP, Type_Conti, "")
    If gB_Enable_LCD Then CustomizeTname = GlbCustomizeSet.ReportTname(Type_LCD, Type_Conti)
    If gB_Enable_RF Then CustomizeTname = GlbCustomizeSet.ReportTname(Type_RF, Type_Conti)
    CustomizeTname = Replace(CustomizeTname, "_X_X_<PinSequence>", "FORCEVDD_X_X")
    CustomizeTname = Replace(CustomizeTname, "<MeasureCase>", "PowerSense")

    For Each Ts In sa_SensePinsAry
        TestNameInput = Replace(CustomizeTname, "<PinName>", Replace(CStr(Ts), "_", ""))
        TestNameInput = TestNameInput + "_0"
        'TestNameInput = Report_TName_From_Instance("V", CStr(Ts), "DC_Continuity_PowerSense", , 0, ForceResult:=tlForceNone)
        TheExec.Flow.TestLimit ResultPower.Pins(Ts), LowLimit, hiLimit, , , , unitVolt, , ForceVal:=Force_V, Tname:=TestNameInput, ForceResults:=tlForceNone
    Next Ts
    '''-----------------------HAC Printing Format-----------------------'''
    DebugPrintFunc ""                                        ' add for Miner 20151103
    
    Exit Function
errHandler: 'Add ErrHandler 2023/08/18
    Call Print_Error_Message(Error_Info, "VBT_LIB_DC_Conti", "PowerSensePins_continuity") 'Add ErrHandler 2023/08/18
    If AbortTest Then Exit Function Else Resume Next 'Add ErrHandler 2023/08/18
End Function


' [20230407][RF] Add DCVI power short
' [20231003][All][Tank] modify after Chihome review
' [20231003][All][CP] Give each pin group their own limit
' [20231016][All][Tank] Use ApplyLevelsTiming to set iFold limit
' [20231108][All][Clyde] remove duplicate gate off code
' [20231124][All][Tank] Fix use flow limit no need check limit count in argument
' [20231124][All][Tank] Fix use flow limit measure pin group get same limit to set currentrange
' [20240110][All][Tank] modify fixcurrentrange argument
Public Function p2p_short_Power_FVMI(allPowerPins As String, _
                     ForceCondition As String, _
                     LowLimit As String, _
                     hiLimit As String, _
                     TestLimitMode As tlLimitForceResults, _
                     FlowLimitForInitIRange As Boolean, _
                     digital_pins As PinList, _
                     Optional FixCurrentRange As String) As Long
On Error GoTo errHandler
   ''++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
    'Testing Method:  Force 0.1V , measure smaller than 199ma,set clamp to 200ma, if higher than 199 ma then fail
    ''+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
    Dim sFuncName As String:: sFuncName = "p2p_short_Power_FVMI"
    Dim PowerPins As String
    Dim p As Variant, Pin_Ary() As String, Pin_dcvi_Ary() As String
    Dim Tname As String
    Dim TempString As String
    Dim PowerSequence As Double
    Dim site As Variant
    Dim pin As New PinList
    Dim FoldLimit As Double
    Dim power_pin_ary() As String
    
    Dim MaxCurr As Double
    Dim MeasRangeVal As Double
    Dim i As Integer, StepNo As Integer, j As Integer, Stop_Step As Integer
    
    Dim pinsequence As Long: pinsequence = 0
    
    Dim FlowLimitsInfo As IFlowLimitsInfo

    Dim val As Double
    Dim PowerPinGroup() As String
    Dim Val_Hi() As String
    Dim Val_Lo() As String
    Dim Low() As String
    Dim High() As String
    Dim Pins() As String, Pin_Cnt As Long
    Dim LowValue As Double
    Dim HighValue As Double
    Dim Lolimit_new() As Double
    Dim HiLimit_new() As Double
    Dim GroupNum As Long: GroupNum = 0
    Dim LoopIndex As Long
    Dim Power_pins As New PinList

    Dim Merge_Type, Slot_Type As String
    Dim A_Slot_Type As String
    Dim Split_Ary() As String
    Dim SattleTime As Double
    Dim WaitTime As Double
    
    Dim p_hexvs As String
    Dim p_uvs As String
    Dim p_dc07 As String
    Dim p_vsm As String
    Dim p_uvs64 As String
    Dim p_uvs256hp As String      '''''UFP_Corr
    Dim p_dc30 As String    'Tank add DC-30
    
    Dim HexVS_Power_data As New PinListData
    Dim UVS_Power_data As New PinListData
    Dim VSM_Power_data As New PinListData
    Dim UVS64_Power_data As New PinListData
    Dim UVS256hp_Power_data As New PinListData      '''''UFP_Corr
    Dim DC07_Power_data As New PinListData
    Dim DC30_Power_data As New PinListData  'Tank add DC-30
    
    Dim CustomizeTname As String
    Dim TmpTname As String
    Dim PinNameCustomize As String
    Dim ForceValtage() As String
    Dim ForceValue As New Dictionary
    Dim mLoop As Long
    Dim ForceV As Double
    Dim StepVoltageMode As Boolean: StepVoltageMode = False
    
    Dim sPin As String
    
    Dim d_OrgSource_PinCurrentRange() As New SiteDouble
    
    Dim s_DCVS_PinList As String
    Dim s_DCVS_PinAry() As String
    
    Dim s_ErrorMsg As String
    Dim sa_AllPinAry() As String
    Dim n_AllPinCnt As Long
    
    Dim s_TempAllMeasPin As String
    Dim s_DCVS_Pins As String
    
    Dim b_isCountTheSame As Boolean
    Dim n_FlowLimitIndex As Long
    
    Dim dict_FixedCurrent As New Dictionary
    
    ''--------------
    ''Datalog Format
    ''--------------
    Set GlbCustomizeSet = New CustomizeSetting
    If gB_Enable_AP Then CustomizeTname = GlbCustomizeSet.ReportTname(Type_AP, Type_Conti, CStr(digital_pins))
    If gB_Enable_LCD Then CustomizeTname = GlbCustomizeSet.ReportTname(Type_LCD, Type_Conti)
    If gB_Enable_RF Then CustomizeTname = GlbCustomizeSet.ReportTname(Type_RF, Type_Conti)
    CustomizeTname = Replace(CustomizeTname, "_<PinSequence>", "")
    CustomizeTname = Replace(CustomizeTname, "_<ForceI>", "")
    
    ''-----------------------
    ''Get H/L limit from flow
    ''-----------------------
    If (FlowLimitForInitIRange = True Or TestLimitMode = tlForceFlow) Then
        Call TheExec.Flow.GetTestLimits(FlowLimitsInfo)
        Call FlowLimitsInfo.GetLowLimits(Val_Lo)
        Call FlowLimitsInfo.GetHighLimits(Val_Hi)

        ReDim Lolimit_new(UBound(Val_Lo))
        ReDim HiLimit_new(UBound(Val_Hi))
        For i = 0 To UBound(Val_Lo)
            If Val_Lo(i) <> "" Then Lolimit_new(i) = CDbl(Val_Lo(i))
        Next i
        For i = 0 To UBound(Val_Hi)
            If Val_Hi(i) <> "" Then HiLimit_new(i) = CDbl(Val_Hi(i))
        Next i
    End If
    
    s_TempAllMeasPin = Replace(allPowerPins, ";", ",")
    PowerPinGroup = Split(allPowerPins, ";")
    ForceValtage = Split(ForceCondition, ";")
    Low = Split(LowLimit, ";")
    High = Split(hiLimit, ";")
    b_isCountTheSame = False
    
    Call FixCurrentRange_StrToDic(FixCurrentRange, dict_FixedCurrent)

    ''-----------------------------------------------------------------------------
    ''The Numbers of PinGroup / ForceValue / HighLimit / LowLimit must be the same.
    ''If their numbers are not the same, it will binout here.
    ''-----------------------------------------------------------------------------
    If UBound(PowerPinGroup) = UBound(ForceValtage) Then
        If (TestLimitMode = tlForceFlow) Then
            GroupNum = UBound(PowerPinGroup)
            b_isCountTheSame = True
        Else
            If UBound(ForceValtage) = UBound(Low) And UBound(Low) = UBound(High) Then
                GroupNum = UBound(PowerPinGroup)
                b_isCountTheSame = True
            End If
        End If
    End If
    If b_isCountTheSame = False Then
        TheExec.Flow.TestLimit resultVal:=1, lowVal:=-1, hiVal:=-1, unit:=unitNone, Tname:=TheExec.DataManager.instancename
        s_ErrorMsg = "The Numbers of PinGroup / ForceValue / HighLimit / LowLimit aren't the same!!"
        Call Print_Error_Message(Error_Warning_Info.Warning_Info, "VBT_LIB_DC_Conti", sFuncName, s_ErrorMsg)
        Exit Function
    End If
    
    TheExec.DataManager.DecomposePinList s_TempAllMeasPin, sa_AllPinAry, n_AllPinCnt
    
    For Each p In sa_AllPinAry
        sPin = LCase(p)
        If gl_GetInstrument_Dic.Exists(sPin) Then
            A_Slot_Type = UCase(gl_GetInstrument_Dic(sPin))
            'classfy power pin
            Select Case A_Slot_Type
                Case glbConstIns_HEXVS:       p_hexvs = CombineStringList(p_hexvs, sPin)
                Case glbConstIns_VSM:         p_vsm = CombineStringList(p_vsm, sPin)
                Case glbConstIns_VHDVS:       p_uvs = CombineStringList(p_uvs, sPin)
                Case glbConstIns_VS5A:        p_uvs64 = CombineStringList(p_uvs64, sPin)
                Case glbConstIns_VS800MA:     p_uvs256hp = CombineStringList(p_uvs256hp, sPin)
                Case glbConstIns_DC07:        p_dc07 = CombineStringList(p_dc07, sPin)
                Case glbConstIns_DC30:        p_dc30 = CombineStringList(p_dc30, sPin)
            End Select
        End If
    '''loop PowerPinGrp'''
    Next p
    
    ''-----------------------------
    '' CurrentLimitBehaviorGateOff
    ''-----------------------------
    If p_hexvs <> "" Then s_DCVS_Pins = CombineStringList(s_DCVS_Pins, p_hexvs)
    If p_vsm <> "" Then s_DCVS_Pins = CombineStringList(s_DCVS_Pins, p_vsm)
    If p_uvs <> "" Then s_DCVS_Pins = CombineStringList(s_DCVS_Pins, p_uvs)
    If p_uvs64 <> "" Then s_DCVS_Pins = CombineStringList(s_DCVS_Pins, p_uvs64)
    If p_uvs256hp <> "" Then s_DCVS_Pins = CombineStringList(s_DCVS_Pins, p_uvs256hp)
    
    If s_DCVS_Pins <> "" Then
        TheHdw.DCVS.Pins(s_DCVS_Pins).Gate = False
    End If
    If p_dc07 <> "" Then
        TheHdw.DCVI.Pins(p_dc07).Gate = False
    End If
    If p_dc30 <> "" Then
        TheHdw.DCVI.Pins(p_dc30).Gate = False
    End If
    
    n_FlowLimitIndex = 0
    For LoopIndex = 0 To GroupNum
        Power_pins = PowerPinGroup(LoopIndex)
        ForceV = ForceValtage(LoopIndex)
        If (TestLimitMode = tlForceFlow) Then
            LowValue = 0
            HighValue = 0
        Else
            LowValue = Low(LoopIndex)
            HighValue = High(LoopIndex)
        End If
        PowerPins = vbNullString

        Set HexVS_Power_data = Nothing
        Set UVS_Power_data = Nothing
        Set VSM_Power_data = Nothing
        Set UVS64_Power_data = Nothing
        Set UVS256hp_Power_data = Nothing
        Set DC07_Power_data = Nothing
        Set DC30_Power_data = Nothing

        TheExec.DataManager.DecomposePinList Power_pins, power_pin_ary(), Pin_Cnt
        ''---------------
        ''Remove N/C pin.
        ''---------------
        For Each p In power_pin_ary
            sPin = LCase(p)
            If gl_GetInstrument_Dic.Exists(sPin) Then       'Remove N/C pin.
                PowerPins = CombineStringList(PowerPins, sPin)
            End If
        Next p
        Pin_Ary = Split(PowerPins, ",")

        If PowerPins <> "" Then
            '---------------------------------------
            'store current range and ifold
            '---------------------------------------
            ReDim d_OrgSource_PinCurrentRange(UBound(Pin_Ary))
        
            For i = 0 To UBound(Pin_Ary)
                sPin = LCase(Pin_Ary(i))
                If LCase(gl_GetInstrumentType_Dic(sPin)) Like "*dcvs*" Then
                    d_OrgSource_PinCurrentRange(i) = TheHdw.DCVS.Pins(sPin).CurrentRange
                ElseIf UCase(gl_GetInstrument_Dic(sPin)) = glbConstIns_DC07 Or UCase(gl_GetInstrument_Dic(sPin)) = glbConstIns_DC30 Then
                    d_OrgSource_PinCurrentRange(i) = TheHdw.DCVI.Pins(sPin).CurrentRange
                End If
            Next i

            '''Only set one time'''
            If 0 = LoopIndex Then
                TheHdw.Digital.ApplyLevelsTiming True, True, True, tlPowered     'SEC DRAM
                ''''Disconnect os_Pins Pin Electronics from Pins in order to connect PPMU's''''
                If digital_pins.value <> "" Then TheHdw.Digital.Pins(digital_pins).Disconnect
                TheHdw.Wait 3 * ms
                WaitTime = 260 * us
            End If

            ''----------------
            ''Set init IRange
            ''----------------
            For i = 0 To UBound(Pin_Ary)
                A_Slot_Type = UCase(gl_GetInstrument_Dic(LCase(Pin_Ary(i))))
                
                Select Case A_Slot_Type
                    Case glbConstIns_HEXVS, glbConstIns_VSM, glbConstIns_VHDVS, glbConstIns_VS5A, glbConstIns_VS800MA:
                    
                        If dict_FixedCurrent.Exists(LCase(Pin_Ary(i))) Then
                            val = dict_FixedCurrent(LCase(Pin_Ary(i)))
                        ElseIf (FlowLimitForInitIRange = True Or TestLimitMode = tlForceFlow) Then
                            val = Abs(HiLimit_new(n_FlowLimitIndex + i))
                        Else
                            val = HighValue
                        End If
                        Call SetCurrentRange(LCase(Pin_Ary(i)), val, WaitTime)
                        
                    Case glbConstIns_DC07, glbConstIns_DC30:
                    Case Else:
                        s_ErrorMsg = "Use power pin:" & Pin_Ary(i) & "No support!!"
                        Call Print_Error_Message(Error_Warning_Info.Warning_Info, "VBT_LIB_DC_Conti", sFuncName, s_ErrorMsg)
                End Select
            Next i

            If (p_hexvs <> "" Or p_vsm <> "" Or p_uvs <> "" Or p_uvs64 <> "" Or p_uvs256hp <> "") Then   '''''UFP_Corr

                TheHdw.Wait WaitTime    '20230406 add set current range wait time

                For i = 0 To UBound(Pin_Ary)
                    ''----------------------------
                    ''force voltage and measure
                    ''----------------------------
''                    If StepVoltageMode = False Then
''                        ForceV = ForceValue(Pin_Ary(i))
''                    End If
                    TheHdw.DCVS.Pins(Pin_Ary(i)).Voltage.Main.value = ForceV
                    TheExec.Datalog.WriteComment Pin_Ary(i) & " - force value = " & CStr(ForceV)
                    
                    A_Slot_Type = UCase(gl_GetInstrument_Dic(LCase(Pin_Ary(i))))
                    If (A_Slot_Type = glbConstIns_HEXVS) Or (A_Slot_Type = glbConstIns_VSM) Or (A_Slot_Type = glbConstIns_VS5A) Or (A_Slot_Type = glbConstIns_VS800MA) Then  '''''UFP_Corr
                        If TheHdw.DCVS.Pins(Pin_Ary(i)).Meter.CurrentRange = 0.01 Then
                            TheHdw.Wait 100 * ms
                        Else
                            TheHdw.Wait 30 * ms
                        End If
                    Else
                        TheHdw.Wait 5 * ms  'align Cyprus TTR
                    End If

                    'classfy power pin
                    Select Case A_Slot_Type
                        Case glbConstIns_HEXVS:
                            HexVS_Power_data.AddPin Pin_Ary(i)
                            HexVS_Power_data.Pins(Pin_Ary(i)) = TheHdw.DCVS.Pins(Pin_Ary(i)).Meter.Read(tlStrobe, 10, , tlDCVSMeterReadingFormatAverage)
                        Case glbConstIns_VSM:
                            VSM_Power_data.AddPin Pin_Ary(i)
                            VSM_Power_data.Pins(Pin_Ary(i)) = TheHdw.DCVS.Pins(Pin_Ary(i)).Meter.Read(tlStrobe, 10, , tlDCVSMeterReadingFormatAverage)
                        Case glbConstIns_VHDVS:
                            UVS_Power_data.AddPin Pin_Ary(i)
                            UVS_Power_data.Pins(Pin_Ary(i)) = TheHdw.DCVS.Pins(Pin_Ary(i)).Meter.Read(tlStrobe, 1)
                        Case glbConstIns_VS5A:
                            UVS64_Power_data.AddPin Pin_Ary(i)
                            UVS64_Power_data.Pins(Pin_Ary(i)) = TheHdw.DCVS.Pins(Pin_Ary(i)).Meter.Read(tlStrobe, 10, , tlDCVSMeterReadingFormatAverage)
                        Case glbConstIns_VS800MA:
                            UVS256hp_Power_data.AddPin Pin_Ary(i)
                            UVS256hp_Power_data.Pins(Pin_Ary(i)) = TheHdw.DCVS.Pins(Pin_Ary(i)).Meter.Read(tlStrobe, 10, , tlDCVSMeterReadingFormatAverage)
                    End Select

                    TheHdw.DCVS.Pins(Pin_Ary(i)).Voltage.Main.value = 0#
                Next i
                ''----------------------------
                ''DataLog result
                ''----------------------------
                pinsequence = 0
                For i = 0 To UBound(Pin_Ary)
''                    If StepVoltageMode = False Then
''                        ForceV = ForceValue(Pin_Ary(i))
''                    End If
                    
                    Tname = "pwr_FVMI_" & Pin_Ary(i)
                    
                    PinNameCustomize = Pin_Ary(i)
                    If gB_Enable_AP Or gB_Enable_RF Then
                        PinNameCustomize = Replace(PinNameCustomize, "_", "")
                    ElseIf gB_Enable_LCD Then
                        PinNameCustomize = Replace(PinNameCustomize, "_", "-")
                    End If
                        
                    TmpTname = Replace(CustomizeTname, "<PinName>", PinNameCustomize)
                    If gB_Enable_AP Or gB_Enable_RF Then TmpTname = TmpTname + "_" + CStr(pinsequence)
                    
                    A_Slot_Type = UCase(gl_GetInstrument_Dic(LCase(Pin_Ary(i))))
                    ''----------------------------
                    ''compare result
                    ''----------------------------
                    Select Case A_Slot_Type
                        Case glbConstIns_HEXVS:       Call DCVS_PowerPinContiCompare(HexVS_Power_data.Pins(Pin_Ary(i)), Pin_Ary(i), TmpTname, ForceV, TestLimitMode, LowValue, HighValue)
                        Case glbConstIns_VSM:         Call DCVS_PowerPinContiCompare(VSM_Power_data.Pins(Pin_Ary(i)), Pin_Ary(i), TmpTname, ForceV, TestLimitMode, LowValue, HighValue)
                        Case glbConstIns_VHDVS:       Call DCVS_PowerPinContiCompare(UVS_Power_data.Pins(Pin_Ary(i)), Pin_Ary(i), TmpTname, ForceV, TestLimitMode, LowValue, HighValue)
                        Case glbConstIns_VS5A:        Call DCVS_PowerPinContiCompare(UVS64_Power_data.Pins(Pin_Ary(i)), Pin_Ary(i), TmpTname, ForceV, TestLimitMode, LowValue, HighValue)
                        Case glbConstIns_VS800MA:     Call DCVS_PowerPinContiCompare(UVS256hp_Power_data.Pins(Pin_Ary(i)), Pin_Ary(i), TmpTname, ForceV, TestLimitMode, LowValue, HighValue)
                    End Select
                    pinsequence = pinsequence + 1
                Next i
            ElseIf (p_dc07 <> "" Or p_dc30 <> "") Then
                If (p_dc07 <> "") Then
                    Pin_dcvi_Ary = Split(p_dc07, ",")
                    TheHdw.DCVI.Pins(p_dc07).SetCurrentAndRange 0.02, 0.02
                    FoldLimit = TheHdw.DCVI.Pins(p_dc07).Current
                    
                    With TheHdw.DCVI.Pins(p_dc07)
                        .Disconnect tlDCVIConnectDefault
                        .Gate = False
                        .mode = tlDCVIModeVoltage
                        .Voltage = 0
                        .CurrentRange.Autorange = True
                        '.CurrentRange.Value = FoldLimit
                        .VoltageRange.Autorange = True
                        .Connect tlDCVIConnectDefault
                        .Gate = True
                        .Meter.mode = tlDCVIMeterCurrent
                        .Meter.CurrentRange.value = FoldLimit
                    End With
                    
                    TheHdw.Wait 3 * ms
                    DC07_Power_data = TheHdw.DCVI.Pins(p_dc07).Meter.Read(tlStrobe, 1, , tlDCVIMeterReadingFormatAverage)
                    '=========================================================================
                    For i = 0 To DC07_Power_data.Pins.Count - 1
                        TmpTname = Replace(CustomizeTname, "<PinName>", Pin_Ary(i))
        ''                UVI80MeasRangeList = TheHdw.DCVI.Pins(DC07_Power_data.Pins(i).Name).Meter.CurrentRange.List
                        If StepVoltageMode = False Then
                            ForceV = ForceValue(Pin_dcvi_Ary(i))
                        End If
                        TheHdw.DCVI.Pins(Pin_dcvi_Ary(i)).Voltage = ForceV
                    
                        TheHdw.DCVI.Pins(DC07_Power_data.Pins(i).name).CurrentRange.Autorange = True
                        TheHdw.Wait 3 * ms
                        DC07_Power_data.Pins(i).value = TheHdw.DCVI.Pins(DC07_Power_data.Pins(i).name).Meter.Read(tlStrobe, 10, , tlDCVIMeterReadingFormatAverage)
                        TheHdw.DCVI.Pins(Pin_dcvi_Ary(i)).Voltage = 0
                        
                        'offline mode simulation
                        If TheExec.TesterMode = testModeOffline Then
                            For Each site In TheExec.sites
                                DC07_Power_data.Pins(Pin_dcvi_Ary(i)).value(site) = 0.01 + Rnd() * 0.0001
                            Next site
                        End If
                    Next i
                    
                    For i = 0 To UBound(Pin_dcvi_Ary)
                        PinNameCustomize = Pin_dcvi_Ary(i)
                        If gB_Enable_AP Or gB_Enable_RF Then
                            PinNameCustomize = Replace(PinNameCustomize, "_", "")
                        ElseIf gB_Enable_LCD Then
                            PinNameCustomize = Replace(PinNameCustomize, "_", "-")
                        End If
                            
                        TmpTname = Replace(CustomizeTname, "<PinName>", PinNameCustomize)
                        If gB_Enable_AP Or gB_Enable_RF Then TmpTname = TmpTname + "_" + CStr(pinsequence)
                        
                        If StepVoltageMode = False Then
                            ForceV = ForceValue(Pin_dcvi_Ary(i))
                        End If
                        If TestLimitMode = tlForceFlow Then
                            TheExec.Flow.TestLimit resultVal:=DC07_Power_data.Pins(Pin_dcvi_Ary(i)), scaletype:=scaleNone, unit:=unitAmp, formatStr:="%.3f", Tname:=TmpTname, _
                            ForceVal:=ForceV, ForceUnit:=unitVolt, ForceResults:=tlForceFlow
                        ElseIf TestLimitMode = tlForceNone Then
                            TheExec.Flow.TestLimit resultVal:=DC07_Power_data.Pins(Pin_dcvi_Ary(i)), lowVal:=LowLimit, hiVal:=hiLimit, scaletype:=scaleNone, unit:=unitAmp, formatStr:="%.3f", Tname:=TmpTname, _
                            ForceVal:=ForceV, ForceUnit:=unitVolt, ForceResults:=tlForceNone
                        End If
                    Next i
                    
                    TheHdw.DCVI.Pins(p_dc07).Voltage = 0
                End If
                If (p_dc30 <> "") Then
                    Pin_dcvi_Ary = Split(p_dc30, ",")
                    TheHdw.DCVI.Pins(p_dc30).SetCurrentAndRange 0.02, 0.02
                    FoldLimit = TheHdw.DCVI.Pins(p_dc30).Current
                    
                    With TheHdw.DCVI.Pins(p_dc30)
                        .Disconnect tlDCVIConnectDefault
                        .Gate = False
                        .mode = tlDCVIModeVoltage
                        .Voltage = 0
                        .CurrentRange.Autorange = True
                        .VoltageRange.Autorange = True
                        .Connect tlDCVIConnectDefault
                        .Gate = True
                        .Meter.mode = tlDCVIMeterCurrent
                        .Meter.CurrentRange.value = FoldLimit
                    End With
                    
                    TheHdw.Wait 3 * ms
                    DC30_Power_data = TheHdw.DCVI.Pins(p_dc30).Meter.Read(tlStrobe, 1, , tlDCVIMeterReadingFormatAverage)
                    
                    For i = 0 To DC30_Power_data.Pins.Count - 1
                        If StepVoltageMode = False Then
                            ForceV = ForceValue(Pin_dcvi_Ary(i))
                        End If
                        TheHdw.DCVI.Pins(Pin_dcvi_Ary(i)).Voltage = ForceV
                        
                        TheHdw.DCVI.Pins(Pin_dcvi_Ary(i)).CurrentRange.Autorange = True
                        TheHdw.Wait 3 * ms
                        
                        DC30_Power_data.Pins(i).value = TheHdw.DCVI.Pins(Pin_dcvi_Ary(i)).Meter.Read(tlStrobe, 10, , tlDCVIMeterReadingFormatAverage)
                        TheHdw.DCVI.Pins(Pin_dcvi_Ary(i)).Voltage = 0
                        
                        If TheExec.TesterMode = testModeOffline Then
                            For Each site In TheExec.sites
                                DC30_Power_data.Pins(Pin_dcvi_Ary(i)).value(site) = 0.01 + Rnd() * 0.0001
                            Next site
                        End If
                    Next i
                    
                    For i = 0 To UBound(Pin_dcvi_Ary)
                        PinNameCustomize = Pin_dcvi_Ary(i)
                        If gB_Enable_AP Or gB_Enable_RF Then
                            PinNameCustomize = Replace(PinNameCustomize, "_", "")
                        ElseIf gB_Enable_LCD Then
                            PinNameCustomize = Replace(PinNameCustomize, "_", "-")
                        End If
                            
                        TmpTname = Replace(CustomizeTname, "<PinName>", PinNameCustomize)
                        If gB_Enable_AP Or gB_Enable_RF Then TmpTname = TmpTname + "_" + CStr(pinsequence)
                        
                        If StepVoltageMode = False Then
                            ForceV = ForceValue(Pin_dcvi_Ary(i))
                        End If
                        If TestLimitMode = tlForceFlow Then
                            TheExec.Flow.TestLimit resultVal:=DC30_Power_data.Pins(Pin_dcvi_Ary(i)), scaletype:=scaleNone, unit:=unitAmp, formatStr:="%.3f", Tname:=TmpTname, _
                            ForceVal:=ForceV, ForceUnit:=unitVolt, ForceResults:=tlForceFlow
                        ElseIf TestLimitMode = tlForceNone Then
                            TheExec.Flow.TestLimit resultVal:=DC30_Power_data.Pins(Pin_dcvi_Ary(i)), lowVal:=LowLimit, hiVal:=hiLimit, scaletype:=scaleNone, unit:=unitAmp, formatStr:="%.3f", Tname:=TmpTname, _
                            ForceVal:=ForceV, ForceUnit:=unitVolt, ForceResults:=tlForceNone
                        End If
                    Next i
                    
                    TheHdw.DCVI.Pins(p_dc30).Voltage = 0
                End If
            End If
            '---------------------------------------
            'Restore current range and ifold
            '---------------------------------------
            For i = 0 To UBound(Pin_Ary)
                sPin = LCase(Pin_Ary(i))
                If LCase(gl_GetInstrumentType_Dic(sPin)) Like "*dcvs*" Then
                    TheHdw.DCVS.Pins(sPin).CurrentRange = d_OrgSource_PinCurrentRange(i)
                ElseIf UCase(gl_GetInstrument_Dic(sPin)) = glbConstIns_DC07 Or UCase(gl_GetInstrument_Dic(sPin)) = glbConstIns_DC30 Then
                    TheHdw.DCVI.Pins(sPin).CurrentRange = d_OrgSource_PinCurrentRange(i)
                End If
            Next i
        End If
        n_FlowLimitIndex = n_FlowLimitIndex + Pin_Cnt
    Next LoopIndex
    '---------------------------------------
    'store current range and ifold
    '---------------------------------------
    
    TheHdw.Digital.ApplyLevelsTiming False, True, False, tlPowered       'SEC DRAM
    DebugPrintFunc ""
    
    Exit Function

errHandler:
    Call Print_Error_Message(Error_Info, "VBT_LIB_DC_Conti", sFuncName)
    If AbortTest Then Exit Function Else Resume Next
End Function


Public Function PPMU_Continuity_IV_Curve(digital_pins As PinList, force_i_S As Double, force_i_E As Double, force_i_Step As Double, LowLimit As Double, hiLimit As Double, TestLimitMode As tlLimitForceResults, Optional Separate_limit As Boolean = False, Optional independt_meas As Boolean) As Long
On Error GoTo errHandler 'Add ErrHandler 2023/08/18
    
    Dim DUTPin As Variant
    Dim Pins() As String, Pin_Cnt As Long
    Dim i As Long
    Dim force_i As Double
    Dim PPMUMeasure As New PinListData
    Dim site As Variant

'    thehdw.DCVS.Pins("AllUvsCP,VDD_CPU").Alarm(tlDCVSAlarmAll) = tlAlarmOff 'chihome

    TheExec.DataManager.DecomposePinList digital_pins, Pins(), Pin_Cnt
    
    For Each DUTPin In Pins
        PPMUMeasure.AddPin (DUTPin)
    Next DUTPin
    
    For force_i = force_i_S To force_i_E Step (force_i_E - force_i_S) / force_i_Step
    

        Dim PinGroup As IPinListData
        Dim Power_conti_volt As New PinListData
        Dim PPMUMeas_HexVs As New PinListData
        Dim Tname As String
    
    
        TheHdw.Digital.ApplyLevelsTiming True, True, True, tlPowered
        'DisconnectVDDCA 'SEC DRAM
        TheHdw.Wait 0.001
        
        '*************   Test digital channel continuity  ******************************
        ''''Disconnect os_Pins Pin Electronics from Pins in order to connect PPMU's''''
        'TheHdw.digital.Pins("Non_Conti_IO").Disconnect
        TheHdw.Digital.Pins(digital_pins).Disconnect
        
        '''''' Connect all os_Pins to ppmu and ppmu force 0v for each one
        If independt_meas = False Then
            With TheHdw.PPMU.Pins(digital_pins)
                .ForceV 0#
                .Connect
                .Gate = tlOn
            End With
        End If
    
    
    
    '    TheExec.DataManager.DecomposePinList digital_pins, Pins(), Pin_Cnt
    
        For Each DUTPin In Pins
    '        PPMUMeasure.AddPin (DUTPin)
    
            If independt_meas = False Then
                With TheHdw.PPMU.Pins(DUTPin)
        ''            .ClampVHi = 1.2
        ''            .ClampVLo = -1
                    .ForceI (force_i)
                End With
            Else
                With TheHdw.PPMU.Pins(DUTPin)
                    .Connect
                    .ForceI (force_i)
                    .Gate = tlOn
                End With
            End If
    
    
            TheHdw.Wait 0.005
            
            DebugPrintFunc_PPMU CStr(DUTPin)
            PPMUMeasure.Pins(DUTPin) = TheHdw.PPMU.Pins(DUTPin).Read(tlPPMUReadMeasurements, 20)
            
            
            'offline mode simulation
            If TheExec.TesterMode = testModeOffline Then
                For Each site In TheExec.sites
    '                If LCase(theexec.DataManager.instanceName) Like "*neg*" Then PPMUMeasure.Pins(DUTPin).Value(Site) = -0.5 + Rnd() * 0.1
    '                If LCase(TheExec.DataManager.instanceName) Like "*pos*" Then PPMUMeasure.Pins(DUTPin).Value(Site) = 0.5 + Rnd() * 0.1
                    PPMUMeasure.Pins(DUTPin).value(site) = 0.5 - Rnd() * 0.01
                Next site
            End If
            
            
            'recover measure dut Pin to 0V before next Pin
            If independt_meas = False Then
                TheHdw.PPMU.Pins(DUTPin).ForceV 0
            Else
                TheHdw.PPMU.Pins(DUTPin).ForceI 0
                TheHdw.PPMU.Pins(DUTPin).Gate = tlOff
                TheHdw.PPMU.Pins(DUTPin).Disconnect
            End If
            'If TheExec.Sites.Active.Count = 0 Then Exit Function 'chihome
    
        Next DUTPin
    
    
        For Each DUTPin In Pins
            If TheExec.DataManager.ChannelType(DUTPin) <> "N/C" Then
                Tname = "Conti1_" & CStr(DUTPin)
                If TheExec.TesterMode = testModeOffline Then
                    TheExec.Flow.TestLimit resultVal:=PPMUMeasure.Pins(DUTPin), lowVal:=LowLimit, hiVal:=0.5, scaletype:=scaleNone, unit:=unitVolt, formatStr:="%.3f", Tname:=Tname, ForceVal:=force_i, ForceUnit:=unitAmp, ForceResults:=tlForceNone
                Else
                    If TestLimitMode = tlForceFlow Then
                            TheExec.Flow.TestLimit resultVal:=PPMUMeasure.Pins(DUTPin), scaletype:=scaleNone, unit:=unitVolt, formatStr:="%.3f", Tname:=Tname, ForceVal:=force_i, ForceUnit:=unitAmp, ForceResults:=tlForceFlow
                    Else: TestLimitMode = tlForceNone
                            TheExec.Flow.TestLimit resultVal:=PPMUMeasure.Pins(DUTPin), lowVal:=LowLimit, hiVal:=hiLimit, scaletype:=scaleNone, unit:=unitVolt, formatStr:="%.3f", Tname:=Tname, ForceVal:=force_i, ForceUnit:=unitAmp, ForceResults:=tlForceNone
                    End If
                End If
            End If

        
        Next DUTPin
        
        If Separate_limit = True Then
            For Each DUTPin In Pins
                If TheExec.DataManager.ChannelType(DUTPin) <> "N/C" Then
                    Tname = "Conti2_" & CStr(DUTPin)
                    If TheExec.TesterMode = testModeOffline Then
                        TheExec.Flow.TestLimit resultVal:=PPMUMeasure.Pins(DUTPin), lowVal:=LowLimit, hiVal:=0.5, scaletype:=scaleNone, unit:=unitVolt, formatStr:="%.3f", Tname:=Tname, ForceVal:=force_i, ForceUnit:=unitAmp, ForceResults:=tlForceNone
                    Else
                        If TestLimitMode = tlForceFlow Then
                            TheExec.Flow.TestLimit resultVal:=PPMUMeasure.Pins(DUTPin), scaletype:=scaleNone, unit:=unitVolt, formatStr:="%.3f", Tname:=Tname, ForceVal:=force_i, ForceUnit:=unitAmp, ForceResults:=tlForceFlow
                        Else
                            TestLimitMode = tlForceNone
                            TheExec.Flow.TestLimit resultVal:=PPMUMeasure.Pins(DUTPin), lowVal:=LowLimit, hiVal:=hiLimit, scaletype:=scaleNone, unit:=unitVolt, formatStr:="%.3f", Tname:=Tname, ForceVal:=force_i, ForceUnit:=unitAmp, ForceResults:=tlForceNone
                        End If
                    End If
                End If
            Next DUTPin
        End If
        
''          'initialize ppmu to suitable clamp
''        With TheHdw.PPMU.Pins("Pins_1p0v")
''            .ClampVHi = TheExec.Specs.Globals("Pins_1p0v_Vch_GLB").ContextValue
''            .ClampVLo = TheExec.Specs.Globals("Pins_1p0v_Vch_GLB").ContextValue
''        End With
''
''        With TheHdw.PPMU.Pins("Pins_1p1v")
''            .ClampVHi = TheExec.Specs.Globals("Pins_1p1v_Vch_GLB").ContextValue
''            .ClampVLo = TheExec.Specs.Globals("Pins_1p1v_Vch_GLB").ContextValue
''        End With
''
''        With TheHdw.PPMU.Pins("Pins_1p8v")
''            .ClampVHi = TheExec.Specs.Globals("Pins_1p8v_Vch_GLB").ContextValue
''            .ClampVLo = TheExec.Specs.Globals("Pins_1p8v_Vch_GLB").ContextValue
''        End With
        
        'Disconnect PPMU from digital channels
        With TheHdw.PPMU.Pins(digital_pins)
            .Gate = tlOff
            .Disconnect
        End With
                
        TheHdw.Digital.Pins(digital_pins).Connect
    
    Next force_i
    
    
Exit Function
errHandler: 'Add ErrHandler 2023/08/18
    Call Print_Error_Message(Error_Info, "VBT_LIB_DC_Conti", "PPMU_Continuity_IV_Curve") 'Add ErrHandler 2023/08/18
    If isDebugMode Then TheExec.AddOutput "Error in Continuity"
    If AbortTest Then Exit Function Else Resume Next 'Add ErrHandler 2023/08/18
End Function

' [20231003][All][Tank] modify after Chihome review
Public Function GndSensePins_continuity(PowerPins As String, digital_pins As String, LowLimit As Double, hiLimit As Double, Power_Force_V As Double, ch_force_i As Double) As Long
On Error GoTo errHandler 'Add ErrHandler 2023/08/18
        
    Dim ResultPower As New PinListData
    
    Dim power_Pins_array() As String
    Dim digital_pins_array() As String
    Dim Ts As Variant
    Dim PowerSeqNum As Long
    Dim DigitalSeqNum As Long
    Dim seqnum As Integer
    Dim seqnum_check As Integer
    Dim i As Long
        
    Dim power_sense As String

    digital_pins_array = Split(digital_pins, ",")
    DigitalSeqNum = UBound(digital_pins_array) + 1
'    DigitalSeqNum = CLng(DigitalSeqNum)
    Call Trim_NC_Pin(digital_pins_array, DigitalSeqNum)
    
    If DigitalSeqNum > 0 Then
        For i = 0 To UBound(digital_pins_array)
            digital_pins = CombineStringList(digital_pins, digital_pins_array(i))
        Next i
    Else
        digital_pins = vbNullString
    End If

    TheHdw.Digital.ApplyLevelsTiming True, True, True, tlPowered
    '////////////////////////////////////////////////////////////////////////////////////////////////////////////////
    power_Pins_array = Split(PowerPins, ",")
    PowerSeqNum = UBound(power_Pins_array)

    digital_pins_array = Split(digital_pins, ",")
    DigitalSeqNum = UBound(digital_pins_array)

    If (PowerSeqNum = DigitalSeqNum) Then
        seqnum_check = 0
    Else
        seqnum_check = 1
    End If
    
    If digital_pins <> "" Then
        Call PPMUInit(digital_pins, ch_force_i, "i")       '20230406 add sub function to init ppmu pin

        TheHdw.DCVS.Pins(PowerPins).Voltage.Main = Power_Force_V

        TheHdw.Wait 0.005
        ResultPower = TheHdw.PPMU(digital_pins).Read(tlPPMUReadMeasurements)
        TheHdw.PPMU.Pins(digital_pins).Gate = tlOff
        TheHdw.PPMU.Pins(digital_pins).Disconnect
        
        TheExec.Flow.TestLimit resultVal:=ResultPower, lowVal:=LowLimit, hiVal:=hiLimit, unit:=unitVolt, ForceVal:=ch_force_i
        DebugPrintFunc ""                                        ' add for Miner 20151103
    End If

Exit Function
errHandler: 'Add ErrHandler 2023/08/18
    Call Print_Error_Message(Error_Info, "VBT_LIB_DC_Conti", "GndSensePins_continuity") 'Add ErrHandler 2023/08/18
    If AbortTest Then Exit Function Else Resume Next 'Add ErrHandler 2023/08/18
End Function

' [20231003][All][CP] Give each pin group their own limit
' [20240110][All][Tank] modify fixcurrentrange argument
Public Function p2p_short_Power_FVMI_VI_Curve(allPowerPins As String, _
                     ForceV_S As Double, _
                     ForceV_E As Double, _
                     ForceV_Step_Count As Double, _
                     LowLimit As String, _
                     hiLimit As String, _
                     digital_pins As PinList, _
                     Optional FixCurrentRange As String) As Long
On Error GoTo errHandler 'Add ErrHandler 2023/08/18

   ''++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
    'Testing Method:  ForceV from ForceV_S to ForceV_E, stepping ForceV_Step_Count. Measure current by using  p2p_short_Power_FVMI()
    ''+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
    
    Dim ForceV As Double
    
    'Force FlowLimitForInitIRange as False.  Not to use flow limit due to unknow stepping
    
    For ForceV = ForceV_S To ForceV_E Step (ForceV_E - ForceV_S) / ForceV_Step_Count
    
        Call p2p_short_Power_FVMI(allPowerPins, CStr(ForceV), LowLimit, hiLimit, 0, False, digital_pins, FixCurrentRange)
    
    Next ForceV

Exit Function 'Add ErrHandler 2023/08/18
errHandler: 'Add ErrHandler 2023/08/18
    Call Print_Error_Message(Error_Info, "VBT_LIB_DC_Conti", "p2p_short_Power_FVMI_VI_Curve") 'Add ErrHandler 2023/08/18
    If AbortTest Then Exit Function Else Resume Next 'Add ErrHandler 2023/08/18
End Function


Public Function RetrieveDictionaryOfDiffPairs()    'Wherever the DcConti module, put it there
On Error GoTo errHandler 'Add ErrHandler 2023/08/18
    Dim Pins() As String, Pin_Cnt As Long, iPin As Long
    Dim DiffGroup  As String: DiffGroup = "All_DiffPairs"                            'T-Autogen will create it."
    TheExec.DataManager.DecomposePinList DiffGroup, Pins(), Pin_Cnt
    DicDiffPairs.RemoveAll
    If Pin_Cnt Mod 2 <> 0 Or Pin_Cnt < 1 Then GoTo errHandler
        For iPin = 0 To Pin_Cnt - 1 Step 2
            DicDiffPairs.Add LCase(CStr(Pins(iPin))), LCase(CStr(Pins(iPin + 1)))
            DicDiffPairs.Add LCase(CStr(Pins(iPin + 1))), LCase(CStr(Pins(iPin)))
        Next iPin
    
    Exit Function
errHandler: 'Add ErrHandler 2023/08/18
    Call Print_Error_Message(Error_Info, "VBT_LIB_DC_Conti", "RetrieveDictionaryOfDiffPairs") 'Add ErrHandler 2023/08/18
End Function


Public Function PPMU_Measure_Contact_Resistance_Corner_Vss(digital_pins As PinList, force_i_S As Double, force_i_E As Double, force_i_Step As Double, LowLimit As Double, hiLimit As Double, TestLimitMode As tlLimitForceResults, PairNum As Double, Optional Separate_limit As Boolean = False, Optional independt_meas As Boolean) As Long
On Error GoTo errHandler 'Add ErrHandler 2023/08/18
    
    Dim DUTPin As Variant
    Dim Pins() As String, Pin_Cnt As Long
    Dim i As Long
    Dim force_i As Double
    Dim PPMUMeasure As New PinListData
    Dim site As Variant
    Dim MeasV_Force50mA As New PinListData: Dim MeasV_Force0mA As New PinListData: Dim Calculate_Contact_R As New PinListData

'    thehdw.DCVS.pins("AllUvsCP,VDD_CPU").Alarm(tlDCVSAlarmAll) = tlAlarmOff 'chihome

    TheExec.DataManager.DecomposePinList digital_pins, Pins(), Pin_Cnt
    
    For Each DUTPin In Pins
        PPMUMeasure.AddPin (DUTPin)
        Calculate_Contact_R.AddPin (DUTPin)
        MeasV_Force50mA.AddPin (DUTPin)
        MeasV_Force0mA.AddPin (DUTPin)
    Next DUTPin
    
    For force_i = force_i_S To force_i_E Step (force_i_E - force_i_S) / force_i_Step
    

        Dim PinGroup As IPinListData
        Dim Power_conti_volt As New PinListData
        Dim PPMUMeas_HexVs As New PinListData
        Dim Tname As String

    
        TheHdw.Digital.ApplyLevelsTiming True, True, True, tlPowered
        'DisconnectVDDCA 'SEC DRAM
        TheHdw.Wait 0.001
        
        '*************   Test digital channel continuity  ******************************
        ''''Disconnect os_pins Pin Electronics from pins in order to connect PPMU's''''
        'TheHdw.digital.Pins("Non_Conti_IO").Disconnect
        TheHdw.Digital.Pins(digital_pins).Disconnect
            
        '''''' Connect all os_pins to ppmu and ppmu force 0v for each one
        If independt_meas = False Then
            With TheHdw.PPMU.Pins(digital_pins)
                .ForceV 0#
                .Connect
                .Gate = tlOn
            End With
        End If

    
    
    '    TheExec.DataManager.DecomposePinList digital_pins, Pins(), pin_cnt
    
        For Each DUTPin In Pins
    '        PPMUMeasure.AddPin (DUTPin)
    
            If independt_meas = False Then
                With TheHdw.PPMU.Pins(DUTPin)
        ''            .ClampVHi = 1.2
        ''            .ClampVLo = -1
                    .ForceI (force_i)
                End With
            Else
                With TheHdw.PPMU.Pins(DUTPin)
                    .Connect
                    .ForceI (force_i)
                    .Gate = tlOn
                End With
            End If
    
    
            TheHdw.Wait 0.005
            
            DebugPrintFunc_PPMU CStr(DUTPin)
            PPMUMeasure.Pins(DUTPin) = TheHdw.PPMU.Pins(DUTPin).Read(tlPPMUReadMeasurements, 20)
    
    '///////////////////////ZB add code///////////////////////////////////////
            For Each site In TheExec.sites
                If force_i = 0.05 Then
                    MeasV_Force50mA.Pins(DUTPin).value = PPMUMeasure.Pins(DUTPin).value
                    TheExec.Datalog.WriteComment "50mA, Pins :" & DUTPin & " site: " & site & MeasV_Force50mA.Pins(DUTPin).value
                Else
                    MeasV_Force0mA.Pins(DUTPin).value = PPMUMeasure.Pins(DUTPin).value
                    TheExec.Datalog.WriteComment "0mA, Pins :" & DUTPin & " site: " & site & MeasV_Force0mA.Pins(DUTPin).value
                End If
            Next site
    '/////////////////////////////////////////////////////////////////////////
            
            'offline mode simulation
            If TheExec.TesterMode = testModeOffline Then
                For Each site In TheExec.sites
    '                If LCase(TheExec.DataManager.instanceName) Like "*neg*" Then PPMUMeasure.Pins(DUTPin).Value(Site) = -0.5 + Rnd() * 0.1
    '                If LCase(TheExec.DataManager.instanceName) Like "*pos*" Then PPMUMeasure.Pins(DUTPin).Value(Site) = 0.5 + Rnd() * 0.1
                    MeasV_Force50mA.Pins(DUTPin).value(site) = 0.5 + Rnd() * 0.1
                    MeasV_Force0mA.Pins(DUTPin).value(site) = Rnd() * 0.1
                Next site
            End If
            
            
            'recover measure dut pin to 0V before next pin
            If independt_meas = False Then
                TheHdw.PPMU.Pins(DUTPin).ForceV 0
            Else
                TheHdw.PPMU.Pins(DUTPin).ForceI 0
                TheHdw.PPMU.Pins(DUTPin).Gate = tlOff
                TheHdw.PPMU.Pins(DUTPin).Disconnect
            End If
            'If TheExec.Sites.Active.Count = 0 Then Exit Function 'chihome
    
        Next DUTPin
        
        
        For Each DUTPin In Pins
            If TheExec.DataManager.ChannelType(DUTPin) <> "N/C" Then
                Tname = "Conti1_" & CStr(DUTPin)
                If TestLimitMode = tlForceFlow Then
'                    TheExec.Flow.TestLimit resultVal:=PPMUMeasure.Pins(DUTPin), ScaleType:=scaleNone, unit:=unitVolt, formatstr:="%.3f", Tname:=Tname, forceVal:=force_i, forceunit:=unitAmp, ForceResults:=tlForceFlow
                Else
                    TestLimitMode = tlForceNone
'                    TheExec.Flow.TestLimit resultVal:=PPMUMeasure.Pins(DUTPin), lowval:=LowLimit, hival:=HiLimit, ScaleType:=scaleNone, unit:=unitVolt, formatstr:="%.3f", Tname:=Tname, forceVal:=force_i, forceunit:=unitAmp, ForceResults:=tlForceNone 'original
                    'TheExec.Flow.TestLimit resultVal:=PPMUMeasure.Pins(DUTPin), ScaleType:=scaleNone, unit:=unitVolt, formatstr:="%.3f", Tname:=Tname, forceVal:=force_i, forceunit:=unitAmp, ForceResults:=tlForceNone 'ZB
                End If
            End If
            
        Next DUTPin
        
        If Separate_limit = True Then
            For Each DUTPin In Pins
                If TheExec.DataManager.ChannelType(DUTPin) <> "N/C" Then
                    Tname = "Conti2_" & CStr(DUTPin)
                    If TestLimitMode = tlForceFlow Then
                        TheExec.Flow.TestLimit resultVal:=PPMUMeasure.Pins(DUTPin), scaletype:=scaleNone, unit:=unitVolt, formatStr:="%.3f", Tname:=Tname, ForceVal:=force_i, ForceUnit:=unitAmp, ForceResults:=tlForceFlow
                    Else
                        TestLimitMode = tlForceNone
                        TheExec.Flow.TestLimit resultVal:=PPMUMeasure.Pins(DUTPin), lowVal:=LowLimit, hiVal:=hiLimit, scaletype:=scaleNone, unit:=unitVolt, formatStr:="%.3f", Tname:=Tname, ForceVal:=force_i, ForceUnit:=unitAmp, ForceResults:=tlForceNone
                    End If
                End If
                
            Next DUTPin
        End If
    
''      'initialize ppmu to suitable clamp
''    With TheHdw.PPMU.Pins("Pins_1p0v")
''        .ClampVHi = TheExec.Specs.Globals("Pins_1p0v_Vch_GLB").ContextValue
''        .ClampVLo = TheExec.Specs.Globals("Pins_1p0v_Vch_GLB").ContextValue
''    End With
''
''    With TheHdw.PPMU.Pins("Pins_1p1v")
''        .ClampVHi = TheExec.Specs.Globals("Pins_1p1v_Vch_GLB").ContextValue
''        .ClampVLo = TheExec.Specs.Globals("Pins_1p1v_Vch_GLB").ContextValue
''    End With
''
''    With TheHdw.PPMU.Pins("Pins_1p8v")
''        .ClampVHi = TheExec.Specs.Globals("Pins_1p8v_Vch_GLB").ContextValue
''        .ClampVLo = TheExec.Specs.Globals("Pins_1p8v_Vch_GLB").ContextValue
''    End With
    
     'Disconnect PPMU from digital channels
        With TheHdw.PPMU.Pins(digital_pins)
            .Gate = tlOff
            .Disconnect
        End With
             
        TheHdw.Digital.Pins(digital_pins).Connect
    
    Next force_i
'///////////////////////////////////ZB add for measure contact resistance//////////////////////////////////////////////////////
    For Each DUTPin In Pins
        For Each site In TheExec.sites
            Calculate_Contact_R.Pins(DUTPin).value(site) = (MeasV_Force50mA.Pins(DUTPin).value(site) - MeasV_Force0mA.Pins(DUTPin).value(site)) / 0.05
        Next site
    Next DUTPin
    
    Dim RakV() As Double
    For Each site In TheExec.sites
        For Each DUTPin In PPMUMeasure.Pins
            'RakV = TheHdw.PPMU.ReadRakValuesByPinnames(DUTPin, Site)
                        Calculate_Contact_R.Pins(DUTPin).value(site) = Calculate_Contact_R.Pins(DUTPin).value(site) - (CurrentJob_Card_RAK.Pins(DUTPin).value(site))
        Next DUTPin
    Next site

    TheExec.Flow.TestLimit resultVal:=Calculate_Contact_R, lowVal:=LowLimit, hiVal:=hiLimit, scaletype:=scaleNone, unit:=unitOhm, PinName:=digital_pins      'ZB add for measure contact resistance
    'TheExec.Flow.TestLimit resultVal:=PPMUMeasure.Pins(DUTPin), ScaleType:=scaleNone, unit:=unitVolt, formatstr:="%.3f", Tname:=Tname, forceVal:=force_i, forceunit:=unitAmp, ForceResults:=tlForceNone 'ZB modify
    '///////////////////////////////////End measure contact resistance//////////////////////////////////////////////////////
    
    Exit Function
errHandler: 'Add ErrHandler 2023/08/18
    Call Print_Error_Message(Error_Info, "VBT_LIB_DC_Conti", "PPMU_Measure_Contact_Resistance_Corner_Vss") 'Add ErrHandler 2023/08/18
    If isDebugMode Then TheExec.AddOutput "Error in Resistance Corner Vss"
    If AbortTest Then Exit Function Else Resume Next 'Add ErrHandler 2023/08/18
End Function

Private Function GetMeasRange(InitResult As Double, MeasRangeList() As Double)
On Error GoTo errHandler 'Add ErrHandler 2023/08/18
    Dim i As Integer
    Dim j As Integer
    Dim MinCurr As Double
    MinCurr = 0.0002
    GetMeasRange = MeasRangeList(UBound(MeasRangeList))
    For i = 0 To UBound(MeasRangeList)
        If (MeasRangeList(i) > 2 * Math.Abs(InitResult)) Then
            If (MeasRangeList(i) < 0.0002) Then
                For j = 0 To UBound(MeasRangeList)
                    If (MeasRangeList(j) >= MinCurr) Then
                        GetMeasRange = MeasRangeList(j)
                        Exit For
                    End If
                Next
            Else
                GetMeasRange = MeasRangeList(i)
            End If
            Exit For
        End If
    Next

Exit Function 'Add ErrHandler 2023/08/18
errHandler: 'Add ErrHandler 2023/08/18
    Call Print_Error_Message(Error_Info, "VBT_LIB_DC_Conti", "GetMeasRange") 'Add ErrHandler 2023/08/18
    If AbortTest Then Exit Function Else Resume Next 'Add ErrHandler 2023/08/18
End Function

' [20230407][RF] Add DCVI power short parallel
' [20230907][T-Don][Carter] Source/Sink Timeout Alarm behavior gate off setting
' [20231003][All][Tank] modify after Chihome review
' [20240110][All][Tank] modify fixcurrentrange argument
' [20240110][All][Tank] modify change limit argument type to string
Public Function p2p_short_Power_FVMI_Parallel(allPowerPins As String, _
                                                ForceCondition As String, _
                                                LowLimit As String, _
                                                hiLimit As String, _
                                                MaxForceV As Double, _
                                                PinGroup_Cnt As Double, _
                                                Step_Voltage_Level As Double, _
                                                TestLimitMode As tlLimitForceResults, _
                                                FlowLimitForInitIRange As Boolean, _
                                                digital_pins As PinList, _
                                                Optional FixCurrentRange As String) As Long
On Error GoTo errHandler
   ''++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
    'Testing Method:  Force 0.1V , measure smaller than 199ma,set clamp to 200ma, if higher than 199 ma then fail
    ''+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
    Dim funcName As String: funcName = "p2p_short_Power_FVMI_Parallel"

    
    Dim p As Variant, Pin_dcvi_Ary() As String
    Dim Tname As String
    Dim site As Variant
    Dim pin As New PinList
    
    Dim FlowLimitsInfo As IFlowLimitsInfo

    Dim val As Double
    Dim Val_Hi() As String
    Dim Val_Lo() As String

    Dim SattleTime As Double
    Dim WaitTime As Double
    
    Dim p_ufp As String
    Dim p_vsm As String
    Dim p_uvs As String
    Dim p_hexvs As String
    Dim p_dc07 As String
    Dim p_dc30 As String
    
    Dim p_ufp_Group() As String
    Dim p_vsm_Group() As String
    Dim p_uvs_Group() As String
    Dim p_hexVs_Group() As String
    Dim p_dc07_Group() As String
    Dim p_dc30_Group() As String
    
    Dim ufp_PowerPindata As New PinListData
    Dim vsm_PowerPindata As New PinListData
    Dim uvs_PowerPindata As New PinListData
    Dim hexVs_PowerPindata As New PinListData
    Dim dc07_PowerPindata As New PinListData
    Dim dc30_PowerPindata As New PinListData
    Dim PowerPindata As New PinListData

    Dim MaxValue As Double
    Dim getMaxValue As Double
    Dim MeasRangeList() As Double
    Dim UVI80MeasRangeList() As Double
    
    '''====================  Re-write
    Dim Num_PinGroup As Integer
    Dim AllPin_Cnt As Integer

    Dim PowerPins_Group() As String
    Dim Current_PinsGroup_ary() As String
    Dim Current_Val_Hi() As String
    Dim Current_Val_Lo() As String
    
    Dim iCount As Integer: iCount = 0
    Dim GroupIndex As Integer: GroupIndex = 0
    Dim PinsIndex As Integer: PinsIndex = 0
    Dim EFUSELimit As Double: EFUSELimit = 0.1
    Dim TempForceV As Double: TempForceV = 0
    Dim StepVoltageMode As Boolean: StepVoltageMode = False
    Dim ForceV As Double
    Dim ForceValue As New Dictionary
    Dim dic_HiLimit As New Dictionary
    Dim dic_LowLimit As New Dictionary
  
    ''' setup by init default
    If MaxForceV = 0 Then MaxForceV = 3
    If PinGroup_Cnt = 0 Then PinGroup_Cnt = 10
    If Step_Voltage_Level = 0 Then Step_Voltage_Level = 0 '0.01
    
    Dim i As Long
    Dim j As Long
    Dim Pin_Cnt As Long
    Dim Pin_Ary() As String
    Dim tmpstr() As String
    
    Dim CustomizeTname As String
    Dim TmpTname As String
    Dim PinNameCustomize As String
    Dim PowerPins As String
    Dim tmp_Pin_Instrument As String
    
    Dim A_Slot_Type As String
    Dim Grp_Val_Hi() As String
    Dim Grp_Val_Lo() As String
    Dim sPin As String
    
    Dim s_tempDCVIPins As String
    Dim s_tempDCVSPins As String
    Dim sa_MeasurePinGroup() As String
    
    Dim sa_tempHiLimit() As String
    Dim sa_tempLowLimit() As String
    Dim sa_Flow_Grp_ValHi() As String
    Dim sa_MeasurePinListAry() As String
    Dim s_ErrorMsg As String
    Dim dict_FixedCurrent As New Dictionary
    
    
    Set GlbCustomizeSet = New CustomizeSetting
    If gB_Enable_AP Then CustomizeTname = GlbCustomizeSet.ReportTname(Type_AP, Type_Conti, CStr(digital_pins))
    If gB_Enable_LCD Then CustomizeTname = GlbCustomizeSet.ReportTname(Type_LCD, Type_Conti)
    If gB_Enable_RF Then
        CustomizeTname = GlbCustomizeSet.ReportTname(Type_RF, Type_Conti)
    End If
    CustomizeTname = Replace(CustomizeTname, "_<PinSequence>", "")
    CustomizeTname = Replace(CustomizeTname, "_<ForceI>", "")
    
    If allPowerPins = "" Then
        TheExec.Datalog.WriteComment "No Input pins."
        GoTo errHandler
    End If

    If (FlowLimitForInitIRange = True) Or (TestLimitMode = tlForceFlow) Then
        Call TheExec.Flow.GetTestLimits(FlowLimitsInfo)
        FlowLimitsInfo.GetHighLimits Val_Hi
        FlowLimitsInfo.GetLowLimits Val_Lo
    End If
    
    sa_MeasurePinGroup = Split(allPowerPins, ";")

    tmpstr = Split(ForceCondition, ";")
    
    Call FixCurrentRange_StrToDic(FixCurrentRange, dict_FixedCurrent)
    
    '==== Check ForceCondition count ====
    If UBound(tmpstr) + 1 > 1 And UBound(tmpstr) <> UBound(sa_MeasurePinGroup) Then
        TheExec.Flow.TestLimit resultVal:=1, lowVal:=-1, hiVal:=-1, unit:=unitNone, Tname:=TheExec.DataManager.instancename
        s_ErrorMsg = "The Numbers of PinGroup / ForceValue aren't the same!!"
        Call Print_Error_Message(Error_Warning_Info.Warning_Info, "VBT_LIB_DC_Conti", "p2p_short_Power_FVMI_Parallel", s_ErrorMsg)
        Exit Function
    End If
    '==== Check ForceCondition count ====
    
    ReDim sa_MeasurePinListAry(UBound(sa_MeasurePinGroup))
    
    '==== Store ForceCondition by each pins ====
    If UBound(tmpstr) + 1 = 1 Then
        ForceV = CDbl(tmpstr(0))
        StepVoltageMode = True
    End If
    
    Set ForceValue = Nothing
    For j = 0 To UBound(sa_MeasurePinGroup)
        TheExec.DataManager.DecomposePinList sa_MeasurePinGroup(j), Pin_Ary, Pin_Cnt
        sa_MeasurePinListAry(j) = Join(Pin_Ary, ",")
        If UBound(tmpstr) + 1 > 1 Then
            For i = 0 To UBound(Pin_Ary)
                If ForceValue.Exists(LCase(Pin_Ary(i))) = False Then ForceValue.Add LCase(Pin_Ary(i)), CDbl(tmpstr(j))
            Next i
        End If
    Next j
    '==== Store ForceCondition by each pins ====
    
    '==== Check limit count ====
    If TestLimitMode = tlForceNone Then
        Set dic_HiLimit = Nothing
        Set dic_LowLimit = Nothing
        
        sa_tempHiLimit = Split(hiLimit, ";")
        sa_tempLowLimit = Split(LowLimit, ";")
        
        If (UBound(sa_MeasurePinGroup) = UBound(sa_tempHiLimit)) And (UBound(sa_MeasurePinGroup) = UBound(sa_tempLowLimit)) Then
            For j = 0 To UBound(sa_MeasurePinGroup)
                'TheExec.DataManager.DecomposePinList sa_MeasurePinGroup(j), Pin_Ary, Pin_Cnt
                Pin_Ary = Split(sa_MeasurePinListAry(j), ",")
                For i = 0 To UBound(Pin_Ary)
                    If dic_HiLimit.Exists(LCase(Pin_Ary(i))) = False Then dic_HiLimit.Add LCase(Pin_Ary(i)), CDbl(sa_tempHiLimit(j))
                    If dic_LowLimit.Exists(LCase(Pin_Ary(i))) = False Then dic_LowLimit.Add LCase(Pin_Ary(i)), CDbl(sa_tempLowLimit(j))
                Next i
            Next j
        Else
            TheExec.Flow.TestLimit resultVal:=1, lowVal:=-1, hiVal:=-1, unit:=unitNone, Tname:=TheExec.DataManager.instancename
            s_ErrorMsg = "The Numbers of PinGroup / HighLimit / LowLimit aren't the same!!"
            Call Print_Error_Message(Error_Warning_Info.Warning_Info, "VBT_LIB_DC_Conti", "p2p_short_Power_FVMI_Parallel", s_ErrorMsg)
            Exit Function
        End If
    End If
    '==== Check limit count ====

'''***** Start - Define Instrument Type *****
    For j = 0 To UBound(sa_MeasurePinGroup)
        'TheExec.DataManager.DecomposePinList sa_MeasurePinGroup(j), Pin_Ary, Pin_Cnt
        Pin_Ary = Split(sa_MeasurePinListAry(j), ",")
        For i = 0 To UBound(Pin_Ary)
            If gl_GetInstrument_Dic.Exists(LCase(Pin_Ary(i))) Then
                A_Slot_Type = UCase(gl_GetInstrument_Dic(LCase(Pin_Ary(i))))
                Select Case A_Slot_Type
                    Case glbConstIns_HEXVS: p_hexvs = CombineStringList(p_hexvs, Pin_Ary(i))
                    Case glbConstIns_VHDVS: p_uvs = CombineStringList(p_uvs, Pin_Ary(i))
                    Case glbConstIns_VSM: p_vsm = CombineStringList(p_vsm, Pin_Ary(i))
                    Case glbConstIns_VS5A: p_ufp = CombineStringList(p_ufp, Pin_Ary(i))
                    Case glbConstIns_VS800MA: p_ufp = CombineStringList(p_ufp, Pin_Ary(i))
                    Case glbConstIns_DC07: p_dc07 = CombineStringList(p_dc07, Pin_Ary(i))
                    Case glbConstIns_DC30: p_dc30 = CombineStringList(p_dc30, Pin_Ary(i))
                End Select
                
                PowerPins = CombineStringList(PowerPins, Pin_Ary(i))
                AllPin_Cnt = AllPin_Cnt + 1
            End If
        Next i
    Next j
'''***** End - Define Instrument Type *****

'''***** Start - Define the number of PinGroup *****
    If PinGroup_Cnt > AllPin_Cnt Then PinGroup_Cnt = AllPin_Cnt
    Pin_Ary = Split(PowerPins, ",")
    
    If AllPin_Cnt <> 0 And PinGroup_Cnt <> 0 Then
        ''' check PinGroup_Cnt is over max voltage
        If (StepVoltageMode = True) And ((ForceV + ((PinGroup_Cnt - 1) * Step_Voltage_Level)) > MaxForceV) Then
            PinGroup_Cnt = CInt((MaxForceV - ForceV) / Step_Voltage_Level)
            TheExec.Datalog.WriteComment " Over max voltage, change PinGroup_Cnt to" & PinGroup_Cnt
        End If
        If AllPin_Cnt Mod PinGroup_Cnt <> 0 Then
            Num_PinGroup = Floor(AllPin_Cnt / PinGroup_Cnt) + 1
        Else
            Num_PinGroup = (AllPin_Cnt / PinGroup_Cnt)
        End If
    End If
'''***** End - Define the number of PinGroup *****

    ReDim Grp_Val_Hi(Num_PinGroup - 1) As String
    ReDim Grp_Val_Lo(Num_PinGroup - 1) As String
    ReDim sa_Flow_Grp_ValHi(Num_PinGroup - 1) As String
    
    ReDim PowerPins_Group(Num_PinGroup - 1) As String
    ReDim p_hexVs_Group(Num_PinGroup - 1) As String
    ReDim p_uvs_Group(Num_PinGroup - 1) As String
    ReDim p_vsm_Group(Num_PinGroup - 1) As String
    ReDim p_ufp_Group(Num_PinGroup - 1) As String
    ReDim p_dc07_Group(Num_PinGroup - 1) As String
    ReDim p_dc30_Group(Num_PinGroup - 1) As String
    
    
    '''Pin_ary is sorting'''
    iCount = 0
    For GroupIndex = 0 To (Num_PinGroup - 1)
        For PinsIndex = 0 To (PinGroup_Cnt - 1)
            If TestLimitMode = tlForceFlow Then
                Grp_Val_Hi(GroupIndex) = CombineStringList(Grp_Val_Hi(GroupIndex), Val_Hi(iCount))
                Grp_Val_Lo(GroupIndex) = CombineStringList(Grp_Val_Lo(GroupIndex), Val_Lo(iCount))
            Else
                Grp_Val_Hi(GroupIndex) = CombineStringList(Grp_Val_Hi(GroupIndex), dic_HiLimit(LCase(Pin_Ary(iCount))))
                Grp_Val_Lo(GroupIndex) = CombineStringList(Grp_Val_Lo(GroupIndex), dic_LowLimit(LCase(Pin_Ary(iCount))))
            End If
            
            If FlowLimitForInitIRange = True Then
                sa_Flow_Grp_ValHi(GroupIndex) = CombineStringList(sa_Flow_Grp_ValHi(GroupIndex), Val_Hi(iCount))        'FlowLimitForInitIRange use
            End If
            
            PowerPins_Group(GroupIndex) = CombineStringList(PowerPins_Group(GroupIndex), Pin_Ary(iCount))
            
            A_Slot_Type = UCase(gl_GetInstrument_Dic(LCase(Pin_Ary(iCount))))
        
            Select Case A_Slot_Type
                Case glbConstIns_HEXVS: p_hexVs_Group(GroupIndex) = CombineStringList(p_hexVs_Group(GroupIndex), Pin_Ary(iCount))
                Case glbConstIns_VHDVS: p_uvs_Group(GroupIndex) = CombineStringList(p_uvs_Group(GroupIndex), Pin_Ary(iCount))
                Case glbConstIns_VSM: p_vsm_Group(GroupIndex) = CombineStringList(p_vsm_Group(GroupIndex), Pin_Ary(iCount))
                Case glbConstIns_VS5A: p_ufp_Group(GroupIndex) = CombineStringList(p_ufp_Group(GroupIndex), Pin_Ary(iCount))
                Case glbConstIns_VS800MA: p_ufp_Group(GroupIndex) = CombineStringList(p_ufp_Group(GroupIndex), Pin_Ary(iCount))
                Case glbConstIns_DC07: p_dc07_Group(GroupIndex) = CombineStringList(p_dc07_Group(GroupIndex), Pin_Ary(iCount))
                Case glbConstIns_DC30: p_dc30_Group(GroupIndex) = CombineStringList(p_dc30_Group(GroupIndex), Pin_Ary(iCount))
            End Select
            
            iCount = iCount + 1
            If (iCount >= AllPin_Cnt) Then Exit For
        Next PinsIndex
        
    Next GroupIndex
    
    s_tempDCVSPins = vbNullString
    s_tempDCVIPins = vbNullString
    
    s_tempDCVSPins = CombineStringList(s_tempDCVSPins, p_ufp)
    s_tempDCVSPins = CombineStringList(s_tempDCVSPins, p_vsm)
    s_tempDCVSPins = CombineStringList(s_tempDCVSPins, p_uvs)
    s_tempDCVSPins = CombineStringList(s_tempDCVSPins, p_hexvs)
    
    s_tempDCVIPins = CombineStringList(s_tempDCVIPins, p_dc07)
    s_tempDCVIPins = CombineStringList(s_tempDCVIPins, p_dc30)
    
    If s_tempDCVSPins <> "" Then TheHdw.DCVS.Pins(s_tempDCVSPins).Gate = False
    If s_tempDCVIPins <> "" Then TheHdw.DCVI.Pins(s_tempDCVIPins).Gate = False
    

    TheHdw.Digital.ApplyLevelsTiming True, True, True, tlPowered     'SEC DRAM
    
    ''''Disconnect os_Pins Pin Electronics from Pins in order to connect PPMU's''''
    If digital_pins.value <> "" Then TheHdw.Digital.Pins(digital_pins).Disconnect
    
    TheHdw.Wait 3 * ms
    WaitTime = 260 * us
    
    '''====================  Re-write
    If s_tempDCVSPins <> "" Then TheHdw.DCVS.Pins(s_tempDCVSPins).Voltage.Main.value = 0# '''Set all power pins to 0 V first
    If s_tempDCVIPins <> "" Then TheHdw.DCVI.Pins(s_tempDCVIPins).Voltage.value = 0# '''Set all power pins to 0 V first
        
    TheHdw.Wait 3 * ms
    
    For GroupIndex = 0 To (Num_PinGroup - 1)
        '''' init PinListData
        
        Set PowerPindata = Nothing
        Set ufp_PowerPindata = Nothing
        Set vsm_PowerPindata = Nothing
        Set uvs_PowerPindata = Nothing
        Set hexVs_PowerPindata = Nothing
        Set dc07_PowerPindata = Nothing
        Set dc30_PowerPindata = Nothing
        
        iCount = 0
        
        Current_Val_Hi = Split(Grp_Val_Hi(GroupIndex), ",")
        Current_Val_Lo = Split(Grp_Val_Lo(GroupIndex), ",")
        Current_PinsGroup_ary = Split(PowerPins_Group(GroupIndex), ",")
        
'''***** Start - Vmain Voltage Setup and raise Step_Voltage_Level per step *****
        For PinsIndex = 0 To UBound(Current_PinsGroup_ary)
            If UCase(Current_PinsGroup_ary(PinsIndex)) Like "*EFUSE*" And ((ForceV + (PinsIndex * Step_Voltage_Level)) >= EFUSELimit) Then
                TheExec.Datalog.WriteComment "Pin: " & Current_PinsGroup_ary(PinsIndex) & " Voltage Level over " & EFUSELimit & " v"
            End If
            
            A_Slot_Type = UCase(gl_GetInstrument_Dic(LCase(Current_PinsGroup_ary(PinsIndex))))
            
            Select Case A_Slot_Type
                Case glbConstIns_HEXVS, glbConstIns_VSM, glbConstIns_VHDVS, glbConstIns_VS5A, glbConstIns_VS800MA
                    If StepVoltageMode = False Then
                        ForceV = ForceValue(LCase(Current_PinsGroup_ary(PinsIndex)))
                        TheHdw.DCVS.Pins(Current_PinsGroup_ary(PinsIndex)).Voltage.Main.value = ForceV
                    Else
                        TheHdw.DCVS.Pins(Current_PinsGroup_ary(PinsIndex)).Voltage.Main.value = ForceV + (PinsIndex * Step_Voltage_Level) ''force V by pin
                    End If
                Case glbConstIns_DC07, glbConstIns_DC30
                    If StepVoltageMode = False Then
                        ForceV = ForceValue(LCase(Current_PinsGroup_ary(PinsIndex)))
                        TheHdw.DCVI.Pins(Current_PinsGroup_ary(PinsIndex)).Voltage.value = ForceV
                    Else
                        TheHdw.DCVI.Pins(Current_PinsGroup_ary(PinsIndex)).Voltage.value = ForceV + (PinsIndex * Step_Voltage_Level) ''force V by pin
                    End If
            End Select
            
        Next PinsIndex
'''***** End - Vmain Voltage Setup and raise Step_Voltage_Level per step *****
        
'''***** Start - CurrentRange Setup *****
        If UBound(Current_Val_Hi) <> UBound(Current_PinsGroup_ary) Then
            TheExec.Datalog.WriteComment "The dimension is different between Current_PinsGroup_ary and Current_Val_Hi. Please check Test Limit in the Flow Sheet."
            GoTo errHandler
        Else
            Call SetGroupCurrentRange(PowerPins_Group(GroupIndex), sa_Flow_Grp_ValHi(GroupIndex), FlowLimitForInitIRange, WaitTime, Grp_Val_Hi(GroupIndex), dict_FixedCurrent)
        End If
        
        Dim ufp_Array() As String
        Dim vsm_Array() As String
        Dim uvs_Array() As String
        Dim hexVs_Array() As String
        Dim dc07_Array() As String
        Dim dc30_Array() As String
        
        If p_ufp_Group(GroupIndex) <> "" Then ufp_Array = Split(p_ufp_Group(GroupIndex), ",")
        If p_vsm_Group(GroupIndex) <> "" Then vsm_Array = Split(p_vsm_Group(GroupIndex), ",")
        If p_uvs_Group(GroupIndex) <> "" Then uvs_Array = Split(p_uvs_Group(GroupIndex), ",")
        If p_hexVs_Group(GroupIndex) <> "" Then hexVs_Array = Split(p_hexVs_Group(GroupIndex), ",")
        If p_dc07_Group(GroupIndex) <> "" Then dc07_Array = Split(p_dc07_Group(GroupIndex), ",")
        If p_dc30_Group(GroupIndex) <> "" Then dc30_Array = Split(p_dc30_Group(GroupIndex), ",")
        
        ReDim Preserve ufp_Array(UBound(Current_PinsGroup_ary))
        ReDim Preserve vsm_Array(UBound(Current_PinsGroup_ary))
        ReDim Preserve uvs_Array(UBound(Current_PinsGroup_ary))
        ReDim Preserve hexVs_Array(UBound(Current_PinsGroup_ary))
        ReDim Preserve dc07_Array(UBound(Current_PinsGroup_ary))
        ReDim Preserve dc30_Array(UBound(Current_PinsGroup_ary))
        For PinsIndex = 0 To UBound(Current_PinsGroup_ary)
            If p_ufp_Group(GroupIndex) <> "" Then
                If ufp_Array(PinsIndex) <> "" Then TheExec.Datalog.WriteComment " Pin:" & ufp_Array(PinsIndex) & " Current Range: " & TheHdw.DCVS.Pins(ufp_Array(PinsIndex)).Meter.CurrentRange
            End If
            If p_vsm_Group(GroupIndex) <> "" Then
                If vsm_Array(PinsIndex) <> "" Then TheExec.Datalog.WriteComment " Pin:" & vsm_Array(PinsIndex) & " Current Range: " & TheHdw.DCVS.Pins(vsm_Array(PinsIndex)).Meter.CurrentRange
            End If
            If p_uvs_Group(GroupIndex) <> "" Then
                If uvs_Array(PinsIndex) <> "" Then TheExec.Datalog.WriteComment " Pin:" & uvs_Array(PinsIndex) & " Current Range: " & TheHdw.DCVS.Pins(uvs_Array(PinsIndex)).Meter.CurrentRange
            End If
            If p_hexVs_Group(GroupIndex) <> "" Then
                If hexVs_Array(PinsIndex) <> "" Then TheExec.Datalog.WriteComment " Pin:" & hexVs_Array(PinsIndex) & " Current Range: " & TheHdw.DCVS.Pins(hexVs_Array(PinsIndex)).Meter.CurrentRange
            End If
            If p_dc07_Group(GroupIndex) <> "" Then
                If dc07_Array(PinsIndex) <> "" Then TheExec.Datalog.WriteComment " Pin:" & dc07_Array(PinsIndex) & " Current Range: " & TheHdw.DCVI.Pins(dc07_Array(PinsIndex)).Meter.CurrentRange
            End If
            If p_dc30_Group(GroupIndex) <> "" Then
                If dc30_Array(PinsIndex) <> "" Then TheExec.Datalog.WriteComment " Pin:" & dc30_Array(PinsIndex) & " Current Range: " & TheHdw.DCVI.Pins(dc30_Array(PinsIndex)).Meter.CurrentRange
            End If
        Next PinsIndex
'''***** End - CurrentRange Setup *****
        TheHdw.Wait WaitTime
'''***** Start - Measurement *****
        If p_ufp_Group(GroupIndex) <> "" Then ufp_PowerPindata = TheHdw.DCVS.Pins(p_ufp_Group(GroupIndex)).Meter.Read(tlStrobe, 10, , tlDCVSMeterReadingFormatAverage)
        If p_vsm_Group(GroupIndex) <> "" Then vsm_PowerPindata = TheHdw.DCVS.Pins(p_vsm_Group(GroupIndex)).Meter.Read(tlStrobe, 10, , tlDCVSMeterReadingFormatAverage)
        If p_uvs_Group(GroupIndex) <> "" Then uvs_PowerPindata = TheHdw.DCVS.Pins(p_uvs_Group(GroupIndex)).Meter.Read(tlStrobe, 1)
        If p_hexVs_Group(GroupIndex) <> "" Then hexVs_PowerPindata = TheHdw.DCVS.Pins(p_hexVs_Group(GroupIndex)).Meter.Read(tlStrobe, 10, , tlDCVSMeterReadingFormatAverage)
        If p_dc07_Group(GroupIndex) <> "" Then dc07_PowerPindata = TheHdw.DCVI.Pins(p_dc07_Group(GroupIndex)).Meter.Read(tlStrobe, 10, , tlDCVIMeterReadingFormatAverage)
        If p_dc30_Group(GroupIndex) <> "" Then dc30_PowerPindata = TheHdw.DCVI.Pins(p_dc30_Group(GroupIndex)).Meter.Read(tlStrobe, 10, , tlDCVIMeterReadingFormatAverage)

'        PowerPindata = TheHdw.DCVS.Pins(PowerPins_Group(GroupIndex)).Meter.Read(tlStrobe, 10, , tlDCVSMeterReadingFormatAverage)
'''***** End - Measurement *****

'''***** Start - DataLog printing *****
        If PowerPins_Group(GroupIndex) <> "" Then
            For PinsIndex = 0 To UBound(Current_PinsGroup_ary)
            
                PinNameCustomize = Current_PinsGroup_ary(PinsIndex)
                If gB_Enable_AP Or gB_Enable_RF Then
                    PinNameCustomize = Replace(PinNameCustomize, "_", "")
                ElseIf gB_Enable_LCD Then
                    PinNameCustomize = Replace(PinNameCustomize, "_", "-")
                End If
                
                TmpTname = Replace(CustomizeTname, "<PinName>", PinNameCustomize)
                If StepVoltageMode = False Then
                    TempForceV = ForceValue(LCase(Current_PinsGroup_ary(PinsIndex)))
                Else
                    TempForceV = (ForceV + (iCount * Step_Voltage_Level)) ''TheHdw.DCVS.Pins(CurrentGroup_HexVS(PinsIndex)).Voltage.Main.Value
                End If

                tmp_Pin_Instrument = UCase(gl_GetInstrument_Dic(LCase(Current_PinsGroup_ary(PinsIndex))))
                
                Select Case tmp_Pin_Instrument
                    Case glbConstIns_HEXVS
                        PowerPindata = hexVs_PowerPindata.Copy(Current_PinsGroup_ary(PinsIndex))
                    Case glbConstIns_VHDVS
                        PowerPindata = uvs_PowerPindata.Copy(Current_PinsGroup_ary(PinsIndex))
                    Case glbConstIns_VSM
                        PowerPindata = vsm_PowerPindata.Copy(Current_PinsGroup_ary(PinsIndex))
                    Case glbConstIns_VS5A
                        PowerPindata = ufp_PowerPindata.Copy(Current_PinsGroup_ary(PinsIndex))
                    Case glbConstIns_VS800MA
                        PowerPindata = ufp_PowerPindata.Copy(Current_PinsGroup_ary(PinsIndex))
                    Case glbConstIns_DC07
                        PowerPindata = dc07_PowerPindata.Copy(Current_PinsGroup_ary(PinsIndex))
                    Case glbConstIns_DC30
                        PowerPindata = dc30_PowerPindata.Copy(Current_PinsGroup_ary(PinsIndex))
                    Case Else
                        Call Print_Error_Message(Error_Info, "VBT_LIB_DC_Conti", "p2p_short_Power_FVMI_Parallel", "Cannot find pin's instrument!")
                End Select
                
                'offline mode simulation
                If TheExec.TesterMode = testModeOffline Then
                    For Each site In TheExec.sites
                        PowerPindata.Pins(Current_PinsGroup_ary(PinsIndex)).value(site) = 0.01 + Rnd() * 0.0001
                    Next site
                End If
                
                If TestLimitMode = tlForceFlow Then
                    Tname = TmpTname & "_" & CStr(GroupIndex)
                End If
                TheExec.Flow.TestLimit resultVal:=PowerPindata.Pins(Current_PinsGroup_ary(PinsIndex)), lowVal:=CDbl(Current_Val_Lo(PinsIndex)), hiVal:=CDbl(Current_Val_Hi(PinsIndex)), scaletype:=scaleNone, unit:=unitAmp, formatStr:="%.3f", Tname:=TmpTname, ForceVal:=TempForceV, ForceUnit:=unitVolt, ForceResults:=tlForceNone
                
                iCount = iCount + 1
                If (iCount >= AllPin_Cnt) Then Exit For
            Next PinsIndex
        End If

'''***** End - DataLog printing *****

        ''' reset this pins group voltage
        If p_ufp_Group(GroupIndex) <> "" Then
            TheHdw.DCVS.Pins(p_ufp_Group(GroupIndex)).Voltage.Main.value = 0#
        End If
        If p_vsm_Group(GroupIndex) <> "" Then
            TheHdw.DCVS.Pins(p_vsm_Group(GroupIndex)).Voltage.Main.value = 0#
        End If
        If p_uvs_Group(GroupIndex) <> "" Then
            TheHdw.DCVS.Pins(p_uvs_Group(GroupIndex)).Voltage.Main.value = 0#
        End If
        If p_hexVs_Group(GroupIndex) <> "" Then
            TheHdw.DCVS.Pins(p_hexVs_Group(GroupIndex)).Voltage.Main.value = 0#
        End If
        If p_dc07_Group(GroupIndex) <> "" Then
            TheHdw.DCVI.Pins(p_dc07_Group(GroupIndex)).Voltage.value = 0#
        End If
        If p_dc30_Group(GroupIndex) <> "" Then
            TheHdw.DCVI.Pins(p_dc30_Group(GroupIndex)).Voltage.value = 0#
        End If
        ''' reset this pins group voltage
        TheHdw.Wait 0.01
        
    Next GroupIndex
    
    TheHdw.Digital.ApplyLevelsTiming False, True, False, tlPowered        'SEC DRAM
    DebugPrintFunc ""
    '''====================  Re-write end
    
Exit Function 'Add ErrHandler 2023/08/18
errHandler: 'Add ErrHandler 2023/08/18
    Call Print_Error_Message(Error_Info, "VBT_LIB_DC_Conti", "p2p_short_Power_FVMI_Parallel") 'Add ErrHandler 2023/08/18
    If AbortTest Then Exit Function Else Resume Next 'Add ErrHandler 2023/08/18
End Function


' [20230620][T-All] No need to gate off group pin when test next group pin.
' [20231003][All][Tank] modify after Chihome review
' [20231003][All][HI] Just test WalkingZ fail pins
' [20231016][All][Tank] Add 2nd limit group
' [20231228][T-All][Tank] Check 1st digital pin NC
' [20240308][T-All][Clyde] Connect all measured pin and force 0 voltage, disconnect after meter
Public Function AutoZ_Continuity(PinName As String, ForceValue As String, Limit_L As String, Limit_H As String, TestLimitMode As tlLimitForceResults, Optional PN_Disconnect As Boolean = False, _
                                    Optional Flag_Open As String = "F_open", Optional Flag_Short As String = "F_short", Optional connect_all_pins As String, _
                                    Optional isTestWalkingZResult As Boolean = False, Optional Limit_L2 As String = vbNullString, Optional Limit_H2 As String = vbNullString) As Long
On Error GoTo errHandler 'Add ErrHandler 2023/08/18

    
    Dim PPMUMeasure As New PinListData
    Dim DUTPin As Variant
    Dim Pins() As String, Pin_Cnt As Long
    Dim i As Long
    Dim Tname As String
    Dim site As Variant
    '////////////////////////////////////////////////
    Dim FlowLimitObj As IFlowLimitsInfo
    Dim Lolimit_new() As Double
    Dim HiLimit_new() As Double
    Dim Lolimit_str() As String
    Dim Hilimit_str() As String

    Dim digital_pins As String
    Dim force_i As Double
    Dim LowLimit As Double
    Dim hiLimit As Double
    Dim LowLimit_2nd As Double
    Dim HiLimit_2nd As Double
    Dim PinGroupName() As String
    Dim ForceI() As String
    Dim Low() As String
    Dim High() As String
    Dim LoopIndex As Long
    Dim GroupNum As Long: GroupNum = 0
    
    Dim pinsequence As Long: pinsequence = 0
    Dim CustomizeTname As String
    Dim PinNameCustomize As String
    Dim TmpTname As String
    
    Dim RetestPinNameAry() As String
    Dim ReTestFlag() As Boolean
    
    Dim s_Instrument_tmp As String
    Dim n_GetFailPinLoopCount As Long
    Dim s_ErrorMsg As String
    Dim n_FlowLimitIndex As Integer
    
    Dim n_PinGroupCnt As Long
    Dim n_ForceCnt As Long
    Dim n_AllPinCnt As Long
    Dim n_LimitCnt_L As Long
    Dim n_LimitCnt_H As Long
    Dim n_LimitCnt_L2 As Long
    Dim n_LimitCnt_H2 As Long
    Dim ad_FlowLimit_L() As Double
    Dim ad_FlowLimit_H() As Double
    Dim ad_FlowLimit_L2() As Double
    Dim ad_FlowLimit_H2() As Double
    Dim as_TempLimitH() As String
    Dim as_TempLimitL() As String
    Dim b_isLimitCntSame As Boolean
    Dim dic_LoopPinGroup As New Dictionary
    Dim dic_CustomizeTname As New Dictionary
    Dim b_isExist2ndLimit As Boolean
    Dim tmpName As String
    Dim s_TmpDigital As String
    Dim instSet As New InstrumentUtility
    
    PinGroupName = Split(PinName, ";")
    ForceI = Split(ForceValue, ";")
    Low = Split(Limit_L, ";")
    High = Split(Limit_H, ";")

    n_PinGroupCnt = UBound(PinGroupName) + 1
    n_ForceCnt = UBound(ForceI) + 1
    Set GlbCustomizeSet = New CustomizeSetting
    
    'Check count with pin and ForceCondition
    If n_PinGroupCnt = n_ForceCnt Then
        GroupNum = UBound(PinGroupName)
    Else
        TheExec.Flow.TestLimit resultVal:=1, lowVal:=-1, hiVal:=-1, unit:=unitNone, Tname:=TheExec.DataManager.instancename
        s_ErrorMsg = "The Count of pins and ForceValue isn't match."
        Call Print_Error_Message(Error_Warning_Info.Warning_Info, "VBT_LIB_DC_Conti", "AutoZ_Continuity", s_ErrorMsg)
        Exit Function
    End If
    
    If TestLimitMode = tlForceFlow Then
        Call TheExec.Flow.GetTestLimits(FlowLimitObj)
        Call FlowLimitObj.GetLowLimits(Lolimit_str)
        Call FlowLimitObj.GetHighLimits(Hilimit_str)

        ReDim Lolimit_new(UBound(Lolimit_str))
        ReDim HiLimit_new(UBound(Hilimit_str))
        
        For i = 0 To UBound(Lolimit_str)
            If Lolimit_str(i) <> "" Then Lolimit_new(i) = CDbl(Lolimit_str(i))
        Next i
        
        For i = 0 To UBound(Hilimit_str)
            If Hilimit_str(i) <> "" Then HiLimit_new(i) = CDbl(Hilimit_str(i))
        Next i
        n_LimitCnt_L = UBound(Lolimit_new) + 1
        n_LimitCnt_H = UBound(HiLimit_new) + 1
    Else
        n_LimitCnt_L = UBound(Low) + 1
        n_LimitCnt_H = UBound(High) + 1
    End If
    
    '==================add FailPin to ary ======================
    ReDim RetestPinNameAry(GroupNum)
    ReDim ReTestFlag(GroupNum)
    
    If isTestWalkingZResult Then
        If gldic_ComposeFailPins.Count > 0 Then
            For LoopIndex = 0 To GroupNum
                ReTestFlag(LoopIndex) = isContiTestFailPinList(PinGroupName(LoopIndex), gldic_ComposeFailPins, RetestPinNameAry(LoopIndex))
                If ReTestFlag(LoopIndex) Then
                    n_GetFailPinLoopCount = n_GetFailPinLoopCount + 1
                End If
            Next LoopIndex
        End If

        If n_GetFailPinLoopCount = 0 Then
            TheExec.Flow.TestLimit resultVal:=1, lowVal:=-1, hiVal:=-1, unit:=unitNone, Tname:=TheExec.DataManager.instancename
            s_ErrorMsg = "Doesn't get fail pins when isTestWalkingZResult = True"
            Call Print_Error_Message(Error_Warning_Info.Warning_Info, "VBT_LIB_DC_Conti", "AutoZ_Continuity", s_ErrorMsg)
            Exit Function
        End If
    End If
    '==================add FailPin to ary======================
    
    '====Check pin test limit not larger then clamp====
    n_AllPinCnt = 0
    n_FlowLimitIndex = 0
    For LoopIndex = 0 To GroupNum
        digital_pins = PinGroupName(LoopIndex)
        LowLimit = Low(LoopIndex)
        hiLimit = High(LoopIndex)
        
        TheExec.DataManager.DecomposePinList digital_pins, Pins(), Pin_Cnt
        n_AllPinCnt = n_AllPinCnt + Pin_Cnt
        dic_LoopPinGroup.Add LoopIndex, Pins
        dic_CustomizeTname.Add LoopIndex, GlbCustomizeSet.ReportTname(Type_AP, Type_Conti, CStr(digital_pins))
        For Each DUTPin In Pins
            If gl_GetInstrument_Dic.Exists(LCase(DUTPin)) Then
                s_Instrument_tmp = LCase(gl_GetInstrument_Dic(LCase(DUTPin)))
                If s_Instrument_tmp = glbConstIns_UP2200 Or s_Instrument_tmp = glbConstIns_UP1600 Then      '''Assume instrument type of each group are the same.
                    If gl_isCheckClampLimit <> ContiClampCheckType.CheckPass Then
                        If True = CheckTestInst_HiLoLimit(LowLimit, hiLimit, Lolimit_new(), HiLimit_new(), Pins(), TestLimitMode, FlowLimitIndex:=n_FlowLimitIndex) Then Exit Function
                    End If
                End If
                Exit For
            End If
        Next DUTPin
    Next LoopIndex
    '====Check pin test limit not larger then clamp====
    
    If Flag_Open Like "" Then Flag_Open = "F_open"
    If Flag_Short Like "" Then Flag_Short = "F_short"
    b_isLimitCntSame = False
    
    '====Check count with pin and limit====
    If TestLimitMode = tlForceNone Then
        If (Limit_L2 = "") And (Limit_H2 = "") Then
            If n_PinGroupCnt = n_LimitCnt_L And n_PinGroupCnt = n_LimitCnt_H Then    'Compare with pin group count
                b_isExist2ndLimit = False
                b_isLimitCntSame = True
            Else
                b_isLimitCntSame = False
            End If
        ElseIf (Limit_L2 <> "") And (Limit_H2 <> "") Then
            as_TempLimitL = Split(Limit_L2, ";")
            as_TempLimitH = Split(Limit_H2, ";")
            n_LimitCnt_L2 = UBound(as_TempLimitL) + 1
            n_LimitCnt_H2 = UBound(as_TempLimitH) + 1
            If n_PinGroupCnt = n_LimitCnt_L2 And n_PinGroupCnt = n_LimitCnt_H2 Then
                ReDim ad_FlowLimit_L2(n_PinGroupCnt - 1)
                ReDim ad_FlowLimit_H2(n_PinGroupCnt - 1)
                For i = 0 To n_PinGroupCnt - 1
                    ad_FlowLimit_L2(i) = CDbl(as_TempLimitL(i))
                    ad_FlowLimit_H2(i) = CDbl(as_TempLimitH(i))
                Next i
                b_isExist2ndLimit = True
                b_isLimitCntSame = True
            Else
                b_isLimitCntSame = False
            End If
        Else
            b_isLimitCntSame = False
        End If
    Else
        If (n_AllPinCnt = n_LimitCnt_L) And (n_AllPinCnt = n_LimitCnt_H) Then
            b_isExist2ndLimit = False
            b_isLimitCntSame = True
        ElseIf (n_AllPinCnt = (n_LimitCnt_L / 2)) And (n_AllPinCnt = (n_LimitCnt_H / 2)) Then
            '2 group flow limit
            '====
                'Use-Limit Pin_A  Lolimit(CompareOpen)  Hilimit(CompareOpen)
                'Use-Limit Pin_A  Lolimit(CompareShort) Hilimit(CompareShort)
                '--
                'Use-Limit Pin_B  Lolimit(CompareOpen)  Hilimit(CompareOpen)
                'Use-Limit Pin_B  Lolimit(CompareShort) Hilimit(CompareShort)
                '--
                'Use-Limit Pin_C  Lolimit(CompareOpen)  Hilimit(CompareOpen)
                'Use-Limit Pin_C  Lolimit(CompareShort) Hilimit(CompareShort)
            '====
            ReDim ad_FlowLimit_L((n_LimitCnt_L / 2) - 1)
            ReDim ad_FlowLimit_H((n_LimitCnt_H / 2) - 1)
            ReDim ad_FlowLimit_L2((n_LimitCnt_L / 2) - 1)
            ReDim ad_FlowLimit_H2((n_LimitCnt_H / 2) - 1)
            For i = 0 To (n_LimitCnt_L / 2) - 1
                ad_FlowLimit_L(i) = Lolimit_new(i * 2)  'Compare OPEN limit
                ad_FlowLimit_H(i) = HiLimit_new(i * 2)  'Compare OPEN limit
                ad_FlowLimit_L2(i) = Lolimit_new(1 + i * 2)  'Compare SHORT limit
                ad_FlowLimit_H2(i) = HiLimit_new(1 + i * 2)  'Compare SHORT limit
            Next i
            
            b_isExist2ndLimit = True
            b_isLimitCntSame = True
        Else
            b_isLimitCntSame = False
        End If
    End If
    '====Check count with pin and limit====
    '====If count different binout====
    If b_isLimitCntSame = False Then
        TheExec.Flow.TestLimit resultVal:=1, lowVal:=-1, hiVal:=-1, unit:=unitNone, Tname:=TheExec.DataManager.instancename
        s_ErrorMsg = "The Count of pins and Limit isn't match."
        Call Print_Error_Message(Error_Warning_Info.Warning_Info, "VBT_LIB_DC_Conti", "AutoZ_Continuity", s_ErrorMsg)
        Exit Function
    End If
    '====If count different binout====
    TheHdw.Digital.ApplyLevelsTiming True, True, True, tlPowered

    pinsequence = 0
    
    If connect_all_pins <> "" Then
        Call PPMUInit(connect_all_pins, 0#, "v")       '20230406 add sub function to init ppmu pin
    End If
    
    Call instSet.Initialize(PinAryToString(PinGroupName))
    ' connect all digital pins and force 0 voltage
    Call instSet.ApplyPower(0, fVoltage, gOn, Connect)
    '20230406 add sub function to init ppmu pin
    Call PPMUInit(instSet.GetIOPins, 0#, "v")
    
    For LoopIndex = 0 To GroupNum
    
        If isTestWalkingZResult Then
            If ReTestFlag(LoopIndex) = True Then
                Pins = Split(RetestPinNameAry(LoopIndex), ",")
            End If
        Else
            Pins = dic_LoopPinGroup(LoopIndex)
        End If
        
        CustomizeTname = dic_CustomizeTname(LoopIndex)
        
        If (isTestWalkingZResult And ReTestFlag(LoopIndex) = True) Or (isTestWalkingZResult = False) Then
        
            force_i = ForceI(LoopIndex)
            LowLimit = Low(LoopIndex)
            hiLimit = High(LoopIndex)
            
            ''' Pin in each group
            For Each DUTPin In Pins
                PPMUMeasure.AddPin (DUTPin)
            
                'FIMV and meter read
                If gl_GetInstrument_Dic.Exists(LCase(CStr(DUTPin))) Then
                    PPMUMeasure.Pins(DUTPin) = Measure_Pin_FIMV(CStr(DUTPin), force_i, PN_Disconnect)
                End If
                    
                'offline mode simulation
                If TheExec.TesterMode = testModeOffline Then
                    For Each site In TheExec.sites
                        If LCase(TheExec.DataManager.instancename) Like "*neg*" Or force_i < 0 Then PPMUMeasure.Pins(DUTPin).value(site) = -0.5 ' Use fixed value for easily checking offline mode.
                        If LCase(TheExec.DataManager.instancename) Like "*pos*" Or force_i > 0 Then PPMUMeasure.Pins(DUTPin).value(site) = 0.5 ' Use fixed value for easily checking offline mode.
                    Next site
                End If
            Next DUTPin
            
            If TestLimitMode = tlForceFlow Then
                tmpName = "ALLPIN"
            Else
                If force_i < 0 Then
                    tmpName = "IONEG"
                Else
                    tmpName = "IOPOS"
                End If
            End If
            CustomizeTname = Replace(CustomizeTname, "<MeasureCase>", tmpName)
            
            For Each DUTPin In Pins
            
                If gl_GetInstrument_Dic.Exists(LCase(CStr(DUTPin))) Then 'if N/C jump next Pin
            
                    PinNameCustomize = CStr(DUTPin)
                    PinNameCustomize = Replace(PinNameCustomize, "_", "")
                    
                    TmpTname = Replace(CustomizeTname, "<PinName>", PinNameCustomize)
                    
                    ''--------------------------------------------------------------------------
                    ''The test limit is from argument.
                    ''AP\RF - one test limit set(HiLimit\LowLimit of H\L.
                    ''LCD - one set test limit for "OPEN". one set test limit for "SHORT"
                    ''--------------------------------------------------------------------------
                    ''--------------------------------------------------------------------------
                    ''The test limit is from flow
                    ''--------------------------------------------------------------------------
                    If b_isExist2ndLimit = True Then    'Have 2nd limit to test
                        '====Open case====
                        TmpTname = Replace(TmpTname, "<PinSequence>", CStr(pinsequence))
                        TmpTname = Replace(TmpTname, "<OpenShort>", "OPEN")
                        If Name_Flag = 1 Then TmpTname = Replace(TmpTname, "OPEN", "OPEN-EOP")
                        If TestLimitMode = tlForceFlow Then
                            hiLimit = ad_FlowLimit_H(pinsequence)
                            LowLimit = ad_FlowLimit_L(pinsequence)
                        End If
                        TheExec.Flow.TestLimit resultVal:=PPMUMeasure.Pins(DUTPin), lowVal:=LowLimit, hiVal:=hiLimit, scaletype:=scaleNone, unit:=unitVolt, formatStr:="%.3f", Tname:=TmpTname, ForceVal:=force_i, ForceUnit:=unitAmp, ForceResults:=tlForceNone
                        Call GlbCustomizeSet.SetContiFlagTwoLimit(hiLimit, LowLimit, PPMUMeasure.Pins(DUTPin), Flag_Open)
                        '====Open case====
                        
                        '====Short case====
                        PinNameCustomize = CStr(DUTPin)
                        PinNameCustomize = Replace(PinNameCustomize, "_", "-")
                        TmpTname = Replace(CustomizeTname, "<PinName>", PinNameCustomize)
                        TmpTname = Replace(TmpTname, "<PinSequence>", CStr(pinsequence))
                        TmpTname = Replace(TmpTname, "<OpenShort>", "SHORT")
                        If Name_Flag = 1 Then Tname = Replace(Tname, "SHORT", "SHORT-EOP")
            
                        If TestLimitMode = tlForceFlow Then
                            HiLimit_2nd = ad_FlowLimit_H2(pinsequence)
                            LowLimit_2nd = ad_FlowLimit_L2(pinsequence)
                        Else
                            HiLimit_2nd = ad_FlowLimit_H2(LoopIndex)
                            LowLimit_2nd = ad_FlowLimit_L2(LoopIndex)
                        End If
                        TheExec.Flow.TestLimit resultVal:=PPMUMeasure.Pins(DUTPin), lowVal:=LowLimit_2nd, hiVal:=HiLimit_2nd, scaletype:=scaleNone, unit:=unitVolt, formatStr:="%.3f", Tname:=TmpTname, ForceVal:=force_i, ForceUnit:=unitAmp, ForceResults:=tlForceNone
                        Call GlbCustomizeSet.SetContiFlagTwoLimit(HiLimit_2nd, LowLimit_2nd, PPMUMeasure.Pins(DUTPin), Flag_Short)
                        '====Short case====
                    Else
                        TmpTname = Replace(TmpTname, "<PinSequence>", CStr(pinsequence))
                        If TestLimitMode = tlForceFlow Then
                            hiLimit = HiLimit_new(pinsequence)
                            LowLimit = Lolimit_new(pinsequence)
                        End If
                        TheExec.Flow.TestLimit resultVal:=PPMUMeasure.Pins(DUTPin), lowVal:=LowLimit, hiVal:=hiLimit, scaletype:=scaleNone, unit:=unitVolt, formatStr:="%.3f", Tname:=TmpTname, ForceVal:=force_i, ForceUnit:=unitAmp, ForceResults:=tlForceNone
                        Call GlbCustomizeSet.SetContiFlagOnelimit(hiLimit, LowLimit, PPMUMeasure.Pins(DUTPin), Flag_Open, Flag_Short)
                    End If
                End If
                pinsequence = pinsequence + 1
            Next DUTPin
            
            Set PPMUMeasure = Nothing
        End If
    Next
        
    If connect_all_pins <> "" Then
        With TheHdw.PPMU.Pins(connect_all_pins)
            '.ForceV 0#
            .Gate = tlOff
            .Disconnect
        End With
    End If
    
    ' Disconnect and gate off measured pins
    Call instSet.ApplyPower(0, fNone, gOff, Disconnect) 'michael fixed reset setting 20240527
    gldic_ComposeFailPins.RemoveAll
    
    DebugPrintFunc ""
    
Exit Function
errHandler: 'Add ErrHandler 2023/08/18
    Call Print_Error_Message(Error_Info, "VBT_LIB_DC_Conti", "AutoZ_Continuity") 'Add ErrHandler 2023/08/18
    If isDebugMode Then TheExec.AddOutput "Error in Continuity"
    If AbortTest Then Exit Function Else Resume Next 'Add ErrHandler 2023/08/18
End Function




' [20230620][T-All] move form hardIP to DC_Conti
' [20231003][All][Tank] modify after Chihome review
Public Function PPMU_Impedance_Function(Sink_Groups As String, Meas_Groups As String, Sink1_Current As Double, Sink2_Current As Double, LowLimit As Double, hiLimit As Double) As Long
  
    Dim Sink_Groups_Array() As String
    Dim Sink_Groups_Num As Double
    Dim Meas_Groups_Array() As String
    Dim Meas_Groups_Num As Double
    Dim ResultPower1 As New PinListData
    Dim ResultPower2 As New PinListData

    Dim site As Variant
    Dim Calculate_Contact_R As New PinListData
    Dim sTotalPinList As String     '20230406 Combine Sink_Groups and Meas_Groups to process together
    
    '210407 resistance monitor
    Dim Delta_R_result As New PinListData
    Dim i As Long
    
    Dim RAK_pins As Variant
    Dim LookUpResult As Variant

    On Error GoTo errHandler

    glb_TestInstance = vbNullString
    glb_TestInstance = UCase(TheExec.DataManager.instancename)
    
    Sink_Groups_Array = Split(Sink_Groups, ",")
    Sink_Groups_Num = UBound(Sink_Groups_Array)
    
    Meas_Groups_Array = Split(Meas_Groups, ",")
    Meas_Groups_Num = UBound(Meas_Groups_Array)
    
    If Sink_Groups_Num <> Meas_Groups_Num Then
        TheExec.Datalog.WriteComment "None Match Pin Num"
    Else
        TheHdw.Digital.ApplyLevelsTiming True, True, True
        'DisconnectVDDCA 'SEC DRAM
        TheHdw.Wait 0.001
        
        sTotalPinList = Sink_Groups & "," & Meas_Groups

        TheHdw.Digital.Pins(sTotalPinList).Disconnect
        
        TheHdw.PPMU.Pins(sTotalPinList).Connect
        
    '========================Sink1======================
        'TheHdw.PPMU.Pins(Sink_Groups).ForceV 0, 0.02       '20230406 no need to ForceV to 0v before ForceI
        TheHdw.PPMU.Pins(Meas_Groups).ForceI 0, 0.00002
        
        TheHdw.PPMU.Pins(sTotalPinList).Gate = tlOn
        
        TheHdw.Wait 0.0005
        TheHdw.PPMU.Pins(Sink_Groups).Gate = tlOff      '20230406 if check ppmu force type(V->I) need to gate off?
        TheHdw.PPMU.Pins(Sink_Groups).ForceI Sink1_Current, Sink1_Current
        TheHdw.PPMU.Pins(Sink_Groups).Gate = tlOn
        TheHdw.Wait 0.0001
        ResultPower1 = TheHdw.PPMU(Meas_Groups).Read(tlPPMUReadMeasurements)
        TheExec.Flow.TestLimit resultVal:=ResultPower1, hiVal:=0.1, unit:=unitVolt, ForceVal:=Sink1_Current, ForceUnit:=unitAmp, Tname:="ForceI1"
        
    '========================Sink2======================
        TheHdw.PPMU.Pins(Sink_Groups).ForceI Sink2_Current, Sink2_Current
        TheHdw.Wait 0.0005
        ResultPower2 = TheHdw.PPMU(Meas_Groups).Read(tlPPMUReadMeasurements)
        TheExec.Flow.TestLimit resultVal:=ResultPower2, hiVal:=1, unit:=unitVolt, ForceVal:=Sink2_Current, ForceUnit:=unitAmp, Tname:="ForceI2"
'
        TheHdw.PPMU.Pins(sTotalPinList).Gate = tlOff

        TheHdw.PPMU.Pins(sTotalPinList).Disconnect
        
        TheHdw.Wait 0.001
          
        Calculate_Contact_R = ResultPower1.Math.Subtract(ResultPower2)
        Calculate_Contact_R = Calculate_Contact_R.Math.divide(Sink1_Current - Sink2_Current).Abs
        
        If gl_Flag_1st_contact_R = False Then 'Instance_name Like "*1ST" Then
            First_Contact_R = Calculate_Contact_R.Copy
            gl_Flag_1st_contact_R = True
        End If
        
        If SheetExist = True Then
            For Each RAK_pins In Calculate_Contact_R.Pins
                'Check measure pin existing in RAK table
                LookUpResult = Filter(Path_2RAK1_PinList, RAK_pins)
                If UBound(LookUpResult) <> -1 Then
                    If TheExec.TesterMode = testModeOffline Then
                        Calculate_Contact_R.Pins(RAK_pins) = Calculate_Contact_R.Pins(RAK_pins).Subtract(Path_2RAK1.Pins(RAK_pins)).Abs
                    Else
                        Calculate_Contact_R.Pins(RAK_pins) = Calculate_Contact_R.Pins(RAK_pins).Subtract(Path_2RAK1.Pins(RAK_pins))
                    End If
                Else
                    TheExec.Datalog.WriteComment "Warning! pin " & RAK_pins & " does not exist in RAK2Path1 table"
                End If
            Next
        End If
       
        'Looping instance for calucating delta R
        If TheExec.Flow.enableWord("Looping_PowerSense_Impedance") = True Then    'Instance_name Like "*LOOPING" Then
            Delta_R_result = Calculate_Contact_R.Math.Subtract(First_Contact_R).Abs
            For i = 0 To Calculate_Contact_R.Pins.Count - 1
                TheExec.Flow.TestLimit resultVal:=Delta_R_result.Pins(i), unit:=unitCustom, customUnit:="Ohm", Tname:="Delta_Resistance", ForceResults:=tlForceFlow
            Next i
            TheHdw.Wait 1  'wait 1 sec
        End If
        
        TheExec.Flow.TestLimit resultVal:=Calculate_Contact_R, lowVal:=LowLimit, hiVal:=hiLimit, unit:=unitCustom, customUnit:="Ohm", Tname:="Contact_R"
        DebugPrintFunc vbNullString
        
    End If
    
    Exit Function
errHandler:
    Call Print_Error_Message(Error_Info, "VBT_LIB_DC_Conti", "PPMU_Impedance_Function")
    If AbortTest Then Exit Function Else Resume Next
End Function

'[20230731][T-Cai][YH] Add new function to check efuse power pin resistance
' [20231003][All][Tank] modify after Chihome review
Public Function PowerPin_CRES_Calculation(Power_pins As PinList, force_i1 As Double, force_i2 As Double, LowLimit As Double, hiLimit As Double, CresLowLimit As Double, CresHiLimit As Double, TestLimitMode As tlLimitForceResults, _
                                            Optional connect_all_pins As String, Optional Flag_Open As String = "F_open", Optional Flag_Short As String = "F_short", Optional Flag_CRES_Fail As String = "F_CRES_Fail") As Long
On Error GoTo errHandler
    Dim DUTPin As Variant
    Dim Pins() As String, Pin_Cnt As Long
    Dim site As Variant
    
    Dim TmpTname As String
    Dim MeasTname As String
    Dim TmpPinName As String
    Dim Actual_forceI1 As Double
    Dim Actual_forceI2 As Double
    
    Dim PowerPin_MeasV1 As New PinListData
    Dim PowerPin_MeasV2 As New PinListData
    
    Dim CRES_CalculationResult As New PinListData
    Dim CRES_InstanceName As String
    
    Dim b_isGetError As Boolean
    Dim s_ErrorMsg As String
    
    Dim FlowLimitObj As IFlowLimitsInfo
    Dim Lolimit_str() As String
    Dim Hilimit_str() As String
    Dim n As Long
    Dim s_TempNC As String
    
    s_TempNC = vbNullString
    
    If TestLimitMode = tlForceFlow Then
        Call TheExec.Flow.GetTestLimits(FlowLimitObj)
        Call FlowLimitObj.GetLowLimits(Lolimit_str)
        Call FlowLimitObj.GetHighLimits(Hilimit_str)
    End If
    
    If Flag_Short = "" Then Flag_Short = "F_short"
    If Flag_Open = "" Then Flag_Open = "F_open"
    If Flag_CRES_Fail = "" Then Flag_CRES_Fail = "F_CRES_Fail"
    Set GlbCustomizeSet = New CustomizeSetting
    
    b_isGetError = False
    
    TheHdw.Digital.ApplyLevelsTiming True, True, True, tlPowered
    TheHdw.Wait 0.001

    '//set connent_all_pins to 0v
    If connect_all_pins <> "" Then
        Call PPMUInit(connect_all_pins, 0#, "v")       '20230406 add sub function to init ppmu pin
    End If
    TheExec.DataManager.DecomposePinList Power_pins, Pins(), Pin_Cnt
    
    If TestLimitMode = tlForceFlow Then
        '//check flow use limit count the same with measure pin
        If UBound(Hilimit_str) = UBound(Pins) And UBound(Lolimit_str) = UBound(Pins) Then
        Else
            b_isGetError = True
        End If
    End If
    
    If b_isGetError = False Then
        For Each DUTPin In Pins
            CRES_CalculationResult.AddPin (DUTPin)
            PowerPin_MeasV1.AddPin (DUTPin)
            PowerPin_MeasV2.AddPin (DUTPin)
            
            TmpPinName = Replace(DUTPin, "_", "")
            TmpTname = "HAC_MeasV_PowerPin_CRES_Continuity_" & TmpPinName & "_AllDigital_X_X"
    
            '//check pin exist
            If gl_GetInstrumentType_Dic.Exists(LCase(DUTPin)) Then
                '//check is DCVS or not
                If UCase(gl_GetInstrumentType_Dic(LCase(DUTPin))) Like "*DCVS*" Then
                    '20210416, add for UFP
                    If glb_TesterType = "Jaguar" Then
        
                        '' First Point ''
                        MeasTname = TmpTname & "_FirstPoint"
                        '// set FoltLimit to let FIMV use
                        
                        'FIMV and meter read
                        PowerPin_MeasV1.Pins(DUTPin) = Measure_Pin_FIMV(CStr(DUTPin), force_i1)

                        'Get actural forceI
                        b_isGetError = Calc_ActuralClampCurrent(CStr(DUTPin), force_i1, Actual_forceI1, s_ErrorMsg)
                        
                        If b_isGetError Then
                            s_ErrorMsg = "Use not define instrument!!!"
                            Exit For
                        End If
                        
                        TheExec.Flow.TestLimit resultVal:=PowerPin_MeasV1.Pins(DUTPin), lowVal:=LowLimit, hiVal:=hiLimit, scaletype:=scaleMilli, unit:=unitVolt, formatStr:="%.3f", Tname:=MeasTname, ForceVal:=Actual_forceI1, ForceUnit:=unitAmp, ForceResults:=tlForceNone
                        'Judge failed open or failed short for Voltage Measurement
                        Call GlbCustomizeSet.SetContiFlagOnelimit(hiLimit, LowLimit, PowerPin_MeasV1.Pins(DUTPin), Flag_Open, Flag_Short)
                        
                        With TheHdw.DCVS.Pins(DUTPin)
                            .Voltage.value = 0#
                        End With
                        
                        '' Second Point ''
                        MeasTname = TmpTname & "_SecondPoint"
                        
                        'FIMV and meter read
                        PowerPin_MeasV2.Pins(DUTPin) = Measure_Pin_FIMV(CStr(DUTPin), force_i2)
                        
                        'Get actural forceI
                        b_isGetError = Calc_ActuralClampCurrent(CStr(DUTPin), force_i2, Actual_forceI2, s_ErrorMsg)
                        
                        If b_isGetError Then
                            s_ErrorMsg = "Use not define instrument!!!"
                            Exit For
                        End If
                        
                        TheExec.Flow.TestLimit resultVal:=PowerPin_MeasV2.Pins(DUTPin), lowVal:=LowLimit, hiVal:=hiLimit, scaletype:=scaleMilli, unit:=unitVolt, formatStr:="%.3f", Tname:=MeasTname, ForceVal:=Actual_forceI2, ForceUnit:=unitAmp, ForceResults:=tlForceNone
                        'Judge failed open or failed short for Voltage Measurement
                        Call GlbCustomizeSet.SetContiFlagOnelimit(hiLimit, LowLimit, PowerPin_MeasV2.Pins(DUTPin), Flag_Open, Flag_Short)
                        
                        'Efuse power pin need to gate off and disconnect after test
                        With TheHdw.DCVS.Pins(DUTPin)
                            .Gate = tlOff
                            .Disconnect
                        End With
                    
                    
                    ElseIf glb_TesterType = "UltraFLEXplus" Then
        
                        ''First Point''
                        MeasTname = TmpTname & "_FirstPoint"
                        
                        'FIMV and meter read
                        PowerPin_MeasV1.Pins(DUTPin) = Measure_Pin_FIMV(CStr(DUTPin), force_i1)
                        
                        'Get actural forceI
                        b_isGetError = Calc_ActuralClampCurrent(CStr(DUTPin), force_i1, Actual_forceI1, s_ErrorMsg)
                        
                        If b_isGetError Then
                            s_ErrorMsg = "Use not define instrument!!!"
                            Exit For
                        End If
                        
                        TheExec.Flow.TestLimit resultVal:=PowerPin_MeasV1.Pins(DUTPin), lowVal:=LowLimit, hiVal:=hiLimit, scaletype:=scaleMilli, unit:=unitVolt, formatStr:="%.3f", Tname:=MeasTname, ForceVal:=Actual_forceI1, ForceUnit:=unitAmp, ForceResults:=tlForceNone
                        'Judge failed open or failed short for Voltage Measurement
                        Call GlbCustomizeSet.SetContiFlagOnelimit(hiLimit, LowLimit, PowerPin_MeasV1.Pins(DUTPin), Flag_Open, Flag_Short)
                        
                        With TheHdw.DCVS.Pins(DUTPin)
                            .mode = tlDCVSModeVoltage
                            .Voltage.value = 0#
                        End With
                        
                        ''Second Point''
                        MeasTname = TmpTname & "_SecondPoint"
                        
                        'FIMV and meter read
                        PowerPin_MeasV2.Pins(DUTPin) = Measure_Pin_FIMV(CStr(DUTPin), force_i2)
                        
                        'Get actural forceI
                        b_isGetError = Calc_ActuralClampCurrent(CStr(DUTPin), force_i2, Actual_forceI2, s_ErrorMsg)
                        
                        If b_isGetError Then
                            s_ErrorMsg = "Use not define instrument!!!"
                            Exit For
                        End If
                        
                        ''Store Value for CRES Calculation''
                        TheExec.Flow.TestLimit resultVal:=PowerPin_MeasV2.Pins(DUTPin), lowVal:=LowLimit, hiVal:=hiLimit, scaletype:=scaleMilli, unit:=unitVolt, formatStr:="%.3f", Tname:=MeasTname, ForceVal:=Actual_forceI2, ForceUnit:=unitAmp, ForceResults:=tlForceNone
                        'Judge failed open or failed short for Voltage Measurement
                        Call GlbCustomizeSet.SetContiFlagOnelimit(hiLimit, LowLimit, PowerPin_MeasV2.Pins(DUTPin), Flag_Open, Flag_Short)
                        
                        'Efuse power pin need to gate off and disconnect after test
                        With TheHdw.DCVS.Pins(DUTPin)
                            .Gate = tlOff
                            .Disconnect
                        End With
                        
                    End If
                Else
                    '//if not DCVS binout
                    b_isGetError = True
                    s_ErrorMsg = "Please check measure pin not in DCVS insturment !!!"
                    Exit For
                End If
            Else
                '//if N/A or not exist store pin name
                s_TempNC = CombineStringList(s_TempNC, CStr(DUTPin))
            End If
    
        Next DUTPin
    End If
    
    If connect_all_pins <> "" Then
        With TheHdw.PPMU.Pins(connect_all_pins)
            '.ForceV 0#
            .Gate = tlOff
            .Disconnect
        End With
    End If
    
    If b_isGetError = False Then
        
        '//when offline set result to low limit let test pass  'offline mode simulation
        If TheExec.TesterMode = testModeOffline Then
            For n = 0 To UBound(Pins)
                If TestLimitMode = tlForceFlow Then
                    CRES_CalculationResult.Pins(Pins(n)) = Lolimit_str(n)
                Else
                    CRES_CalculationResult.Pins(Pins(n)) = CresLowLimit
                End If
            Next n
        Else
            ''CRES Calculation''
            CRES_CalculationResult = PowerPin_MeasV2.Math.Subtract(PowerPin_MeasV1)     '// V2 - V1
            CRES_CalculationResult = CRES_CalculationResult.Math.divide(Actual_forceI2 - Actual_forceI1).Abs        '// resultR = (V2 - V1) / (I2 - I1)
        End If
    
            
        For Each DUTPin In Pins
            
            CRES_InstanceName = "HAC_CalcR_PowerPin_CRES_Continuity_" & Replace(DUTPin, "_", "") & "_ForceCurrent_N" & Format(Abs(Actual_forceI1) * 1000, "0") & "mA_N" & Format(Abs(Actual_forceI2) * 1000, "0") & "mA"
            
            If TestLimitMode = tlForceFlow Then
                TheExec.Flow.TestLimit resultVal:=CRES_CalculationResult.Pins(DUTPin), unit:=unitCustom, customUnit:="Ohm", Tname:=CRES_InstanceName, ForceResults:=tlForceFlow
            ElseIf TestLimitMode = tlForceNone Then
                TheExec.Flow.TestLimit resultVal:=CRES_CalculationResult.Pins(DUTPin), lowVal:=CresLowLimit, hiVal:=CresHiLimit, unit:=unitCustom, customUnit:="Ohm", Tname:=CRES_InstanceName
                Call GlbCustomizeSet.SetContiFlagOnelimit(CresHiLimit, CresLowLimit, CRES_CalculationResult.Pins(DUTPin), Flag_CRES_Fail, Flag_CRES_Fail)
            End If
            
        Next DUTPin
    Else
        '//if get error just binout
        For Each site In TheExec.sites
            TheExec.Flow.TestLimit resultVal:=-1, lowVal:=1, hiVal:=1, unit:=unitCustom, Tname:=CRES_InstanceName
        Next site
        Call Print_Error_Message(Warning_Info, "VBT_LIB_DC_Conti", "PowerPin_CRES_Calculation_new", s_ErrorMsg)
    End If
    
    If s_TempNC <> "" Then
        s_ErrorMsg = "PowerPin_CRES_Calculation Exist N/C pin : " & s_TempNC
        Call Print_Error_Message(Warning_Info, "VBT_LIB_DC_Conti", "PowerPin_CRES_Calculation_new", s_ErrorMsg)
    End If
    
    DebugPrintFunc ""

    Exit Function
errHandler:
    Call Print_Error_Message(Error_Info, "VBT_LIB_DC_Conti", "PowerPin_CRES_Calculation")
    If AbortTest Then Exit Function Else Resume Next
End Function
