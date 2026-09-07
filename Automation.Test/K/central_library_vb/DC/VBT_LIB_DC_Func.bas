Attribute VB_Name = "VBT_LIB_DC_Func"
Option Explicit

Public Type GetShmooVihVil
    Vih As New SiteDouble
    Vil As New SiteDouble
End Type

Public Type Pin_parameter
    Voltage As Double
    Start_voltage As Double
    End_voltage As Double
End Type
Public Type Pins_detail
    PinName As String
    Step_value As Double
    Step_value_up As Double
    Step_value_down As Double
    Start_voltage As Double
    Start_voltage_up As Double
    Start_voltage_down As Double
    Pin_rise As Boolean
    LatchUp_Final_Value As Double
    Gate_check As New SiteBoolean
    PA_G1G2_Print_Enable As Boolean  'add for PA power pin print 20210507
    Stress_Grp As Integer ''(Only 1 or 2)
    P_Consumption As New SiteDouble
    Stress_C As New SiteDouble
End Type

Public GPIO_Vih_Vil(100) As GetShmooVihVil

Public g_CFG_GPIO_PF_Val_LV As New PinListData '' 0 is pass, 1 is fail
Public g_CFG_GPIO_PF_Val_NV As New PinListData '' 0 is pass, 1 is fail
Public g_CFG_GPIO_PF_Val_HV As New PinListData '' 0 is pass, 1 is fail
Public GPIO_driver_result As New SiteDouble
Public F_EVS_Invalid_Category As Boolean
Public F_start_profile As Boolean

'Revision History:
'V0.0 initial bring up

' This module should be used for VBT Tests.  All functions in this module
' will be available to be used from the Test Instance sheet.
' Additional modules may be added as needed (all starting with "VBT_").
'
' The required signature for a VBT Test is:
'
' Public Function FuncName(<arglist>) As Long
'    where <arglist> is any list of arguments supported by VBT Tests.
'
' See online help for supported argument types in VBT Tests.
'
'
' It is highly suggested to use error handlers in VBT Tests.  A sample
' VBT Test with a suggeseted error handler is shown below:
'
' Function FuncName() As Long
'     On Error GoTo errHandler
'
'     Exit Function
' errHandler:
'     If AbortTest Then Exit Function Else Resume Next
' End Function
Public Function DC_Func_WriteFuncResult(Optional PerPinFailLog As Boolean = False) As Long
On Error GoTo errHandler 'Add ErrHandler 2023/08/18
    Dim site As Variant
    Dim TestNumber As Long
    Dim FailCount As New PinListData
    Dim AllPins As PinList
    Dim pin As Variant
    Dim Pins() As String
    Dim Pin_Cnt As Long
    
     TheExec.DataManager.DecomposePinList "All_Digital", Pins(), Pin_Cnt
    'AllPins = "AllPins"
    ' allpins.Value=
  
    For Each site In TheExec.sites
        TestNumber = TheExec.sites.item(site).TestNumber
        If TheHdw.Digital.Patgen.PatternBurstPassed(site) Then  'collect pattern burst result
            Call TheExec.Datalog.WriteFunctionalResult(site, TestNumber, logTestPass)
        Else
            Call TheExec.Datalog.WriteFunctionalResult(site, TestNumber, logTestFail)
            TheExec.sites.item(site).testResult = siteFail

            
            If PerPinFailLog Then    'per pin fail log collection
                For Each pin In Pins
                'thehdw.Digital.Pins(pin)
                    If TheExec.DataManager.ChannelType(pin) <> "N/C" Then
                        FailCount = TheHdw.Digital.Pins(pin).FailCount
                        If FailCount <> 0 Then TheExec.Datalog.WriteComment "===> Pin " & pin & " Fail count =" & FailCount
                    End If
                Next pin
            End If
        End If
        TheExec.sites.item(site).TestNumber = TestNumber + 1
    Next site

Exit Function 'Add ErrHandler 2023/08/18
errHandler: 'Add ErrHandler 2023/08/18
    Call Print_Error_Message(Error_Info, "VBT_LIB_DC_Func", "DC_Func_WriteFuncResult") 'Add ErrHandler 2023/08/18
    If AbortTest Then Exit Function Else Resume Next 'Add ErrHandler 2023/08/18
End Function


' [20231108][All][Carter] Restore delete function
Public Function IO_HardIP_PPMU_Measure_I_TTR(TestPinArrayIV() As String, _
                                             TestSeqNum As Integer, _
                                             TestSeqNumIdx As Long, _
                                             ForceSequenceArray() As String, _
                                             k As Long, Pat As Variant, _
                                             Flag_SingleLimit As Boolean, _
                                             HighLimitVal As Double, _
                                             LowLimitVal As Double, _
                                             TestLimitPerPin_VIR As String, _
                                             TestIrange() As String, _
                                    Optional VDD_IO_1p2 As String = vbNullString, _
                                    Optional VDD_IO_1p8 As String = vbNullString, _
                                    Optional CUS_Str_MainProgram As String, _
                                    Optional PPMU_TestLimit_TTR As Boolean, _
                                    Optional MeasIGrpPinCnt_Str As String = vbNullString, _
                                    Optional CFG_GPIO_Pins As String, _
                                    Optional CFGTest_FirstSequence As Boolean, _
                                    Optional Rtn_MeasCurr As PinListData, Optional Temp_index As Long) As Long
On Error GoTo errHandler 'Add ErrHandler 2023/08/18
'VDD_IO_1p2 = "VDDIO12_GRP5"
'VDD_IO_1p8 = "VDDIO18_GRP"
    Dim funcName As String:: funcName = "IO_HardIP_PPMU_Measure_I_TTR"
    
    Dim MeasureValue As New PinListData
    Dim TestNum As Long
    Dim pin As Variant
    Dim pinG As Variant '20230111
    Dim p As Long
    Dim ForceV  As Double
    Dim MeasureValueGrp As New PinListData
    Dim PinName As String
    Dim Vdiff As Double


 ''========================================================================================
    '' 20150202 - Range Check
    Dim RangeCheck_HighLimitVal() As Double
    Dim RangeCheck_LowLimitVal() As Double
    If Range_Check_Enable_Word = True Then
        Call GetFlowSingleUseLimit(RangeCheck_HighLimitVal, RangeCheck_LowLimitVal)
    End If
    ''========================================================================================
    
    ''=========================================================================================================
    '' 20160108 - Add rule to cover force value is different
    Dim ForceByPin() As String
    Dim ForceValByPin() As String
''    Dim ForceValIdx As Integer
    Dim Measure_I_Range() As String
    Dim MeasurePin As String
    Dim MI_Range_Index As Long
    Dim i, j As Long
    Dim OutputTname As String
    
    
'    If LCase(TheExec.DataManager.InstanceName) Like "*lv*" Then
'        Vratio = 0.9
'    ElseIf LCase(TheExec.DataManager.InstanceName) Like "*hv*" Then
'        Vratio = 1.1
'    Else
'        Vratio = 1
'    End If
    
    MI_Range_Index = 0
    
    '' Force Pin
    If UBound(TestPinArrayIV) = 0 Then
        ForceByPin = Split(TestPinArrayIV(0), ",")
        MeasurePin = TestPinArrayIV(0)
        TheHdw.Digital.Pins(TestPinArrayIV(0)).Disconnect
    Else
        ForceByPin = Split(TestPinArrayIV(TestSeqNumIdx), ",")
        MeasurePin = TestPinArrayIV(TestSeqNumIdx)
        TheHdw.Digital.Pins(TestPinArrayIV(TestSeqNumIdx)).Disconnect
    End If
    
    '' Force Volt value
    If UBound(ForceSequenceArray) = 0 Then
        ForceValByPin = Split(ForceSequenceArray(0), ",")
    Else
        ForceValByPin = Split(ForceSequenceArray(TestSeqNumIdx), ",")
    End If
    
    '' Measure Current range
    If UBound(TestIrange) = 0 Then
        Measure_I_Range = Split(TestIrange(0), ",")
        Dim MeasurePinArry() As String
        MeasurePinArry = Split(MeasurePin, ",")    'expend I range for all point
        If UBound(Measure_I_Range) = 0 Then
            ReDim Preserve Measure_I_Range(UBound(MeasurePinArry))
            'ReDim Measure_I_Range(UBound(MeasurePinArry))
            For i = 0 To UBound(MeasurePinArry)
                Measure_I_Range(i) = Measure_I_Range(0)
            Next i
        End If
    Else
        Measure_I_Range = Split(TestIrange(TestSeqNumIdx), ",")
    End If
    
    ''=========================================================================================================
    '' 20150108 - Check number whether differrent between measure current range and force pin, add defalut value to let input number are the same.
    Call VIR_CheckTestCondition_Measure_I_R_Z("I", ForceByPin, Measure_I_Range)
     '' 20150111 - Check force value is the same or different
    ' Dim i As Long
    Dim b_ForceDiffVolt As Boolean
    Dim PastVal As Double
    b_ForceDiffVolt = False
    For i = 0 To UBound(ForceValByPin)
        If i <> 0 Then
            If ForceValByPin(i) <> PastVal Then
                b_ForceDiffVolt = True
                Exit For
            End If
        End If
        
        PastVal = (ForceValByPin(i)) '' use global para need to use the Evalute
    Next i
    '==== MeasIPingroup support per test seqeunce define ==== -- 20230109
    Dim MeasIGrpPinCnt  As Integer
    If MeasIGrpPinCnt_Str = "" Then
        MeasIGrpPinCnt = 1
    ElseIf InStr(MeasIGrpPinCnt_Str, ",") = 0 Then
        MeasIGrpPinCnt = CInt(MeasIGrpPinCnt_Str)
    Else
        MeasIGrpPinCnt = CInt(Split(MeasIGrpPinCnt_Str, ",")(TestSeqNum))
    End If
    
    Dim MeasIGrpCnt As Integer
    Dim TestPinArray() As String
    Dim PinCount As Long
    Dim TestPinStr As String
    Dim site As Variant 'Carter, 20240304
    For Each pin In ForceByPin
        TheExec.DataManager.DecomposePinList pin, TestPinArray, PinCount
    
        If PinCount Mod MeasIGrpPinCnt <> 0 Then
            MeasIGrpCnt = PinCount \ MeasIGrpPinCnt + 1
        Else
            MeasIGrpCnt = PinCount / MeasIGrpPinCnt
        End If
        
        If b_ForceDiffVolt = False Then
            If ForceValByPin(0) > 0.8 Then
'                If LCase(pin) Like "*1p2*" Then
                    Vdiff = TheHdw.DCVS.Pins(VDD_IO_1p2).Voltage.value - ForceValByPin(0) ''vddio12_grp!!!
'                Else
'                    Vdiff = TheHdw.DCVS.Pins(VDD_IO_1p8).Voltage.value - ForceValByPin(0) ''vddio18_grp!!!
'                End If
            Else
                Vdiff = ForceValByPin(0)
            End If
        Else
            If ForceValByPin(MI_Range_Index) > 0.8 Then
'                If LCase(pin) Like "*1p2*" Then
                    Vdiff = TheHdw.DCVS.Pins(VDD_IO_1p2).Voltage.value - ForceValByPin(MI_Range_Index)
'                Else
'                    Vdiff = TheHdw.DCVS.Pins(VDD_IO_1p8).Voltage.value - ForceValByPin(MI_Range_Index)
'                End If
            Else
                Vdiff = ForceValByPin(MI_Range_Index)
            End If
        End If

        For i = 0 To MeasIGrpCnt - 1
            TestPinStr = vbNullString
            
            For j = 0 To MeasIGrpPinCnt - 1
                If (i * MeasIGrpPinCnt + j) < PinCount Then
                    TestPinStr = TestPinStr & "," & TestPinArray(i * MeasIGrpPinCnt + j)
                    MeasureValue.AddPin (TestPinArray(i * MeasIGrpPinCnt + j))
                End If
            Next j
            TestPinStr = right(TestPinStr, Len(TestPinStr) - 1)
    
            With TheHdw.PPMU.Pins(TestPinStr)
            '' 20150615 - Force 0 mA before expected force value to solve over clamp issue.
                .ForceI 0, 0.05
                .Connect
                .Gate = tlOn
                If b_ForceDiffVolt = False Then
                    .ForceV ForceValByPin(0), Measure_I_Range(MI_Range_Index)
'                    If ForceValByPin(0) > 0.7 Then ''161219
'                        Vdiff = TheHdw.DCVS.Pins("VDDIO18_GRP").Voltage.Value - ForceValByPin(0) ''vddio18_grp!!!
'                    Else
'                        Vdiff = ForceValByPin(0)
'                    End If
                Else
                    .ForceV ForceValByPin(MI_Range_Index), Measure_I_Range(MI_Range_Index)
'                    If ForceValByPin(0) > 0.7 Then
'                        Vdiff = TheHdw.DCVS.Pins("VDDIO18_GRP").Voltage.Value - ForceValByPin(MI_Range_Index)
'                    Else
'                        Vdiff = ForceValByPin(MI_Range_Index)
'                    End If
                End If
                '' 20160108 - Only keep 1 force value but current range can be different for force pin
            End With
            
            If PPMU_TestLimit_TTR = False Then TheExec.Datalog.WriteComment "Pin = " & (TestPinStr & " Measure Current Range = " & TheHdw.PPMU.Pins(pin).MeasureCurrentRange)
            
            TheHdw.Wait (100 * us)
            
            MeasureValueGrp = TheHdw.PPMU.Pins(TestPinStr).Read(tlPPMUReadMeasurements, 10)
            
            For j = 0 To MeasIGrpPinCnt - 1
                If (i * MeasIGrpPinCnt + j) < PinCount Then
                    PinName = TestPinArray(i * MeasIGrpPinCnt + j)
                    
                    ' 20161207 Add RAK
                    MeasureValue.Pins(PinName) = MeasureValueGrp.Pins(PinName).Multiply(CurrentJob_Card_RAK.Pins(PinName)).Abs.Negate.Add(Vdiff).Invert.Multiply(Vdiff).Multiply(MeasureValueGrp.Pins(PinName))
                    If TheExec.TesterMode = testModeOffline Then MeasureValue.Pins(PinName) = 0.0005 + Rnd() * 0.0001
                End If
            Next j
            
            TheHdw.PPMU.Pins(TestPinStr).Disconnect
        Next i
        MI_Range_Index = MI_Range_Index + 1
    Next pin
    
    
    If UBound(ForceSequenceArray) <> 0 Then
        If ForceSequenceArray(TestSeqNum) = "" Then
            ForceSequenceArray(TestSeqNum) = 0
        End If
    End If
    
    
'''    TheHdw.Wait (100 * us)
    
    '' 20160112 - Comment this
''    MeasureValue = TheHdw.PPMU.Pins(TestPinArrayIV(TestSeqNumIdx)).Read(tlPPMUReadMeasurements, 10)
    DebugPrintFunc_PPMU CStr(MeasurePin)
    
    ''20150904 - Move to CUS_VIR_MainProgram_MeasureI
''    If CUS_Str_MainProgram <> "" Then
''        Call CUS_VIR_MainProgram_MeasureI(CUS_Str_MainProgram, VIR_MI_AFTER_MEASUREMENT, MeasureValue)
''    End If

    
    Dim TestNameInput As String
    TestNameInput = "Curr_meas_" + CStr(TestSeqNum)
        
    ''20151103 print force condition
    Call Print_Force_Condition("i", MeasureValue)
    
''    ''20160112 - Force value index for test limit if force voltage value is different
''    Dim ForceVal_Index As Long
''    ForceVal_Index = 0
                
    Flag_SingleLimit = True '''Driver strength test using single limit.

    If PPMU_TestLimit_TTR = True Then
        Dim Lowlimitval_temp As Double
        Dim Hilimitval_temp As Double
        TheExec.Flow.TestLimitIndex = TheExec.Flow.TestLimitIndex + 1
        
        Lowlimitval_temp = GetLowLimitFromFlow
        Hilimitval_temp = GetHiLimitFromFlow
        For Each site In TheExec.sites.Active
            For p = 0 To MeasureValue.Pins.Count - 1
                If MeasureValue.Pins(p).value > Hilimitval_temp Or MeasureValue.Pins(p).value < Lowlimitval_temp Then
                    If gl_UseStandardTestName_Flag = True Then                        'Roger add
                        TestNameInput = Report_TName_From_Instance("I", MeasureValue.Pins(p), OutputTname, TestSeqNum, p)
                    End If
                    TheExec.Datalog.WriteComment "Pin = " & (MeasureValue.Pins(p) & " Measure Current Range = " & TheHdw.PPMU.Pins(MeasureValue.Pins(p)).MeasureCurrentRange)
                    TheExec.Flow.TestLimit MeasureValue.Pins(p), Lowlimitval_temp, Hilimitval_temp, , , , unitAmp, , Tname:=TestNameInput, ForceResults:=tlForceNone
                End If
            Next p
        Next site

        'TheExec.Flow.TestLimitIndex = TheExec.Flow.TestLimitIndex + 1
    ElseIf Flag_SingleLimit = True Then
        If b_ForceDiffVolt = False Then

            For Each pinG In ForceByPin
                TheExec.DataManager.DecomposePinList pinG, TestPinArray, PinCount
                p = 0
                For Each pin In TestPinArray
                    p = p + 1
                    TheExec.Flow.TestLimitIndex = Temp_index  'restore
                    If EnableDigitalTestLimitTTR = True Then
                        TestNameInput = Report_TName_From_Instance("I", "X", OutputTname, CInt(Instance_Data.TestSeqNum), p)
                    Else
                        TestNameInput = Report_TName_From_Instance("I", MeasureValue.Pins(pin), OutputTname, TestSeqNum, p)
                    End If
                    '191223 CW for IO single limit-------
                    TheExec.Flow.TestLimit MeasureValue.Pins(pin), , , scaletype:=scaleNone, unit:=unitAmp, Tname:=TestNameInput, ForceVal:=FormatNumber(TheHdw.PPMU(MeasureValue.Pins(pin)).Voltage.value, 3), ForceUnit:=unitVolt, ForceResults:=tlForceFlow
                    '------------------------------------
                Next pin
                
                Temp_index = TheExec.Flow.TestLimitIndex 'return to next seq
            Next pinG

        Else
            For p = 0 To MeasureValue.Pins.Count - 1
                If gl_UseStandardTestName_Flag = True Then                        'Roger add
                    TestNameInput = Report_TName_From_Instance("I", MeasureValue.Pins(p), OutputTname, TestSeqNum, p)
                End If
                If p = 0 Then
                    TheExec.Flow.TestLimit MeasureValue.Pins(p), LowLimitVal, HighLimitVal, scaletype:=scaleNone, unit:=unitAmp, Tname:=TestNameInput, ForceVal:=FormatNumber(TheHdw.PPMU(MeasureValue.Pins(p)).Voltage.value, 3), ForceUnit:=unitVolt, ForceResults:=tlForceFlow
                Else
                    TheExec.Flow.TestLimit MeasureValue.Pins(p), GetLowLimitFromFlow, GetHiLimitFromFlow, scaletype:=scaleNone, unit:=unitAmp, Tname:=TestNameInput, ForceVal:=FormatNumber(TheHdw.PPMU(MeasureValue.Pins(p)).Voltage.value, 3), ForceUnit:=unitVolt, ForceResults:=tlForceNone
                End If
            
            Next p
            
            Temp_index = TheExec.Flow.TestLimitIndex 'return to next seq 20230111
            
        End If

        
    Else
        For p = 0 To MeasureValue.Pins.Count - 1
            If EnableDigitalTestLimitTTR = True Then
                TestNameInput = Report_TName_From_Instance("I", "X", OutputTname, TestSeqNum, p)
            Else
                TestNameInput = Report_TName_From_Instance("I", MeasureValue.Pins(p), OutputTname, TestSeqNum, p)
            End If
            TheExec.Flow.TestLimit MeasureValue.Pins(p), , , scaletype:=scaleNone, unit:=unitAmp, Tname:=TestNameInput, ForceVal:=FormatNumber(TheHdw.PPMU(MeasureValue.Pins(p)).Voltage.value, 3), ForceUnit:=unitVolt, ForceResults:=tlForceFlow
        Next p

    End If
    '----------------------------for store name 180525----------------------------20230111
    Set Rtn_MeasCurr = MeasureValue
    '----------------------------------------------------------------------------------------
  
Exit Function 'Add ErrHandler 2023/08/18
errHandler: 'Add ErrHandler 2023/08/18
    Call Print_Error_Message(Error_Info, "VBT_LIB_DC_Func", "IO_HardIP_PPMU_Measure_I_TTR") 'Add ErrHandler 2023/08/18
    If AbortTest Then Exit Function Else Resume Next 'Add ErrHandler 2023/08/18
End Function


' [20231108][All][Carter] Restore delete function
Public Function IO_HardIP_PPMU_Measure_V_TTR(TestPinArrayIV() As String, _
                                             TestSeqNum As Integer, _
                                             TestSeqNumIdx As Long, _
                                             ForceSequenceArray() As String, _
                                             k As Long, Pat As Variant, _
                                             Flag_SingleLimit As Boolean, _
                                             HighLimitVal As Double, _
                                             LowLimitVal As Double, _
                                             TestLimitPerPin_VIR As String, _
                                        ByRef ReturnMeasVolt As PinListData, _
                                    Optional InstSpecialSetting As InstrumentSpecialSetup = 0, _
                                    Optional RAK_Flag As Enum_RAK = 0, _
                                    Optional CUS_Str_MainProgram As String = vbNullString, _
                                    Optional MeasVGrpPinCnt As Integer) As Long
On Error GoTo errHandler 'Add ErrHandler 2023/08/18

    Dim MeasureValue As New PinListData
    Dim Force_idx As Integer
    Dim site As Variant
    Dim TestNum As Long
    Dim pin  As Variant
    
    Dim p As Long
    Dim ForceV  As Double
    Dim ForceByPin() As String
    Dim ForceValByPin() As String
    Dim ForceValIdx As Integer
    Dim IdxV As Integer
    Dim MeasurePinStr As String
    Dim i, j As Long
    Dim MeasureValueGrp As New PinListData
    Dim PinName As String
    Dim PinArr() As String
    Dim PinCount As Long
    Dim MeasVGrpCnt As Integer
    Dim TestPinStr As String
    Dim TestPinArrayIVsize As Long

    ''========================================================================================
    '' 20150202 - Range Check
    Dim RangeCheck_HighLimitVal() As Double
    Dim RangeCheck_LowLimitVal() As Double
    Dim TempMeasVal_PerPin(100) As New PinListData
    If Range_Check_Enable_Word = True Then
        Call GetFlowSingleUseLimit(RangeCheck_HighLimitVal, RangeCheck_LowLimitVal)
    End If
    
    If UBound(ForceSequenceArray) = 0 Then
        ForceValByPin = Split(ForceSequenceArray(0), ",")
    Else
        ForceValByPin = Split(ForceSequenceArray(TestSeqNumIdx), ",")
    End If
    
    ForceValIdx = 0
    
    If UBound(TestPinArrayIV) = 0 Then
        ForceByPin = Split(TestPinArrayIV(0), ",")
        MeasurePinStr = TestPinArrayIV(0)        '20160224 add to allow every seq with the same pins
        'TheHdw.Digital.Pins(TestPinArrayIV(0)).Disconnect
    Else
        ForceByPin = Split(TestPinArrayIV(TestSeqNumIdx), ",")
        MeasurePinStr = TestPinArrayIV(TestSeqNumIdx)        '20160224 add to allow every seq with the same pins
        'TheHdw.Digital.Pins(TestPinArrayIV(TestSeqNumIdx)).Disconnect
    End If
    '' 20150721 - Apply force I value from Stored_MeasI_PPMU,
    ''                - Can not coexist between stored value and force value at the same sequence
    ''
    Dim b_IsNumeral As Boolean
    Dim b_UseStoredForceVal As Boolean
    
    b_IsNumeral = ContentIsNumeral(ForceValByPin(0))
    If b_IsNumeral Then
        b_UseStoredForceVal = False
    Else
        b_UseStoredForceVal = True
    End If
    Dim ForceValI As Double
    If b_UseStoredForceVal = False Then '' 20150721 - Normal usage
        If MeasVGrpPinCnt > 0 And InstSpecialSetting <> InstrumentSpecialSetup.PPMU_SerialMeasurement Then '' Split pin group
            For Each pin In ForceByPin
                TheExec.DataManager.DecomposePinList pin, PinArr, PinCount
        
                If PinCount Mod MeasVGrpPinCnt <> 0 Then
                    MeasVGrpCnt = PinCount \ MeasVGrpPinCnt + 1
                Else
                    MeasVGrpCnt = PinCount / MeasVGrpPinCnt
                End If
                
                For i = 0 To MeasVGrpCnt - 1
                    TestPinStr = vbNullString
                    
                    For j = 0 To MeasVGrpPinCnt - 1
                        If (i * MeasVGrpPinCnt + j) < PinCount Then
                            TestPinStr = TestPinStr & "," & PinArr(i * MeasVGrpPinCnt + j)
                            MeasureValue.AddPin (PinArr(i * MeasVGrpPinCnt + j))
                        End If
                    Next j
                    TestPinStr = right(TestPinStr, Len(TestPinStr) - 1)
                
                    With TheHdw.PPMU.Pins(TestPinStr)
                        '' 20150615 - Force 0 mA before expected force value to solve over clamp issue.
                        .ForceI 0, 0
                        .Connect
                        .Gate = tlOn
                        
                        If UBound(ForceValByPin) = 0 Then
                            .ForceI ForceValByPin(0), ForceValByPin(0)
                            ForceValI = ForceValByPin(0)
                        ElseIf ForceValByPin(ForceValIdx) <> "" Then
                            .ForceI ForceValByPin(ForceValIdx), ForceValByPin(ForceValIdx)
                            ForceValI = ForceValByPin(ForceValIdx)
                        Else:
                            .ForceI 0
                            ForceValI = 0
                        End If
                    End With
                    
                    TheHdw.Wait (1 * ms)
                    DebugPrintFunc_PPMU CStr(TestPinStr)
                    MeasureValueGrp = TheHdw.PPMU.Pins(TestPinStr).Read(tlPPMUReadMeasurements, 10)
                    
                    For j = 0 To MeasVGrpPinCnt - 1
                        If (i * MeasVGrpPinCnt + j) < PinCount Then
                            PinName = PinArr(i * MeasVGrpPinCnt + j)
                            MeasureValue.Pins(PinName) = MeasureValueGrp.Pins(PinName)
                            If TheExec.TesterMode = testModeOffline Then MeasureValue.Pins(PinName) = 0.005 + Rnd() * 0.001
                        End If
                    Next j
                    
                    TheHdw.PPMU.Pins(TestPinStr).Disconnect
                Next i
                ForceValIdx = ForceValIdx + 1
            Next pin
        Else '' No split pin group
            '' 20150721 - Normal usage
            If InstSpecialSetting = InstrumentSpecialSetup.PPMU_SerialMeasurement Then
                TheExec.DataManager.DecomposePinList MeasurePinStr, PinArr, PinCount
                TheHdw.PPMU.Pins(MeasurePinStr).ForceI 0, 0
        
                For Each pin In PinArr
                    MeasureValue.AddPin (pin)
                    TheHdw.PPMU.Pins(pin).ForceI ForceValByPin(0), Abs(ForceValByPin(0))
                    ForceValI = ForceValByPin(0)
                    TheHdw.Wait 0.001
                    DebugPrintFunc_PPMU CStr(pin)
                    MeasureValue.Pins(pin) = TheHdw.PPMU.Pins(pin).Read(tlPPMUReadMeasurements, 10)
                    TheHdw.PPMU.Pins(pin).ForceI 0, 0
                    TheHdw.PPMU.Pins(pin).Disconnect
                Next pin
            Else
                For Each pin In ForceByPin
                    With TheHdw.PPMU.Pins(pin)
                        '' 20150615 - Force 0 mA before expected force value to solve over clamp issue.
                        .ForceI 0, 0
                        .Connect
                        .Gate = tlOn
                        
                        If UBound(ForceValByPin) = 0 Then
                            .ForceI ForceValByPin(0), ForceValByPin(0)
                            ForceValI = ForceValByPin(0)
                        ElseIf ForceValByPin(ForceValIdx) <> "" Then
                            .ForceI ForceValByPin(ForceValIdx), ForceValByPin(ForceValIdx)
                             ForceValI = ForceValByPin(ForceValIdx)
                        Else:
                            .ForceI 0
                            ForceValI = 0
                        End If
                    End With
                    ForceValIdx = ForceValIdx + 1
                Next pin
            
                TheHdw.Wait (1 * ms)
                DebugPrintFunc_PPMU CStr(MeasurePinStr)
                MeasureValue = TheHdw.PPMU.Pins(MeasurePinStr).Read(tlPPMUReadMeasurements, 10)
                If TheExec.TesterMode = testModeOffline Then MeasureValue = 0.005 + Rnd() * 0.001
            End If
        End If
    
    Else
        '' 20150721 - Apply stored value
        Dim AfterformulaVal_PPMU As New PinListData
''        Call CUS_FormulaCalc(Stored_MeasI_PPMU, AfterformulaVal_PPMU)
        
        '' 20150721 - Store ForceValue for each site for test limit use.
        Dim TestPinMaxNum As Integer
        TestPinMaxNum = UBound(ForceByPin)
        ReDim StoreForceI(TestPinMaxNum) As New SiteDouble
        
        If MeasVGrpPinCnt > 0 And InstSpecialSetting <> InstrumentSpecialSetup.PPMU_SerialMeasurement Then '' Split pin group
            For Each pin In ForceByPin
                TheExec.DataManager.DecomposePinList pin, PinArr, PinCount
        
                If PinCount Mod MeasVGrpPinCnt <> 0 Then
                    MeasVGrpCnt = PinCount \ MeasVGrpPinCnt + 1
                Else
                    MeasVGrpCnt = PinCount / MeasVGrpPinCnt
                End If
                
                For i = 0 To MeasVGrpCnt - 1
                    TestPinStr = vbNullString
                    
                    For j = 0 To MeasVGrpPinCnt - 1
                        If (i * MeasVGrpPinCnt + j) < PinCount Then
                            TestPinStr = TestPinStr & "," & PinArr(i * MeasVGrpPinCnt + j)
                            MeasureValue.AddPin (PinArr(i * MeasVGrpPinCnt + j))
                        End If
                    Next j
                    TestPinStr = right(TestPinStr, Len(TestPinStr) - 1)
                    
                    For Each site In TheExec.sites.Active
                        With TheHdw.PPMU.Pins(TestPinStr)
                            If UBound(ForceValByPin) = 0 Then
                                .ForceI AfterformulaVal_PPMU.Pins(ForceValByPin(0)).value(site), AfterformulaVal_PPMU.Pins(ForceValByPin(0)).value(site)
                                StoreForceI(ForceValIdx) = AfterformulaVal_PPMU.Pins(ForceValByPin(ForceValIdx)).value(site)
                            ElseIf ForceValByPin(ForceValIdx) <> "" Then
                                .ForceI AfterformulaVal_PPMU.Pins(ForceValByPin(ForceValIdx)).value(site), AfterformulaVal_PPMU.Pins(ForceValByPin(ForceValIdx)).value(site)
                                StoreForceI(ForceValIdx) = AfterformulaVal_PPMU.Pins(ForceValByPin(ForceValIdx)).value(site)
                            Else
                                .ForceI 0
                            End If
                        End With
                    Next site
                    
                    TheHdw.Wait (1 * ms)
                    DebugPrintFunc_PPMU CStr(TestPinStr)
                    MeasureValueGrp = TheHdw.PPMU.Pins(TestPinStr).Read(tlPPMUReadMeasurements, 10)
                    
                    For j = 0 To MeasVGrpPinCnt - 1
                        If (i * MeasVGrpPinCnt + j) < PinCount Then
                            PinName = PinArr(i * MeasVGrpPinCnt + j)
                            MeasureValue.Pins(PinName) = MeasureValueGrp.Pins(PinName)
                            If TheExec.TesterMode = testModeOffline Then MeasureValue.Pins(PinName) = 0.005 + Rnd() * 0.001
                        End If
                    Next j
                    
                    TheHdw.PPMU.Pins(TestPinStr).Disconnect
                Next i
                ForceValIdx = ForceValIdx + 1
            Next pin
        Else '' No split pin group
            If InstSpecialSetting = InstrumentSpecialSetup.PPMU_SerialMeasurement Then
                TheExec.DataManager.DecomposePinList MeasurePinStr, PinArr, PinCount
                TheHdw.PPMU.Pins(MeasurePinStr).ForceI 0, 0
        
                For Each pin In PinArr
                    MeasureValue.AddPin (pin)
                    TheHdw.PPMU.Pins(pin).ForceI ForceValByPin(0), Abs(ForceValByPin(0))
                    ForceValI = ForceValByPin(0)
                    TheHdw.Wait 0.001
                    DebugPrintFunc_PPMU CStr(pin)
                    MeasureValue.Pins(pin) = TheHdw.PPMU.Pins(pin).Read(tlPPMUReadMeasurements, 10)
                    TheHdw.PPMU.Pins(pin).ForceI 0, 0
                    TheHdw.PPMU.Pins(pin).Disconnect
                Next pin
            Else
                For Each pin In ForceByPin
                    For Each site In TheExec.sites.Active
                        With TheHdw.PPMU.Pins(pin)
                            If UBound(ForceValByPin) = 0 Then
                                .ForceI AfterformulaVal_PPMU.Pins(ForceValByPin(0)).value(site), AfterformulaVal_PPMU.Pins(ForceValByPin(0)).value(site)
                                
                                StoreForceI(ForceValIdx) = AfterformulaVal_PPMU.Pins(ForceValByPin(ForceValIdx)).value(site)
                                
                            ElseIf ForceValByPin(ForceValIdx) <> "" Then
                                .ForceI AfterformulaVal_PPMU.Pins(ForceValByPin(ForceValIdx)).value(site), AfterformulaVal_PPMU.Pins(ForceValByPin(ForceValIdx)).value(site)
                                
                                StoreForceI(ForceValIdx) = AfterformulaVal_PPMU.Pins(ForceValByPin(ForceValIdx)).value(site)
                            Else
                                .ForceI 0
                            End If
                        End With
                    Next site
                    ForceValIdx = ForceValIdx + 1
                Next pin
            
                TheHdw.Wait (1 * ms)
                DebugPrintFunc_PPMU CStr(MeasurePinStr)
                MeasureValue = TheHdw.PPMU.Pins(MeasurePinStr).Read(tlPPMUReadMeasurements, 10)
                If TheExec.TesterMode = testModeOffline Then MeasureValue = 0.005 + Rnd() * 0.001
            End If
        End If
    End If

    If UBound(ForceSequenceArray) <> 0 Then
        If ForceSequenceArray(TestSeqNum) = "" Then
            ForceSequenceArray(TestSeqNum) = 0
        End If
    End If

    For Each site In TheExec.sites.Active
        TestNum = TheExec.sites.item(site).TestNumber
    Next site

'    TheHdw.Wait (1 * ms)
'
'    If InstSpecialSetting = DigitalConnectPPMU2 Then
'        TheHdw.PPMU.AllowPPMUFuncRelayConnection (True)
'        TheHdw.PPMU.Pins(MeasurePinStr).ForceI 0, 0.0002 '20160224 add to allow every seq with the same pins
'        TheHdw.Digital.Pins(MeasurePinStr).Connect
'    End If
'
'    If InstSpecialSetting = PPMU_SerialMeasurement Then
'        Dim PinArr() As String
'        Dim PinCount As Long
'
'        TheExec.DataManager.DecomposePinList MeasurePinStr, PinArr, PinCount
'        TheHdw.PPMU.Pins(MeasurePinStr).ForceI 0, 0
'
'        For Each Pin In PinArr
'            MeasureValue.AddPin (Pin)
'            TheHdw.PPMU.Pins(Pin).ForceI ForceValByPin(0), Abs(ForceValByPin(0))
'            TheHdw.Wait 0.001
'            DebugPrintFunc_PPMU CStr(Pin)
'            MeasureValue.Pins(Pin) = TheHdw.PPMU.Pins(Pin).Read(tlPPMUReadMeasurements, 10)
'            TheHdw.PPMU.Pins(Pin).ForceI 0, 0
'            TheHdw.PPMU.Pins(Pin).Disconnect
'        Next Pin
'    Else
'        DebugPrintFunc_PPMU CStr(MeasurePinStr)
'        MeasureValue = TheHdw.PPMU.Pins(MeasurePinStr).Read(tlPPMUReadMeasurements, 10)
'    End If
    
    '' Calculate RAK
    Dim RakV() As Double
    If RAK_Flag = True Then
        For Each site In TheExec.sites
            For Each pin In MeasureValue.Pins
    
                'RakV = TheHdw.PPMU.ReadRakValuesByPinnames(Pin, Site)
                
                If InStr(TheExec.CurrentChanMap, "CP") <> 0 Then
                    MeasureValue.Pins(pin).value(site) = MeasureValue.Pins(pin).value(site) - ForceValI * (CP_Card_RAK.Pins(pin).value(site))
                Else
                    MeasureValue.Pins(pin).value(site) = MeasureValue.Pins(pin).value(site) - ForceValI * (FT_Card_RAK.Pins(pin).value(site))
                End If

            Next pin
        Next site
    End If
      
    '' 20150728 - Add return measure volt to main function.

    ReturnMeasVolt = MeasureValue
    
    Force_idx = TestSeqNum
    If UBound(ForceSequenceArray) = 0 Then
        Force_idx = 0
    End If

    Dim TestNameInput As String
    TestNameInput = "Volt_meas_" + CStr(TestSeqNum)
    
    '''20151103 print force condition
    Call Print_Force_Condition("v", MeasureValue)
    
    '' 20150721 - Test limit for force stored value
    Dim ForceIndex As Integer
    ForceIndex = 0
    If b_UseStoredForceVal = True Then
        For Each pin In ForceByPin
            'For Each Site In TheExec.Sites
            If CurrentJobName_L Like "*char*" Then
                Disable_Inst_pinname_in_PTR
                TheExec.Flow.TestLimit MeasureValue.Pins(pin), , , scaletype:=scaleNone, unit:=unitVolt, formatStr:="%.3f", ForceVal:=StoreForceI(ForceIndex).value(site), ForceUnit:=unitAmp, ForceResults:=tlForceFlow
                Enable_Inst_pinname_in_PTR
            Else
                TheExec.Flow.TestLimit MeasureValue.Pins(pin), , , scaletype:=scaleNone, unit:=unitVolt, formatStr:="%.3f", Tname:=TestNameInput, ForceVal:=StoreForceI(ForceIndex).value(site), ForceUnit:=unitAmp, ForceResults:=tlForceFlow
            
            End If
            'Next Site
            ForceIndex = ForceIndex + 1
        Next pin
      '  Exit Function
    'End If
    ElseIf InStr(CUS_Str_MainProgram, "DDR_VOHL") <> 0 Then
         
            Dim HiLimitVal As Integer
            Dim LoLimitVal As Integer
            HiLimitVal = 0: LoLimitVal = 0
            If CUS_Str_MainProgram = "DDR_VOHL_1" Then
                HiLimitVal = 132: LoLimitVal = 108
            ElseIf CUS_Str_MainProgram = "DDR_VOHL_2" Then
                If TestSeqNumIdx = 0 Then HiLimitVal = 42: LoLimitVal = 38
                If TestSeqNumIdx = 1 Then HiLimitVal = 176: LoLimitVal = 144
            ElseIf CUS_Str_MainProgram = "DDR_VOHL_3" Then
                HiLimitVal = 264: LoLimitVal = 216
            Else
              
              If TestSeqNumIdx = 0 Then HiLimitVal = 132: LoLimitVal = 108
            End If
         
            If TestSeqNumIdx = 0 Then
                TheExec.Flow.TestLimit MeasureValue, , , scaletype:=scaleNone, unit:=unitVolt, formatStr:="%.3f", Tname:=TestNameInput, ForceUnit:=unitAmp, ForceResults:=tlForceNone, ForceVal:=ForceValI 'AfterformulaVal_PPMU.Pins(p).Value(Site)'
                TheExec.Flow.TestLimit MeasureValue.Math.divide(ForceValI), LoLimitVal, HiLimitVal, scaletype:=scaleNone, unit:=unitCustom, formatStr:="%.3f", Tname:=TestNameInput, ForceUnit:=unitAmp, ForceResults:=tlForceNone, ForceVal:=ForceValI, customUnit:="ohm" 'AfterformulaVal_PPMU.Pins(p).Value(Site)'
            Else
                TheExec.Flow.TestLimit MeasureValue, , , scaletype:=scaleNone, unit:=unitVolt, formatStr:="%.3f", Tname:=TestNameInput, ForceUnit:=unitAmp, ForceResults:=tlForceNone, ForceVal:=ForceValI 'AfterformulaVal_PPMU.Pins(p).Value(Site)'
                TheExec.Flow.TestLimit MeasureValue.Math.Subtract(1.1).divide(ForceValI).Abs, LoLimitVal, HiLimitVal, scaletype:=scaleNone, unit:=unitCustom, formatStr:="%.3f", Tname:=TestNameInput, ForceUnit:=unitAmp, ForceResults:=tlForceNone, ForceVal:=ForceValI, customUnit:="ohm" 'AfterformulaVal_PPMU.Pins(p).Value(Site)'
            End If
        
        ElseIf Flag_SingleLimit = True Then
            If LCase(currentJobName) Like "*char*" Then
                TheExec.Flow.TestLimit MeasureValue, LowLimitVal, HighLimitVal, scaletype:=scaleNone, unit:=unitVolt, formatStr:="%.3f", ForceUnit:=unitAmp, ForceResults:=tlForceFlow, ForceVal:=ForceValI  'AfterformulaVal_PPMU.Pins(p).Value(Site)'
            Else
                TheExec.Flow.TestLimit MeasureValue, LowLimitVal, HighLimitVal, scaletype:=scaleNone, unit:=unitVolt, formatStr:="%.3f", Tname:=TestNameInput, ForceUnit:=unitAmp, ForceResults:=tlForceNone, ForceVal:=ForceValI 'AfterformulaVal_PPMU.Pins(p).Value(Site)'
            End If
        '' 20150202 - Range Check
        If Range_Check_Enable_Word = True Then
            Call CheckRangesAndClamps(MeasureValue, "V", RangeCheck_HighLimitVal(gl_UseLimitCheck_Counter), RangeCheck_LowLimitVal(gl_UseLimitCheck_Counter))
            gl_UseLimitCheck_Counter = gl_UseLimitCheck_Counter + 1
        End If
    
    Else
        If mid(TestLimitPerPin_VIR, 1, 1) = "F" And UBound(ForceValByPin) = 0 Then
            If LCase(currentJobName) Like "*char*" Then
                TheExec.Flow.TestLimit MeasureValue, , , scaletype:=scaleNone, unit:=unitVolt, formatStr:="%.3f", ForceVal:=ForceSequenceArray(Force_idx), ForceUnit:=unitAmp, ForceResults:=tlForceFlow
            Else
                        '' TheExec.Flow.TestLimit MeasureValue, , , ScaleType:=scaleNone, unit:=unitVolt, formatstr:="%.3f", Tname:=TestNameInput + CStr(TestSeqNum) + "_" + CStr(k - 1) + "@COND:PATTERN=" + PATT_ExculdePath(Pat), forceVal:=ForceSequenceArray(Force_Idx), forceunit:=unitAmp, ForceResults:=tlForceFlow
                        TheExec.Flow.TestLimit MeasureValue, , , scaletype:=scaleNone, unit:=unitVolt, formatStr:="%.3f", Tname:=TestNameInput, ForceVal:=ForceSequenceArray(Force_idx), ForceUnit:=unitAmp, ForceResults:=tlForceFlow
            End If
            '' 20150202 - Range Check
            If Range_Check_Enable_Word = True Then
                Call CheckRangesAndClamps(MeasureValue, "V", RangeCheck_HighLimitVal(gl_UseLimitCheck_Counter), RangeCheck_LowLimitVal(gl_UseLimitCheck_Counter))
                gl_UseLimitCheck_Counter = gl_UseLimitCheck_Counter + 1
            End If
        Else
            IdxV = 0
            For p = 0 To MeasureValue.Pins.Count - 1
                If CurrentJobName_L Like "*char*" Then
                    Disable_Inst_pinname_in_PTR
                    TheExec.Flow.TestLimit MeasureValue.Pins(p), , , scaletype:=scaleNone, unit:=unitVolt, formatStr:="%.3f", ForceVal:=ForceValByPin(IdxV), ForceUnit:=unitAmp, ForceResults:=tlForceFlow
                    Enable_Inst_pinname_in_PTR
                Else
                    TheExec.Flow.TestLimit MeasureValue.Pins(p), , , scaletype:=scaleNone, unit:=unitVolt, formatStr:="%.3f", Tname:=TestNameInput, ForceVal:=ForceValByPin(IdxV), ForceUnit:=unitAmp, ForceResults:=tlForceFlow
                End If
                
                If UBound(ForceValByPin) = 0 Then
                    IdxV = 0
                Else
                    IdxV = IdxV + 1
                End If
                
                '' 20150202 - Range Check
                If Range_Check_Enable_Word = True Then
                    TempMeasVal_PerPin(p).AddPin (MeasureValue.Pins(p))
                    TempMeasVal_PerPin(p).Pins(MeasureValue.Pins(p)) = MeasureValue.Pins(p)
                    Call CheckRangesAndClamps(TempMeasVal_PerPin(p), "V", RangeCheck_HighLimitVal(gl_UseLimitCheck_Counter), RangeCheck_LowLimitVal(gl_UseLimitCheck_Counter))
                    gl_UseLimitCheck_Counter = gl_UseLimitCheck_Counter + 1
                End If
            Next p
        End If
    End If
        
    ' 20160105: Steph added for Refbuf test (Autogen) --- start
'    Call CUS_VFI_MeasureVolt(CUS_Str_MainProgram, MeasureValue, TestSeqNum, Pat)
    ' 20160105: Steph added for Refbuf test (Autogen) --- end
    
Exit Function 'Add ErrHandler 2023/08/18
errHandler: 'Add ErrHandler 2023/08/18
    Call Print_Error_Message(Error_Info, "VBT_LIB_DC_Func", "IO_HardIP_PPMU_Measure_V_TTR") 'Add ErrHandler 2023/08/18
    If AbortTest Then Exit Function Else Resume Next 'Add ErrHandler 2023/08/18
End Function


' [20230809][All][Neil] add init glb_TestInstance
' [20231108][All][Carter] Restore delete function
Public Function Meas_VIR_IO_Universal_func_GPIO_TTR(Optional patset As Pattern, Optional TestSequence As String, Optional CPUA_Flag_In_Pat As Boolean, _
                                    Optional DisableComparePins As PinList, Optional DisableConnectPins As PinList, Optional DisableFRC As Boolean = False, Optional FRCPortName As String, _
                                    Optional Measure_Pin_PPMU As String, Optional ForceV As String, Optional ForceI As String, Optional MeasureI_Range As String = "0.05", _
                                    Optional DigCap_Pin As PinList, Optional DigCap_DataWidth As Long, Optional DigCap_Sample_Size As Long, Optional DigCap_DSPWaveSetting As CalculateMethodSetup = 0, _
                                    Optional DigSrc_pin As PinList, Optional DigSrc_DataWidth As Long, Optional DigSrc_Sample_Size As Long, Optional DigSrc_Equation As String, Optional DigSrc_Assignment As String, Optional DigSrc_FlowForLoopIntegerName As String = vbNullString, _
                                    Optional InstSpecialSetting As InstrumentSpecialSetup = 0, Optional PPMU_TestLimit_TTR As Boolean = False, Optional RAK_Flag As Enum_RAK = 0, _
                                    Optional CUS_Str_MainProgram As String = vbNullString, Optional CUS_Str_DigCapData As String = vbNullString, Optional CUS_Str_DigSrcData As String, _
                                    Optional Flag_SingleLimit As Boolean = False, Optional TestLimitPerPin_VIR As String = "FFF", _
                                    Optional InterFunc_PrePat As InterposeName, Optional InterFuncArgs_PrePat As String, _
                                    Optional InterFunc_PostPat As InterposeName, Optional InterFuncArgs_PostPat As String, _
                                    Optional CharInputString As String, Optional ForceFunctional_Flag As Boolean = False, Optional MeasIGrpPinCnt_Str As String = vbNullString, Optional Meas_StoreName As String, Optional Calc_Eqn As String, Optional KeepEmptyLimit As Boolean = False, _
                                    Optional CFG_GPIO_Pins As String, Optional MeasVGrpPinCnt As Integer = 0, Optional VDD_IO_1p2 As String, Optional VDD_IO_1p8 As String, Optional Validating_ As Boolean) As Long
On Error GoTo errHandler 'Add ErrHandler 2023/08/18
                                    
''20191028, Add VDD_IO_1p2 and VDD_IO_1p8 for calculating differential voltage, Carter
''20150924 - Remove argument as below
''Remove ,Optional MeasOnHalt As Boolean
''Remove ,Optional InterFunc_SequenceN As InterposeName
''Remove ,Optional InterFuncArgs_SequenceN As String

''==================================================================================
'' 20150621 - Check with CCWu: FRCPortName As String, Optional DisableFRC As Boolean = False not use in this function
'' 20150717 - Impedance measurement by using 2 point measure method, Define "Z" for TestSequence - On going
''                - EX: Pin1, Pin2 + Pin3, Pin4     V1, V2 + V3, V4
''                - V1 and V2 use for Pin1 of impedence measurement
''                - V1 and V2 use for Pin2 of impedence measurement
'' 20150717 - Get I from previous item and apply the current value to next item, use enum for the feature
''                - EX: TestSequence: "V,V,V"
''                  If second V want to apply calcuated I value that Force I value argument should be "0,keyword,0"
'' 20150727 - MeasureI_Range is use for test sequence "I", "R" and "Z"
''==================================================================================
    
    Dim PatCount As Long
    Dim k As Long
    Dim TestOptLen As Integer
    Dim TestSequenceArray() As String, ForceISequenceArray() As String, ForceVSequenceArray() As String
    Dim TestOption As Variant
    Dim Ts As Variant
    Dim TestSeqNum As Integer
    Dim PatMeas As String
    Dim TestPinArrayIV() As String
    Dim TestIrange() As String
    Dim TestSeqNumIdx As Long
    Dim InDSPWave As New DSPWave
    Dim OutDspWave() As New DSPWave
    Dim ShowDec As String
    Dim ShowOut As String
    
    ''20141223
    Dim site As Variant
    Dim PattArray() As String
    Dim Pat As String
    Dim patt As Variant
    
    Dim HighLimitVal() As Double
    Dim LowLimitVal() As Double
    
    ''20150728
    Dim ReturnMeasVolt As New PinListData
    Dim FlowForLoopName() As String    ' Sequences : Code , Voltage , Loop Index
    
    Dim i, j As Integer

    ''20160821-Add judgement by 7.75mA for Fuse
    Dim CFGTestPins(0) As String
    Dim CFGTest_FirstSequence As Boolean
    CFGTest_FirstSequence = True
    
    
    If Validating_ Then 'Carter, 20190315
        Call PrLoadPattern(patset.value)
        Exit Function    ' Exit after validation
    End If

    glb_TestInstance = vbNullString
    glb_TestInstance = UCase(TheExec.DataManager.instancename)

    Shmoo_Pattern = patset.value

    Call tl_PinListDataSort(True)
    ''========================================================================================
    '' 20150121 - Range Check
    If Range_Check_Enable_Word = True Then         'Change to "Range_Check_Enable_Word" by Martin 20151225  for TTR
        If TheExec.DataManager.MemberIndex = 0 Then
            gl_UseLimitCheck_Counter = 0
        End If
    End If
    ''========================================================================================
    
    '' 20160111 - Check input condition for measure I, R and Z.
''    Call CheckCondition_Measure_I_R_Z(TestSequence, Measure_Pin_PPMU, ForceV, MeasureI_Range)
    
    TestSequenceArray = Split(TestSequence, ",")
    
    If ForceI = "" Then ForceI = 0
    If ForceV = "" Then ForceV = 0
    
    Call GetFlowTName
    
    '----------------------------for store name 180525----------------------------
    Dim StoreIndex As Long
    Dim MeasStoreName_Ary() As String
    Dim Interpose_PreMeas_Ary() As String
    Dim MeasPinAry_F_Differential() As String
    Dim MeasPinAry_F() As String
    Dim MeasPinAry_V() As String
    
    Call VFI_ProcessInputString(TestSequence, "", Measure_Pin_PPMU, "", "", MeasureI_Range, Meas_StoreName, "", _
                                            ForceV, ForceI, _
                                            TestSequenceArray(), MeasPinAry_V(), TestPinArrayIV(), MeasPinAry_F(), _
                                            MeasPinAry_F_Differential(), TestIrange(), MeasStoreName_Ary(), Interpose_PreMeas_Ary(), ForceVSequenceArray(), ForceISequenceArray())
    '----------------------------for store name 180525----------------------------
    
    Call ProcessInputToGLB(patset)
    
    ForceISequenceArray = Split(ForceI, "+")
    ForceVSequenceArray = Split(ForceV, "+")
        '''''''''''Apply DC spec'''''''''''
    For i = 0 To UBound(ForceVSequenceArray)
        ForceVSequenceArray(i) = EvaluateEachBlock(ForceVSequenceArray(i)) ''zhhuangf
    Next i
    '''''''''''Apply DC spec'''''''''''
    '' 20150812-Decompose DigCap_Pin by ","
    Dim DigCap_Pin_Ary() As String
    DigCap_Pin_Ary = Split(DigCap_Pin, ",")
    FlowForLoopName = Split(DigSrc_FlowForLoopIntegerName, ",")
    
    Char_Test_Name_Curr_Loc = 0 'init char datalog test name index
    
    TestPinArrayIV = Split((Measure_Pin_PPMU), "+")
    
    If MeasureI_Range = "" Then MeasureI_Range = "50e-3"
    
    TestIrange = Split(MeasureI_Range, "+")
    
    TheHdw.Digital.ApplyLevelsTiming True, True, True, tlPowered
  
    If (DisableConnectPins <> "") Then TheHdw.Digital.Pins(DisableConnectPins).Disconnect
    If (DisableComparePins <> "") Then TheHdw.Digital.Pins(DisableComparePins).DisableCompare = True
    
    If InterFunc_PrePat <> "" Then Call Interpose(InterFunc_PrePat, InterFuncArgs_PrePat)
    
    '20151028  CUS_MeasV_And_CalR -- TYCHENGG
    ''ex.  CUS_Str_MainProgram = "CalR;1.88;RVOH,RVOL,RVOH"
    ''========================================================================================
    Dim CUS_CalR_VDD As Double
    Dim CUS_CalR_Seq_Ary() As String
    Dim CUS_CalR_Arg_Ary() As String
    If (CUS_Str_MainProgram <> "") Then
        If (UCase(CUS_Str_MainProgram) Like "*CALR*") Then
            CUS_CalR_Arg_Ary = Split(CUS_Str_MainProgram, ";")
            CUS_CalR_VDD = CDbl(CUS_CalR_Arg_Ary(1))
            CUS_CalR_Seq_Ary = Split(CUS_CalR_Arg_Ary(2), ",")
        End If
    End If
    ''========================================================================================
    
    '' 20150625 - Apply Char setup
    If CharInputString <> "" Then
        Call SetForceCondition(CharInputString)
    End If
    
    TheHdw.Patterns(patset).Load
    Call PATT_GetPatListFromPatternSet(patset.value, PattArray, PatCount)
    Call TheHdw.Digital.Patgen.Continue(0, cpuA + cpuB + cpuC + cpuD)
    TheHdw.Digital.Patgen.HaltMode = tlHaltOnOpcode
    If KeepEmptyLimit = True Then
        Call GetFlowSingleUseLimit_KeepEmpty(HighLimitVal, LowLimitVal)  ''20141223
    Else
        Call GetFlowSingleUseLimit(HighLimitVal, LowLimitVal)  ''20141223
    End If
    For Each patt In PattArray
        Pat = CStr(patt)
        PatMeas = Pat
         
        If DigSrc_Sample_Size <> 0 Then
             
             '' 20150810 - Source dssc index by For opcode from flow table
            If DigSrc_FlowForLoopIntegerName <> "" Then        '20151201
                If FlowForLoopName(0) <> "" Then        '20151201
                    Call DSSCSrcBitFromFlowForLoop(FlowForLoopName(0), DigSrc_DataWidth, DigSrc_Equation, DigSrc_Assignment)
                End If
            End If
            
            For Each site In TheExec.sites.Active
                ''20150708- Add customize string for digsrc data process
                Call Create_DigSrc_Data(DigSrc_pin, DigSrc_DataWidth, DigSrc_Sample_Size, DigSrc_Equation, DigSrc_Assignment, InDSPWave, site, CUS_Str_DigSrcData)
            Next site
            
            Call SetupDigSrcDspWave(PatMeas, DigSrc_pin, "Meas_src", DigSrc_Sample_Size, InDSPWave)
        End If

        ''20150812- Add program to setup multiple DigCap_Pin.
        If DigCap_Sample_Size <> 0 Then
            Dim DigCap_Pin_Num As Integer
            DigCap_Pin_Num = UBound(DigCap_Pin_Ary)
            ReDim OutDspWave(DigCap_Pin_Num) As New DSPWave
            For i = 0 To DigCap_Pin_Num
                TheExec.Datalog.WriteComment ("======== Setup Dig Cap Test for " & DigCap_Pin_Ary(i) & " ========")
                OutDspWave(i).CreateConstant 0, DigCap_Sample_Size
                DigCap_Pin.value = DigCap_Pin_Ary(i)
                Call DigCapSetup(Pat, DigCap_Pin, "Meas_cap", DigCap_Sample_Size, OutDspWave(i))
            Next i
        End If
          
        Call TheHdw.Patterns(Pat).start
        TestSeqNum = 0
        
        Dim FlowTestName() As String
        Dim Temp_index As Long
        Temp_index = 0
        
        For Each Ts In TestSequenceArray
            ''20150907 - Only need CPUA_Flag_In_Pat to do control
            If (CPUA_Flag_In_Pat) Then
                Call TheHdw.Digital.Patgen.FlagWait(cpuA, 0) 'Meas during CPUA loop
            Else
                Call TheHdw.Digital.Patgen.HaltWait 'Pattern without CPUA loop
            End If
            
            TestOptLen = Len(Ts)
            
            TestSeqNumIdx = TestSeqNum
            For k = 1 To TestOptLen
                TestOption = mid(Ts, k, 1)
                
                '' 20160106 - If "ForceFunctional_Flag" = True to let TestOption = "N" to make the test instance only run functional test
                If ForceFunctional_Flag = True Then
                    TestOption = "N"
                End If
                
                If (Measure_Pin_PPMU <> "") Then
                    Call Meas_VIR_IO_PreSetupBeforeMeasurement(TestPinArrayIV, TestSeqNumIdx)
                    
                    Select Case UCase(TestOption)
                    
                        Case "V"
                        
'                            Call IO_HardIP_PPMU_Measure_V_TTR(TestPinArrayIV, TestSeqNum, TestSeqNumIdx, ForceISequenceArray, _
'                                    k, pat, Flag_SingleLimit, HighLimitVal(TestSeqNum), LowLimitVal(TestSeqNum), TestLimitPerPin_VIR, ReturnMeasVolt, _
'                                    SpecialCalcValSetting, InstSpecialSetting, RAK_Flag, CUS_Str_MainProgram, MeasVGrpPinCnt)
 
                             ''20151028  CUS_MeasV_And_CalR -- TYCHENGG
                            ''========================================================================================
                            If (UCase(CUS_Str_MainProgram) Like "*CALR*") Then
                                Call CUS_VIR_MainProgram_MeasV_CalR(TestPinArrayIV, TestSeqNum, CUS_CalR_Seq_Ary, ForceISequenceArray, ReturnMeasVolt, CUS_CalR_VDD)
                            End If
                            ''========================================================================================
                            
                        Case "I"
                            Dim Rtn_MeasCurr As New PinListData
                            If DisableFRC = True Then FreeRunClk_Disable (FRCPortName)
                            
                            Call IO_HardIP_PPMU_Measure_I_TTR(TestPinArrayIV, TestSeqNum, TestSeqNumIdx, ForceVSequenceArray, _
                                    k, Pat, Flag_SingleLimit, HighLimitVal(TestSeqNum), LowLimitVal(TestSeqNum), TestLimitPerPin_VIR, TestIrange, VDD_IO_1p2, VDD_IO_1p8, _
                                    CUS_Str_MainProgram, PPMU_TestLimit_TTR, MeasIGrpPinCnt_Str, CFG_GPIO_Pins, CFGTest_FirstSequence, Rtn_MeasCurr, Temp_index)

                            If CFGTest_FirstSequence = True Then CFGTest_FirstSequence = False
                            
                            '==== Store measure current result ==== -- 20230109
                            If Meas_StoreName <> "" Then
                                If MeasStoreName_Ary(TestSeqNum) <> "" Then
                                    Call AddStoredMeasurement(MeasStoreName_Ary(TestSeqNum), Rtn_MeasCurr)
                                    StoreIndex = StoreIndex + 1
                                End If
                            End If
                            
                        Case Else
                            TheExec.Datalog.WriteComment "Error Test Option, please select V, I"
                    
                    End Select
                    
                    Call Meas_VIR_IO_PostSetupAfterMeasurement(TestPinArrayIV, TestSeqNumIdx)
                End If
            Next k
            
            TestSeqNum = TestSeqNum + 1
            
            If (CPUA_Flag_In_Pat) Then
                Call TheHdw.Digital.Patgen.Continue(0, cpuA) 'Jump out CPUA loop
            End If
        Next Ts
        '==== Add support post calculate function ==== -- 20230109
        If Calc_Eqn <> "" Then
            Call ProcessCalcEquation(Calc_Eqn)
        End If
        
        If debugPrintEnable = True Then
            TheExec.Datalog.WriteComment "  Pattern(" & PatCount & "): " & Pat & ""
        End If
                
        TheHdw.Digital.Patgen.HaltWait ' haltwait at patten end
        PatCount = PatCount + 1
        
        'This function is no longer exist
'        If DigCap_Sample_Size <> 0 Then
'            Call HardIP_Digcap_Print(CUS_Str_DigCapData, OutDspWave, DigCap_Sample_Size, DigCap_DataWidth, ShowDec, ShowOut, , DigCap_DSPWaveSetting) '''change
'        End If
    Next patt
    
    If InterFunc_PostPat <> "" Then Call Interpose(InterFunc_PostPat, InterFuncArgs_PostPat)
    
    If (DisableConnectPins <> "") Then TheHdw.Digital.Pins(DisableConnectPins).Connect
    If DisableComparePins <> "" Then TheHdw.Digital.Pins(DisableComparePins).DisableCompare = False
    
    If DisableFRC = True Then
        Call FreeRunclk_Enable(FRCPortName)
    End If

    Call HardIP_WriteFuncResult(, , glb_TestInstance)  '20230111
    
    DebugPrintFunc patset.value  ' print all debug information
    
    '' 20150728 - Print Char setup for power pins.
    If CurrentJobName_U Like "*CHAR*" Then
        If CharInputString <> "" Then
          '  Call PrintCharPowerSet(CharInputString)
        End If
    End If

Exit Function 'Add ErrHandler 2023/08/18
errHandler: 'Add ErrHandler 2023/08/18
    Call Print_Error_Message(Error_Info, "VBT_LIB_DC_Func", "Meas_VIR_IO_Universal_func_GPIO_TTR") 'Add ErrHandler 2023/08/18
    If AbortTest Then Exit Function Else Resume Next 'Add ErrHandler 2023/08/18
End Function


' [20231108][All][Carter] Restore delete function
Public Function EvaluateEachBlock(ForceVSeq As String) As String ''zhhuangf
On Error GoTo errHandler 'Add ErrHandler 2023/08/18
    Dim i As Integer, j As Integer
        ''_VDDIO18_GRP_VAR_H
    Dim ForceVarray() As String
    If InStr(ForceVSeq, ",") Then
        Dim tempV1() As String
        Dim Variable_Spec As String
        ForceVarray = Split(ForceVSeq, ",")
        For i = 0 To UBound(ForceVarray)
            If InStr(ForceVarray(i), "*") Then
                tempV1 = Split(ForceVarray(i), "*")
                For j = 0 To UBound(tempV1)
                    If InStr(tempV1(j), "_") Then
                        Variable_Spec = right(tempV1(j), Len(tempV1(j)) - 1)
                        tempV1(j) = CStr(TheExec.Specs.DC.item(Variable_Spec).ContextValue)
                    End If
                    If j = 0 Then
                        ForceVarray(i) = tempV1(0)
                    Else
                        ForceVarray(i) = ForceVarray(i) & "*" & tempV1(j)
                    End If
                Next j
            ElseIf InStr(ForceVarray(i), "/") Then
                tempV1 = Split(ForceVarray(i), "/")
                For j = 0 To UBound(tempV1)
                    If InStr(tempV1(j), "_") Then
                        Variable_Spec = right(tempV1(j), Len(tempV1(j)) - 1)
                        tempV1(j) = CStr(TheExec.Specs.DC.item(Variable_Spec).ContextValue)
                    End If
                    If j = 0 Then
                        ForceVarray(i) = tempV1(0)
                    Else
                        ForceVarray(i) = ForceVarray(i) & "/" & tempV1(j)
                    End If
                Next j
            ElseIf InStr(ForceVSeq, "_") Then
                Variable_Spec = right(ForceVSeq, ForceVSeq - 1)
                ForceVSeq = CStr(TheExec.Specs.DC.item(Variable_Spec).ContextValue)
            End If
            If i = 0 Then
                EvaluateEachBlock = Evaluate(ForceVarray(0))
            Else
                EvaluateEachBlock = EvaluateEachBlock & "," & Evaluate(ForceVarray(i))
            End If
        Next i
    Else
        If InStr(ForceVSeq, "*") Then
            tempV1 = Split(ForceVSeq, "*")
            For j = 0 To UBound(tempV1)
                If InStr(tempV1(j), "_") Then
                    Variable_Spec = right(tempV1(j), Len(tempV1(j)) - 1)
                    tempV1(j) = CStr(TheExec.Specs.DC.item(Variable_Spec).ContextValue)
                End If
                If j = 0 Then
                    ForceVSeq = tempV1(0)
                Else
                    ForceVSeq = ForceVSeq & "*" & tempV1(j)
                End If
            Next j
        ElseIf InStr(ForceVSeq, "/") Then
            tempV1 = Split(ForceVSeq, "/")
            For j = 0 To UBound(tempV1)
                If InStr(tempV1(j), "_") Then
                    Variable_Spec = right(tempV1(j), Len(tempV1(j)) - 1)
                    tempV1(j) = CStr(TheExec.Specs.DC.item(Variable_Spec).ContextValue)
                End If
                If j = 0 Then
                    ForceVSeq = ForceVSeq & tempV1(0)
                Else
                    ForceVSeq = ForceVSeq & "/" & tempV1(j)
                End If
            Next j
        ElseIf InStr(ForceVSeq, "_") Then
            Variable_Spec = right(ForceVSeq, ForceVSeq - 1)
            ForceVSeq = CStr(TheExec.Specs.DC.item(Variable_Spec).ContextValue)
        End If
        EvaluateEachBlock = Evaluate(ForceVSeq)
    End If
    
Exit Function 'Add ErrHandler 2023/08/18
errHandler: 'Add ErrHandler 2023/08/18
    Call Print_Error_Message(Error_Info, "VBT_LIB_DC_Func", "EvaluateEachBlock") 'Add ErrHandler 2023/08/18
    If AbortTest Then Exit Function Else Resume Next 'Add ErrHandler 2023/08/18
End Function


' [20231108][All][Carter] Restore delete function
Public Function GetFlowSingleUseLimit_KeepEmpty(ByRef d_HighLimitVal() As Double, ByRef d_LowLimitVal() As Double) As Double
On Error GoTo errHandler 'Add ErrHandler 2023/08/18
    ' Get the limits info
    Dim FlowLimitsInfo As IFlowLimitsInfo
    Dim HighLimitValArray() As String
    Dim LowLimitValArray() As String
    Dim HighLimitArraySize As Long
    Dim LowLimitArraySize As Long
    
    Dim i As Integer
    
    Call TheExec.Flow.GetTestLimits(FlowLimitsInfo)
    
    
    ' TheExec.Flow.GetTestLimits
    If FlowLimitsInfo Is Nothing Then
    ' i = i
    '  End If
        ReDim d_HighLimitVal(0) As Double
        ReDim d_LowLimitVal(0) As Double
        ReDim HighLimitValArray(0) As String
        ReDim LowLimitValArray(0) As String
        d_HighLimitVal(0) = 0
        d_LowLimitVal(0) = 0
        HighLimitValArray(0) = 0
        LowLimitValArray(0) = 0
    Else
        Call FlowLimitsInfo.GetHighLimits(HighLimitValArray())
        Call FlowLimitsInfo.GetLowLimits(LowLimitValArray())
        
        HighLimitArraySize = UBound(HighLimitValArray)
        ReDim d_HighLimitVal(HighLimitArraySize) As Double
        LowLimitArraySize = UBound(LowLimitValArray)
        ReDim d_LowLimitVal(LowLimitArraySize) As Double
        
    End If
    For i = 0 To HighLimitArraySize
        If (HighLimitValArray(i)) = "" Then HighLimitValArray(i) = -123456.123456 '''ZHHUANGF 20160728
        d_HighLimitVal(i) = CDbl(HighLimitValArray(i))
    Next i
        
    For i = 0 To LowLimitArraySize
        If LowLimitValArray(i) = "" Then LowLimitValArray(i) = -123456.123456
        
        d_LowLimitVal(i) = CDbl(LowLimitValArray(i))
    Next i
        
Exit Function 'Add ErrHandler 2023/08/18
errHandler: 'Add ErrHandler 2023/08/18
    Call Print_Error_Message(Error_Info, "VBT_LIB_DC_Func", "GetFlowSingleUseLimit_KeepEmpty") 'Add ErrHandler 2023/08/18
    If AbortTest Then Exit Function Else Resume Next 'Add ErrHandler 2023/08/18
End Function


''20230223: Modidfied to add EVS can ramp pin group voltage by parallel .
' [20230407][RF] Add DCVI EVS
' [20231228][T-All][Tank] Add EVS parallel IVCurve
Public Function EVS_Static_Power_Ramp(dc_spec As String, S_WaitTime As Double, Power_pin As String, Optional Step_number As Integer = 10, Optional Rising_Delay_time As Double = 0.02, Optional Looping_Contorl As Boolean = False, Optional Looping_Range As Double = 0.4, Optional Looping_Index_Name As String = vbNullString, Optional Looping_Max_Steps_Name As String = vbNullString, Optional Open_LatchUp_measure As Boolean = False, Optional Multi_Function As Boolean = False, Optional Mulit_EVS_Index As Integer = 1, Optional Test_time_breakdown As Boolean = False, _
                                    Optional Cooling_Time As Double = 0, Optional TotalPwrLimit As Double = 20, _
                                    Optional Flag_Serial As Boolean = True, Optional Parallel_Pin_Voltage As String = vbNullString) As Long

On Error GoTo errHandler
    If Looping_Contorl And (Looping_Max_Steps_Name = "" Or Looping_Index_Name = "") Then
        TheExec.Datalog.WriteComment "Error!! Please make sure you fill in argument both Looping_Max_Steps_Name and Looping_Index_Name "
        GoTo errHandler
    End If
    Dim funcName As String:: funcName = "EVS_Static_Power_Ramp"
    Dim inst_name As String
    'Dim Pin As Variant
    Dim i As Long
    Dim Spec As Variant
    Dim Spec_Var As String
    Dim pin As Variant
    Dim power_pin_ary() As String
    Dim Dc_spec_type As String
    Dim EVS_Detail_value() As Pins_detail ''Define each pin information
    Dim Mulit_Loop_EVS As Integer
    Dim TExec_Before_Pat As Double
    Dim Total_time As Double
    Dim EVS_Sequence As String
    Dim EVS_Type As String
    Dim d_PowerPin_ApplyVol() As Double
    Dim CntPins As Long
    
    glb_TestInstance = vbNullString
    glb_TestInstance = UCase(TheExec.DataManager.instancename)
    
    If gl_EnableCurrentProfile Or gl_EnableVoltageProfile Then
        ProfileCollect_Start glb_TestInstance
    End If
    
    EVS_Sequence = "Total_time"
    Call Test_time_breakdown_Start(Total_time, Test_time_breakdown, EVS_Sequence)
    
    Call TheExec.DataManager.DecomposePinList(Power_pin, power_pin_ary(), CntPins)
    
    Dim Filter_Power_Ary() As String
    Dim sPinName As String
    Dim nTempCnt As Long
    Dim site As Variant 'Carter, 20240304
    nTempCnt = 0
    For Each pin In power_pin_ary
        sPinName = LCase(pin)
        If gl_GetInstrumentType_Dic.Exists(sPinName) Then
            If LCase(gl_GetInstrumentType_Dic(sPinName)) Like "*dcvs*" Then
                ReDim Preserve Filter_Power_Ary(nTempCnt)
                Filter_Power_Ary(nTempCnt) = sPinName
                nTempCnt = nTempCnt + 1
            ElseIf LCase(gl_GetInstrumentType_Dic(sPinName)) Like "*dcvi*" Then
                ReDim Preserve Filter_Power_Ary(nTempCnt)
                Filter_Power_Ary(nTempCnt) = sPinName
                nTempCnt = nTempCnt + 1
            End If
        End If
    Next pin
    
    ReDim EVS_Detail_value(UBound(Filter_Power_Ary))

    '// ----Multi EVS ----
    If Multi_Function = False Then
        Mulit_EVS_Index = 1
        EVS_Type = "Regular-EVS"
    Else
        Mulit_EVS_Index = Mulit_EVS_Index
        EVS_Type = "Multi-EVS"
    End If
    S_WaitTime = S_WaitTime / Mulit_EVS_Index
    
    Dim TExec_up As Double
    Dim TExec_down As Double
    If (Flag_Serial = False) And (Parallel_Pin_Voltage <> "") Then
        ''Flag_Serial = False
        ''Parallel_Pin_Voltage = 1.1:VDD_SOC,VDD_GPU;1.0:VDDIO12
        EVS_Sequence = "EVS_Pre_Setting"
        Call Test_time_breakdown_Start(TExec_Before_Pat, Test_time_breakdown, EVS_Sequence)
        Call Test_time_breakdown_End(TExec_Before_Pat, Test_time_breakdown, EVS_Sequence, Mulit_Loop_EVS, S_WaitTime)

        For Mulit_Loop_EVS = 0 To Mulit_EVS_Index - 1

            EVS_Sequence = "Evs_Ramp_UP"
            Call Test_time_breakdown_Start(TExec_up, Test_time_breakdown, EVS_Sequence)
            Call Evs_Ramp_UPorDown_Parallel(Parallel_Pin_Voltage, "UP", d_PowerPin_ApplyVol(), S_WaitTime, glb_EVS_Disable_Printout, Open_LatchUp_measure, Step_number, Rising_Delay_time, TotalPwrLimit)
            Call Test_time_breakdown_End(TExec_up, Test_time_breakdown, EVS_Sequence, Mulit_Loop_EVS, S_WaitTime)

            DebugPrintFunc ""

            EVS_Sequence = "Evs_Ramp_DOWN"
            Call Test_time_breakdown_Start(TExec_down, Test_time_breakdown, EVS_Sequence)
            Call Evs_Ramp_UPorDown_Parallel(Parallel_Pin_Voltage, "DOWN", d_PowerPin_ApplyVol(), S_WaitTime, glb_EVS_Disable_Printout, Open_LatchUp_measure, Step_number, Rising_Delay_time, TotalPwrLimit)
            Call Test_time_breakdown_End(TExec_down, Test_time_breakdown, EVS_Sequence, Mulit_Loop_EVS, S_WaitTime)

            TheHdw.Wait Cooling_Time

        Next Mulit_Loop_EVS

    Else
    
        'TExec_Before_Pat = theexec.Timer(0)
        EVS_Sequence = "EVS_Pre_Setting"
        Call Test_time_breakdown_Start(TExec_Before_Pat, Test_time_breakdown, EVS_Sequence)
        Call EVS_Pre_Setting(dc_spec, Filter_Power_Ary, EVS_Detail_value, Step_number, Looping_Contorl, Looping_Range, Looping_Index_Name, Looping_Max_Steps_Name, Open_LatchUp_measure)
        Call Test_time_breakdown_End(TExec_Before_Pat, Test_time_breakdown, EVS_Sequence, Mulit_Loop_EVS, S_WaitTime)
        
        Call PrintIFoldInfo(Power_pin)    '2022.10.19 add print suggest iFold limit(use UVS256-HP, UVS256 need)
        
        If F_EVS_Invalid_Category = True Then
            ''------------If the EVS stress voltage is smaller than Pattern Setup Voltage------------
            ''------------Did not run the EVS Stress Instance------------
            For Each site In TheExec.sites
                TheExec.sites.item(site).FlagState("F_EVS_Fail_Stop") = logicTrue
            Next site
        Else
            'TExec_Before_Pat = theexec.Timer(TExec_Before_Pat)
            'theexec.DataLog.WriteComment "EVS_Pre_Setting : " + Format(TExec_Before_Pat * 1000#, "##0.000") + " msec"
            'evs pre
            'ProfileMarkLeave (preEVSseup)
            If Open_LatchUp_measure = False Then 'Or Looping_Contorl = False Then
                    'TheHdw.Alarms.StartMonitoringAlarms
            End If
            TheExec.Datalog.Setup.Shared.ascii.Columns.EnableCustomWidths = True
            TheExec.Datalog.Setup.Shared.ascii.Columns.Parametric.TestName.Width = 170
            TheExec.Datalog.Setup.Shared.ascii.Columns.Functional.TestName.Width = 100
            TheExec.Datalog.Setup.Shared.ascii.Columns.Functional.Pattern.Width = 75
            TheExec.Datalog.ApplySetup
    
            For Mulit_Loop_EVS = 0 To Mulit_EVS_Index - 1
                'TExec_Before_Pat = TheExec.Timer(0)
                EVS_Sequence = "Evs_Ramp_UP"
                'TExec_Before_Pat = 0
                Call Test_time_breakdown_Start(TExec_up, Test_time_breakdown, EVS_Sequence)
                Call Evs_Ramp_UPorDown(EVS_Detail_value, "UP", S_WaitTime, Filter_Power_Ary, Step_number, Rising_Delay_time, Open_LatchUp_measure, _
                                        Looping_Contorl, Looping_Index_Name, Looping_Max_Steps_Name, TotalPwrLimit)
                Call Test_time_breakdown_End(TExec_up, Test_time_breakdown, EVS_Sequence, Mulit_Loop_EVS, S_WaitTime)
                'TExec_Before_Pat = TheExec.Timer(TExec_Before_Pat) - S_WaitTime
                'TheExec.DataLog.WriteComment "EVS_Ramp_up" & Mulit_Loop_EVS + 1 & " : " + Format(TExec_Before_Pat * 1000#, "##0.000") + " msec"
                'TheExec.DataLog.WriteComment "EVS Stress time: " & S_WaitTime
                DebugPrintFunc ""
                'TExec_Before_Pat = TheExec.Timer(0)
                EVS_Sequence = "Evs_Ramp_DOWN"

                Call Test_time_breakdown_Start(TExec_down, Test_time_breakdown, EVS_Sequence)
                Call Evs_Ramp_UPorDown(EVS_Detail_value, "DOWN", S_WaitTime, Filter_Power_Ary, Step_number, Rising_Delay_time, Open_LatchUp_measure, _
                                        Looping_Contorl, Looping_Index_Name, Looping_Max_Steps_Name, TotalPwrLimit)
                Call Test_time_breakdown_End(TExec_down, Test_time_breakdown, EVS_Sequence, Mulit_Loop_EVS, S_WaitTime)
                
                TheHdw.Wait Cooling_Time
                'TExec_Before_Pat = TheExec.Timer(TExec_Before_Pat)
                'TheExec.DataLog.WriteComment "EVS_Ramp_down" & Mulit_Loop_EVS + 1 & " : " + Format(TExec_Before_Pat * 1000#, "##0.000") + " msec"
            Next Mulit_Loop_EVS
                
            Dim pin_count As Integer
            Dim Print_Looping_Inf As String
            Dim Pins As String
            Dim EVS_Level As Double
            Dim EVS_Instance_name As String
            Dim Gatecheck As Boolean
            Dim EVS_Gate_check As String
            Dim Die_X_location As New SiteLong
            Dim Die_Y_location As New SiteLong
            Dim m_InstanceName As String
            m_InstanceName = TheExec.DataManager.instancename

            Print_Looping_Inf = vbNullString
            EVS_Instance_name = m_InstanceName
            If Looping_Contorl = True Then
                TheHdw.Digital.ApplyLevelsTiming True, True, True, tlPowered
                TheHdw.Wait 0.002
                TheExec.Datalog.WriteComment "-------------------------EVS Collection mode Start-------------------------"
                TheExec.Datalog.WriteComment "Instance name:" & EVS_Instance_name
                For Each site In TheExec.sites
                    pin_count = 0
                    Print_Looping_Inf = vbNullString
                    Die_X_location = XCoord(site)
                    Die_Y_location = YCoord(site)
                    For Each pin In power_pin_ary
                        'If TheHdw.DCVS.Pins(Pin).Gate = False Then
                        If EVS_Detail_value(pin_count).Gate_check(site) = False Then
                                EVS_Gate_check = "Alarm:True"
                        Else
                                EVS_Gate_check = "Alarm:False"
                        End If
                        Pins = EVS_Detail_value(pin_count).PinName
                        EVS_Level = EVS_Detail_value(pin_count).LatchUp_Final_Value
                        Print_Looping_Inf = Print_Looping_Inf + Pins & ":" & EVS_Level & ";" & EVS_Gate_check & ";" ', Stress time:" & S_WaitTime & "(msec);"
                        pin_count = pin_count + 1
                    Next pin
                    TheExec.Datalog.WriteComment "EVS results: SITE" & site & ";" & "X,Y: " & Die_X_location & "," & Die_Y_location & ";" & Print_Looping_Inf & EVS_Type    '& Chr(10)
                Next site
                'TheExec.DataLog.WriteComment "EVS pin's Level: " & Print_Looping_Inf
                TheExec.Datalog.WriteComment "-------------------------EVS Collection mode End-------------------------"
            End If
                
            EVS_Sequence = "Total_time"
            Call Test_time_breakdown_End(Total_time, Test_time_breakdown, EVS_Sequence, Mulit_Loop_EVS, S_WaitTime)
            
    ''    If Open_LatchUp_measure = False Then 'Or Looping_Contorl = False Then
    ''        'TheHdw.Alarms.StopMonitoringAlarms
    ''        'TheHdw.Alarms.CloseAlarmWindow
    ''    End If
            
            TheHdw.Alarms.Check
            TheExec.Flow.TestLimit Cooling_Time, 0, 99, tlSignGreaterEqual, tlSignLessEqual, scaletype:=scaleNone, unit:=unitCustom, Tname:="EVS Cooling time", PinName:="EVS Cooling time", customUnit:="sec" 'BurstResult=1:Pass
    ''    TheExec.Flow.TestLimit 0, 0, 0, , , , , , , , , , , , , tlForceNone ''NOTE:JCIDE ADDS ON THE PLATFORM UFP
        
        End If
    End If
    'Total_time = TheExec.Timer(Total_time)
    'TheExec.DataLog.WriteComment "Total test time : " + Format(Total_time * 1000#, "##0.000") + " msec"
    If gl_EnableCurrentProfile Or gl_EnableVoltageProfile Then
        ProfileCollect_End glb_TestInstance
    End If
    
    Exit Function

errHandler:
    Call Print_Error_Message(Error_Info, "VBT_LIB_DC_Func.bas", "EVS_Static_Power_Ramp")
    If AbortTest Then Exit Function Else Resume Next
End Function

'[20230407][RF] Add DCVI EVS
Public Function EVS_Pre_Setting(dc_spec As String, power_pin_ary() As String, ByRef EVS_Detail_value() As Pins_detail, Optional Step_number As Integer = 10, Optional Looping_Control As Boolean, Optional Looping_Range As Double = 0.4, Optional Looping_Index_Name As String, Optional Looping_Max_Steps_Name As String, Optional Open_LatchUp_measure As Boolean)
On Error GoTo errHandler
    Dim Spec As Variant, Spec_Var As String, pin As Variant, Dc_spec_type As String
    Dim Step_value As Double
    Dim pin_count As Integer
    Dim End_Value As Double
    Dim Start_Value As Double
    Dim diff_value As Double
    Dim Looping_Index As Double, Looping_Max_Step As Double, Looping_Step_range As Double
    Dim Gap_Persent As Double
    Dim EVS_Index As Integer, Evs_End_Index As Integer
    Dim sPinName As String
    pin_count = 0  '  initial value
    Evs_End_Index = Step_number
 
    '\\\\\\\\\\\\\\\\\\\Calculate EVS Ramp up/down information\\\\\\\\\\\\\\\\\\\
    For Each pin In power_pin_ary
        sPinName = LCase(pin)
        Spec_Var = pin & "_VAR" '& Dc_spec_type
        
        If LCase(gl_GetInstrumentType_Dic(sPinName)) Like "*dcvs*" Then
            Start_Value = TheHdw.DCVS.Pins(sPinName).Voltage.value ''Start Voltage is read by hardware
        ElseIf LCase(gl_GetInstrumentType_Dic(sPinName)) Like "*dcvi*" Then
            Start_Value = TheHdw.DCVI.Pins(sPinName).Voltage ''Start Voltage is read by hardware
        Else
        End If
        
        End_Value = TheExec.Specs.DC.item(Spec_Var).Categories.item(dc_spec).max.value

        ''------------Check if the EVS stress voltage is larger than Pattern Setup Voltage------------
        If sPinName Like UCase("*efuse*") Then
            TheExec.Datalog.WriteComment "EVS Stress voltage includes Efuse Pin, Please check the argument with EVS test instance for Pin " & pin
            F_EVS_Invalid_Category = True
            Exit For
        Else
            If Start_Value > End_Value Then
                TheExec.Datalog.WriteComment "EVS Stress voltage is lower than Pattern Setup Voltage, Please check the DC Spec for Pin " & pin
                F_EVS_Invalid_Category = True
            End If
        End If
        
        ''------------Check if the EVS stress voltage is larger than Pattern Setup Voltage------------


        ''//////////////////////////Looping Function for EVS experiment//////////////////////////
        '' If we want to apply this method , we need to additionly change the EVS flow and assign varient at main flow
        ''
        '' Example as below:
        ''//////// Main flow ////////
        '' |     Opcode     |      Parameter     |
        '' | assign-integer |     Max_step 5     | <= Assign integer at IGXL Main flow
        '' | create-integer | EVS_Looping_INDEX  | <= Create integer at IGXL Main flow
        ''//////// Main flow ////////
        
        ''//////// EVS flow ////////
        '' |     Opcode    |                             Parameter                                 |
        '' |      For      | EVS_Looping_INDEX=0; EVS_Looping_INDEX< Max_Step; EVS_Looping_INDEX++ | <= For index for looping ,Start at 0
        '' |      test     |                         CpuSaEVS_Static_*pp**                         | <= Run Pattern
        '' |      test     |                        EVS_Static_Power_Up_CpuSa                        | <= Power up
        '' |      test     |                      EVS_Static_Power_Down_CpuSa                      | <= Power down
        '' |      Next     |                          EVS_Looping_INDEX                            |
        ''//////// EVS flow ////////
        
        '' Instance setting for "EVS_Static_Power_Up(Down)_CpuSa
        '' Addtional setting argument Looping_Control = true, Looping_Index_Name = "EVS_Looping_INDEX" and Looping_Max_Steps_Name = "Max_Step"

        If Looping_Control And Looping_Index_Name <> "" And Looping_Max_Steps_Name <> "" Then  ''Only need to change final value of Power Up
            Looping_Index = TheExec.Flow.var(Looping_Index_Name).value '' get index from flow
            Looping_Max_Step = TheExec.Flow.var(Looping_Max_Steps_Name).value '' get max step from flow
            Looping_Step_range = Looping_Range / (Looping_Max_Step - 1) '' get max step from flow
            End_Value = TheExec.Specs.DC.item(Spec_Var).Categories.item(dc_spec).max.value ''' update end value for each looping
            'Start_Value = End_Value - Looping_Range
            Dim Looping_Start_Voltage As Double
            Looping_Start_Voltage = End_Value - Looping_Range
            End_Value = Looping_Start_Voltage + Looping_Step_range * Looping_Index
        End If
        If Not (Looping_Control) And Looping_Index_Name <> "" And Looping_Max_Steps_Name <> "" Then
        '' You can keep the opcode "For" on IGXL flow and trun looping_Control into false.
        '' If you still have Looping_Max_Steps_Name , it will set it to 1 which means it won't loop on the IGXL flow anymore.
            TheExec.Flow.var(Looping_Max_Steps_Name).value = 1
'        ElseIf Not (Looping_Control) And Looping_Index_Name <> "" And Looping_Max_Steps_Name <> "" Then
'            TheExec.Flow.var(Looping_Max_Steps_Name).Value = 1
        End If
        ''//////////////////////////Looping Function for EVS experiment//////////////////////////
        
        EVS_Detail_value(pin_count).LatchUp_Final_Value = Format(End_Value, "0.00")
        diff_value = End_Value - Start_Value
        Gap_Persent = Abs(diff_value / Start_Value) '''+-0.5%
        If Gap_Persent < 0.005 Then 'small than 0.5 % pin don't need to rise at all
            EVS_Detail_value(pin_count).Pin_rise = False
        Else
        
            Step_value = diff_value / Evs_End_Index ' calculate step value by end index

            EVS_Detail_value(pin_count).Start_voltage = Start_Value
            EVS_Detail_value(pin_count).Start_voltage_up = Start_Value
            EVS_Detail_value(pin_count).Start_voltage_down = End_Value
            
            EVS_Detail_value(pin_count).Step_value = Step_value
            EVS_Detail_value(pin_count).Step_value_up = Step_value
            EVS_Detail_value(pin_count).Step_value_down = Step_value * (-1)
            
            EVS_Detail_value(pin_count).Pin_rise = True
            EVS_Detail_value(pin_count).PinName = CStr(pin)
            'Pin_detail_dict.Add CStr(Pin), Temp_value
        End If
        pin_count = pin_count + 1
    Next pin
    Exit Function
    '\\\\\\\\\\\\\\\\\\\ Calculate EVS Ramp up/down information\\\\\\\\\\\\\\\\\\\
errHandler:
    Call Print_Error_Message(Error_Warning_Info.Error_Info, "VBT_LIB_DC_Func", "EVS_Pre_Setting")
    If AbortTest Then Exit Function Else Resume Next
End Function

''20230223: Modidfied to add when specific site pin alarm, then gate off "All_Power" in site.
'[20230407][RF] Add DCVI EVS
'[20230731][T-Pal][Tank] Add meter reed value need use PinListData to store
Public Function Evs_Ramp_UPorDown(EVS_Detail_value() As Pins_detail, Direction As String, S_WaitTime As Double, power_pin_ary() As String, _
                                Optional Step_number As Integer = 10, Optional Rising_Delay_time As Double = 0.02, Optional Open_LatchUp_measure As Boolean = False, _
                                Optional Looping_Control As Boolean = False, Optional Looping_Index_Name As String = vbNullString, Optional Looping_Max_Steps_Name As String = vbNullString, _
                                Optional TotalPwrLimit As Double)

    On Error GoTo errHandler
    Dim funcName As String:: funcName = "Evs_Ramp_UPorDown"
    
    Dim pin_count As Integer
    Dim PowerPin_Final_Value As New PinListData
    Dim ForceV As Double
    Dim LatchUp_measure_Value As New PinListData
    Dim EVS_Index As Integer, Evs_End_Index As Integer
    Dim Gatecheck As Boolean
    Dim power_pin_value As Double
    Dim Latch_Up_name As String
    Dim pin As Variant
    Dim EVS_Index_Value As Integer
    Dim EVS_End_Index_Value As Integer:: Evs_End_Index = Step_number
    Dim ifold As Double
    Dim m_InstanceName As String
    Dim Total_Power_Name As String
    Dim Power_Name As String
    m_InstanceName = TheExec.DataManager.instancename
    Dim Power As New PinListData
    Dim TotalPower As New SiteDouble
    Dim CorePower_Cnt As Long
    Dim All_Core_Power() As String
    Dim Core_power As String
    Dim All_Power As Variant
    Dim alarm_occur As Boolean
    Dim sPinName As String
    Dim sSlotType As String
    Dim site As Variant 'Carter, 20240304
    If TotalPwrLimit = 0 Then TotalPwrLimit = 20
        
    For EVS_Index = 1 To Evs_End_Index
        If UCase(Direction) = "UP" Then
            pin_count = 0 '' reset pin index
        ElseIf UCase(Direction) = "DOWN" Then
            pin_count = UBound(power_pin_ary())
        End If
        
        For Each pin In power_pin_ary
            Dim PinName As String
            sPinName = LCase(pin)
            If UCase(Direction) = "UP" Then
                EVS_Detail_value(pin_count).Start_voltage = EVS_Detail_value(pin_count).Start_voltage_up
                EVS_Detail_value(pin_count).Step_value = EVS_Detail_value(pin_count).Step_value_up
            ElseIf UCase(Direction) = "DOWN" Then
                EVS_Detail_value(pin_count).Start_voltage = EVS_Detail_value(pin_count).Start_voltage_down
                EVS_Detail_value(pin_count).Step_value = EVS_Detail_value(pin_count).Step_value_down
            End If
            
            PinName = LCase(EVS_Detail_value(pin_count).PinName)
            sSlotType = UCase(gl_GetInstrument_Dic(PinName))
            If EVS_Detail_value(pin_count).Pin_rise Then
                ForceV = EVS_Detail_value(pin_count).Start_voltage + EVS_Index * EVS_Detail_value(pin_count).Step_value
                
                If LCase(gl_GetInstrumentType_Dic(PinName)) Like "*dcvs*" Then
                    TheHdw.DCVS.Pins(PinName).Voltage.value = ForceV
                ElseIf LCase(gl_GetInstrumentType_Dic(PinName)) Like "*dcvi*" Then
                    TheHdw.DCVI.Pins(PinName).Voltage = ForceV
                Else
                    Call Print_Error_Message(Error_Warning_Info.Warning_Info, "VBT_LIB_DC_Func", "Evs_Ramp_UPorDown", "Pin did not exist in gl_GetInstrumentType_Dic!!")
                End If
        
                If EVS_Index = Evs_End_Index Then
                    Power.AddPin (PinName)
                    PowerPin_Final_Value.AddPin (PinName)
                    
                    If LCase(gl_GetInstrumentType_Dic(PinName)) Like "*dcvs*" Then
                        PowerPin_Final_Value.Pins(PinName) = TheHdw.DCVS.Pins(PinName).Voltage.value
                    ElseIf LCase(gl_GetInstrumentType_Dic(PinName)) Like "*dcvi*" Then
                        PowerPin_Final_Value.Pins(PinName) = TheHdw.DCVI.Pins(PinName).Voltage.value
                    Else
                        Call Print_Error_Message(Error_Warning_Info.Warning_Info, "VBT_LIB_DC_Func", "Evs_Ramp_UPorDown", "Pin did not exist in gl_GetInstrumentType_Dic!!")
                    End If
                    
                End If
            End If
            
            If Open_LatchUp_measure Then
                If LCase(gl_GetInstrumentType_Dic(PinName)) Like "*dcvs*" Then
                    TheHdw.DCVS.Pins(PinName).Meter.mode = tlDCVSMeterCurrent
                ElseIf LCase(gl_GetInstrumentType_Dic(PinName)) Like "*dcvi*" Then
                    TheHdw.DCVI.Pins(PinName).Meter.mode = tlDCVIMeterCurrent
                Else
                    Call Print_Error_Message(Error_Warning_Info.Warning_Info, "VBT_LIB_DC_Func", "Evs_Ramp_UPorDown", "Pin did not exist in gl_GetInstrumentType_Dic!!")
                End If
                            
          '' Print out measure current value if want to collect the Latch up data
                If TheExec.enableWord("CurrentProfile") = True Or TheExec.Flow.enableWord("VoltageProfile") = True Or Profile_byflow = True Then
                    LatchUp_measure_Value.AddPin (PinName)
                    LatchUp_measure_Value.Pins(PinName) = 0.001
                Else
                    Select Case sSlotType
                        Case glbConstIns_VS5A, glbConstIns_HEXVS, glbConstIns_VSM:
                            TheHdw.DCVS.Pins(PinName).Meter.mode = tlDCVSMeterCurrent
                            LatchUp_measure_Value = TheHdw.DCVS.Pins(PinName).Meter.Read(tlStrobe, 10, , tlDCVSMeterReadingFormatAverage)
                        Case glbConstIns_VS800MA, glbConstIns_VHDVS:    'UVS256 and UVS256-HP sample size = 1
                            TheHdw.DCVS.Pins(PinName).Meter.mode = tlDCVSMeterCurrent
                            LatchUp_measure_Value = TheHdw.DCVS.Pins(PinName).Meter.Read(tlStrobe, 1)
                        Case glbConstIns_DC07, glbConstIns_DC30, glbConstIns_DC75:
                            TheHdw.DCVI.Pins(PinName).Meter.mode = tlDCVIMeterCurrent
                            LatchUp_measure_Value = TheHdw.DCVI.Pins(PinName).Meter.Read(tlStrobe, 10, , tlDCVIMeterReadingFormatAverage)
                        Case Else
                            Call Print_Error_Message(Error_Warning_Info.Warning_Info, "VBT_LIB_DC_Func", "Evs_Ramp_UPorDown", "Slot type did not define!!")
                    End Select
                    
                End If
                Latch_Up_name = m_InstanceName & "_" & "Latch_up_data_MeasI_" & Replace(CStr(EVS_Detail_value(pin_count).LatchUp_Final_Value), ".", "p") & "V"
                TheExec.Flow.TestLimit resultVal:=LatchUp_measure_Value.Pins(PinName), PinName:=PinName, scaletype:=scaleNone, unit:=unitAmp, formatStr:="%.3f", Tname:=Latch_Up_name, ForceVal:=ForceV, ForceUnit:=unitVolt, ForceResults:=tlForceNone
            Else
                If EVS_Index = Evs_End_Index And UCase(Direction) = "UP" Then
                    If EVS_Detail_value(pin_count).Pin_rise Then
                        If TheExec.enableWord("CurrentProfile") = True Or TheExec.Flow.enableWord("VoltageProfile") = True Or Profile_byflow = True Then
                            LatchUp_measure_Value.AddPin (PinName)
                            LatchUp_measure_Value.Pins(PinName) = 0.001
                        Else
                            Select Case sSlotType
                                Case glbConstIns_VS5A, glbConstIns_HEXVS, glbConstIns_VSM:
                                    TheHdw.DCVS.Pins(PinName).Meter.mode = tlDCVSMeterCurrent
                                    LatchUp_measure_Value = TheHdw.DCVS.Pins(PinName).Meter.Read(tlStrobe, 10, , tlDCVSMeterReadingFormatAverage)
                                Case glbConstIns_VS800MA, glbConstIns_VHDVS:    'UVS256 and UVS256-HP sample size = 1
                                    TheHdw.DCVS.Pins(PinName).Meter.mode = tlDCVSMeterCurrent
                                    LatchUp_measure_Value = TheHdw.DCVS.Pins(PinName).Meter.Read(tlStrobe, 1)
                                Case glbConstIns_DC07, glbConstIns_DC30, glbConstIns_DC75:
                                    TheHdw.DCVI.Pins(PinName).Meter.mode = tlDCVIMeterCurrent
                                    LatchUp_measure_Value = TheHdw.DCVI.Pins(PinName).Meter.Read(tlStrobe, 10, , tlDCVIMeterReadingFormatAverage)
                                Case Else
                                    Call Print_Error_Message(Error_Warning_Info.Warning_Info, "VBT_LIB_DC_Func", "Evs_Ramp_UPorDown", "Slot type did not define!!")
                            End Select
                            
                        End If
                        Latch_Up_name = m_InstanceName & "_" & "Latch_up_data_MeasI_" & Replace(CStr(EVS_Detail_value(pin_count).LatchUp_Final_Value), ".", "p") & "V"
                        Total_Power_Name = "Total_Power_" & m_InstanceName
                        TheExec.Flow.TestLimit resultVal:=LatchUp_measure_Value.Pins(PinName), PinName:=PinName, scaletype:=scaleNone, unit:=unitAmp, formatStr:="%.3f", Tname:=Latch_Up_name, ForceVal:=ForceV, ForceUnit:=unitVolt, ForceResults:=tlForceNone
                        Power.Pins(PinName) = LatchUp_measure_Value.Pins(PinName).Multiply(PowerPin_Final_Value.Pins(PinName))
                        TotalPower = TotalPower.Add(Power.Pins(PinName))
                    End If
                End If
            
            End If
         '' Print out measure current value if want to collect the Latch up data

            If UCase(Direction) = "UP" Then
                pin_count = pin_count + 1
            ElseIf UCase(Direction) = "DOWN" Then
                pin_count = pin_count - 1
            End If
        Next pin
        
        TheHdw.Wait Rising_Delay_time 'delay time of each ramp up
        
    Next EVS_Index
    '\\\\\\\\\\\\\\\\\\\ End Power up/Down \\\\\\\\\\\\\\\\\
    pin_count = 0 'reset pin count
    '///////////stress time after power up///////////
    If UCase(Direction) = "UP" Then
                ''default = 20
        TheExec.Flow.TestLimit TotalPower, 0, TotalPwrLimit, tlSignGreaterEqual, tlSignLessEqual, scaletype:=scaleNone, unit:=unitCustom, Tname:=Total_Power_Name, PinName:="Total Power", customUnit:="W"      'BurstResult=1:Pass

''      '-----------------------------------------------------------
''      'if need more bin info then active below VBT(T-cre request)
''      '-----------------------------------------------------------
''        Dim F_EVS_POWER As String
''        For Each site In TheExec.sites
''            If TotalPower(site) > TotalPwrLimit Then    ''original 150
''                TheExec.sites(site).FlagState("F_EVS_POWER") = logicTrue
''            Else
''                TheExec.sites(site).FlagState("F_EVS_POWER") = logicFalse
''            End If
''        Next site
''

        TheHdw.Wait S_WaitTime

        TheExec.Datalog.WriteComment ""
        TheExec.Datalog.WriteComment "-------------------------EVS Power ramp start-------------------------"
        TheExec.Flow.TestLimit S_WaitTime, 0, 99, tlSignGreaterEqual, tlSignLessEqual, scaletype:=scaleNone, unit:=unitCustom, Tname:="EVS stress time", PinName:="EVS stress time", customUnit:="sec" 'BurstResult=1:Pass
    '///////////Print out final value of each pin (Power up and down)///////////
        For Each pin In power_pin_ary
            If EVS_Detail_value(pin_count).Pin_rise Then
                TheExec.Flow.TestLimit PowerPin_Final_Value.Pins(pin), , , tlSignGreaterEqual, tlSignLessEqual, scaletype:=scaleNone, unit:=unitCustom, customUnit:="V"      'BurstResult=1:Pass
            End If
            pin_count = pin_count + 1
        Next pin
        pin_count = 0 ' MQ for rest pincount
        For Each pin In power_pin_ary
            If EVS_Detail_value(pin_count).Pin_rise Then
                Power_Name = m_InstanceName & "_" & "Power"
                TheExec.Flow.TestLimit Power.Pins(pin), , , tlSignGreaterEqual, tlSignLessEqual, scaletype:=scaleNone, unit:=unitCustom, Tname:=Power_Name, customUnit:="W"     'BurstResult=1:Pass
            End If
            pin_count = pin_count + 1
        Next pin
    '///////////Print out final value of each pin (Power up and down)///////////
    '///////////Alarm check after stress during power up///////////
        For Each site In TheExec.sites
            pin_count = 0
            For Each pin In power_pin_ary
                sPinName = LCase(pin)
                If LCase(gl_GetInstrumentType_Dic(sPinName)) Like "*dcvs*" Then
                    Gatecheck = TheHdw.DCVS.Pins(sPinName).Gate
                    EVS_Detail_value(pin_count).Gate_check(site) = Gatecheck
                    If TheHdw.DCVS.Pins(sPinName).Gate = False Then
    '                    Dim power_pin_value As Double
                        power_pin_value = TheHdw.DCVS.Pins(sPinName).Voltage.value
                        TheExec.Flow.TestLimit power_pin_value, , , tlSignGreaterEqual, tlSignLessEqual, scaletype:=scaleNone, unit:=unitCustom, customUnit:="V", PinName:=sPinName, Tname:="Vlotage_PatternAlarm_After_Stress"
                        TheExec.Flow.TestLimit Gatecheck, 1, 1, tlSignGreaterEqual, tlSignLessEqual, scaletype:=scaleNone, Tname:="Gate_PatternAlarm_After_Stress", PinName:=sPinName
                    End If
                ElseIf LCase(gl_GetInstrumentType_Dic(sPinName)) Like "*dcvi*" Then
                    Gatecheck = TheHdw.DCVI.Pins(sPinName).Gate
                    EVS_Detail_value(pin_count).Gate_check(site) = Gatecheck
                    If TheHdw.DCVI.Pins(sPinName).Gate = False Then
    '                    Dim power_pin_value As Double
                        power_pin_value = TheHdw.DCVI.Pins(sPinName).Voltage
                        TheExec.Flow.TestLimit power_pin_value, , , tlSignGreaterEqual, tlSignLessEqual, scaletype:=scaleNone, unit:=unitCustom, customUnit:="V", PinName:=sPinName, Tname:="Vlotage_PatternAlarm_After_Stress"
                        TheExec.Flow.TestLimit Gatecheck, 1, 1, tlSignGreaterEqual, tlSignLessEqual, scaletype:=scaleNone, Tname:="Gate_PatternAlarm_After_Stress", PinName:=sPinName
                    End If
                Else
                    Call Print_Error_Message(Error_Warning_Info.Warning_Info, "VBT_LIB_DC_Func", "Evs_Ramp_UPorDown", "Pin did not exist in gl_GetInstrumentType_Dic!!")
                End If
                
                pin_count = pin_count + 1
            Next pin
            If alarm_occur = True Then '20221208 Jayden, gate off all stress rail if alarm occured.
                TheHdw.DCVS.Pins("All_Power").Gate = False
                TheHdw.Wait 0.0001
                alarm_occur = False
            End If
        Next site

'        ///////////Alarm check after stress during power up///////////
    ElseIf UCase(Direction) = "DOWN" Then
    '///////////Alarm check after stress during power down///////////
        
        
        
        'Core_power = "VDD_AVE,VDD_DCS_DDR,VDD_DISP,VDD_ECPU,VDD_GPU,VDD_PCPU,VDD_SOC,VDD_SRAM_ANE,VDD_SRAM_CPU,VDD_SRAM_GPU,VDD_SRAM_SOC,VDD_LOW,VDD_FIXED"
        'All_Core_Power() = Split(Core_power, ",")
        TheExec.DataManager.DecomposePinList "CorePower", All_Core_Power, CorePower_Cnt

        For Each site In TheExec.sites
            'For Each Pin In power_pin_ary
            For Each All_Power In All_Core_Power
                sPinName = LCase(All_Power)
                If LCase(gl_GetInstrumentType_Dic(sPinName)) Like "*dcvs*" Then
                    Gatecheck = TheHdw.DCVS.Pins(All_Power).Gate
                    If TheHdw.DCVS.Pins(All_Power).Gate = False Then
                        TheExec.Flow.TestLimit Gatecheck, -1, -1, tlSignGreaterEqual, tlSignLessEqual, scaletype:=scaleNone, Tname:="Gate_PatternAlarm_After_EVS_ramp_down", PinName:=All_Power
                    End If
                ElseIf LCase(gl_GetInstrumentType_Dic(sPinName)) Like "*dcvi*" Then
                    Gatecheck = TheHdw.DCVS.Pins(All_Power).Gate
                    If TheHdw.DCVI.Pins(All_Power).Gate = False Then
                        TheExec.Flow.TestLimit Gatecheck, -1, -1, tlSignGreaterEqual, tlSignLessEqual, scaletype:=scaleNone, Tname:="Gate_PatternAlarm_After_EVS_ramp_down", PinName:=All_Power
                    End If
                Else
                    Call Print_Error_Message(Error_Warning_Info.Warning_Info, "VBT_LIB_DC_Func", "Evs_Ramp_UPorDown", "Pin did not exist in gl_GetInstrumentType_Dic!!")
                End If
            Next All_Power
            If alarm_occur = True Then '20221208 Jayden, gate off all stress rail if alarm occured.
                TheHdw.DCVS.Pins("All_Power").Gate = False
                TheHdw.Wait 0.0001
                alarm_occur = False
            End If
        Next site
        TheExec.Datalog.WriteComment ""
        TheExec.Datalog.WriteComment "--------------------------EVS Power ramp end--------------------------"
    
    End If
    '///////////stress time after power up///////////
    Exit Function

errHandler:
    Call Print_Error_Message(Error_Warning_Info.Error_Info, "VBT_LIB_DC_Func", "Evs_Ramp_UPorDown")
    If AbortTest Then Exit Function Else Resume Next
End Function


Public Function Test_time_breakdown_End(ByRef Timer As Double, Test_time_breakdown As Boolean, EVS_Sequence As String, Mulit_Loop_EVS As Integer, S_WaitTime As Double)
On Error GoTo errHandler 'Add ErrHandler 2023/08/18

    If Test_time_breakdown = True Then
        Timer = TheExec.Timer(Timer)
        If EVS_Sequence = "EVS_Pre_Setting" Then
            TheExec.Datalog.WriteComment "EVS_Pre_Setting : " + Format(Timer * 1000#, "##0.000") + " msec"
        End If
        If EVS_Sequence = "Evs_Ramp_UP" Then
            TheExec.Datalog.WriteComment "EVS_Ramp_up" & Mulit_Loop_EVS + 1 & " : " + Format((Timer - S_WaitTime) * 1000#, "##0.000") + " msec"
        End If
        If EVS_Sequence = "Evs_Ramp_DOWN" Then
            TheExec.Datalog.WriteComment "EVS_Ramp_down" & Mulit_Loop_EVS + 1 & " : " + Format(Timer * 1000#, "##0.000") + " msec"
        End If
        If EVS_Sequence = "Total_time" Then
            TheExec.Datalog.WriteComment "Total test time : " + Format(Timer * 1000#, "##0.000") + " msec"
        End If
    End If
    
Exit Function 'Add ErrHandler 2023/08/18
errHandler: 'Add ErrHandler 2023/08/18
    Call Print_Error_Message(Error_Info, "VBT_LIB_DC_Func", "Test_time_breakdown_End") 'Add ErrHandler 2023/08/18
    If AbortTest Then Exit Function Else Resume Next 'Add ErrHandler 2023/08/18
End Function

Public Function Test_time_breakdown_Start(ByRef Timer As Double, Test_time_breakdown As Boolean, EVS_Sequence As String)
On Error GoTo errHandler 'Add ErrHandler 2023/08/18

    If Test_time_breakdown = True Then
        Timer = TheExec.Timer(0)
    End If
    
Exit Function 'Add ErrHandler 2023/08/18
errHandler: 'Add ErrHandler 2023/08/18
    Call Print_Error_Message(Error_Info, "VBT_LIB_DC_Func", "Test_time_breakdown_Start") 'Add ErrHandler 2023/08/18
    If AbortTest Then Exit Function Else Resume Next 'Add ErrHandler 2023/08/18
End Function


Public Function EVS_Static_Power_Ramp_PAEVS(dc_spec As String, S_WaitTime As Double, Power_pin As String, Optional Step_number As Integer = 10, Optional Rising_Delay_time As Double = 0.02, Optional Looping_Contorl As Boolean = False, Optional Looping_Range As Double = 0.4, Optional Looping_Index_Name As String = vbNullString, Optional Looping_Max_Steps_Name As String = vbNullString, Optional Open_LatchUp_measure As Boolean = False, Optional Multi_Function As Boolean = False, Optional Mulit_EVS_Index As Integer = 1, Optional Test_time_breakdown As Boolean = False, Optional Cooling_Time As Double = 0, Optional Block As String, Optional PowerPin1 As String, Optional PowerPin2 As String, Optional PA_Enable As Boolean = True) As Long
On Error GoTo errHandler
    If Looping_Contorl And (Looping_Max_Steps_Name = "" Or Looping_Index_Name = "") Then
        TheExec.Datalog.WriteComment "Error!! Please make sure you fill in argument both Looping_Max_Steps_Name and Looping_Index_Name "
        GoTo errHandler
    End If
    Dim funcName As String:: funcName = "EVS_Static_Power_Ramp"
    Dim inst_name As String
    'Dim Pin As Variant
    Dim i As Long
    Dim Spec As Variant
    Dim Spec_Var As String
    Dim pin As Variant
    Dim power_pin_ary() As String
    Dim power_pin_cnt As Long
    Dim Dc_spec_type As String
    Dim EVS_Detail_value() As Pins_detail ''Define each pin information
    
    Dim EVS_Detail_value1() As Pins_detail ''Define each pin information
    Dim EVS_Detail_value2() As Pins_detail ''Define each pin information
    Dim Mulit_Loop_EVS As Integer
    Dim TExec_Before_Pat As Double
    Dim Total_time As Double
    Dim EVS_Sequence As String
    Dim EVS_Type As String
    'If Test_time_breakdown = True Then
        'Total_time = TheExec.Timer(0)
    'End If
    '======================PACAVS info
    Dim PACAVSi As Integer
    
    Dim PAi As Integer, PA_Counter As Integer
    
    Dim pin_count As Integer
    Dim Print_Looping_Inf As String
    Dim Pins As String
    Dim EVS_Level As Double
    Dim EVS_Instance_name As String
    Dim Gatecheck As Boolean
    Dim EVS_Gate_check As String
    Dim Die_X_location As New SiteLong
    Dim Die_Y_location As New SiteLong
    Dim m_InstanceName As String
    
    
    Dim TExec_up As Double
    Dim TExec_down As Double
    
    Dim PABol As Boolean
    Dim s_ErrorMsg As String
    Dim site As Variant 'Carter, 20240304
    
    
    s_ErrorMsg = vbNullString
    PABol = False
    
    TheExec.DataManager.DecomposePinList Power_pin, power_pin_ary(), power_pin_cnt      'power_pin_ary = Split(Power_pin, ",")
    If power_pin_cnt <> 0 Then
        If PowerPin1 <> "" And PowerPin2 <> "" Then PABol = True
        If PABol Then
            PACAVSi = 3 ' 2
        Else
            PACAVSi = 1
        End If
        
        'F_start_profile = True
        TheHdw.Digital.ApplyLevelsTiming True, True, True, tlPowered    '''Add for applying the prestress condition
        
        EVS_Sequence = "Total_time"
        Call Test_time_breakdown_Start(Total_time, Test_time_breakdown, EVS_Sequence)
        
    
'        Call Test_time_breakdown_Start(TExec_Before_Pat, Test_time_breakdown, EVS_Sequence)
        
        TheExec.Datalog.Setup.Shared.ascii.Columns.EnableCustomWidths = True
        TheExec.Datalog.Setup.Shared.ascii.Columns.Parametric.TestName.Width = 170
        TheExec.Datalog.Setup.Shared.ascii.Columns.Functional.TestName.Width = 100
        TheExec.Datalog.Setup.Shared.ascii.Columns.Functional.Pattern.Width = 75
        TheExec.Datalog.ApplySetup
        SetCurrentRange_From_Ifold Power_pin 'apply irange w/ ifold
        
        ReDim EVS_Detail_value(UBound(power_pin_ary))
        '=====pre EVS set up
        EVS_Sequence = "EVS_Pre_Setting"
        Call Test_time_breakdown_Start(TExec_Before_Pat, Test_time_breakdown, EVS_Sequence)
        Call EVS_Pre_Setting_wPACAVS(dc_spec, power_pin_ary, EVS_Detail_value, Step_number, Looping_Contorl, Looping_Range, Looping_Index_Name, Looping_Max_Steps_Name, Open_LatchUp_measure, PowerPin1, PowerPin2)
        Call Test_time_breakdown_PAEVS_End(TExec_Before_Pat, Test_time_breakdown, EVS_Sequence, Mulit_Loop_EVS, S_WaitTime, Cooling_Time, Block)
    
        
        
        '// ----Multi EVS ----
        If Multi_Function = False Then
            Mulit_EVS_Index = 1
            EVS_Type = "Regular-EVS"
        Else
            Mulit_EVS_Index = Mulit_EVS_Index
            EVS_Type = "Multi-EVS"
        End If
        S_WaitTime = S_WaitTime / Mulit_EVS_Index
    
        If Open_LatchUp_measure = False Then 'Or Looping_Contorl = False Then
            'TheHdw.Alarms.StartMonitoringAlarms
        End If

        
        For Mulit_Loop_EVS = 0 To Mulit_EVS_Index - 1
            'TExec_Before_Pat = TheExec.Timer(0)
            
            For PAi = 1 To PACAVSi
            
                If PAi = 1 Then
                    EVS_Sequence = "Evs_Ramp_UP"

                    Call Test_time_breakdown_Start(TExec_up, Test_time_breakdown, EVS_Sequence)
                    Call Evs_Ramp_UPorDown_PAEVS(EVS_Detail_value, "UP", S_WaitTime, power_pin_ary, Step_number, Rising_Delay_time, Open_LatchUp_measure, Looping_Contorl, Looping_Index_Name, Looping_Max_Steps_Name, Cooling_Time, Mulit_Loop_EVS, Mulit_EVS_Index, PAi, PA_Enable)
                    
                    Call Test_time_breakdown_PAEVS_End(TExec_up, Test_time_breakdown, EVS_Sequence, Mulit_Loop_EVS, S_WaitTime, Cooling_Time, Block)
            
                ElseIf PAi = 2 Then
                    EVS_Sequence = "Evs_Ramp_DOWN"
                    
                    Call Test_time_breakdown_Start(TExec_down, Test_time_breakdown, EVS_Sequence)
                
                    Call Evs_Ramp_UPorDown_PAEVS(EVS_Detail_value, "DOWN", S_WaitTime, power_pin_ary, Step_number, Rising_Delay_time, Open_LatchUp_measure, Looping_Contorl, Looping_Index_Name, Looping_Max_Steps_Name, Cooling_Time, Mulit_Loop_EVS, Mulit_EVS_Index, PAi, PA_Enable)
        
                    Call Test_time_breakdown_PAEVS_End(TExec_down, Test_time_breakdown, EVS_Sequence, Mulit_Loop_EVS, S_WaitTime, Cooling_Time, Block)
                End If
                '=====EVS stress wait
                If PAi = 1 And PA_Enable = False Then
                    EVS_Sequence = "Evs_Ramp_DOWN"
                    Call Test_time_breakdown_Start(TExec_down, Test_time_breakdown, EVS_Sequence)
                
                    Call Evs_Ramp_UPorDown_PAEVS(EVS_Detail_value, "DOWN", S_WaitTime, power_pin_ary, Step_number, Rising_Delay_time, Open_LatchUp_measure, Looping_Contorl, Looping_Index_Name, Looping_Max_Steps_Name, Cooling_Time, Mulit_Loop_EVS, Mulit_EVS_Index, PAi, PA_Enable)
        
                    Call Test_time_breakdown_PAEVS_End(TExec_down, Test_time_breakdown, EVS_Sequence, Mulit_Loop_EVS, S_WaitTime, Cooling_Time, Block)
                End If
                'PA_Counter = PA_Counter + 1
                    
            '    TheExec.Flow.TestLimit S_WaitTime, 0, 99, tlSignGreaterEqual, tlSignLessEqual, scaletype:=scaleNone, Unit:=unitCustom, Tname:="EVS_stress_time" & "_" & Block & "_" & Mulit_Loop_EVS + 1, PinName:="EVS_stress_time", customUnit:="sec" 'BurstResult=1:Pass
                DebugPrintFunc ""
                'TExec_Before_Pat = TheExec.Timer(0)
                If (Mulit_Loop_EVS = Mulit_EVS_Index - 1) And PAi = PACAVSi And PA_Enable = True Then
                    EVS_Sequence = "Evs_Ramp_UP"
                    Call Test_time_breakdown_Start(TExec_down, Test_time_breakdown, EVS_Sequence)
                
                    Call Evs_Ramp_UPorDown_PAEVS(EVS_Detail_value, "UP", S_WaitTime, power_pin_ary, Step_number, Rising_Delay_time, Open_LatchUp_measure, Looping_Contorl, Looping_Index_Name, Looping_Max_Steps_Name, Cooling_Time, Mulit_Loop_EVS, Mulit_EVS_Index, PAi, PA_Enable)
        
                    Call Test_time_breakdown_PAEVS_End(TExec_down, Test_time_breakdown, EVS_Sequence, Mulit_Loop_EVS, S_WaitTime, Cooling_Time, Block)
                End If
                'PA_Counter = PA_Counter + 1
                
            '===============0912 EVS Alarm================
    
                'For Each site In TheExec.sites
                '    For Each pin In power_pin_ary
                '        If TheHdw.DCVS.Pins(pin).Gate = False Then
                '            TheHdw.DCVS.Pins("All_Power").Gate = False
                '            TheHdw.Wait 0.0001
                '            Exit For
                '        End If
                '    Next pin
                'Next site
    
                '===============0912 EVS Alarm================
                
                
            Next PAi
    
            'TheHdw.Wait Cooling_Time
        
        Next Mulit_Loop_EVS
        'TheHdw.Wait Cooling_Time
        m_InstanceName = TheExec.DataManager.instancename
        Print_Looping_Inf = vbNullString
        EVS_Instance_name = m_InstanceName
        If Looping_Contorl = True Then
            TheHdw.Digital.ApplyLevelsTiming True, True, True, tlPowered
            TheHdw.Wait 0.002
            TheExec.Datalog.WriteComment "-------------------------EVS Collection mode Start-------------------------"
            TheExec.Datalog.WriteComment "Instance name:" & EVS_Instance_name
            For Each site In TheExec.sites
                pin_count = 0
                Print_Looping_Inf = vbNullString
                Die_X_location = XCoord(site)
                Die_Y_location = YCoord(site)
                For Each pin In power_pin_ary
                    'If TheHdw.DCVS.Pins(Pin).Gate = False Then
                    If EVS_Detail_value(pin_count).Gate_check(site) = False Then
                        EVS_Gate_check = "Alarm:True"
                    Else
                        EVS_Gate_check = "Alarm:False"
                    End If
                    Pins = EVS_Detail_value(pin_count).PinName
                    EVS_Level = EVS_Detail_value(pin_count).LatchUp_Final_Value
                    Print_Looping_Inf = Print_Looping_Inf + Pins & ":" & EVS_Level & ";" & EVS_Gate_check & ";" ', Stress time:" & S_WaitTime & "(msec);"
                    pin_count = pin_count + 1
                Next pin
                TheExec.Datalog.WriteComment "EVS results: SITE" & site & ";" & "X,Y: " & Die_X_location & "," & Die_Y_location & ";" & Print_Looping_Inf & EVS_Type    '& Chr(10)
            Next site
            'TheExec.DataLog.WriteComment "EVS pin's Level: " & Print_Looping_Inf
            TheExec.Datalog.WriteComment "-------------------------EVS Collection mode End-------------------------"
        End If
        
        If Open_LatchUp_measure = False Then 'Or Looping_Contorl = False Then
            'TheHdw.Alarms.StopMonitoringAlarms
            'TheHdw.Alarms.CloseAlarmWindow
        End If
        
        TheHdw.Alarms.Check
        TheExec.Flow.TestLimit Cooling_Time, 0, 99, tlSignGreaterEqual, tlSignLessEqual, scaletype:=scaleNone, unit:=unitCustom, Tname:="EVS_Cooling_time" & "_" & Block, PinName:="EVS_Cooling_time", customUnit:="sec" 'BurstResult=1:Pass
        TheExec.Flow.TestLimit 0, 0, 0, , , , , , , , , , , , , tlForceNone
        
        F_start_profile = False
        
        EVS_Sequence = "Total_time"
        Call Test_time_breakdown_PAEVS_End(Total_time, Test_time_breakdown, EVS_Sequence, Mulit_Loop_EVS, S_WaitTime, Cooling_Time, Block)
    
    Else
        TheExec.Flow.TestLimit resultVal:=1, lowVal:=-1, hiVal:=-1, unit:=unitNone, Tname:=TheExec.DataManager.instancename
        s_ErrorMsg = "Measure pin not exist."
        Call Print_Error_Message(Error_Warning_Info.Warning_Info, "VBT_LIB_DC_Func", "EVS_Static_Power_Ramp_PAEVS", s_ErrorMsg)
    End If
    
    Exit Function

errHandler:
    Call Print_Error_Message(Error_Warning_Info.Error_Info, "VBT_LIB_DC_Func", "EVS_Static_Power_Ramp_PAEVS")
    If AbortTest Then Exit Function Else Resume Next
End Function

Public Function SetCurrentRange_From_Ifold(powerPin As String) '211118
On Error GoTo errHandler 'Add ErrHandler 2023/08/18

    Dim Ifold_source As Double
    Dim PowerPins() As String
    Dim PinCnt As Long
    Dim i As Integer
    TheExec.DataManager.DecomposePinList powerPin, PowerPins(), PinCnt
    
    For i = 0 To UBound(PowerPins)
        If TheExec.DataManager.ChannelType(PowerPins(i)) <> "N/C" Then
            Dim targetifold As Double, Maxifold As Double
            targetifold = TheHdw.DCVS.Pins(PowerPins(i)).CurrentLimit.Source.FoldLimit.level.value
            Maxifold = TheHdw.DCVS.Pins(PowerPins(i)).CurrentRange.max
            ''Debug.Print PowerPins(i) & "," & targetifold & "," & Maxifold
            If targetifold > Maxifold Then targetifold = Maxifold
            
            Ifold_source = targetifold
            TheHdw.DCVS.Pins(PowerPins(i)).CurrentRange.value = Ifold_source
        End If
    Next
    Wait 0.1
    
Exit Function 'Add ErrHandler 2023/08/18
errHandler: 'Add ErrHandler 2023/08/18
    Call Print_Error_Message(Error_Info, "VBT_LIB_DC_Func", "SetCurrentRange_From_Ifold") 'Add ErrHandler 2023/08/18
    If AbortTest Then Exit Function Else Resume Next 'Add ErrHandler 2023/08/18
End Function


Public Function EVS_Pre_Setting_wPACAVS(dc_spec As String, power_pin_ary() As String, ByRef EVS_Detail_value() As Pins_detail, Optional Step_number As Integer = 10, Optional Looping_Control As Boolean, Optional Looping_Range As Double = 0.4, Optional Looping_Index_Name As String, Optional Looping_Max_Steps_Name As String, Optional Open_LatchUp_measure As Boolean, Optional G1 As String, Optional G2 As String)
On Error GoTo errHandler
    Dim Spec As Variant, Spec_Var As String, pin As Variant, Dc_spec_type As String
    Dim Step_value As Double
    Dim pin_count As Integer
    Dim End_Value As Double
    Dim Start_Value As Double
    Dim diff_value As Double
    Dim Looping_Index As Double, Looping_Max_Step As Double, Looping_Step_range As Double
    Dim Gap_Persent As Double
    Dim EVS_Index As Integer, Evs_End_Index As Integer
    pin_count = 0  '  initial value
    Evs_End_Index = Step_number
    Dim SP As Variant, t As String
    Dim Gpin As Variant
    Dim Dict_PGrp As Scripting.Dictionary
    Set Dict_PGrp = New Scripting.Dictionary
    Dim PABol As Boolean
    Dim G1Ary() As String, G1cnt As Long
    Dim G2Ary() As String, G2cnt As Long
    PABol = False
'''''Define Power group
    If G1 <> "" And G2 <> "" Then
        PABol = True
        TheExec.DataManager.DecomposePinList G1, G1Ary, G1cnt
        TheExec.DataManager.DecomposePinList G2, G2Ary, G2cnt
        '====================G1====================
        For Each Gpin In G1Ary
            If Dict_PGrp.Exists(UCase(Gpin)) Then
                TheExec.Datalog.WriteComment "Error!!!Please help to check PA CAVS power pin group!!!Duplicate pin!!!!!"
            Else
                Dict_PGrp.Add UCase(Gpin), 1
            End If
        Next Gpin
        '====================G2====================
        For Each Gpin In G2Ary
            If Dict_PGrp.Exists(UCase(Gpin)) Then
                TheExec.Datalog.WriteComment "Error!!!Please help to check PA CAVS power pin group!!!Duplicate pin!!!!!"
            Else
                Dict_PGrp.Add UCase(Gpin), 2
            End If
        Next Gpin
    End If
    
    
    
    '\\\\\\\\\\\\\\\\\\\Calculate EVS Ramp up/down information\\\\\\\\\\\\\\\\\\\
    
    For Each SP In TheExec.Specs.DC.Categories(dc_spec).SpecList
        SP = LCase(SP)
        If SP Like "*_var_bi" Then
            Dc_spec_type = "_BI"
        ElseIf SP Like "*_var_sc" Then
            Dc_spec_type = "_SC"
        ElseIf SP Like "*_var_e" Then
            Dc_spec_type = "_E"
        ElseIf SP Like "*_var" Then
            Dc_spec_type = vbNullString
        Else
            TheExec.ErrorLogMessage "DC spec " & SP & " is not ended with _VAR in " & TheExec.DataManager.instancename
        End If
    Next SP
    
    For Each pin In power_pin_ary
'        If Open_LatchUp_measure Then
'            TheHdw.DCVS.Pins(Pin).Meter.mode = tlDCVSMeterCurrent
'            TheHdw.DCVS.Pins(Pin).Meter.CurrentRange = 1
'        End If
'        Dim DC_Spec_Var As String
'        If Pin Like "*SRAM*" Then
'            DC_Spec_Var = "_VAR_"
'        Else
'            DC_Spec_Var = "_VOP_VAR_"
'        End If
            
        Spec_Var = pin & "_VAR" & Dc_spec_type
        Start_Value = TheHdw.DCVS.Pins(pin).Voltage.value ''Start Voltage is read by hardware
        End_Value = TheExec.Specs.DC.item(Spec_Var).Categories.item(dc_spec).max.value

        ''//////////////////////////Looping Function for EVS experiment//////////////////////////
        '' If we want to apply this method , we need to additionly change the EVS flow and assign varient at main flow
        ''
        '' Example as below:
        ''//////// Main flow ////////
        '' |     Opcode     |      Parameter     |
        '' | assign-integer |     Max_step 5     | <= Assign integer at IGXL Main flow
        '' | create-integer | EVS_Looping_INDEX  | <= Create integer at IGXL Main flow
        ''//////// Main flow ////////
        
        ''//////// EVS flow ////////
        '' |     Opcode    |                             Parameter                                 |
        '' |      For      | EVS_Looping_INDEX=0; EVS_Looping_INDEX< Max_Step; EVS_Looping_INDEX++ | <= For index for looping ,Start at 0
        '' |      test     |                         CpuSaEVS_Static_*pp**                         | <= Run Pattern
        '' |      test     |                        EVS_Static_Power_Up_CpuSa                        | <= Power up
        '' |      test     |                      EVS_Static_Power_Down_CpuSa                      | <= Power down
        '' |      Next     |                          EVS_Looping_INDEX                            |
        ''//////// EVS flow ////////
        
        '' Instance setting for "EVS_Static_Power_Up(Down)_CpuSa
        '' Addtional setting argument Looping_Control = true, Looping_Index_Name = "EVS_Looping_INDEX" and Looping_Max_Steps_Name = "Max_Step"
'        If Looping_Control And Looping_Index_Name <> "" And Looping_Max_Steps_Name <> "" Then  ''Only need to change final value of Power Up
'            Looping_Index = TheExec.Flow.var(Looping_Index_Name).Value '' get index from flow
'            Looping_Max_Step = TheExec.Flow.var(Looping_Max_Steps_Name).Value '' get max step from flow
'            End_Value = Start_Value + (Looping_Index + 1) * ((End_Value - Start_Value) / Looping_Max_Step) ''' update end value for each looping
'        End If
        If Looping_Control And Looping_Index_Name <> "" And Looping_Max_Steps_Name <> "" Then  ''Only need to change final value of Power Up
            Looping_Index = TheExec.Flow.var(Looping_Index_Name).value '' get index from flow
            Looping_Max_Step = TheExec.Flow.var(Looping_Max_Steps_Name).value '' get max step from flow
            Looping_Step_range = Looping_Range / (Looping_Max_Step - 1) '' get max step from flow
            End_Value = TheExec.Specs.DC.item(Spec_Var).Categories.item(dc_spec).max.value ''' update end value for each looping
            'Start_Value = End_Value - Looping_Range
            Dim Looping_Start_Voltage As Double
            Looping_Start_Voltage = End_Value - Looping_Range
            End_Value = Looping_Start_Voltage + Looping_Step_range * Looping_Index
        End If
        'Looping_Max_Step = 6
        If Not (Looping_Control) And Looping_Index_Name <> "" And Looping_Max_Steps_Name <> "" Then
        '' You can keep the opcode "For" on IGXL flow and trun looping_Control into false.
        '' If you still have Looping_Max_Steps_Name , it will set it to 1 which means it won't loop on the IGXL flow anymore.
            TheExec.Flow.var(Looping_Max_Steps_Name).value = 1
'        ElseIf Not (Looping_Control) And Looping_Index_Name <> "" And Looping_Max_Steps_Name <> "" Then
'            TheExec.Flow.var(Looping_Max_Steps_Name).Value = 1
        End If
        ''//////////////////////////Looping Function for EVS experiment//////////////////////////
        
        EVS_Detail_value(pin_count).LatchUp_Final_Value = Format(End_Value, "0.00")
        diff_value = End_Value - Start_Value
        Gap_Persent = Abs(diff_value / Start_Value) '''+-0.5%
        If Gap_Persent < 0.005 Then 'small than 0.5 % pin don't need to rise at all
            EVS_Detail_value(pin_count).Pin_rise = False
        Else
        
            Step_value = diff_value / Evs_End_Index ' calculate step value by end index

            EVS_Detail_value(pin_count).Start_voltage = Start_Value
            EVS_Detail_value(pin_count).Start_voltage_up = Start_Value
            EVS_Detail_value(pin_count).Start_voltage_down = End_Value
            
            EVS_Detail_value(pin_count).Step_value = Step_value
            EVS_Detail_value(pin_count).Step_value_up = Step_value
            EVS_Detail_value(pin_count).Step_value_down = Step_value * (-1)
            
            EVS_Detail_value(pin_count).Pin_rise = True
            EVS_Detail_value(pin_count).PinName = CStr(pin)
            EVS_Detail_value(pin_count).PA_G1G2_Print_Enable = True 'add for PA power pin print 20210507
            ''=====================Add PA EVS grouping into EVS detail information=====================
            ''If no exist treat as G1....
            If PABol Then
                EVS_Detail_value(pin_count).Stress_Grp = IIf(Dict_PGrp.Exists(UCase(pin)), Dict_PGrp(UCase(pin)), 1)
            Else
            '''' if don't need PA CAVS put 1.
                EVS_Detail_value(pin_count).Stress_Grp = 1
            End If
            
        End If
        pin_count = pin_count + 1
    Next pin
    '\\\\\\\\\\\\\\\\\\\ Calculate EVS Ramp up/down information\\\\\\\\\\\\\\\\\\\
    Exit Function

errHandler:
    Call Print_Error_Message(Error_Warning_Info.Error_Info, "VBT_LIB_DC_Func", "EVS_Pre_Setting_wPACAVS")
    If AbortTest Then Exit Function Else Resume Next
End Function

''20230223: Modidfied to add when specific site pin alarm, then gate off "All_Power" in site.
'[20230731][T-Pal][Tank] Add meter reed value need use PinListData to store
Public Function Evs_Ramp_UPorDown_PAEVS(EVS_Detail_value() As Pins_detail, Direction As String, S_WaitTime As Double, power_pin_ary() As String, Optional Step_number As Integer = 10, Optional Rising_Delay_time As Double = 0.02, Optional Open_LatchUp_measure As Boolean = False, Optional Looping_Control As Boolean = False, Optional Looping_Index_Name As String = vbNullString, Optional Looping_Max_Steps_Name As String = vbNullString, Optional Cooling_Time As Double, Optional Mulit_Loop_EVS As Integer, Optional Mulit_EVS_Index As Integer, Optional PA_Counter As Integer, Optional PA_Enable As Boolean)
On Error GoTo errHandler
    Dim pin_count As Integer
    Dim PowerPin_Final_Value As New PinListData
    Dim ForceV As Double
    Dim LatchUp_measure_Value As New PinListData
    Dim EVS_Index As Integer, Evs_End_Index As Integer
    Dim Gatecheck As Boolean
    Dim power_pin_value As Double
    Dim Latch_Up_name As String
    Dim pin As Variant
    Dim EVS_Index_Value As Integer
    Dim EVS_End_Index_Value As Integer:: Evs_End_Index = Step_number
    Dim ifold As Double
    Dim m_InstanceName As String
    Dim Total_Power_Name As String
    Dim Power_Name As String
    m_InstanceName = TheExec.DataManager.instancename
    Dim Power As New PinListData
    Dim TotalPower As New SiteDouble
    Dim TotalPower_G1 As New SiteDouble  'MS 20210422
    Dim TotalPower_G2 As New SiteDouble  'MS 20210422
    Dim All_power_CurrentQ As New SiteDouble
    Dim ALL_PWR_PRINT As Boolean    'add for PA power pin print 20210507
    Dim alarm_occur As Boolean
    Dim PinName As String
    Dim sSlotType As String
    Dim CorePower_Cnt As Long
    Dim All_Core_Power() As String
    Dim p As Variant
    Dim site As Variant 'Carter, 20240304
    ALL_PWR_PRINT = False    'add for PA power pin print 20210507
    
    For EVS_Index = 1 To Evs_End_Index
'''        If UCase(Direction) = "UP" Then    'MS 20210422 mask
'''            pin_count = 0 '' reset pin index    'MS 20210422 mask
'''        ElseIf UCase(Direction) = "DOWN" Then    'MS 20210422 mask
'''            pin_count = UBound(power_pin_ary())    'MS 20210422 mask
'''        End If    'MS 20210422 mask
        pin_count = 0 '' reset pin index    'MS 20210422
        For Each pin In power_pin_ary
            
            'TheHdw.DCVS.Pins(Pin).CurrentRange = ifold
            'ifold = TheHdw.DCVS.Pins(Pin).CurrentLimit.Source.FoldLimit.Level
            If (PA_Counter Mod 2 = 1) And Mulit_Loop_EVS = 0 And PA_Counter <> 3 Then
                '' start point only care G1 up
                If UCase(Direction) = "UP" Then  'MS 20210422
                    If EVS_Detail_value(pin_count).Stress_Grp = 1 Then  'MS 20210422
                        EVS_Detail_value(pin_count).Start_voltage = EVS_Detail_value(pin_count).Start_voltage_up  'MS 20210422
                        EVS_Detail_value(pin_count).Step_value = EVS_Detail_value(pin_count).Step_value_up  'MS 20210422

                    ElseIf EVS_Detail_value(pin_count).Stress_Grp = 2 Then  'MS 20210422
                        EVS_Detail_value(pin_count).Start_voltage = EVS_Detail_value(pin_count).Start_voltage_up  'MS 20210422
                        EVS_Detail_value(pin_count).Step_value = 0  'MS 20210422
                        If ALL_PWR_PRINT = False Then                'add for PA power pin print 20210507
                            EVS_Detail_value(pin_count).PA_G1G2_Print_Enable = False
                        
                        End If
                    End If  'MS 20210422
                ElseIf UCase(Direction) = "DOWN" Then  'MS 20210422
                    If EVS_Detail_value(pin_count).Stress_Grp = 1 Then  'MS 20210422
                        EVS_Detail_value(pin_count).Start_voltage = EVS_Detail_value(pin_count).Start_voltage_down  'MS 20210520
                        EVS_Detail_value(pin_count).Step_value = EVS_Detail_value(pin_count).Step_value_down  'MS 20210520
                        If ALL_PWR_PRINT = False Then                'add for PA power pin print 20210520
                            EVS_Detail_value(pin_count).PA_G1G2_Print_Enable = False
                        
                        End If
                    End If
                End If  'MS 20210520
              
                
            ElseIf (PA_Counter Mod 2 = 1) And Mulit_Loop_EVS > 0 And PA_Counter <> 3 Then
                ' Interactive point for ramping
                ' odd G1 up, G2 down
                If UCase(Direction) = "UP" Then  'MS 20210422
                    If EVS_Detail_value(pin_count).Stress_Grp = 1 Then  'MS 20210422
                        EVS_Detail_value(pin_count).Start_voltage = EVS_Detail_value(pin_count).Start_voltage_up  'MS 20210422
                        EVS_Detail_value(pin_count).Step_value = EVS_Detail_value(pin_count).Step_value_up  'MS 20210422
                        
                    ElseIf EVS_Detail_value(pin_count).Stress_Grp = 2 Then  'MS 20210422
                        EVS_Detail_value(pin_count).Start_voltage = EVS_Detail_value(pin_count).Start_voltage_down  'MS 20210422
                        EVS_Detail_value(pin_count).Step_value = EVS_Detail_value(pin_count).Step_value_down  'MS 20210422
                        If ALL_PWR_PRINT = False Then                'add for PA power pin print 20210507
                            EVS_Detail_value(pin_count).PA_G1G2_Print_Enable = False
                        
                        End If
                    End If  'MS 20210422
                ElseIf UCase(Direction) = "DOWN" Then  'MS 20210422
                    If EVS_Detail_value(pin_count).Stress_Grp = 1 Then  'MS 20210422
                        EVS_Detail_value(pin_count).Start_voltage = EVS_Detail_value(pin_count).Start_voltage_down  'MS 20210520
                        EVS_Detail_value(pin_count).Step_value = EVS_Detail_value(pin_count).Step_value_down  'MS 20210520
                        If ALL_PWR_PRINT = False Then                'add for PA power pin print 20210520
                            EVS_Detail_value(pin_count).PA_G1G2_Print_Enable = False
                        
                        End If
                    End If
                End If  'MS 20210520

            ElseIf (PA_Counter Mod 2 = 0) Then 'And Mulit_Loop_EVS < Mulit_EVS_Index - 1 Then 'And Mulit_Loop_EVS <> Mulit_EVS_Index Then
                '' Interactive point for ramping
                '' even G1 down, G2 up
                If UCase(Direction) = "DOWN" Then  'MS 20210422
                    If EVS_Detail_value(pin_count).Stress_Grp = 1 Then    'MS 20210422
                         EVS_Detail_value(pin_count).Start_voltage = EVS_Detail_value(pin_count).Start_voltage_down  'MS 20210422
                         EVS_Detail_value(pin_count).Step_value = EVS_Detail_value(pin_count).Step_value_down    'MS 20210422
                         If ALL_PWR_PRINT = False Then              'add for PA power pin print 20210507
                            EVS_Detail_value(pin_count).PA_G1G2_Print_Enable = False
                        
                        End If
                    ElseIf EVS_Detail_value(pin_count).Stress_Grp = 2 Then
                         EVS_Detail_value(pin_count).Start_voltage = EVS_Detail_value(pin_count).Start_voltage_up  'MS 20210422
                         EVS_Detail_value(pin_count).Step_value = EVS_Detail_value(pin_count).Step_value_up    'MS 20210422
                         
                    End If    'MS 20210422
                End If  'MS 20210422
                
            ElseIf (PA_Counter = 3) And Mulit_Loop_EVS = Mulit_EVS_Index - 1 Then
                '' Last point only care G2 Down
                'If UCase(Direction) = "UP" Then  'MS 20210422
                    Direction = "END"
                    If EVS_Detail_value(pin_count).Stress_Grp = 2 Then    'MS 20210422
                         EVS_Detail_value(pin_count).Start_voltage = EVS_Detail_value(pin_count).Start_voltage_down  'MS 20210422
                         EVS_Detail_value(pin_count).Step_value = EVS_Detail_value(pin_count).Step_value_down    'MS 20210422
                    ElseIf EVS_Detail_value(pin_count).Stress_Grp = 1 And PA_Enable = False Then
                         EVS_Detail_value(pin_count).Start_voltage = EVS_Detail_value(pin_count).Start_voltage_down  'MS 20210422
                         EVS_Detail_value(pin_count).Step_value = EVS_Detail_value(pin_count).Step_value_down  'MS 20210422
                    ElseIf EVS_Detail_value(pin_count).Stress_Grp = 1 And PA_Enable = True Then
                         EVS_Detail_value(pin_count).Start_voltage = EVS_Detail_value(pin_count).Start_voltage_up  'MS 20210422
                         EVS_Detail_value(pin_count).Step_value = 0    'MS 20210422
                    End If    'MS 20210422
                'End If  'MS 20210422
            Else
                
            End If
'            If UCase(Direction) = "UP" Then
'                EVS_Detail_value(pin_count).Start_voltage = EVS_Detail_value(pin_count).Start_voltage_up
'                EVS_Detail_value(pin_count).Step_value = EVS_Detail_value(pin_count).Step_value_up
'            ElseIf UCase(Direction) = "DOWN" Then
'                EVS_Detail_value(pin_count).Start_voltage = EVS_Detail_value(pin_count).Start_voltage_down
'                EVS_Detail_value(pin_count).Step_value = EVS_Detail_value(pin_count).Step_value_down
'            End If
            
            PinName = LCase(EVS_Detail_value(pin_count).PinName)
            sSlotType = UCase(gl_GetInstrument_Dic(PinName))
            If EVS_Detail_value(pin_count).Pin_rise Then
                ForceV = EVS_Detail_value(pin_count).Start_voltage + EVS_Index * EVS_Detail_value(pin_count).Step_value
                TheHdw.DCVS.Pins(PinName).Voltage.value = ForceV
        
                If EVS_Index = Evs_End_Index Then
                    Power.AddPin (PinName)
                    PowerPin_Final_Value.AddPin (PinName)
                    PowerPin_Final_Value.Pins(PinName) = TheHdw.DCVS.Pins(PinName).Voltage.value
                End If
            End If
            
            If Open_LatchUp_measure Then
                TheHdw.DCVS.Pins(PinName).Meter.mode = tlDCVSMeterCurrent
'                If Pin = "VDD_SOC" Then
'                    TheHdw.DCVS.Pins(PinName).Meter.CurrentRange = 15
'                Else
'                    TheHdw.DCVS.Pins(PinName).Meter.CurrentRange = 1
'                End If
                            
                '' Print out measure current value if want to collect the Latch up data
                Select Case sSlotType
                    Case glbConstIns_VS5A, glbConstIns_HEXVS, glbConstIns_VSM:
                        LatchUp_measure_Value = TheHdw.DCVS.Pins(PinName).Meter.Read(tlStrobe, 10, , tlDCVSMeterReadingFormatAverage)
                    Case glbConstIns_VS800MA, glbConstIns_VHDVS:    'UVS256 and UVS256-HP sample size = 1
                        LatchUp_measure_Value = TheHdw.DCVS.Pins(PinName).Meter.Read(tlStrobe, 1)
                    Case Else
                        Call Print_Error_Message(Error_Warning_Info.Warning_Info, "VBT_LIB_DC_Func", "Evs_Ramp_UPorDown_PAEVS", "Slot type did not define!!")
                End Select
                Latch_Up_name = m_InstanceName & "_" & "Latch_up_data_" & Replace(CStr(EVS_Detail_value(pin_count).LatchUp_Final_Value), ".", "p") & "V"
                TheExec.Flow.TestLimit resultVal:=LatchUp_measure_Value.Pins(PinName), PinName:=PinName, scaletype:=scaleNone, unit:=unitAmp, formatStr:="%.3f", Tname:=Latch_Up_name, ForceVal:=ForceV, ForceUnit:=unitVolt, ForceResults:=tlForceNone
            Else
                If EVS_Index = Evs_End_Index And UCase(Direction) <> "END" Then    '  And UCase(Direction) = "UP" Then    'MS 20210422 mask
                    If EVS_Detail_value(pin_count).Pin_rise Then
                        If F_start_profile = False Then
                            If EVS_Detail_value(pin_count).PA_G1G2_Print_Enable = True Then    'add for PA power pin print 20210507
                                TheHdw.DCVS.Pins(PinName).Meter.mode = tlDCVSMeterCurrent
                                Select Case sSlotType
                                    Case glbConstIns_VS5A, glbConstIns_HEXVS, glbConstIns_VSM:
                                        LatchUp_measure_Value = TheHdw.DCVS.Pins(PinName).Meter.Read(tlStrobe, 10, , tlDCVSMeterReadingFormatAverage)
                                    Case glbConstIns_VS800MA, glbConstIns_VHDVS:    'UVS256 and UVS256-HP sample size = 1
                                        LatchUp_measure_Value = TheHdw.DCVS.Pins(PinName).Meter.Read(tlStrobe, 1)
                                    Case Else
                                        Call Print_Error_Message(Error_Warning_Info.Warning_Info, "VBT_LIB_DC_Func", "Evs_Ramp_UPorDown_PAEVS", "Slot type did not define!!")
                                End Select
                                EVS_Detail_value(pin_count).Stress_C = LatchUp_measure_Value.Pins(PinName)    ''add 20210504
                                Latch_Up_name = m_InstanceName & "_" & "Latch_up_MeasI_" & Replace(CStr(EVS_Detail_value(pin_count).LatchUp_Final_Value), ".", "p") & "V"
                                Total_Power_Name = "Total_Power_" & m_InstanceName
                                TheExec.Flow.TestLimit resultVal:=LatchUp_measure_Value.Pins(PinName), PinName:=PinName, scaletype:=scaleNone, unit:=unitAmp, formatStr:="%.3f", Tname:=Latch_Up_name, ForceVal:=ForceV, ForceUnit:=unitVolt, ForceResults:=tlForceNone
                                Power.Pins(PinName) = LatchUp_measure_Value.Pins(PinName).Multiply(PowerPin_Final_Value.Pins(PinName))
                                EVS_Detail_value(pin_count).P_Consumption = Power.Pins(PinName) '' ''add 20210504
                                'TotalPower = TotalPower.Add(Power.Pins(PinName))
                                If EVS_Detail_value(pin_count).Stress_Grp = 1 Then      'MS 20210422
                                    TotalPower_G1 = TotalPower_G1.Add(Power.Pins(PinName)) 'MS 20210422
                                ElseIf EVS_Detail_value(pin_count).Stress_Grp = 2 Then 'MS 20210422
                                    TotalPower_G2 = TotalPower_G2.Add(Power.Pins(PinName)) 'MS 20210422
                                End If                                                                              'MS 20210422
                            End If
                        End If
                    End If
                End If
            End If
         '' Print out measure current value if want to collect the Latch up data

'            If UCase(Direction) = "UP" Then
'                pin_count = pin_count + 1
'            ElseIf UCase(Direction) = "DOWN" Then
'                pin_count = pin_count - 1
'            End If
            pin_count = pin_count + 1
        Next pin
        
        TheHdw.Wait Rising_Delay_time 'delay time of each ramp up
        
    Next EVS_Index
    '\\\\\\\\\\\\\\\\\\\ End Power up/Down \\\\\\\\\\\\\\\\\
    pin_count = 0 'reset pin count
    '///////////stress time after power up///////////

    If (PA_Counter Mod 2 = 1) And UCase(Direction) = "UP" Then
    
        If F_start_profile = False Then
            'If TheExec.DataManager.instanceName Like "*Gfx*" Then
                'TheExec.Flow.TestLimit TotalPower, 0, 210, tlSignGreaterEqual, tlSignLessEqual, scaletype:=scaleNone, Unit:=unitCustom, Tname:=Total_Power_Name, PinName:="Total Power", customUnit:="W"    'BurstResult=1:Pass
            'Else
                'TheExec.Flow.TestLimit TotalPower, 0, 210, tlSignGreaterEqual, tlSignLessEqual, scaletype:=scaleNone, Unit:=unitCustom, Tname:=Total_Power_Name, PinName:="Total Power", customUnit:="W"    'BurstResult=1:Pass    'MS 20210422 mask
            TheExec.Flow.TestLimit TotalPower_G1, 0, 270, tlSignGreaterEqual, tlSignLessEqual, scaletype:=scaleNone, unit:=unitCustom, Tname:=Total_Power_Name, PinName:="Total Power_G1", customUnit:="W"    'BurstResult=1:Pass  'MS 20210422
            'End If
        End If
        TheHdw.Wait S_WaitTime
        '''Core_power
        
        
        ''''Dim CorePower_Cnt As Long
        ''''Dim all_core_power() As String
        TheExec.DataManager.DecomposePinList "CorePower", All_Core_Power, CorePower_Cnt

        
        '''' Dim p As Variant
''''        For Each p In all_core_power
''''        TheHdw.DCVS.Pins(p).Meter.mode = tlDCVSMeterCurrent
''''        All_power_CurrentQ = TheHdw.DCVS.Pins(p).Meter.Read(tlStrobe, 10, , tlDCVSMeterReadingFormatAverage)
''''        TheExec.Flow.TestLimit resultVal:=All_power_CurrentQ, PinName:=p, scaletype:=scaleNone, Unit:=unitAmp, formatStr:="%.3f", Tname:="Ramp_EVS_Current", forceVal:=ForceV, ForceUnit:=unitVolt, ForceResults:=tlForceNone
''''        Next p
        '''Core_power
        TheExec.Datalog.WriteComment ""
        TheExec.Datalog.WriteComment "-------------------------EVS Power ramp start-------------------------"
''        TheExec.Flow.TestLimit S_WaitTime, 0, 99, tlSignGreaterEqual, tlSignLessEqual, scaletype:=scaleNone, Unit:=unitCustom, Tname:="EVS_stress_time", PinName:="EVS_stress_time", customUnit:="sec" 'BurstResult=1:Pass
    '///////////Print out final value of each pin (Power up and down)///////////
        For Each pin In power_pin_ary
            If EVS_Detail_value(pin_count).Pin_rise And EVS_Detail_value(pin_count).PA_G1G2_Print_Enable = True Then 'add for PA power pin print 20210507
            
                TheExec.Flow.TestLimit PowerPin_Final_Value.Pins(pin), , , tlSignGreaterEqual, tlSignLessEqual, scaletype:=scaleNone, unit:=unitCustom, customUnit:="V"    'BurstResult=1:Pass
            End If
            pin_count = pin_count + 1
        Next pin
        pin_count = 0 ' MQ for rest pincount
        For Each pin In power_pin_ary
            If EVS_Detail_value(pin_count).Pin_rise And EVS_Detail_value(pin_count).PA_G1G2_Print_Enable = True Then 'add for PA power pin print 20210507
                If F_start_profile = False Then
                    Power_Name = m_InstanceName & "_" & "Power"
                    TheExec.Flow.TestLimit Power.Pins(pin), , , tlSignGreaterEqual, tlSignLessEqual, scaletype:=scaleNone, unit:=unitCustom, Tname:=Power_Name, customUnit:="W"    'BurstResult=1:Pass
                End If
            End If
            EVS_Detail_value(pin_count).PA_G1G2_Print_Enable = True    'add for PA power pin print 20210507
            pin_count = pin_count + 1
        Next pin
    '///////////Print out final value of each pin (Power up and down)///////////
    '///////////Alarm check after stress during power up///////////
        For Each site In TheExec.sites
            pin_count = 0
            For Each pin In power_pin_ary
            Gatecheck = TheHdw.DCVS.Pins(pin).Gate
            EVS_Detail_value(pin_count).Gate_check(site) = Gatecheck
            If TheHdw.DCVS.Pins(pin).Gate = False Then
'                Dim power_pin_value As Double
                power_pin_value = TheHdw.DCVS.Pins(pin).Voltage.value
                TheExec.Flow.TestLimit power_pin_value, , , tlSignGreaterEqual, tlSignLessEqual, scaletype:=scaleNone, unit:=unitCustom, customUnit:="V", PinName:=pin, Tname:="Vlotage_PatternAlarm_After_Stress"
                TheExec.Flow.TestLimit Gatecheck, 1, 1, tlSignGreaterEqual, tlSignLessEqual, scaletype:=scaleNone, Tname:="Gate_PatternAlarm_After_Stress", PinName:=pin
                alarm_occur = True
            End If
            pin_count = pin_count + 1
        Next pin
            If alarm_occur = True Then '20221208 Jayden, gate off all stress rail if alarm occured.
                TheHdw.DCVS.Pins("All_Power").Gate = False
                TheHdw.Wait 0.0001
                alarm_occur = False
            End If
        Next site
        TheExec.Flow.TestLimit S_WaitTime, 0, 99, tlSignGreaterEqual, tlSignLessEqual, scaletype:=scaleNone, unit:=unitCustom, Tname:="EVS_stress_time" & "_" & Block & "_" & Mulit_Loop_EVS + 1, PinName:="G1 EVS_stress_time", customUnit:="sec" 'BurstResult=1:Pass

'        ///////////Alarm check after stress during power up///////////
    
'************************************************************************************

    ElseIf (PA_Counter Mod 2 = 0) And UCase(Direction) = "DOWN" Then
    
        If F_start_profile = False Then
            'If TheExec.DataManager.instanceName Like "*Gfx*" Then
                'TheExec.Flow.TestLimit TotalPower, 0, 210, tlSignGreaterEqual, tlSignLessEqual, scaletype:=scaleNone, Unit:=unitCustom, Tname:=Total_Power_Name, PinName:="Total Power", customUnit:="W"    'BurstResult=1:Pass
            'Else
                'TheExec.Flow.TestLimit TotalPower, 0, 210, tlSignGreaterEqual, tlSignLessEqual, scaletype:=scaleNone, Unit:=unitCustom, Tname:=Total_Power_Name, PinName:="Total Power", customUnit:="W"    'BurstResult=1:Pass    'MS 20210422 mask
            TheExec.Flow.TestLimit TotalPower_G2, 0, 270, tlSignGreaterEqual, tlSignLessEqual, scaletype:=scaleNone, unit:=unitCustom, Tname:=Total_Power_Name, PinName:="Total Power_G2", customUnit:="W"    'BurstResult=1:Pass  'MS 20210422
            'End If
        End If
        TheHdw.Wait S_WaitTime
        '''Core_power
        
        
        'Dim CorePower_Cnt As Long
        ' Dim all_core_power() As String
        TheExec.DataManager.DecomposePinList "CorePower", All_Core_Power, CorePower_Cnt

        
        ' Dim p As Variant
''''        For Each p In all_core_power
''''        TheHdw.DCVS.Pins(p).Meter.mode = tlDCVSMeterCurrent
''''        All_power_CurrentQ = TheHdw.DCVS.Pins(p).Meter.Read(tlStrobe, 10, , tlDCVSMeterReadingFormatAverage)
''''        TheExec.Flow.TestLimit resultVal:=All_power_CurrentQ, PinName:=p, scaletype:=scaleNone, Unit:=unitAmp, formatStr:="%.3f", Tname:="Ramp_EVS_Current", forceVal:=ForceV, ForceUnit:=unitVolt, ForceResults:=tlForceNone
''''        Next p
        '''Core_power
        TheExec.Datalog.WriteComment ""
        TheExec.Datalog.WriteComment "-------------------------EVS Power ramp start-------------------------"
''        TheExec.Flow.TestLimit S_WaitTime, 0, 99, tlSignGreaterEqual, tlSignLessEqual, scaletype:=scaleNone, Unit:=unitCustom, Tname:="EVS_stress_time", PinName:="EVS_stress_time", customUnit:="sec" 'BurstResult=1:Pass
    '///////////Print out final value of each pin (Power up and down)///////////
        For Each pin In power_pin_ary
            If EVS_Detail_value(pin_count).Pin_rise And EVS_Detail_value(pin_count).PA_G1G2_Print_Enable = True Then 'add for PA power pin print 20210507
                TheExec.Flow.TestLimit PowerPin_Final_Value.Pins(pin), , , tlSignGreaterEqual, tlSignLessEqual, scaletype:=scaleNone, unit:=unitCustom, customUnit:="V"    'BurstResult=1:Pass
            End If
            pin_count = pin_count + 1
        Next pin
        pin_count = 0 ' MQ for rest pincount
        For Each pin In power_pin_ary
            If EVS_Detail_value(pin_count).Pin_rise And EVS_Detail_value(pin_count).PA_G1G2_Print_Enable = True Then  'add for PA power pin print 20210507
                If F_start_profile = False Then
                    Power_Name = m_InstanceName & "_" & "Power"
                    TheExec.Flow.TestLimit Power.Pins(pin), , , tlSignGreaterEqual, tlSignLessEqual, scaletype:=scaleNone, unit:=unitCustom, Tname:=Power_Name, customUnit:="W"    'BurstResult=1:Pass
                End If
            End If
            EVS_Detail_value(pin_count).PA_G1G2_Print_Enable = True    'add for PA power pin print 20210507
            pin_count = pin_count + 1
        Next pin
    '///////////Print out final value of each pin (Power up and down)///////////
    '///////////Alarm check after stress during power up///////////
        For Each site In TheExec.sites
            pin_count = 0
            For Each pin In power_pin_ary
            Gatecheck = TheHdw.DCVS.Pins(pin).Gate
            EVS_Detail_value(pin_count).Gate_check(site) = Gatecheck
                If TheHdw.DCVS.Pins(pin).Gate = False Then
'                    Dim power_pin_value As Double
                    power_pin_value = TheHdw.DCVS.Pins(pin).Voltage.value
                    TheExec.Flow.TestLimit power_pin_value, , , tlSignGreaterEqual, tlSignLessEqual, scaletype:=scaleNone, unit:=unitCustom, customUnit:="V", PinName:=pin, Tname:="Vlotage_PatternAlarm_After_Stress"
                    TheExec.Flow.TestLimit Gatecheck, 1, 1, tlSignGreaterEqual, tlSignLessEqual, scaletype:=scaleNone, Tname:="Gate_PatternAlarm_After_Stress", PinName:=pin
                    alarm_occur = True
                End If
                pin_count = pin_count + 1
            Next pin
            If alarm_occur = True Then '20221208 Jayden, gate off all stress rail if alarm occured.
                TheHdw.DCVS.Pins("All_Power").Gate = False
                TheHdw.Wait 0.0001
                alarm_occur = False
            End If
        Next site
        TheExec.Flow.TestLimit S_WaitTime, 0, 99, tlSignGreaterEqual, tlSignLessEqual, scaletype:=scaleNone, unit:=unitCustom, Tname:="EVS_stress_time" & "_" & Block & "_" & Mulit_Loop_EVS + 1, PinName:="G2 EVS_stress_time", customUnit:="sec" 'BurstResult=1:Pass

'        ///////////Alarm check after stress during power up///////////
    
    
''''''    ElseIf UCase(Direction) = "DOWN" Then
''''''                    '///////////Alarm check after stress during power down///////////
''''''                        'Dim all_core_power() As String
''''''                        Dim Core_power() As String
''''''                        Dim All_power As Variant
''''''                        'Dim CorePower_Cnt As Long
''''''                        'Core_power = "VDD_AVE,VDD_DCS_DDR,VDD_DISP,VDD_ECPU,VDD_GPU,VDD_PCPU,VDD_SOC,VDD_SRAM_ANE,VDD_SRAM_CPU,VDD_SRAM_GPU,VDD_SRAM_SOC,VDD_LOW,VDD_FIXED"
''''''                        'All_Core_Power() = Split(Core_power, ",")
''''''                        TheExec.DataManager.DecomposePinList "CorePower", all_core_power, CorePower_Cnt
''''''
''''''                        For Each Site In TheExec.sites
''''''                            'For Each Pin In power_pin_ary
''''''                            For Each All_power In all_core_power
''''''                            Gatecheck = thehdw.DCVS.Pins(All_power).Gate
''''''                                If thehdw.DCVS.Pins(All_power).Gate = False Then
''''''                                    TheExec.Flow.TestLimit Gatecheck, -1, -1, tlSignGreaterEqual, tlSignLessEqual, scaletype:=scaleNone, Tname:="Gate_PatternAlarm_After_EVS_ramp_down", PinName:=All_power
''''''                                End If
''''''                            Next All_power
''''''                        Next Site
''''''                        TheExec.Datalog.WriteComment ""
''''''                        TheExec.Datalog.WriteComment "--------------------------EVS Power ramp end--------------------------"
    ElseIf UCase(Direction) = "END" Or (UCase(Direction) = "DOWN" And PA_Enable = False) Then
        TheExec.DataManager.DecomposePinList "CorePower", All_Core_Power, CorePower_Cnt
        For Each site In TheExec.sites
            pin_count = 0
            For Each pin In power_pin_ary
                Gatecheck = TheHdw.DCVS.Pins(pin).Gate
                EVS_Detail_value(pin_count).Gate_check(site) = Gatecheck
                If TheHdw.DCVS.Pins(pin).Gate = False Then
                    'Dim power_pin_value As Double
                    power_pin_value = TheHdw.DCVS.Pins(pin).Voltage.value
                    TheExec.Flow.TestLimit power_pin_value, , , tlSignGreaterEqual, tlSignLessEqual, scaletype:=scaleNone, unit:=unitCustom, customUnit:="V", PinName:=pin, Tname:="Vlotage_PatternAlarm_After_Stress"
                    TheExec.Flow.TestLimit Gatecheck, 1, 1, tlSignGreaterEqual, tlSignLessEqual, scaletype:=scaleNone, Tname:="Gate_PatternAlarm_After_Stress", PinName:=pin
                    alarm_occur = True
                End If
                EVS_Detail_value(pin_count).PA_G1G2_Print_Enable = True    'add for PA power pin print 20210507
                pin_count = pin_count + 1
            Next pin
            If alarm_occur = True Then '20221208 Jayden, gate off all stress rail if alarm occured.
                TheHdw.DCVS.Pins("All_Power").Gate = False
                TheHdw.Wait 0.0001
                alarm_occur = False
            End If
        Next site
    End If
     TheExec.Datalog.WriteComment "--------------------------EVS Power ramp end--------------------------"
    '///////////stress time after power up///////////
    Exit Function

errHandler:
    Call Print_Error_Message(Error_Warning_Info.Error_Info, "VBT_LIB_DC_Func", "Evs_Ramp_UPorDown_PAEVS")
    If AbortTest Then Exit Function Else Resume Next
End Function

Public Function Test_time_breakdown_PAEVS_End(ByRef Timer As Double, Test_time_breakdown As Boolean, EVS_Sequence As String, Mulit_Loop_EVS As Integer, S_WaitTime As Double, Cooling_Time As Double, Block As String)

On Error GoTo errHandler 'Add ErrHandler 2023/08/18
    If Test_time_breakdown = True Then
        Timer = TheExec.Timer(Timer)
        If EVS_Sequence = "EVS_Pre_Setting" Then
'            TheExec.Datalog.WriteComment "EVS_Pre_Setting : " + Format(Timer * 1000#, "##0.000") + " msec"
            TheExec.Flow.TestLimit Format(Timer, "##0.000000"), , , tlSignGreaterEqual, tlSignLessEqual, scaletype:=scaleNone, unit:=unitCustom, Tname:="EVS_Pre_Setting_time" & "_" & Block, PinName:="EVS_Pre_Setting_time", customUnit:="sec" 'BurstResult=1:Pass
        End If
        If EVS_Sequence = "Evs_Ramp_UP" Then
            If Mulit_Loop_EVS <> 0 Then
'                TheExec.Datalog.WriteComment "EVS_Ramp_up" & Mulit_Loop_EVS + 1 & " : " + Format((Timer - S_WaitTime - Cooling_Time) * 1000#, "##0.000") + " msec"
                TheExec.Flow.TestLimit Format((Timer - S_WaitTime - Cooling_Time), "##0.000000"), , , tlSignGreaterEqual, tlSignLessEqual, scaletype:=scaleNone, unit:=unitCustom, Tname:="Evs_Ramp_UP_time" & "_" & Block & "_" & Mulit_Loop_EVS + 1, PinName:="Evs_Ramp_UP_time", customUnit:="sec" 'BurstResult=1:Pass
            Else
'                TheExec.Datalog.WriteComment "EVS_Ramp_up" & Mulit_Loop_EVS + 1 & " : " + Format((Timer - S_WaitTime) * 1000#, "##0.000") + " msec"
                TheExec.Flow.TestLimit Format((Timer - S_WaitTime), "##0.000000"), , , tlSignGreaterEqual, tlSignLessEqual, scaletype:=scaleNone, unit:=unitCustom, Tname:="Evs_Ramp_UP_time" & "_" & Block & "_" & Mulit_Loop_EVS + 1, PinName:="Evs_Ramp_UP_time", customUnit:="sec" 'BurstResult=1:Pass
            End If
        End If
        If EVS_Sequence = "Evs_Ramp_DOWN" Then
'            TheExec.Datalog.WriteComment "EVS_Ramp_down" & Mulit_Loop_EVS + 1 & " : " + Format(Timer * 1000#, "##0.000") + " msec"
            TheExec.Flow.TestLimit Format(Timer, "##0.000000"), , , tlSignGreaterEqual, tlSignLessEqual, scaletype:=scaleNone, unit:=unitCustom, Tname:="Evs_Ramp_DOWN_time" & "_" & Block & "_" & Mulit_Loop_EVS + 1, PinName:="Evs_Ramp_DOWN_time", customUnit:="sec" 'BurstResult=1:Pass
        End If
        If EVS_Sequence = "Total_time" Then
'            TheExec.Datalog.WriteComment "Total test time : " + Format(Timer * 1000#, "##0.000") + " msec"
            TheExec.Flow.TestLimit Format(Timer, "##0.000000"), , , tlSignGreaterEqual, tlSignLessEqual, scaletype:=scaleNone, unit:=unitCustom, Tname:="Total_time" & "_" & Block, PinName:="Total_time", customUnit:="sec"    'BurstResult=1:Pass
        End If
    End If
    
Exit Function 'Add ErrHandler 2023/08/18
errHandler: 'Add ErrHandler 2023/08/18
    Call Print_Error_Message(Error_Info, "VBT_LIB_DC_Func", "Test_time_breakdown_PAEVS_End") 'Add ErrHandler 2023/08/18
    If AbortTest Then Exit Function Else Resume Next 'Add ErrHandler 2023/08/18
End Function

